Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Turn off DRC
pcbDocObj.TransactionStart(epcbDRCModeNone)

' Get all groups
Dim groupColl
Set groupColl = pcbDocObj.Groups(epcbSelectAll, epcbGroupAll, False)

' Get main group name
Dim pcbName, mainGroupName
pcbName = pcbDocObj.Name
mainGroupName = Left(pcbName, Len(pcbName) - 4)

' Delete all sub groups
' Can't delete main group, any attempt to delete the main group will result in a runtime error.
Dim groupObj
Dim cmpColl, cmpObj, cmpStatusArr
Set cmpStatusArr = CreateObject("System.Collections.ArrayList")
For Each groupObj In groupColl
    If groupObj.IsValid And groupObj.Name <> mainGroupName Then
        Set cmpColl = groupObj.Components
        ' Group can't be deleted if any component is fixed or locked
        For Each cmpObj In cmpColl
            If cmpObj.FixLock <> EPcbFixLockNone Then
                ' cmpStatusArr.Add(cmpObj)
                ' cmpStatusArr.Add(cmpObj.FixLock)
                cmpObj.FixLock = EPcbFixLockNone
            End If
        Next
    End If
Next
For Each groupObj In groupColl
    If groupObj.IsValid And groupObj.Name <> mainGroupName Then
        groupObj.Delete 
    End If
Next
' Recover components FixLock status
Dim i, arrSize
arrSize = cmpStatusArr.Count - 1
For i = 0 To arrSize Step 2
    cmpStatusArr(i).FixLock = cmpStatusArr(i+1)
Next

' All the changes should be kept
pcbDocObj.TransactionEnd(True)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Server validation function
Private Function ValidateServer(doc) 
    dim key, licenseServer, licenseToken

    ' Ask Expedition's document for the key
    key = doc.Validate(0)

    ' Get license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")

    ' Ask the license server for the license token
    licenseToken = licenseServer.GetToken(key)

    ' Release license server
    Set licenseServer = nothing

    ' Turn off error messages.  Validate may fail if the token is incorrect
    On Error Resume Next
    Err.Clear

    ' Ask the document to validate the license token
    doc.Validate(licenseToken)
    If Err Then
        ValidateServer = 0    
    Else
        ValidateServer = 1
    End If
End Function