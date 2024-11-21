Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

' New a object filter that contains all objects
Dim objfilterObj
Set objfilterObj = pcbAppObj.Utility.NewObjectFilter()

' Get selected object ids
Dim selObjectIdList
selObjectIdList = pcbDocObj.SelectedObjectIds(objfilterObj)

pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart(epcbDRCModeNone)
' Unfix or unlock objects
Dim i, selObj
For i = 0 To UBound(selObjectIdList)
    Set selObj = pcbDocObj.FindObjectById(selObjectIdList(i))
    ' Check errors
    On Error Resume Next
    Err.Clear()
    selObj.FixLock = EPcbFixLockNone

    If Err.Number Then
	    ' Check errors
	    MsgBox("Unsupported Object !")
	End If
    Err.Clear()
Next
pcbAppObj.Gui.CursorBusy(False)
pcbDocObj.TransactionEnd()

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
