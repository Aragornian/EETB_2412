Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' Get the GUI object
Dim pcbGuiObj
Set pcbGuiObj = pcbAppObj.Gui

' License the document
ValidateServer(pcbDocObj)

pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart

If pcbAppObj.LockServer = True Then

    ' Get the display control object
    Dim displayCtrlObj, utilityObj
    Set displayCtrlObj = pcbDocObj.ActiveViewEx.DisplayControl
    Set utilityObj = pcbAppObj.Utility

    ' Color netclasses
    Dim nclassColl, nclass, nclassName
    Set nclassColl = pcbDocObj.NetClasses("Z*")

    Dim j: j = 0

    ' Match netclass
    Dim z45Re, z50Re, z85Re, z90Re, z100Re
    Set z45Re = New RegExp
    z45Re.Pattern = "Z*45*"
    Set z50Re = New RegExp
    z50Re.Pattern = "Z*5\d[0]*"
    Set z85Re = New RegExp
    z85Re.Pattern = "Z*85*"
    Set z90Re = New RegExp
    z90Re.Pattern = "Z*9\d[0]*"
    Set z100Re = New RegExp
    z100Re.Pattern = "Z*1\d[0]\d[0]*"

    For Each nclass In nclassColl
        nclassName = nclass.Name
        ' 100Ohm impedance
        If z100Re.Test(nclassName) Then
            With displayCtrlObj
                .Lock
                .Global.Color( "[NetClass]." & nclassName ) = utilityObj.NewColorPattern( 51, 153, 255, 100, 0, False, False ) 
                .Visible( "[NetClass]." & nclassName ) = epcbGraphicsItemStateOnEnabled
                .Unlock
            End With
        End If
        ' 90Ohm impedance
        If z90Re.Test(nclassName) Then
            With displayCtrlObj
                .Lock
                .Global.Color( "[NetClass]." & nclassName ) = utilityObj.NewColorPattern( 204, 236, 255, 100, 0, False, False ) 
                .Visible( "[NetClass]." & nclassName ) = epcbGraphicsItemStateOnEnabled
                .Unlock
            End With
        End If
        ' 85Ohm impedance
        If z85Re.Test(nclassName) Then
            With displayCtrlObj
                .Lock
                .Global.Color( "[NetClass]." & nclassName ) = utilityObj.NewColorPattern( 0, 128, 128, 100, 0, False, False ) 
                .Visible( "[NetClass]." & nclassName ) = epcbGraphicsItemStateOnEnabled
                .Unlock
            End With
        End If
        ' 50Ohm impedance
        If z50Re.Test(nclassName) Then
            With displayCtrlObj
                .Lock
                .Global.Color( "[NetClass]." & nclassName ) = utilityObj.NewColorPattern( 102, 0, 51, 100, 0, False, False ) 
                .Visible( "[NetClass]." & nclassName ) = epcbGraphicsItemStateOnEnabled
                .Unlock
            End With
        End If
        ' 45Ohm impedance
        If z45Re.Test(nclassName) Then
            With displayCtrlObj
                .Lock
                .Global.Color( "[NetClass]." & nclassName ) = utilityObj.NewColorPattern( 214, 0, 147, 100, 0, False, False ) 
                .Visible( "[NetClass]." & nclassName ) = epcbGraphicsItemStateOnEnabled
                .Unlock
            End With
        End If
    Next

    Set z45Re = Nothing
    Set z50Re = Nothing
    Set z85Re = Nothing
    Set z90Re = Nothing
    Set z100Re = Nothing

    ' Refresh view
    ' pcbGuiObj.ProcessCommand("View->Fit All")
    pcbGuiObj.ProcessCommand("View->Previous View")
    pcbGuiObj.ProcessCommand("View->Next View")

    pcbAppObj.UnlockServer          
	pcbDocObj.TransactionEnd        
	pcbAppObj.Gui.CursorBusy(False)

End If

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
