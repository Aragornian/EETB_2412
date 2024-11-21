Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

' Variable to hold the Command object
Dim cmdObj

' Register a new command
Set cmdObj = pcbAppObj.Gui.RegisterCommand("Custom Place Via")

' Create an array to contain points
Dim pntsArr()
Dim i : i = 0
Dim drawObj

' Attach events to the command object
Call Scripting.AttachEvents(cmdObj, "cmdObj")

' Keep the script from exiting.
Scripting.DontExit = True

' Command event handlers

Function cmdObj_OnTerminate()
    ' Release the command
    Set cmdObj = Nothing
End Function

' Function cmdObj_OnMouseMove(button, flags, x, y)

'     pcbAppObj.Gui.ProcessCommand("View->Unhighlight All")
'     Dim positionObj
'     Set positionObj = pcbDocObj.Pick(x, y, x, y, epcbObjectClassAll, Nothing, True, False)

'     Dim pos
'     For Each pos In positionObj
'         pos.Highlighted = True
'     Next

' End Function

' Function cmdObj_OnMouseMove(button, flags, x, y)

'     pcbAppObj.Gui.ProcessCommand("View->Unhighlight All")
'     Dim positionObj
'     Set positionObj = pcbDocObj.Pick(x, y, x, y, epcbObjectClassAll, Nothing, True, False)

'     Dim pos
'     For Each pos In positionObj
'         pos.Highlighted = True
'     Next

' End Function

Function cmdObj_OnMouseClk(button, flags, x, y)

    ReDim Preserve pntsArr(2,i)
    pntsArr(0,i) = x
    pntsArr(1,i) = y
    pntsArr(2,i) = 0

    i = i + 1

    If i > 1 Then
        If IsEmpty(drawObj) Then
            Set drawObj = pcbDocObj.PutFabricationLayerGfxEx(epcbFabAssembly, epcbSideTop, 0, i, pntsArr, False, Nothing, Nothing, epcbUnitCurrent)
        Else
            drawObj.Delete
            Set drawObj = pcbDocObj.PutFabricationLayerGfxEx(epcbFabAssembly, epcbSideTop, 0, i, pntsArr, False, Nothing, Nothing, epcbUnitCurrent)
        End If
    End If

End Function



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
