' From BasicCAE
Option Explicit


Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

Dim arg ,ScriptArgs
Set ScriptArgs = ScriptHelper.Arguments
arg = ScriptArgs.Item(3)
REM msgbox arg
' License the document
ValidateServer(pcbDocObj)

' Add the type library so that we can use enums
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
'Scripting.AddTypeLibrary("MGCSDD.KeyBindings")
Dim pcbView
Set pcbView = pcbDocObj.ActiveView
Dim ViewExt
Set ViewExt = pcbView.Extrema
Dim ViewWidth,ViewHeight
ViewWidth = ViewExt.MaxX - ViewExt.MinX
ViewHeight = ViewExt.MaxY - ViewExt.MinY
Dim posX,posY
posX = pcbView.MousePositionX
posY = pcbView.MousePositionY

    ' Lock the server for better performance
    pcbAppObj.LockServer
    
	Select Case arg
	Case 1
		pcbView.SetExtents posX - ViewWidth,posY - ViewHeight,posX + ViewWidth,posY + ViewHeight
	Case 2
		pcbView.SetExtents posX - ViewWidth/4,posY - ViewHeight/4,posX + ViewWidth/4,posY + ViewHeight/4
	Case 3
		pcbView.SetExtentsToBoard
	Case 4
		pcbView.Pan pcbView.MousePositionX,pcbView.MousePositionY
	End Select
    
    ' Unlock the server
    pcbAppObj.UnlockServer
    

Private Function ValidateServer(doc)
    
    dim key, licenseServer, licenseToken

    ' Ask Expedition抯 document for the key
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