'Acceptable Usage Policy
'
'  This software is NOT officially supported by Mentor Graphics.
'
'  ####################################################################
'  ####################################################################
'  ## The following  software  is  "freeware" which  Mentor Graphics ##
'  ## Corporation  provides as a courtesy  to our users.  "freeware" ##
'  ## is provided  "as is" and  Mentor  Graphics makes no warranties ##
'  ## with  respect  to "freeware",  either  expressed  or  implied, ##
'  ## including any implied warranties of merchantability or fitness ##
'  ## for a particular purpose.                                      ##
'  ####################################################################
'  ####################################################################
'
' This script <add description here>

Option Explicit     ' This means that all variables must be declared using the 'Dim' statement

'---------------------------------------
'Begin Main program
'---------------------------------------
' Get the application object
Dim pcbApp, dia, but
'Set pcbApp = GetObject(, "MGCPCB.ExpeditionPCBApplication")
Set pcbApp = Application

' Get the active document
Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

' License the document
If (ValidateServer(pcbDoc) = 1) Then

    ' add a reference to the MGCPCB type library in order to use enums
    Scripting.AddTypeLibrary ("MGCPCB.ExpeditionPCBApplication")

    ' add a reference to the MS common dialog type library in order to use enums
    Scripting.AddTypeLibrary ("MSComDlg.CommonDialog")
        
    pcbApp.LockServer
    dia = pcbApp.Gui.ProcessCommand("Route->Assign Net Name") '33050
        
    Dim annDlg
    Set annDlg = pcbApp.Gui.FindDialog("Assign Net Name")
    Set but = annDlg.FindButton("Graphic Selection Mode")

    If Not (but Is Nothing) Then
        but.Click                   ' click the button
    End If

    ' MsgBox ("Click on the Target Net then click on the Copied Trace. Hit <ESC> When done")
    Write("Click on the Target Net then click on the Copied Trace. Hit <ESC> When done")
    'Msgbox("Hit Escape to end")
    pcbApp.UnlockServer
End If

'Msgbox("Could not validate the server. Exiting program.")

'---------------------------------------
'End Main program
'---------------------------------------

'---------------------------------------
'Add local functions, subroutines,
'and methods here
'---------------------------------------

'---------------------------------------
'Local functions
'---------------------------------------

' Write message in Addin windows
Sub Write(sMsg)
	If Not IsObject(pcbApp) Then
		Echo sMsg
	ElseIf Not pcbApp.Addins("Message Window") Is Nothing Then
		pcbApp.Addins("Message Window").Control.AddTab("Output").AppendText sMsg & vbCrLf
        pcbApp.Addins("Message Window").Control.ActivateTab("Output")
	End If
End Sub

'---------------------------------------
' Begin Validate Server Function
'---------------------------------------
Private Function ValidateServer(doc)

    Dim key, licenseServer, licenseToken

    ' Ask Expedition�s document for the key
    key = doc.Validate(0)

    ' Get license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")

    ' Ask the license server for the license token
    licenseToken = licenseServer.GetToken(key)

    ' Release license server
    Set licenseServer = Nothing

    ' Turn off error messages.  Validate may fail if the token is incorrect
    On Error Resume Next
    Err.Clear

    ' Ask the document to validate the license token
    doc.Validate (licenseToken)

    If Err Then
        ValidateServer = 0
    Else
        ValidateServer = 1
    End If

End Function
'---------------------------------------
' End Validate Server Function
'---------------------------------------
