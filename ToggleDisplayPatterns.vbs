Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

' Get the display control object
Dim displayCtrlObj
Set displayCtrlObj = pcbDocObj.ActiveViewEx.DisplayControl

' Toggle display patterns
Dim currentPatternEnum
currentPatternEnum = displayCtrlObj.Option("Option.FillPatterns")
Select Case currentPatternEnum
    Case epcbGraphicsItemStateOnEnabled
        displayCtrlObj.Option("Option.FillPatterns") = epcbGraphicsItemStateOffEnabled
    Case epcbGraphicsItemStateOffEnabled
        displayCtrlObj.Option("Option.FillPatterns") = epcbGraphicsItemStateOnEnabled
End Select

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
