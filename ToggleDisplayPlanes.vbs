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

' Toggle display planes
Dim planesDisplayEnum
planesDisplayEnum = displayCtrlObj.Option("Option.Planes.Data.Fill")
Select Case planesDisplayEnum
    Case epcbGraphicsItemStateOnEnabled
        displayCtrlObj.Option("Option.Planes.Data.Fill") = epcbGraphicsItemStateOffEnabled
    Case epcbGraphicsItemStateOffEnabled
        displayCtrlObj.Option("Option.Planes.Enabled") = epcbGraphicsItemStateOffEnabled
    Case epcbGraphicsItemStateOnNotEnabled
        displayCtrlObj.Option("Option.Planes.Enabled") = epcbGraphicsItemStateOnEnabled
    Case epcbGraphicsItemStateOffNotEnabled
        displayCtrlObj.Option("Option.Planes.Data.Fill") = epcbGraphicsItemStateOnEnabled
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
