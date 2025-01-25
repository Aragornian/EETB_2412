Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

Dim pinObj, pinColl, pinLayer
Set pinColl = pcbDocObj.Pins(epcbSelectSelected)

' Sets the tie leg type for the SMD pin.
' Turn off DRC
pcbDocObj.TransactionStart(epcbDRCModeNone)

For Each pinObj In pinColl
    pinLayer = pinObj.Layer
    ' pinObj.TieLegRotation(pinLayer) = (epcbTieLegRotationFixed0 Or epcbTieLegRotationFixed90)
    ' pinObj.TieLegWidth(pinLayer, epcbUnitMM) = 0.15
    pinObj.TieLegClearance(pinLayer, epcbUnitMM) = 0.1
    pinObj.TieLegType(pinLayer) = epcbTieLegNone
Next

' All the changes should be kept
pcbDocObj.TransactionEnd(True)

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