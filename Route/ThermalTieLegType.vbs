Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' Get user arguments
Dim arg, scriptArgs
Set scriptArgs = ScriptHelper.Arguments
arg = ScriptArgs.Item(3)

' License the document
ValidateServer(pcbDocObj)

Dim pinColl
Set pinColl = pcbDocObj.Pins(epcbSelectSelected)

' Sets the tie leg type for the SMD pins.
Select Case arg
    Case 1
        ' Thermal TieLegFour
        Call SetPinsThermalTieLegType(pinColl, epcbTieLegFour, 0.1, 0.15, (epcbTieLegRotationFixed0 Or epcbTieLegRotationFixed90))
    Case 2
        ' Thermal TieLegNone
        Call SetPinsThermalTieLegType(pinColl, epcbTieLegNone, 0.1, 0.15, (epcbTieLegRotationFixed0 Or epcbTieLegRotationFixed90))
    Case 3
        ' REMOVE_THERMAL_OVERRIDE
        ' Then thermal type is defined by plane parameters
        pcbAppObj.Gui.ProcessCommand(33672)
End Select

' Set smd pins thermal_tieleg type function
Sub SetPinsThermalTieLegType(pinColl, TieLegType, TieLegClearance, TieLegWidth, TieLegRotation)

    ' Turn off DRC
    pcbDocObj.TransactionStart(epcbDRCModeNone)

    Dim pinObj, pinLayer
    For Each pinObj In pinColl
        pinLayer = pinObj.Layer
        pinObj.TieLegRotation(pinLayer) = TieLegRotation
        pinObj.TieLegWidth(pinLayer, epcbUnitMM) = TieLegWidth
        pinObj.TieLegClearance(pinLayer, epcbUnitMM) = TieLegClearance
        pinObj.TieLegType(pinLayer) = TieLegType
    Next

    ' All the changes should be kept
    pcbDocObj.TransactionEnd(True)
    
End Sub

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