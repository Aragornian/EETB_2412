Option Explicit

' Add type library
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

Dim viaLayerRangeColl, viaLayerRangeObj
Set viaLayerRangeColl = pcbDocObj.LayerRanges
' msgbox viaLayerRangeColl.count
Dim viapadName, viaPadstacksColl, viaPadStackObj

For Each viaLayerRangeObj In viaLayerRangeColl
    ' msgbox viaLayerRangeObj.Name
    viapadName = viaLayerRangeObj.ViaPadstackName
    ' msgbox viapadName
    ' Set viaPadstacksColl = pcbDocObj.Padstacks("*" & viapadName & "*")
    ' msgbox viaPadstacksColl.Count
    ' For Each viaPadStackObj In viaPadstacksColl
    '     msgbox viaPadStackObj.Name
    ' Next
Next
Dim vpobj
Set viaPadstacksColl = pcbDocObj.Padstacks("*V0.4D0.2*")
For Each viaPadStackObj In viaPadstacksColl
    Set vpobj = viaPadStackObj
Next
Dim setupp
Set setupp = pcbDocObj.SetupParameter
Dim viaspanobj
Set viaspanobj = setupp.PutViaSpan(4,9,vpobj)
msgbox viaspanobj.Name
viaspanobj.SameNetViaClearanceOverride(viaspanobj, epcbUnitCurrent) = 0.1
viaspanobj.delay = 10
' msgbox   viaspanobj.SameNetViaClearanceOverride(viaspanobj, epcbUnitCurrent)
' pcbDocObj.UpdatePhysicalReuseInstances()
Call setupp.PutViaSpan(4,9,vpobj)


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
