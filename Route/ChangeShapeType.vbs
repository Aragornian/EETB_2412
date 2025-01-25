Option Explicit

' Add type library
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

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

' Change plane to conductiveshape
pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart(epcbDRCModeNone)

Dim shapeObj, shapeColl, pntsArr, numPoints, activeLayer, shapeNetObj
Select Case arg
    Case 1
        ' Change plane shape to conductive shape
        Set shapeColl = pcbDocObj.PlaneShapes(epcbSelectSelected,0)
        activeLayer = pcbAppObj.Gui.ActiveRouteLayer
        For Each shapeObj In shapeColl
            pntsArr = shapeObj.Geometry.PointsArray
            numPoints = UBound(pntsArr,2) + 1
            Set shapeNetObj = shapeObj.Net
            Call pcbDocObj.PutConductiveArea(activeLayer,shapeNetObj,numPoints,pntsArr,Nothing)
            shapeObj.Delete()
        Next
    Case 2
        ' Change conductive shape to plane shape
        Set shapeColl = pcbDocObj.ConductiveAreas(epcbSelectSelected,0)
        activeLayer = pcbAppObj.Gui.ActiveRouteLayer
        For Each shapeObj In shapeColl
            pntsArr = shapeObj.Geometry.PointsArray
            numPoints = UBound(pntsArr,2) + 1
            Set shapeNetObj = shapeObj.Net
            Call pcbDocObj.PutPlaneShape(activeLayer,numPoints,pntsArr,shapeNetObj,,,,,Nothing)
            shapeObj.Delete()
        Next
End Select

pcbDocObj.TransactionEnd()
pcbAppObj.Gui.CursorBusy(False)

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
