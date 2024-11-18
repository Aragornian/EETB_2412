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

' Change plane to conductiveshape
pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart(epcbDRCModeNone)
If pcbAppObj.LockServer() = True Then
    Dim planeShapeObj, planeShapesColl, pntsArr, numPoints, activeLayer, planeShapeNetObj

    Set planeShapesColl = pcbDocObj.PlaneShapes(epcbSelectSelected,0)
    activeLayer = pcbAppObj.Gui.ActiveRouteLayer
    For Each planeShapeObj In planeShapesColl
        pntsArr = planeShapeObj.Geometry.PointsArray
        numPoints = UBound(pntsArr,2) + 1
        Set planeShapeNetObj = planeShapeObj.Net
        Call pcbDocObj.PutConductiveArea(activeLayer,planeShapeNetObj,numPoints,pntsArr,Nothing)
        planeShapeObj.Delete()
    Next
    pcbDocObj.TransactionEnd()
    pcbAppObj.Gui.CursorBusy(False)
    pcbAppObj.UnlockServer()
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
