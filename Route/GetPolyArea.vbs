Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' and add its typelib to give us access to the public constants 
Scripting.AddTypeLibrary ("MGCPCBEngines.MaskEngine") 
Scripting.AddTypeLibrary ("MGCPCB.Application")

' License the document
ValidateServer(pcbDocObj)

pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart(epcbDRCModeNone)

' New a object filter that contains all objects
Dim objfilterObj
Set objfilterObj = pcbAppObj.Utility.NewObjectFilter()
objfilterObj.ExcludeIfParentSelect = True

' Get selected object ids
Dim selObjectIdList
selObjectIdList = pcbDocObj.SelectedObjectIds(objfilterObj)

If UBound(selObjectIdList) = -1 Then
   Call MsgBox("Please select polygons or components", vbCritical, "Get Polygon Area")
Else 
    ' get masking engine 
    Dim mskeng
    Set mskeng = CreateObject("MGCPCBEngines.MaskEngine") 
    mskeng.currentUnit = emeUnitMM

    Dim i, selObj, pnts
    Dim mskPoly, shapePoly, mskPolyArea, totalPolyArea, polyAreaStr, mskCompArea, totalCompArea, compAreaStr
    totalPolyArea = 0
    totalCompArea = 0
    For i = 0 To UBound(selObjectIdList)
        Set selObj = pcbDocObj.FindObjectById(selObjectIdList(i))
        Dim objType
        objType = selObj.Type
        If objType = epcbCompGeneral And selObj.Name <> "Board Outline" Then
            ' Area of component, define by placementoutline
            Dim compMaxArea: compMaxArea = 0
            Dim placeoutlinesColl, placeoutlinesObj
            Set placeoutlinesColl = selObj.PlacementOutlines
            If placeoutlinesColl.Count = 0 Then
                MsgBox("no placeoutline found in " & selObj.Name & " , fail to get component area !") 
            Else
                ' Create a mask to hold a polygon data
                Set mskPoly = mskeng.Masks.Add
                ' Get collection of shapes on that mask, this will initially be empty
                Set shapePoly = mskPoly.Shapes
                For Each placeoutlinesObj In placeoutlinesColl
                    ' Add a mask shape by pointsarray
                    pnts = placeoutlinesObj.Geometry.PointsArray
                    Call shapePoly.AddByPointsArray(1 + UBound(pnts, 2), pnts)
                    ' Compute shape area
                    mskCompArea = mskPoly.Area(emeUnitMM)
                    If mskCompArea > compMaxArea Then 
                        compMaxArea = mskCompArea
                    End If
                    shapePoly.Delete()
                Next
                totalCompArea = totalCompArea + compMaxArea
            End If
        Else 
            ' Check runtime errors here
            On Error Resume Next
            Err.Clear()

            ' Create a mask to hold a polygon data
            Set mskPoly = mskeng.Masks.Add
            ' Get collection of shapes on that mask, this will initially be empty
            Set shapePoly = mskPoly.Shapes
            ' Add a mask shape by pointsarray
            pnts = selObj.Geometry.PointsArray
            Call shapePoly.AddByPointsArray(1 + UBound(pnts, 2), pnts)
            ' Compute shape area
            mskPolyArea = mskPoly.Area(emeUnitMM)
            ' If polyArea = 0 ,then some errors occur
            If mskPolyArea <> 0 Then
                totalPolyArea = totalPolyArea + mskPolyArea
                polyAreaStr = polyAreaStr & selObj.Name & " area = " & FormatNumber(mskPolyArea, 2) & " mm" & Chr(10)
            End If
            shapePoly.Delete()

            If Err.Number Then
                ' Check errors
                MsgBox(selObj.Name & " is an Unsupported Object !")
            End If
            Err.Clear()
        End If  
    Next
    ' Components are selected, show their area
    If totalCompArea <> 0 Then
        compAreaStr = "Total components area = " & FormatNumber(totalCompArea, 2) & Chr(10) & Chr(10)
    End If
    ' Polygon are selected, show their area
    If totalPolyArea <> 0 Then
        polyAreaStr = "Total Polygon area = " & FormatNumber(totalPolyArea, 2) & " mm" & Chr(10) & Chr(10) & polyAreaStr
    End If
    If totalCompArea <> 0 Or totalPolyArea <> 0 Then
        Call MsgBox(compAreaStr & polyAreaStr, vbInformation, "Get Area")
    End If
    ' release the mask engine now we are finished with it
    Set mskeng = Nothing
End If

pcbAppObj.Gui.CursorBusy(False)
pcbDocObj.TransactionEnd()

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
