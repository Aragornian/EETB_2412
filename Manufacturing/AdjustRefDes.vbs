Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' get masking engine 
Dim mskeng
Set mskeng = CreateObject("MGCPCBEngines.MaskEngine") 
' and add its typelib to give us access to the public constants 
Scripting.AddTypeLibrary ("MGCPCBEngines.MaskEngine") 
Scripting.AddTypeLibrary ("MGCPCB.Application")

' License the document
ValidateServer(pcbDocObj)

pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart(epcbDRCModeNone)

' Lock server for better performance
If pcbAppObj.LockServer = True Then
    Dim compsColl, compObj
    ' Set compsColl = pcbDocObj.Components(epcbSelectSelected, epcbCompGeneral, epcbCelltypePackage)
    Set compsColl = pcbDocObj.Components(epcbSelectAll, epcbCompGeneral, epcbCelltypePackage)

    Dim compCenterX, compCenterY, compOrientation, compSide, compRefDes
    Dim fabTextColl, fabTextObj, fabTextNameLen, fabTextFormat
    Dim fabTextSetHight, fabTextSetWidth
    Dim placementoutlinesColl, placementoutlineObj
    Dim pnts, pntsArr, newFabGfxObj, radian, originalExtrema, originalDx, originalDy, dx, dy
    Dim rotatedPointArr()
    Dim errCompStr: errCompStr = ""

    For Each compObj In compsColl
        ' Get component infomation
        compCenterX = compObj.CenterX(epcbUnitCurrent)
        compCenterY = compObj.CenterY(epcbUnitCurrent)
        compOrientation = compObj.Orientation(epcbAngleUnitDegrees)
        radian = compOrientation*3.14159/180
        compSide = compObj.Side
        compRefDes = compObj.RefDes

        ' Get fabrication silkscreen text of component
        Set fabTextColl = compObj.FabricationLayerTexts(epcbFabSilkscreen, epcbSelectAll, epcbSideTopOrBottom)

        ' Formate silkscreen refdes
        For Each fabTextObj In fabTextColl
            ' Pass to next iteration if silkscreen text is not the same as component refdes
            If fabTextObj.Name = compRefDes Then
                Set fabTextFormat = fabTextObj.Format
                ' Formate text
                With fabTextFormat
                    .Font = "Microsoft Sans Serif"
                    .AspectRatio = 1
                    .PenWidth(epcbUnitCurrent) = 0.01
                    .Bold = False
                    .Italic = False
                    .Underline = False
                    .HorizontalJust = epcbJustifyHCenter
                    .VerticalJust = epcbJustifyVCenter
                End With
                ' Adjust text position
                fabTextObj.PositionX = compCenterX
                fabTextObj.PositionY = compCenterY
                ' Mirror text if component is on bottom side
                If compSide = epcbSideBottom Then
                    fabTextFormat.Mirrored = True
                Else
                    fabTextFormat.Mirrored = False
                End If
                ' Adjust text orientation for Reading convenience
                ' PlacementOutlines must exist in component
                ' Check runtime errors here
                On Error Resume Next
                Err.Clear()
                Set placementoutlinesColl = compObj.PlacementOutlines
                Set placementoutlineObj = GetPlacementOutlineObj(compObj)
                pntsArr = placementoutlineObj.Geometry.PointsArray
                pnts = UBound(pntsArr,2) + 1
                ReDim rotatedPointArr(2, UBound(pntsArr,2))
                ' Get points array of component with 0 degrees rotation
                Dim i
                For i = 0 To UBound(pntsArr,2)
                    rotatedPointArr(0,i) = pntsArr(0,i) * Cos(-radian) - pntsArr(1,i) * Sin(-radian)
                    rotatedPointArr(1,i) = pntsArr(0,i) * Sin(-radian) + pntsArr(1,i) * Cos(-radian)
                    rotatedPointArr(2,i) = pntsArr(2,i)
                Next
                ' New a assembly fabrication figure to get the extrema of points array
                Set newFabGfxObj = pcbDocObj.PutFabricationLayerGfx(epcbFabAssembly, compSide, 0, pnts, rotatedPointArr, True, Nothing, epcbUnitCurrent)
                Set originalExtrema = newFabGfxObj.Extrema
                originalDx = originalExtrema.MaxX - originalExtrema.MinX
                originalDy = originalExtrema.MaxY - originalExtrema.MinY
                dx = CDbl(FormatNumber(originalDx,2))
                dy = CDbl(FormatNumber(originalDy,2))
                newFabGfxObj.Delete()
                ' Set orientation of component
                ' Text orientation always along the long side
                If compOrientation > 0 And compOrientation <= 90 Then
                    If dx >= dy Then
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation
                    Else
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation + 270
                    End If
                ElseIf compOrientation > 90 And compOrientation <= 180 Then
                    If dx >= dy Then
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation + 180
                    Else
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation - 90 
                    End If
                ElseIf compOrientation > 180 And compOrientation <= 270 Then
                    If dx >= dy Then
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation - 180
                    Else
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation + 90
                    End If
                Else
                    If dx >= dy Then
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation
                    Else
                        fabTextFormat.Orientation(epcbAngleUnitDegrees) = compOrientation + 90
                    End If   
                End If
                ' Adjust text size
                fabTextNameLen = Len(fabTextObj.Name)
                If dx >= dy Then
                    fabTextSetHight = originalDy * 0.6
                    fabTextSetWidth = fabTextNameLen * fabTextSetHight
                    While fabTextSetWidth > originalDx
                        fabTextSetHight = fabTextSetHight * 0.9
                        fabTextSetWidth = fabTextSetWidth * 0.9
                    Wend
                    fabTextFormat.Height = fabTextSetHight
                Else   
                    fabTextSetHight = originalDx * 0.6
                    fabTextSetWidth = fabTextNameLen * fabTextSetHight
                    While fabTextSetWidth > originalDy
                        fabTextSetHight = fabTextSetHight * 0.9
                        fabTextSetWidth = fabTextSetWidth * 0.9
                    Wend
                    fabTextFormat.Height = fabTextSetHight
                End If
                If Err.Number Then
                    errCompStr = errCompStr & compObj.Name & Chr(10)
                End If
            End If
        Next
    Next
    If Err.Number Then
        ' Throw an error information
        MsgBox("Can't find placementoutlines, fail to Adjust silkscreen :" & Chr(10) & errCompStr)
    End If
    Err.Clear()
    pcbAppObj.UnlockServer          
    pcbAppObj.Gui.CursorBusy(False)
    pcbDocObj.TransactionEnd()
End If

' Get component max placementoutline object function
Function GetPlacementOutlineObj(compObj)

    Dim mskPoly, shapePoly, mskArea, beforeArea, pnts
    Dim placeoutlinesObj

    Set placementoutlinesColl = compObj.PlacementOutlines
    beforeArea = 0

    ' Create a mask to hold a polygon data
    Set mskPoly = mskeng.Masks.Add
    ' Get collection of shapes on that mask, this will initially be empty
    Set shapePoly = mskPoly.Shapes
    For Each placeoutlinesObj In placementoutlinesColl
        ' Add a mask shape by pointsarray
        pnts = placeoutlinesObj.Geometry.PointsArray
        Call shapePoly.AddByPointsArray(1 + UBound(pnts, 2), pnts)
        ' Compute shape area
        mskArea = mskPoly.Area(emeUnitMM)
        If mskArea > beforeArea Then 
            beforeArea = mskArea
            Set GetPlacementOutlineObj = placeoutlinesObj
        End If
        shapePoly.Delete()
    Next

End Function

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
