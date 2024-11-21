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

pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart(epcbDRCModeNone)

' Lock server for better performance
If pcbAppObj.LockServer = True Then
    ' Get edit control object
    Dim editControlObj
    Set editControlObj = pcbDocObj.EditorControl

    Dim padsColl, padObj, padShapeArr, padLayerRangeSColl
    padShapeArr = Array(epcbRectPad, epcbOblongPad, epcbRoundPad, epcbSquarePad, epcbCustomPad)

    On Error Resume Next
    Err.Clear()
    Dim i
    For i = 0 To UBound(padShapeArr)
        Set padsColl = editControlObj.Pads(padShapeArr(i))
        For Each padObj In padsColl
            ' Returns the valid layer ranges that can be used for the via span
            Set padLayerRangeSColl = editControlObj.PadEntryViaSpansForPad(padObj)
            ' Set via spans allowed under a pad or pads
            editControlObj.SelPadsViaPosRules(padObj, padShapeArr(i), epcbECAllowViasUnderPad) = True
            editControlObj.SelPadsViaPosRules(padObj, padShapeArr(i), epcbECAllowOffPadOrigin) = True
            If padShapeArr(i) = epcbRectPad Or padShapeArr(i) = epcbOblongPad Then
                editControlObj.SelPadsViaPosRules(padObj, padShapeArr(i), epcbECAlignViaOnLongAxis) = False
            End If
            If padShapeArr(i) <> epcbCustomPad Then
                editControlObj.SelPadsViaPosRules(padObj, padShapeArr(i), epcbECLocateViaAtPadEdge) = False
                editControlObj.SelPadsViaPosRules(padObj, padShapeArr(i), epcbECKeepViaCenterInsidePad) = False
                ' editControlObj.SelPadsViaPosRules(padObj, padShapeArr(i), epcbECKeepViaPadInsidePad) = False
            End If
            editControlObj.PadEntryAllowedPadsViaSpan(padObj, padShapeArr(i)) = padLayerRangeSColl
            ' Set the trace positioning rules for a pad
            editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECExtendedPadEntry) = True
            editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECAllowOddAngle) = True
            If padShapeArr(i) = epcbRectPad Then 
                editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceLongCtr) = True
                editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceLongEdge) = True
                editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceShortCtr) = True
                editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceShortEdge) = True
                editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceCorner) = True
            ElseIf padShapeArr(i) = epcbOblongPad Then
                    editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceCorner) = True
                    editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceLongCtr) = True
                    editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceShortCtr) = True
            ElseIf padShapeArr(i) = epcbSquarePad Then
                    editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceLongCtr) = True
                    editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceLongEdge) = True
                    ' editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceShortCtr) = True
                    ' editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceShortEdge) = True
                    editControlObj.SelPadsTracePosRules(padObj, padShapeArr(i), epcbECPadTraceCorner) = True
            End If
            ' Set Pad Entry global rules (applicable to only rect and oblong pads)
            ' PadEntryGlobalRules method is disabled in Xpedition Layout Team Server
            If pcbAppObj.IsXtremeClient() = False And pcbAppObj.IsXtremeServer() = False Then
                If padShapeArr(i) = epcbRectPad Or padShapeArr(i) = epcbOblongPad Then
                    editControlObj.PadEntryGlobalRules(padShapeArr(i), epcbECPadAdjacentCorner) = True
                    editControlObj.PadEntryGlobalRules(padShapeArr(i), epcbECPadAdjacentLong) = True
                End If
            End If
        Next
    Next
    If Err.Number Then
        Write(Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description)
        Err.Clear()
    Else
        Write("Set Pad Entry Successfully! --- " & Now())
    End If
    pcbAppObj.UnlockServer          
    pcbAppObj.Gui.CursorBusy(False)
    pcbDocObj.TransactionEnd()
End If

' Write message in Addin windows
Sub Write(sMsg)
	If Not IsObject(pcbAppObj) Then
		Echo sMsg
	ElseIf Not pcbAppObj.Addins("Message Window") Is Nothing Then
		pcbAppObj.Addins("Message Window").Control.AddTab("Set Pad Entry").AppendText sMsg & vbCrLf
        pcbAppObj.Addins("Message Window").Control.ActivateTab("Set Pad Entry")
	End If
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
