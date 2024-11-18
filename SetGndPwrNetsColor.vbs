Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' Get the GUI object
Dim pcbGuiObj
Set pcbGuiObj = pcbAppObj.Gui

' License the document
ValidateServer(pcbDocObj)

pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart

If pcbAppObj.LockServer = True Then

    ' Get the display control object
    Dim displayCtrlObj, utilityObj
    Set displayCtrlObj = pcbDocObj.ActiveViewEx.DisplayControl
    Set utilityObj = pcbAppObj.Utility

    ' Color GND net
    Dim gndNetsColl, gndNet, gndNetName
    Set gndNetsColl = pcbDocObj.Nets(epcbSelectAll,False,"*GND*")

    For Each gndNet In gndNetsColl
        gndNetName = gndNet.Name
        With displayCtrlObj
            .Lock
            .Global.Color( "[Net]." & gndNetName ) = utilityObj.NewColorPattern( 95, 95, 95, 100, 0, False, False ) 
            .Visible( "[Net]." & gndNetName ) = epcbGraphicsItemStateOnEnabled 
            .Unlock
        End With
    Next

    ' Color power net
    Dim powerNetsArr: powerNetsArr = Array("*VREG*", "*VDD*", "*VCC*", "*VBUS*", "*VSIM*", "*VPH*", "*PWR*", "*VBAT*")
    Dim powerNetsColl, powerNet, powerNetName
    Dim i, j: j = 0
    Dim color, redColor, greenColor, blueColor
    For i = 0 To UBound(powerNetsArr)
        Set powerNetsColl = pcbDocObj.Nets(epcbSelectAll,False,powerNetsArr(i))
        For Each powerNet In powerNetsColl
            powerNetName = powerNet.Name
            color = ChooseColor(j)
            redColor = color(0)
            greenColor = color(1)
            blueColor = color(2)
            With displayCtrlObj
                .Lock
                .Global.Color( "[Net]." & powerNetName ) = utilityObj.NewColorPattern( redColor, greenColor, blueColor, 100, 0, False, False ) 
                .Visible( "[Net]." & powerNetName ) = epcbGraphicsItemStateOnEnabled
                .Unlock
            End With
            If j < 24 Then
                j = j + 1
            Else
                j = 0
            End If
        Next
    Next

    ' Refresh view
    ' pcbGuiObj.ProcessCommand("View->Fit All")
    pcbGuiObj.ProcessCommand("View->Previous View")
    pcbGuiObj.ProcessCommand("View->Next View")

    pcbAppObj.UnlockServer          
	pcbDocObj.TransactionEnd        
	pcbAppObj.Gui.CursorBusy(False)

End If

' Choose color function
Function ChooseColor(colorIndex)

    ' Color list (R,G,B)
    Dim colorsList, colorArr(2)
    colorsList = Array(153, 204, 255, 204,  51,   0, 128, 128,   0,   0, 255, 255, 255,   0,   0,_ 
                     153,  51, 255, 153, 153, 255, 102, 102, 153, 255,  51, 204, 255, 102, 153,_
                     153,  51, 102, 102, 102,  51, 204,   0, 153,  51, 153, 102, 102, 255, 153,_
                     102,   0,  51,  51, 102, 204,   0, 153, 153,  51, 153,  51,  51, 102,   0,_
                     153, 102,  51,  153, 51,   0, 128,   0, 128, 204, 153, 255, 255, 102,   0)
    colorArr(0) = colorsList(colorIndex * 3 + 0)
    colorArr(1) = colorsList(colorIndex * 3 + 1)
    colorArr(2) = colorsList(colorIndex * 3 + 2)
    ChooseColor = colorArr

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
