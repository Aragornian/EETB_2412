'*****************************************************************************
' Unpublished work. Copyright 2019 Siemens
'
' This material contains trade secrets or otherwise confidential information
' owned by Siemens Industry Software Inc. or its affiliates (collectively,
' "SISW"), or its licensors. Access to and use of this information is strictly
' limited as set forth in the Customer's applicable agreements with SISW.
'*****************************************************************************
Option Explicit

Dim sComVersion : sComVersion = Scripting.GetEnvVariable("EXP_PROG_ID_VER") : If sComVersion = "" Then sComVersion = "1"
Dim jobName, masterPath 
Dim args
Dim argsCnt
Dim outputFile

Set args = ScriptHelper.Arguments
argsCnt = args.count
If argsCnt = 3 Or argsCnt = 4 Then
    jobName = args.item(3)
Else
    ' Call from Expedition
    Dim app
    Dim docObj
    ' Set app = GetObject(, "MGCPCB.ExpeditionPCBApplication." & sComVersion)
    Set app = Application
    Set docObj = GetLicensedDoc(app)
    jobName = docObj.FullName
    masterPath = docObj.MasterPath
End If

If app.LockServer = True Then  
    app.Gui.CursorBusy(True)

    ' Run export dxf
    ExportDXF
    
    app.Gui.CursorBusy(False)
    app.UnlockServer
End If

'************************************************************************
'*** Main Function
'************************************************************************
Sub ExportDXF()
    ' Create DXFExport Engine object
    Dim oDXFExport
    Set oDXFExport = CreateObject("MGCPCBEngines.DXFExport." & sComVersion)
    If oDXFExport is Nothing Then
        MsgBox "Failed to create DXF Export object"
        Exit Sub
    End If
    Scripting.AddTypeLibrary("MGCPCBEngines.DXFExport." & sComVersion)
    
    ' Set output file path and scheme file path
    Dim schemeFile
    schemeFile = masterPath + "/Config/" + "UsrDXFExportBot.edxf"
    outputFile = masterPath + "/Output/DXFExport/" + "DXF_BOT.dxf"

    ' Create a FileSystemObject
    Dim fileSysObj 
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    ' Delete old files
    If fileSysObj.FileExists(schemeFile) = True Then
        Call fileSysObj.DeleteFile(schemeFile,True)
    End If

    oDXFExport.DesignFileName = jobName
    oDXFExport.Scheme(0) = "UsrDXFExportBot"
    oDXFExport.OutputFileName = outputFile
    Call  SetDXFExportSettings(oDXFExport)
    ' Run the DXFExport Engine
    On Error Resume Next
    oDXFExport.Go()
    On Error Goto 0
    
    ' Check errors
    Dim oError, oErrors
    Set oErrors = oDXFExport.Errors
    For Each oError In oErrors
    	Write oError.ErrorString
    Next

    Write "Finished DXF_Bottom Export"
    Write outputFile
End Sub

Sub Write(sMsg)
	If Not IsObject(app) Then
		Echo sMsg
	ElseIf Not app.Addins("Message Window") Is Nothing Then
		app.Addins("Message Window").Control.AddTab("RunDXFExport").AppendText sMsg & vbCrLf
	End If
End Sub

'*************************
' set DXFExportSettings
'*************************
Sub SetDXFExportSettings(oDXFExport)
    oDXFExport.FillPads = False
    oDXFExport.MirrorOutput = False
    oDXFExport.DXFUnits = eengUnitMM
    Dim nLayerCount
    nLayerCount = oDXFExport.LayerCount
    Call SetDXFExportElements(oDXFExport)
End Sub

'*************************
' set DXFExportElements
'*************************
Sub SetDXFExportElements(oDXFExport)
    ' Board elements
    Call oDXFExport.PutElementExport(eengElementTypeBoardOutline, "",0)
    Call oDXFExport.PutElementExport(eengElementTypeBoardCavity,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeContour,"",0)
    
    ' Part Assembly Top
    Call oDXFExport.PutElementExport(eengElementTypeAssemblyOutlineBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeAssemblyPartNumberBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeAssemblyRefDesBottom,"",0)
    
    ' Fiducial Top
    ' Call oDXFExport.PutElementExport(eengElementTypeFiducialPadTeardropBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeFiducialPadBottom,"",0)
    
    ' Call oDXFExport.PutElementExport(eengElementTypeGeneralPanelCellBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeGeneratedSilkscreenBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeGlueSpotBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeInsertionOutlineBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeJumperWireBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeManufacturingOutline,"",0)
    
    ' MountingHole
    Call oDXFExport.PutElementExport(eengElementTypeMountingHole,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeMountingHolePad,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeMountingHoleTearDrop,"",0)
    
    ' Part elements
    ' Call oDXFExport.PutElementExport(eengElementTypeNamedSoldermask,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeNamedSoldermaskPads,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypePanelIdentifierBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypePartPads,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypePartPadSMDTeardropsBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypePartPadSMDBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypePartPadsTeardrops,"",0)
    Call oDXFExport.PutElementExport(eengElementTypePartPinHoles,"",0)
    
    ' Part PinNumbers Top
    Call oDXFExport.PutElementExport(eengElementTypePinNumbersBottom,"",0)
    
    ' Placementoutline Top
    Call oDXFExport.PutElementExport(eengElementTypePlacementOutlinesBottom,"",0)
    
    ' Silkscreen Top
    Call oDXFExport.PutElementExport(eengElementTypeSilkscreenOutlineBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeSilkscreenPartNumberBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeSilkscreenRefDesBottom,"",0)
    
    ' Soldermask Top
    ' Call oDXFExport.PutElementExport(eengElementTypeSoldermaskPadsTeardropBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSoldermaskPadsBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSoldermaskBottom,"",0)
    
    ' Solderpaste Top
    ' Call oDXFExport.PutElementExport(eengElementTypeSolderpastePadsTeardropBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSolderpastePadsBottom,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSolderpasteBottom,"",0)
    
    ' Test Coupon Top
    ' Call oDXFExport.PutElementExport(eengElementTypeTestCouponBottom,"",0)
    
    ' Test Point Top
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointAssemblyRefDesBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointPadTeardropBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointPadBottom,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointSilkscreenRefDesBottom,"",0)
    
    ' Tooling Holes
    Call oDXFExport.PutElementExport(eengElementTypeToolingHoles,"",0)

    ' Via Holes
    ' Call oDXFExport.PutElementExport(eengElementTypeViaHoles,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeViaPads,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeViaPadsTeardrops,"",0)
End Sub

Public Function GetLicensedDoc(appObj)
    On Error Resume Next
    Dim key, licenseServer, licenseToken, docObj
    Set GetLicensedDoc = Nothing
    ' collect the active document
    Set docObj = appObj.ActiveDocument
    If Err Then
        Call appObj.Gui.StatusBarText("No active document: " & Err.Description, epcbStatusFieldError)
        Exit Function
    End If
    ' Ask Expedition�s document for the key
    key = docObj.Validate(0)
    ' Get token from license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application." & sComVersion)
    licenseToken = licenseServer.GetToken(key)
    Set licenseServer = Nothing
    ' Ask the document to validate the license token
    Err.Clear
    Call docObj.Validate(licenseToken)
    If Err Then
        Call appObj.Gui.StatusBarText("No active document license: " & Err.Description, epcbStatusFieldError)
        Exit Function
    End If
    ' everything is OK, return document
    Set GetLicensedDoc = docObj
End Function