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
    schemeFile = masterPath + "/Config/" + "UsrDXFExportTop.edxf"
    outputFile = masterPath + "/Output/DXFExport/" + "DXF_TOP.dxf"

    ' Create a FileSystemObject
    Dim fileSysObj 
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    ' Delete old files
    If fileSysObj.FileExists(schemeFile) = True Then
        Call fileSysObj.DeleteFile(schemeFile,True)
    End If

    oDXFExport.DesignFileName = jobName
    oDXFExport.Scheme(0) = "UsrDXFExportTop"
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

    MsgBox "Finished DXF Export"
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
    Call oDXFExport.PutElementExport(eengElementTypeAssemblyOutlineTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeAssemblyPartNumberTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeAssemblyRefDesTop,"",0)
    
    ' Fiducial Top
    ' Call oDXFExport.PutElementExport(eengElementTypeFiducialPadTeardropTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeFiducialPadTop,"",0)
    
    ' Call oDXFExport.PutElementExport(eengElementTypeGeneralPanelCellTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeGeneratedSilkscreenTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeGlueSpotTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeInsertionOutlineTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeJumperWireTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeManufacturingOutline,"",0)
    
    ' MountingHole
    Call oDXFExport.PutElementExport(eengElementTypeMountingHole,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeMountingHolePad,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeMountingHoleTearDrop,"",0)
    
    ' Part elements
    ' Call oDXFExport.PutElementExport(eengElementTypeNamedSoldermask,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeNamedSoldermaskPads,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypePanelIdentifierTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypePartPads,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypePartPadSMDTeardropsTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypePartPadSMDTOP,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypePartPadsTeardrops,"",0)
    Call oDXFExport.PutElementExport(eengElementTypePartPinHoles,"",0)
    
    ' Part PinNumbers Top
    Call oDXFExport.PutElementExport(eengElementTypePinNumbersTop,"",0)
    
    ' Placementoutline Top
    Call oDXFExport.PutElementExport(eengElementTypePlacementOutlinesTop,"",0)
    
    ' Silkscreen Top
    Call oDXFExport.PutElementExport(eengElementTypeSilkscreenOutlineTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeSilkscreenPartNumberTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeSilkscreenRefDesTop,"",0)
    
    ' Soldermask Top
    ' Call oDXFExport.PutElementExport(eengElementTypeSoldermaskPadsTeardropTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSoldermaskPadsTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSoldermaskTop,"",0)
    
    ' Solderpaste Top
    ' Call oDXFExport.PutElementExport(eengElementTypeSolderpastePadsTeardropTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSolderpastePadsTop,"",0)
    Call oDXFExport.PutElementExport(eengElementTypeSolderpasteTop,"",0)
    
    ' Test Coupon Top
    ' Call oDXFExport.PutElementExport(eengElementTypeTestCouponTop,"",0)
    
    ' Test Point Top
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointAssemblyRefDesTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointPadTeardropTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointPadTop,"",0)
    ' Call oDXFExport.PutElementExport(eengElementTypeTestpointSilkscreenRefDesTop,"",0)
    
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