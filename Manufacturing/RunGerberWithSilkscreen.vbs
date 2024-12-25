Option Explicit 
 
Dim args, app, docObj, jobName, jobConfigPath, masterPath
Set args = ScriptHelper.Arguments 
If args.Count = 3 Then 
    ' Call from command line 
    ' mgcscript RunGerber.vbs [job name] 
    jobName = args.item(3) 
Else 
    ' Call from Xpedition Layout 
    ' Set app = GetObject(, "MGCPCB.ExpeditionPCBApplication") 
    Set app = Application
    Set docObj = GetLicensedDoc(app) 
    jobName = docObj.FullName 
    masterPath = docObj.MasterPath
End If 
Scripting.AddTypeLibrary("MGCPCBEngines.Gerber")
Scripting.AddTypeLibrary("Scripting.FileSystemObject")

' Dim gMacro, dx, dy, xCopies, yCopies, dxCopiesX, dyCopiesY, gbrMatchineFormatFileName, FlashPadsFlag
Const gMacro = FALSE
Const dx = 0
Const dy = 0
Const xCopies = 1
Const yCopies = 1
Const dxCopiesX = 0
Const dyCopiesY = 0
Const gbrMatchineFormatFileName = "UserGerberMachineFile.gmf"
Const FlashPadsFlag = TRUE

' Create a FileSystemObject
Dim fileSysObj, gbrOutputPath
gbrOutputPath = masterPath + "Output\Gerber"
MsgBox gbrOutputPath
Set fileSysObj = CreateObject("Scripting.FileSystemObject")
' Delete existed file
If fileSysObj.FolderExists(gbrOutputPath) = True Then
    fileSysObj.DeleteFolder(gbrOutputPath)
End If

If app.LockServer = True Then  
    app.Gui.CursorBusy(True)
    
    ' Save file
    app.Gui.SuppressTrivialDialogs = True
    app.Gui.ProcessCommand("File->Save")
    app.Gui.SuppressTrivialDialogs = False
    ' Run Gerber engine
    Call RunGerber() 

    app.Gui.CursorBusy(False)
    app.UnlockServer
End If
 
' MsgBox "Output Successfully!", VBInformation, "RunGerber.vbs" 
 
'************************************************************************ 
Sub RunGerber() 
    'On Error Resume Next 
    
    ' Create Gerber Engine object 
    Dim oGerberEngine 
    Set oGerberEngine = CreateObject("MGCPCBEngines.Gerber") 
 
    ' Set the design file name 
    oGerberEngine.DesignFileName = jobName 
 
    ' Setup files for Gerber output 
    Call SetupGerberParameters(oGerberEngine)
    Call SetupMachineFormat(oGerberEngine) 
    Call SetupOutputFiles(oGerberEngine)
 
    ' Run the Gerber Engine 
    oGerberEngine.Go
    ' Write gerber plot setup file for the Gerber engine to load the Gerber Output dialog box
    WriteGerberSetupFile(oGerberEngine)

    ' Check errors
    Dim oError, oErrors
    Set oErrors = oGerberEngine.Errors
    For Each oError In oErrors
    	Write oError.ErrorString
    Next
    If oErrors.Count = 0 Then
        Write("Export Gerber Data to Folder " & masterPath & "Output\Gerber\")
        Write("Export Gerber Successfully! " & Now())
    End If
End Sub 

Sub Write(sMsg)
	If Not IsObject(app) Then
		Echo sMsg
	ElseIf Not app.Addins("Message Window") Is Nothing Then
		app.Addins("Message Window").Control.AddTab("RunGerber").AppendText sMsg & vbCrLf
        app.Addins("Message Window").Control.ActivateTab("RunGerber")
	End If
End Sub

Sub SetupGerberParameters(oGerberEngine)
    'On Error Resume Next 
    oGerberEngine.GenerateMacros = gMacro
    oGerberEngine.OffsetX(eengUnitCurrent) = dx
    oGerberEngine.OffsetY(eengUnitCurrent) = dy
    oGerberEngine.XAxisCopies = xCopies
    oGerberEngine.YAxisCopies = yCopies
    oGerberEngine.SpaceBetweenCopiesX(eengUnitCurrent) = dxCopiesX
    oGerberEngine.SpaceBetweenCopiesY(eengUnitCurrent) = dyCopiesY
    ' oGerberEngine.IncludeVariantData = FALSE
End Sub
 
Sub SetupMachineFormat(oGerberEngine) 
    'On Error Resume Next 
    Dim oGerberMachineFormat 
    Set oGerberMachineFormat = oGerberEngine.MachineFormat 
    oGerberMachineFormat.FileName = "./Config/" & gbrMatchineFormatFileName
    oGerberMachineFormat.DataType = eengGerber274X 
    oGerberMachineFormat.DataMode = TRUE 
    oGerberMachineFormat.StepMode = eengStepAbsolute 
    oGerberMachineFormat.DataFormatLeadingDigits = 2 
    oGerberMachineFormat.DataFormatTrailingDigits = 6
    oGerberMachineFormat.ZeroTruncation = eengZeroTruncationLeading 
    oGerberMachineFormat.CharacterSet = eengCharacterSetASCII 
    oGerberMachineFormat.ArcStyle = eengArcStyleQuadrant 
    oGerberMachineFormat.Delimiter = eengDataDelimiterStar 
    oGerberMachineFormat.Comments = TRUE 
    oGerberMachineFormat.SequenceNumbering = FALSE 
    oGerberMachineFormat.Unit = eengUnitInch 
    oGerberMachineFormat.PolygonFillMethod = eengPolygonFillRaster 
    oGerberMachineFormat.RecordLength = 0 
End Sub 
 
Sub SetupOutputFiles(oGerberEngine) 
    'On Error Resume Next  
    Dim oGerberOutputFiles, oGOFAddEach, oCLCount, oBoardItems 
    Set oGerberOutputFiles = oGerberEngine.OutputFiles 
    oCLCount = docObj.LayerCount
    '----------------------------------------------------------------------'
    '                       EtchLayer                                      '
    '----------------------------------------------------------------------'  
    Dim counterInt 
    For counterInt = 1 to oCLCount Step 1
        ' Set Gerber Output FileName
        Dim oGerberOutputFileName
        oGerberOutputFileName = "EtchLayer" + CStr(counterInt) + ".gdo"

        Set oGOFAddEach = oGerberOutputFiles.Add(oGerberOutputFileName) 
        oGOFAddEach.FlashPads = FlashPadsFlag 
        oGOFAddEach.HeaderText = "Gerber Output File" + oGerberOutputFileName
        oGOFAddEach.TrailerText = "Gerber Output File" + oGerberOutputFileName
        oGOFAddEach.ProcessUnconnectedPads = TRUE 
        'oGOFAddEach.DCodeMappingFileName = "C:\MentorGraphics\EEVX.2.10\SDD_HOME\standard\config\pcb\sample.dmf" 
        
        ' Add board items individually
        Set oBoardItems = oGOFAddEach.BoardItems  
        oBoardItems.Add eengBoardItemBoardCavities               
        oBoardItems.Add eengBoardItemBoardOutline               
        oBoardItems.Add eengBoardItemContours                                          
        ' Add all cell types individually 
        Call AddCellTypes(oGOFAddEach)     
        ' Add all conductive items individually 
        oGOFAddEach.ConductiveLayer = counterInt
        Call AddConductiveItems(oGOFAddEach)
    Next
    '----------------------------------------------------------------------'
    '                       Soldermask_Top                                 '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("SoldermaskTop.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File SoldermaskTop.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File SoldermaskTop.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems 
    oBoardItems.Add eengBoardItemBoardCavities               
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours
    oBoardItems.Add eengBoardItemSoldermaskTop                                             
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach) 
    '----------------------------------------------------------------------'
    '                       Soldermask_Bottom                              '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("SoldermaskBottom.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File SoldermaskBottom.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File SoldermaskBottom.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems 
    oBoardItems.Add eengBoardItemBoardCavities               
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours
    oBoardItems.Add eengBoardItemSoldermaskBottom                                           
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach) 
    '----------------------------------------------------------------------'
    '                       SolderPaste_Top                                '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("SolderPasteTop.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File SolderPasteTop.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File SolderPasteTop.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems            
    oBoardItems.Add eengBoardItemBoardCavities               
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours
    oBoardItems.Add eengBoardItemSolderpasteTop                                           
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach) 
    '----------------------------------------------------------------------'
    '                       SolderPaste_Bottom                             '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("SolderPasteBottom.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File SolderPasteBottom.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File SolderPasteBottom.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems            
    oBoardItems.Add eengBoardItemBoardCavities             
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours
    oBoardItems.Add eengBoardItemSolderpasteBottom                                            
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach) 
    '----------------------------------------------------------------------'
    '                       Silkscreen_Top                                 '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("SilkscreenTop.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File SilkscreenTop.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File SilkscreenTop.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems            
    oBoardItems.Add eengBoardItemBoardCavities             
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours
    oBoardItems.Add eengBoardItemGeneratedSilkscreenTop                                            
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach) 
    '----------------------------------------------------------------------'
    '                       Silkscreen_Bottom                              '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("SilkscreenBottom.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File SilkscreenBottom.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File SilkscreenBottom.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems            
    oBoardItems.Add eengBoardItemBoardCavities             
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours
    oBoardItems.Add eengBoardItemGeneratedSilkscreenBottom                                            
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach)
    '----------------------------------------------------------------------'
    '                       Board_Outline                                  '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("BoardOutline.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File BoardOutline.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File BoardOutline.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems            
    oBoardItems.Add eengBoardItemBoardCavities               
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours                                         
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach) 
    '----------------------------------------------------------------------'
    '                       DrillDrawingThrough                            '
    '----------------------------------------------------------------------'
    Set oGOFAddEach = oGerberOutputFiles.Add("DrillDrawingThrough.gdo")
    oGOFAddEach.HeaderText = "Gerber Output File DrillDrawingThrough.gdo" 
    oGOFAddEach.TrailerText = "Gerber Output File DrillDrawingThrough.gdo"
    ' Add board items individually 
    Set oBoardItems = oGOFAddEach.BoardItems  
    oBoardItems.Add eengBoardItemBoardCavities               
    oBoardItems.Add eengBoardItemBoardOutline               
    oBoardItems.Add eengBoardItemContours 
    oBoardItems.Add eengBoardItemDrillDrawingThru                                 
    ' Add all cell types individually 
    Call AddCellTypes(oGOFAddEach) 
    '----------------------------------------------------------------------'
    '                       DrillDrawing Separate                          '
    '----------------------------------------------------------------------'
    ' Dim uNCDrillChartName,uobj
    ' For Each uobj in docObj.UserLayers
    '     uNCDrillChartName = uobj.Name
    '     If Instr(uNCDrillChartName,"NC Drill Chart Span") Then
    '         Set oGOFAddEach = oGerberOutputFiles.Add(uNCDrillChartName + ".gdo")
    '         oGOFAddEach.HeaderText = "Gerber Output File " + uNCDrillChartName + ".gdo"
    '         oGOFAddEach.TrailerText = "Gerber Output File " + uNCDrillChartName + ".gdo"
    '         ' Add all board items individually 
    '         Set oBoardItems = oGOFAddEach.BoardItems            
    '         oBoardItems.Add eengBoardItemBoardCavities               
    '         oBoardItems.Add eengBoardItemBoardOutline               
    '         oBoardItems.Add eengBoardItemContours 
    '         'oBoardItems.Add eengBoardItemDrillDrawingThru
    '         ' Add all cell types individually 
    '         Call AddCellTypes(oGOFAddEach)
    '         ' Add all user layer NC Drill Chart individually
    '         Dim oNCDrillChart
    '         Set oNCDrillChart = oGOFAddEach.UserLayers
    '         oNCDrillChart.Add uNCDrillChartName
    '     End If 
    ' Next
End Sub

Sub AddBoardItems(oGOFAddEach) 
    Dim oBoardItems 
    Set oBoardItems = oGOFAddEach.BoardItems 
    ' Add all items using "All" constants
    ' oBoardItems.Add eengBoardItemAll

    ' Add all board items individually 
    ' Call oBoardItems.Add (eengBoardItemBoardObstruct) 
    ' oBoardItems.Add eengBoardItemBendAreas
    ' oBoardItems.Add eengBoardItemBoardCavities
    ' oBoardItems.Add eengBoardItemFlexLayers
    ' oBoardItems.Add eengBoardItemFlexPadLayers
    ' oBoardItems.Add eengBoardItemBoardObstruct
    ' oBoardItems.Add eengBoardItemBoardOrigin
    oBoardItems.Add eengBoardItemBoardOutline               
    ' oBoardItems.Add eengBoardItemContours                   
    ' oBoardItems.Add eengBoardItemDetailViews                
    ' oBoardItems.Add eengBoardItemDocumentation              
    ' oBoardItems.Add eengBoardItemDrcWindow                  
    ' oBoardItems.Add eengBoardItemDrillDrawingThru           
    ' oBoardItems.Add eengBoardItemGeneratedSilkscreenBottom  
    ' oBoardItems.Add eengBoardItemGeneratedSilkscreenTop     
    ' oBoardItems.Add eengBoardItemManufacturingOutline       
    ' oBoardItems.Add eengBoardItemMountingHoles              
    ' oBoardItems.Add eengBoardItemNCDrillOrigin              
    ' oBoardItems.Add eengBoardItemPanelBorder                
    ' oBoardItems.Add eengBoardItemPanelHole                  
    ' oBoardItems.Add eengBoardItemPanelOrigin                
    ' oBoardItems.Add eengBoardItemPanelOutline               
    ' oBoardItems.Add eengBoardItemPartHoles                  
    ' oBoardItems.Add eengBoardItemPlacementObstructBottom    
    ' oBoardItems.Add eengBoardItemPlacementObstructTop       
    ' oBoardItems.Add eengBoardItemRoomBottom                 
    ' oBoardItems.Add eengBoardItemRoomTop                    
    ' oBoardItems.Add eengBoardItemRouteBorder                
    ' oBoardItems.Add eengBoardItemScoringLine                
    ' oBoardItems.Add eengBoardItemShearingHoles              
    ' oBoardItems.Add eengBoardItemShearingLine               
    ' oBoardItems.Add eengBoardItemSoldermaskBottom           
    ' oBoardItems.Add eengBoardItemSoldermaskTop              
    ' oBoardItems.Add eengBoardItemSolderpasteBottom          
    ' oBoardItems.Add eengBoardItemSolderpasteTop
    ' oBoardItems.Add eengBoardItemTestFixtureOutline             
    ' oBoardItems.Add eengBoardItemToolingHoles
End Sub

Sub AddCellTypes(oGOFAddEach)
    Dim oCellTypes 
    Set oCellTypes = oGOFAddEach.CellTypes 
    ' Add all items using "All" constants
    ' oCellTypes.Add eengCellTypeAll

    ' Add all cell types individually 
    ' oCellTypes.Add eengCellTypeBadBoardIdentifier       
    ' oCellTypes.Add eengCellTypeBreakAwayTab             
    oCellTypes.Add eengCellTypeBuried                   
    oCellTypes.Add eengCellTypeConnector                
    oCellTypes.Add eengCellTypeDiscreteAxial            
    oCellTypes.Add eengCellTypeDiscreteChip             
    oCellTypes.Add eengCellTypeDiscreteOther            
    oCellTypes.Add eengCellTypeDiscreteRadial           
    oCellTypes.Add eengCellTypeEdgeConnector     
    ' oCellTypes.Add eengCellTypeElectricalTestIdentifier 
    oCellTypes.Add eengCellTypeEmbeddedPassive                  
    oCellTypes.Add eengCellTypeGeneral  
    ' oCellTypes.Add eengCellTypeGeneralPackageCell                  
    ' oCellTypes.Add eengCellTypeGeneralPanelCell         
    oCellTypes.Add eengCellTypeGraphic                  
    oCellTypes.Add eengCellTypeICBareDie                
    oCellTypes.Add eengCellTypeICBGA                    
    oCellTypes.Add eengCellTypeICDIP                    
    oCellTypes.Add eengCellTypeICFlipChip               
    oCellTypes.Add eengCellTypeICLCC                    
    oCellTypes.Add eengCellTypeICOther                  
    oCellTypes.Add eengCellTypeICPGA                    
    oCellTypes.Add eengCellTypeICPLCC                   
    oCellTypes.Add eengCellTypeICSIP                    
    oCellTypes.Add eengCellTypeICSOIC                   
    oCellTypes.Add eengCellTypeJumper                   
    oCellTypes.Add eengCellTypeMechanical               
    ' oCellTypes.Add eengCellTypePanelIdentifier          
    ' oCellTypes.Add eengCellTypePanelStiffener           
    ' oCellTypes.Add eengCellTypeRegistrationPin          
    ' oCellTypes.Add eengCellTypeRegistrationPinGrid      
    ' oCellTypes.Add eengCellTypeSkyhook                  
    ' oCellTypes.Add eengCellTypeSolderPalette            
    ' oCellTypes.Add eengCellTypeTestCoupon               
    oCellTypes.Add eengCellTypeTestPoint                              
End Sub

Sub AddCellItems(oGOFAddEach)
    ' oGOFAddEach.CellItemsSide = eengCellSideTop 
    Dim oCellItems 
    Set oCellItems = oGOFAddEach.CellItems 
    ' Add all items using "All" constants
    ' oCellItems.Add eengCellItemAll

    ' Add all cell items individually 
    oCellItems.Add eengCellItemAssemblyOutline               
    oCellItems.Add eengCellItemAssemblyPartNumber            
    oCellItems.Add eengCellItemAssemblyReferenceDesignator  
    oCellItems.Add eengCellItemBondWires   
    ' oCellItems.Add eengCellItemCapacitorDielectric   
    ' oCellItems.Add eengCellItemCapacitorPlate1
    ' oCellItems.Add eengCellItemCapacitorPlate2   
    ' oCellItems.Add eengCellItemCellOrigin                    
    ' oCellItems.Add eengCellItemGlueSpots     
    oCellItems.Add eengCellItemDiePins   
    oCellItems.Add eengCellItemInsertionOutline    
    ' oCellItems.Add eengCellItemMezzanineCapacitorPlate1            
    ' oCellItems.Add eengCellItemMezzanineCapacitorPlate2            
    ' oCellItems.Add eengCellItemMezzanineViaPad            
    oCellItems.Add eengCellItemPlacementOutline   
    ' oCellItems.Add eengCellItemResistorMask              
    ' oCellItems.Add eengCellItemResistorOverglaze              
    oCellItems.Add eengCellItemSilkScreenOutline             
    oCellItems.Add eengCellItemSilkScreenPartNumber          
    oCellItems.Add eengCellItemSilkScreenReferenceDesignator 
    ' oCellItems.Add eengCellItemTestPointObstruct
    ' oCellItems.Add eengCellItemTPAssemblyReferenceDesignator
    ' oCellItems.Add eengCellItemTPSilkscreenReferenceDesignator
End Sub

Sub AddConductiveItems(oGOFAddEach)  
    Dim oConductiveItems 
    Set oConductiveItems = oGOFAddEach.ConductiveItems 
    ' Add all items using "All" constants
    ' oConductiveItems.Add eengCondItemAll 

    ' Add all conductive items indivisually
    oConductiveItems.Add eengCondItemActualPlaneShapes
    oConductiveItems.Add eengCondItemBondPads
    oConductiveItems.Add eengCondItemCopperBalancing
    oConductiveItems.Add eengCondItemEmbeddedCapacitorPads
    oConductiveItems.Add eengCondItemEmbeddedResistorPads  
    oConductiveItems.Add eengCondItemEtchedText        
    oConductiveItems.Add eengCondItemFiducials
    oConductiveItems.Add eengCondItemMountingHoles
    oConductiveItems.Add eengCondItemPartHoles    
    oConductiveItems.Add eengCondItemPartPadsSMD       
    oConductiveItems.Add eengCondItemPartPadsTest      
    oConductiveItems.Add eengCondItemPartPadsThru      
    ' oConductiveItems.Add eengCondItemPinNumbers        
    oConductiveItems.Add eengCondItemPlaneData         
    ' oConductiveItems.Add eengCondItemPlaneNoConnect    
    ' oConductiveItems.Add eengCondItemPlaneObstruct     
    ' oConductiveItems.Add eengCondItemPlaneShape        
    oConductiveItems.Add eengCondItemResistorAreas     
    ' oConductiveItems.Add eengCondItemRouteObstruct     
    ' oConductiveItems.Add eengCondItemRuleArea          
    oConductiveItems.Add eengCondItemTraces            
    oConductiveItems.Add eengCondItemViaHoles          
    oConductiveItems.Add eengCondItemViaPads 
End Sub

Sub AddUserDefinedLayers(oGOFAddEach)
    Dim oUserDefinedLayers
    Set oUserDefinedLayers = oGOFAddEach.UserLayers
    oUserDefinedLayers.Add "NC Drill Chart Span 1-2"
    '...
End Sub

' Write Gerber plot setup file，this is the .gpf file that the Gerber engine uses to process output data and to load the Gerber Output dialog box. 
Sub WriteGerberSetupFile(oGerberEngine)
    ' Create a FileSystemObject
    Dim fileSysObj, setupFile
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    Set setupFile = fileSysObj.CreateTextFile(".\Config\UserGerberPlotSetup.gpf", True)

    ' For chinese OS
    Dim weekdayEngArr, monthEngArr
    weekdayEngArr = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    monthEngArr = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

    Dim gMacroStatus, gOutputFileColl, gOutputFile
    Set gOutputFileColl = oGerberEngine.OutputFiles

    setupFile.WriteLine (".FILETYPE GerberPlotSetupFile")
    setupFile.WriteLine (".VERSION ""VB99.0""")
    setupFile.WriteLine (".CREATOR ""Expedition PCB Gerber Output""")
    setupFile.WriteLine (".DATE " & """" & weekdayEngArr(WeekDay(Date()) - 1) & " " & monthEngArr(Month(Date()) - 1) & " " & Day(Date()) & " " & Year(Date()) & " " & Time() & """")
    ' setupFile.WriteLine (".DATE ""Tuesday, September 26, 2023 02:11 PM""")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".BaseUnits ""1NM""")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".GerberOutputDir "".\\Output\\Gerber\\""")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".DataOffset " & dx & " " & dy)
    setupFile.WriteLine (".DataCopies " & xCopies & " " & yCopies)
    setupFile.WriteLine (" ..OriginOffset " & dxCopiesX & " " & dyCopiesY)
    setupFile.WriteLine ("")
    If gMacro = FALSE Then
        gMacroStatus = "No"
    Else 
        gMacroStatus = "Yes"
    End If
    setupFile.WriteLine (".GenerateMacro " & gMacroStatus)
    setupFile.WriteLine ("")
    setupFile.WriteLine (".VariantsOutput No")
    setupFile.WriteLine (".DCodeToApertureFile ""Sys: GerberD-Codes.dac""")
    setupFile.WriteLine (".GerberMachineFormatFile ""Loc: " & gbrMatchineFormatFileName & """")
    setupFile.WriteLine ("")
    ' write infomation
    Dim gFileNameStr, gFileName, FP, conductiveRe, cLayer
    Set conductiveRe = new RegExp
    conductiveRe.Pattern = "EtchLayer*"
    For Each gOutputFile In gOutputFileColl
        gFileNameStr = Split(gOutputFile.Name, "Gerber\")
        gFileName = gFileNameStr(1)
        setupFile.WriteLine (".GerberOutputFile """ & gFileName & """")
        setupFile.WriteLine (" ..ProcessFile Yes")
        If FlashPadsFlag = TRUE Then
            FP = "Yes"
        Else
            FP = "No"
        End If
        setupFile.WriteLine (" ..FlashPads " & FP)
        setupFile.WriteLine (" ..GerberOutputPath "".\\Output\\Gerber\\" & gFileName & """")
        setupFile.WriteLine (" ..HeaderText")
        setupFile.WriteLine ("   ...CommentLine ""Mentor Graphics Example Gerber Output Definition""")
        ' Conductive items
        If conductiveRe.Test(gFileName) Then
            cLayer = Split(Split(gFileName, ".gdo")(0), "EtchLayer")(1)
            setupFile.WriteLine (" ..ConductiveLayer " & cLayer)
            setupFile.WriteLine ("   ...ConductiveItem BondPads")
            setupFile.WriteLine ("   ...ConductiveItem EtchedText")
            setupFile.WriteLine ("   ...ConductiveItem Fiducials")
            setupFile.WriteLine ("   ...ConductiveItem MountingHolesByLayer")
            setupFile.WriteLine ("   ...ConductiveItem PartHolesByLayer")
            setupFile.WriteLine ("   ...ConductiveItem PartPadsSMD")
            If cLayer = 1 Or cLayer = CStr(docObj.LayerCount) Then
                setupFile.WriteLine ("   ...ConductiveItem PartPadsTest")
            End If
            setupFile.WriteLine ("   ...ConductiveItem PartPadsThru")
            setupFile.WriteLine ("   ...ConductiveItem PlaneData")
            setupFile.WriteLine ("   ...ConductiveItem ResistorAreas")
            setupFile.WriteLine ("   ...ConductiveItem Traces")
            setupFile.WriteLine ("   ...ConductiveItem ViaHoles")
            setupFile.WriteLine ("   ...ConductiveItem ViaPads")
            setupFile.WriteLine ("   ...ConductiveItem CopperBalancing")
            setupFile.WriteLine ("   ...ConductiveItem EmbeddedCapPads")
            setupFile.WriteLine ("   ...ConductiveItem EmbeddedResPads")
            setupFile.WriteLine ("   ...ProcessUnconnectedPads Yes")
        End If
        ' Board items
        setupFile.WriteLine (" ..BoardItem BoardCavities")
        setupFile.WriteLine (" ..BoardItem BoardOutline")
        setupFile.WriteLine (" ..BoardItem Contours")
        setupFile.WriteLine (" ..ContourLayerList """"")
        setupFile.WriteLine ("   ...ContourLayerName ""Through Board""")
        Select Case gFileName
            Case "SoldermaskTop.gdo"
                setupFile.WriteLine (" ..BoardItem SoldermaskTop")
            Case "SoldermaskBottom.gdo"
                setupFile.WriteLine (" ..BoardItem SoldermaskBottom")
            Case "SolderPasteTop.gdo"
                setupFile.WriteLine (" ..BoardItem SolderpasteTop")
            Case "SolderPasteBottom.gdo"
                setupFile.WriteLine (" ..BoardItem SolderPasteBottom")
            Case "SilkscreenTop.gdo"
                setupFile.WriteLine (" ..BoardItem AlteredSilkscreenTop")
            Case "SilkscreenBottom.gdo"
                setupFile.WriteLine (" ..BoardItem AlteredSilkscreenBottom")
            Case "DrillDrawingThrough.gdo"
                setupFile.WriteLine (" ..BoardItem DrillDrawingThru")
        End Select
        ' Cell types
        setupFile.WriteLine (" ..CellType Buried")
        setupFile.WriteLine (" ..CellType Connector")
        setupFile.WriteLine (" ..CellType DiscreteAxial")
        setupFile.WriteLine (" ..CellType DiscreteChip")
        setupFile.WriteLine (" ..CellType DiscreteOther")
        setupFile.WriteLine (" ..CellType DiscreteRadial")
        setupFile.WriteLine (" ..CellType EdgeConnector")
        setupFile.WriteLine (" ..CellType EmbeddedPassive")
        setupFile.WriteLine (" ..CellType General")
        setupFile.WriteLine (" ..CellType Graphic")
        setupFile.WriteLine (" ..CellType ICBareDie")
        setupFile.WriteLine (" ..CellType ICBGA")
        setupFile.WriteLine (" ..CellType ICDIP")
        setupFile.WriteLine (" ..CellType ICFlipChip")
        setupFile.WriteLine (" ..CellType ICLCC")
        setupFile.WriteLine (" ..CellType ICOther")
        setupFile.WriteLine (" ..CellType ICPGA")
        setupFile.WriteLine (" ..CellType ICPLCC")
        setupFile.WriteLine (" ..CellType ICSIP")
        setupFile.WriteLine (" ..CellType ICSOIC")
        setupFile.WriteLine (" ..CellType Jumper")
        setupFile.WriteLine (" ..CellType Mechanical")
        setupFile.WriteLine (" ..CellType TestPoint")
        setupFile.WriteLine ("")
    Next
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
    ' Ask Xpedition Layout’s document for the key 
    key = docObj.Validate(0) 
    ' Get token from license server 
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application") 
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