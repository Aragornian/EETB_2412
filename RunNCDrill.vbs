Option Explicit 
 
Dim args, app, docObj, jobName, jobConfigPath, masterPath
Set args = ScriptHelper.Arguments 
If args.Count = 3 Then 
    ' Call from command line 
    ' mgcscript RunNCDrill.vbs [job name] 
    jobName = args.item(3) 
Else 
    ' Call from Xpedition Layout 
    ' Set app = GetObject(, "MGCPCB.ExpeditionPCBApplication") 
    Set app = Application
    Set docObj = GetLicensedDoc(app)
    jobName = docObj.FullName 
    masterPath = docObj.MasterPath
End If 

Scripting.AddTypeLibrary("MGCPCBEngines.NCDrill") 
Scripting.AddTypeLibrary("Scripting.FileSystemObject")

Dim ncDrillOutputDirectory
ncDrillOutputDirectory = masterPath + "Output\NCDrill\"

' Drill parameters
' Dim drillSweepAxis,drillBandWidth, drillPDHLargerThan, mfFileName
Const drillSweepAxis = eengSweepaxisHorizontal
Const drillBandWidth = 0.1
Const drillPDHLargerThan = 1
Const mfFileName = "UserDrillMachineFormat.dff"
' Chart formats
' Dim aSOSC, aDC, dSOSL, fFontName, fName, fSize, lSpace, pWidth, pLD, pTD, pTolerance, nTolerance, hLS, vLS, inContours, sNotes, textFTC, chartTitle, chartUnit
Const aSOSC = TRUE
Const aDC = TRUE
Const dSOSL = FALSE
Const fFontName = "MentorGDT"
Const fName = "VeriBest Gerber 0" 
Const fSize = 2.54
Const lSpace = 1.27
Const pWidth = 0.127
Const pLD = 1
Const pTD = 4
Const pTolerance = 0.01
Const nTolerance = 0.01
Const hLS = TRUE
Const vLS = TRUE
Const inContours = TRUE
Const sNotes = "notes"
Const textFTC = TRUE
Const chartTitle = "title"
Const chartUnit = eengChartUnitsMM

If app.LockServer = True Then  
    app.Gui.CursorBusy(True)
    
    ' Save file
    app.Gui.SuppressTrivialDialogs = True
    app.Gui.ProcessCommand("File->Save")
    app.Gui.SuppressTrivialDialogs = False
    ' Run NCDrill engine
    ' -----------------------
    ' Call RunNCDrill() 
    ' -----------------------

    ' Run NCDrill by NC Drill Generation dialog
    ' -----------------------
    Call RunNCDrillByDialog()
    ' -----------------------

    app.Gui.CursorBusy(False)
    app.UnlockServer
End If
 
' MsgBox "fini", VBInformation, "RunNCDrill.vbs" 

'************************************************************************ 
Sub RunNCDrill() 
    'On Error Resume Next 
 
    ' Create NCDrill Engine object 
    Dim oNCDrillEngine 
    Set oNCDrillEngine = CreateObject("MGCPCBEngines.NCDrill") 
 
    ' Set the design file name 
    oNCDrillEngine.DesignFileName = jobName 
    oNCDrillEngine.OutputDirectory = ncDrillOutputDirectory

    Call SetupMachineFormat(oNCDrillEngine) 
    Call SetupParameters(oNCDrillEngine) 
    Call SetupChart(oNCDrillEngine)

    ' Run the NCDrill Engine
    On Error Resume Next
    oNCDrillEngine.Go
    On Error Goto 0
     
    ' Check errors
    Dim oError, oErrors
    Set oErrors = oNCDrillEngine.Errors
    For Each oError In oErrors
    	Write oError.ErrorString
    Next
    If oErrors.Count = 0 Then
        Write("Export NCDrill Data to Folder " & ncDrillOutputDirectory)
        Write("Export NCDrill Successfully! " & Now())
    End If
End Sub 

Sub RunNCDrillByDialog()
    'On Error Resume Next 
    Dim ncDrillDialogBtnObj   

    ' Create NCDrill Engine object 
    Dim oNCDrillEngine 
    Set oNCDrillEngine = CreateObject("MGCPCBEngines.NCDrill") 

    ' Set the design file name 
    oNCDrillEngine.DesignFileName = jobName 
    oNCDrillEngine.OutputDirectory = ncDrillOutputDirectory

    Call SetupMachineFormat(oNCDrillEngine) 
    Call SetupParameters(oNCDrillEngine) 
    Call SetupChart(oNCDrillEngine)

    ' Overwrite configuration files first
    Call WriteSetupFile(oNCDrillEngine, 0)
    Call WriteSetupFile(oNCDrillEngine, 1)

    app.Gui.ProcessCommand("Output->NC Drill")
    Set ncDrillDialogBtnObj = app.Gui.FindDialog("NC Drill Generation").FindButton("OK")
    ncDrillDialogBtnObj.Click
    ' On Error Goto 0

End Sub

Sub Write(sMsg)
	If Not IsObject(app) Then
		Echo sMsg
	ElseIf Not app.Addins("Message Window") Is Nothing Then
		app.Addins("Message Window").Control.AddTab("RunNCDrill").AppendText sMsg & vbCrLf
        app.Addins("Message Window").Control.ActivateTab("RunNCDrill")
	End If
End Sub
 
Sub SetupMachineFormat(oNCDrillEngine) 
    'On Error Resume Next 
 
    Dim oNCDrillMF 
    Set oNCDrillMF = oNCDrillEngine.MachineFormat 
     
    Dim rtn 
    ' rtn = MsgBox("English or Metric" & vbCrLf & "(Yes for English, No for Metric)?", VBYesNo + VBQuestion + vbDefaultButton2, "Select Drill Format") 
    ' Use metric unit for drill format
    rtn = VBNo
    If rtn = VBYes Then  
        ' Match DrillEnglish.dff 
        oNCDrillMF.DataType                 = eengExcellon 
        oNCDrillMF.Unit                     = eengUnitInch 
        oNCDrillMF.DataFormatLeadingDigits  = 2 
        oNCDrillMF.DataFormatTrailingDigits = 4 
        oNCDrillMF.StepMode                 = eengStepAbsolute 
        oNCDrillMF.ZeroTruncation           = eengZeroTruncationTrailing 
        oNCDrillMF.DataMode                 = TRUE 
        oNCDrillMF.ArcStyle                 = eengArcStyleRadius 
        oNCDrillMF.SequenceNumbering        = FALSE 
        oNCDrillMF.CharacterSet             = eengCharacterSetASCII 
        oNCDrillMF.Delimiter                = "" 
        oNCDrillMF.Comments                 = TRUE 
        oNCDrillMF.CommentStr               = "; " 
        oNCDrillMF.RecordLength             = 0  
    Else     
        ' Match DrillMetric.dff 
        oNCDrillMF.DataType                 = eengExcellon 
        oNCDrillMF.Unit                     = eengUnitMM 
        oNCDrillMF.DataFormatLeadingDigits  = 3 
        oNCDrillMF.DataFormatTrailingDigits = 3 
        oNCDrillMF.StepMode                 = eengStepAbsolute 
        oNCDrillMF.ZeroTruncation           = eengZeroTruncationLeading 
        oNCDrillMF.DataMode                 = TRUE 
        oNCDrillMF.ArcStyle                 = eengArcStyleRadius 
        oNCDrillMF.SequenceNumbering        = FALSE 
        oNCDrillMF.CharacterSet             = eengCharacterSetASCII 
        oNCDrillMF.Delimiter                = "" 
        oNCDrillMF.Comments                 = TRUE 
        oNCDrillMF.CommentStr               = "; " 
        oNCDrillMF.RecordLength             = 0 
    End If 
     
    ' Optional, where to write the format file 
    oNCDrillMF.FileName = "./Config/" + mfFileName
End Sub 
 
Sub SetupParameters(oNCDrillEngine) 
    'On Error Resume Next 
    Dim oNCDrillParameters 
    Set oNCDrillParameters = oNCDrillEngine.Parameters 
     
    ' Same as the options in "Drill Options"  
    ' tab of the "NC Drill Generation" dialog 
    oNCDrillParameters.SweepAxis = eengSweepaxisHorizontal 
    oNCDrillParameters.Bandwidth(eengUnitMM) = drillBandWidth 
    ' oNCDrillParameters.PreDrillHolesLargerThan(eengUnitMM) = drillPDHLargerThan 
    oNCDrillParameters.OutputFileExtension = ".ncd" 
	Call oNCDrillParameters.ClearFileHeader() 
	Call oNCDrillParameters.ClearFileNotes() 
    Call oNCDrillParameters.AddFileHeader("header") 
    Call oNCDrillParameters.AddFileNotes("notes") 
End Sub 
 
Sub SetupChart(oNCDrillEngine) 
    'On Error Resume Next 
    Dim oNCDrillChart, oNCDrillChartColumns, oNCDrillHoles
    Set oNCDrillChart = oNCDrillEngine.Chart 

    ' Set drill symbol for drill holes
    Set oNCDrillHoles = oNCDrillChart.DrillHoles
    Call SetupChartHoleSymbols(oNCDrillHoles)

    ' Same as the options in "Drill Chart Options" 
    ' tab of the "NC Drill Generation" dialog 
    oNCDrillChart.AllSpansOnSingleChart = aSOSC
    oNCDrillChart.AssignDrillCharacters = aDC
    oNCDrillChart.DrillSymbolsOnSeparateLayers = dSOSL
    ' oNCDrillChart.FCFFontName = fFontName
    oNCDrillChart.FontName = fName
    oNCDrillChart.FontSize(eengUnitMM) =  fSize 
    oNCDrillChart.LineSpacing(eengUnitMM) = lSpace 
    oNCDrillChart.PenWidth(eengUnitMM) =  pWidth 
    oNCDrillChart.PrecisionLeadingDigits = pLD 
    oNCDrillChart.PrecisionTrailingDigits = pTD
    oNCDrillChart.PositiveTolerance(eengUnitMM) =  pTolerance 
    oNCDrillChart.NegativeTolerance(eengUnitMM) =  nTolerance
    oNCDrillChart.HorizontalLineSeparator = hLS
    oNCDrillChart.VerticalLineSeparator = vLS
    oNCDrillChart.IncludeContours = inContours
    oNCDrillChart.SpecialNotes = sNotes
    oNCDrillChart.TextFormatTitleCase = textFTC
    oNCDrillChart.Title = chartTitle 
    oNCDrillChart.UnitEx = chartUnit 
    Call oNCDrillChart.ResetTolerance()
    
    ' Add drill chart columns
    Set oNCDrillChartColumns = oNCDrillChart.Columns 
    oNCDrillChartColumns.Add eengColumnSymbol 
    oNCDrillChartColumns.Add eengColumnDiameter 
    oNCDrillChartColumns.Add eengColumnTolerance 
    oNCDrillChartColumns.Add eengColumnPlated 
    oNCDrillChartColumns.Add eengColumnPunched 
    oNCDrillChartColumns.Add eengColumnHolename 
    oNCDrillChartColumns.Add eengColumnQuantity 
    oNCDrillChartColumns.Add eengColumnSpan
    ' oNCDrillChartColumns.Add eengColumnType
    ' oNCDrillChartColumns.Add eengColumnUser
    ' oNCDrillChartColumns.Add eengColumnDiameterRange
    Call SetupChartColumn(oNCDrillChartColumns)
End Sub 

Sub SetupChartColumn(oNCDrillChartColumns)
    Dim oNCDrillChartColumn
    For Each oNCDrillChartColumn In oNCDrillChartColumns
        oNCDrillChartColumn.Alignment = eengColumnAlignCenter
        ' oNCDrillChartColumn.DisplayName = oNCDrillChartColumn.Name
        ' MsgBox LCase(oNCDrillChartColumn.DisplayName)
    Next
End Sub

Sub SetupChartHoleSymbols(oNCDrillHoles)
    Dim oNCDrillHole
    For Each oNCDrillHole In oNCDrillHoles
        ' oNCDrillHole.CombineCharAndSize = FALSE
        ' oNCDrillHole.DrillSymbol = eengDrillSymbolCircle
        oNCDrillHole.DrillSymbolSize(eengUnitMM) = 0.5
        oNCDrillHole.DrillSymbolType = eengDrillSymbolTypeAutomatic
    Next
End Sub

' Write configuration file
' 0 ---     .hkp setup file
' 1 ---     .dcs scheme file
Sub WriteSetupFile(oNCDrillEngine, isSchemeFile)
    ' Create a FileSystemObject
    Dim fileSysObj 
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")

    Dim setupFile
    If isSchemeFile Then
        Set setupFile = fileSysObj.CreateTextFile(".\Config\UserDrill.dsf", True)
    Else
        Set setupFile = fileSysObj.CreateTextFile(".\Config\DrillPreferences.hkp", True)
    End If

    Dim unit, drillSA, textCase, columnStr, columnDisplayAsStr, columnAlignmentStr
    columnStr = ""
    columnDisplayAsStr = ""
    columnAlignmentStr = ""

    Dim oNCDrillHoles, oNCDrillHole, oNCDrillChartColumns, oNCDrillChartColumn, contourExist
    Set oNCDrillHoles = oNCDrillEngine.Chart.DrillHoles
    Set oNCDrillChartColumns = oNCDrillEngine.Chart.Columns
    ' contourExist = 0

    ' For chinese OS
    Dim weekdayEngArr, monthEngArr
    weekdayEngArr = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    monthEngArr = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

    setupFile.WriteLine (".FILETYPE NC_DRILL_PREFERENCES")
    setupFile.WriteLine (".VERSION ""2.0""")
    setupFile.WriteLine (".CREATOR ""Mentor Graphics Corporation""")
    ' setupFile.WriteLine (".DATE ""Wednesday October 30 14:15:59 2024""")
    setupFile.WriteLine (".DATE " & """" & weekdayEngArr(WeekDay(Date()) - 1) & " " & monthEngArr(Month(Date()) - 1) & " " & Day(Date()) & " " & Time() & " " & Year(Date()) & """")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".Units mm")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".DrillOutputDir "".\\Output\\NCDrill\\""")
    ' setupFile.WriteLine (".DrillOutputDir " & """" & ncDrillOutputDirectory & """")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".SelectMillingLayer NO")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".DataOffset (0, 0)")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".StepAndRepeat")
    setupFile.WriteLine ("  ..DataCopies   (1, 1)")
    setupFile.WriteLine ("  ..OriginOffset (0, 0)")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".Optimize")
    Select Case drillSweepAxis
        Case eengSweepaxisHorizontal
            drillSA = "HORIZONTAL"
        Case eengSweepaxisVertical
            drillSA = "VERTICAL"
    End Select
    setupFile.WriteLine ("  ..SweepAxis " & drillSA)
    setupFile.WriteLine ("  ..BandWidth " & drillBandWidth)
    setupFile.WriteLine ("  ..SortByBoard NO")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".PreDrill OFF")
    ' setupFile.WriteLine ("  ..HoleSize " & drillPDHLargerThan)
    setupFile.WriteLine ("  ..HoleSize 0")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".OutputFiles")
    setupFile.WriteLine ("  ..MachineFormat  "".\\Config\\" & mfFileName & """")
    setupFile.WriteLine ("  ..MachineFileExt "".ncd""")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".DrillSymbols")
    setupFile.WriteLine ("  ..AutoAssign PREFER_CHARACTERS")
    setupFile.WriteLine ("")
    setupFile.WriteLine (".DrillCharts")
    Select Case chartUnit
        Case eengChartUnitsMM
            unit = "MM"
        Case eengChartUnitsUM
            unit = "UM"
        Case eengChartUnitsInch
            unit = "IN"
        Case eengChartUnitsMils
            unit = "TH"
    End Select
    setupFile.WriteLine ("  ..ChartUnits " & unit)
    Dim alignmentStr
    For Each oNCDrillChartColumn In oNCDrillChartColumns
        columnStr = columnStr + oNCDrillChartColumn.Name + " "
        columnDisplayAsStr = columnDisplayAsStr + """" + oNCDrillChartColumn.Name + """" + " "
        Select Case oNCDrillChartColumn.Alignment
            Case eengColumnAlignCenter 
                alignmentStr = "CENTER"
            Case eengColumnAlignLeft
                alignmentStr = "LEFT"
            Case eengColumnAlignRight
                alignmentStr = "RIGHT"
        End Select
        columnAlignmentStr = columnAlignmentStr + alignmentStr + " "
    Next
    setupFile.WriteLine ("  ..Columns " & columnStr)
    setupFile.WriteLine ("  ..DisplayAs " & columnDisplayAsStr)
    setupFile.WriteLine ("  ..Alignment " & columnAlignmentStr)
    setupFile.WriteLine ("  ..FontName         " & """" & fName & """")
    setupFile.WriteLine ("  ..FontSize          " & fSize)
    setupFile.WriteLine ("  ..LineSpacing       " & lSpace)
    setupFile.WriteLine ("  ..PenWidth          " & pWidth)
    setupFile.WriteLine ("  ..PosTolerance      -1")
    setupFile.WriteLine ("  ..NegTolerance      1")
    setupFile.WriteLine ("  ..Precision         " & pLD & ", " & pTD)
    setupFile.WriteLine ("  ..SingleChart       " & ConvertFlagToInt(aSOSC))
    setupFile.WriteLine ("  ..SeparateLayers    " & ConvertFlagToInt(dSOSL))
    setupFile.WriteLine ("  ..HorizontalLSep    " & ConvertFlagToInt(hLS))
    setupFile.WriteLine ("  ..VerticalLSep      " & ConvertFlagToInt(vLS))
    If ConvertFlagToInt(textFTC) = 1 Then
        textCase = 0
    Else
        textCase = 1
    End If
    setupFile.WriteLine ("  ..UpperCaseText     " & textCase)
    setupFile.WriteLine ("  ..IncludeContours   " & ConvertFlagToInt(inContours))
    setupFile.WriteLine ("")
    setupFile.WriteLine ("")
    setupFile.WriteLine ("")
    setupFile.WriteLine ("  ..FCFFontName ""MentorGDT""")
    setupFile.WriteLine ("")
    setupFile.WriteLine ("  ..FCFData")
    For Each oNCDrillHole In oNCDrillHoles
        setupFile.WriteLine ("   ...HoleFCF " & """" & oNCDrillHole.Name & """")
        setupFile.WriteLine ("    ....valFCF """" """" """" """" """"")
        ' If oNCDrillHole.Name = "No Tool Contour" Then
        '     contourExist = 1
        ' End If
    Next
    If isSchemeFile = 0 Then
        setupFile.WriteLine (".SymbolScheme    ""UserDrill""")
        setupFile.WriteLine (".SchemeLocation    ""Loc""") 
    End If
End Sub

Function ConvertFlagToInt(Flag)
    If Flag = TRUE Then
        ConvertFlagToInt = 1
    Else 
        ConvertFlagToInt = 0
    End If 
End Function
 
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