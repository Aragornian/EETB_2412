'(Declarations)
' Created 12-17-2013
Option Explicit

Sub runCommand(oBtn)
	
	Dim k
	Dim tmpArr
	
	For Each k In dctFormButtons
	
		tmpArr = dctFormButtons(k)
		
		If k = oBtn.Text Then
		
			If tmpArr(2) = "SUBROUTINE" Then ' this button will run a subroutine
				
'				MsgOut "Run " & tmpArr(2) & " : " & tmpArr(3)
				Execute tmpArr(3)
				
			End If
			
			If tmpArr(2) = "PATH" Then ' this button will run another script
			
'				MsgOut "Run " & tmpArr(2) & " : " & tmpArr(3)
				Call ExecCommand(tmpArr(3))
				
			End If
			
			If tmpArr(2) = "KEYIN" Then ' this button will run a keyin
			
'				MsgOut "Run " & tmpArr(2) & " : " & tmpArr(3)
				Call ExecKeyin(tmpArr(3))
				
			End If
			
			If tmpArr(2) = "HTML" Then ' this button will run a keyin
			
'				MsgOut "Run " & tmpArr(2) & " : " & tmpArr(3)
				Call ExecHtml(tmpArr(3))
				
			End If

			Exit Sub
		
		End If
		
	Next	
	
End Sub


'========================================================================
' build main category tabstrip
'========================================================================
Sub loadFormBtns()

	Dim  i, sBtn
	
	For Each sBtn In dctFormButtons
	
		Dim Arr
		Arr = dctFormButtons(sBtn)

		Call displayFormButtons(formCnt, sBtn, Arr(1), Arr(0))
		
		formCnt = formCnt + 1
		
	Next
	
End Sub

' ===========================================
' Category button location, visibility, label etc
' ===========================================
Sub displayFormButtons(ButtonId, sLabel, sToolTip, bEnable)

	Dim btnHT: btnHT = 40
	Dim btnWD: btnWD = 80
	Dim btnLT: btnLT = 10
	Dim btnRT: btnRT = 90
	Select Case ButtonId
		Case 1
			Call ConfigureButtons(f1, 60, 100, btnLT, btnRT, sLabel, sToolTip, bEnable)
		Case 2
			Call ConfigureButtons(f2, f1.Top, f1.Bottom, btnLT+btnWD, btnRT+btnWD, sLabel, sToolTip, bEnable)
		Case 3
			Call ConfigureButtons(f3, f2.Top, f2.Bottom, f2.Left+btnWD, f2.Right+btnWD, sLabel, sToolTip, bEnable)
		Case 4
			Call ConfigureButtons(f4, f1.Top+btnHT, f1.Bottom+btnHT, btnLT, btnRT, sLabel, sToolTip, bEnable)
		Case 5
			Call ConfigureButtons(f5, f4.Top, f4.Bottom, f4.Left+btnWD, f4.Right+btnWD, sLabel, sToolTip, bEnable)
		Case 6
			Call ConfigureButtons(f6, f5.Top, f5.Bottom, f5.Left+btnWD, f5.Right+btnWD, sLabel, sToolTip, bEnable)

	End Select
	
End Sub

Function attributeWalk(node)
    Dim i, spcr, parentNode, tabID, sFormButton, cmdType, bEnable, sPath
    parentNode=""
    For i=1 To indent
        spcr = spcr + "-"
    Next
	Dim attrib, tmpStr, tmpKey, tmpValue
	cmdName=""
	cmdEn=0 
	cmdStr=""
	Dim tCntr: tCntr = 1

	For Each attrib In node.attributes
		tmpStr = ""
		If attrib.value = "FlipChipViewer" Then	
			FlipChipMenuType = FlipChipViewer
		Elseif attrib.value = "FlipChip" Then	
			FlipChipMenuType = FlipChip
		End If
		If node.nodeName = "CATEGORY" And FlipChipMenuType = mainAppName Then ' get main command categories
			If attrib.name = "NAME" Then ' attribute name of each category (route, DRC Utilities etc)
				catButton = attrib.value
			End If
			If attrib.name = "ENABLE" Then ' is category going to be used?
				If attrib.value = 1 Then
					' this category is required to add to main tabstrip
					dctCats.Add cntr, catButton
					cntr=cntr+1
				End If
			End If
		End If
		
		If node.parentNode.nodeName = "FORM" And FlipChipMenuType = mainAppName Then	
			
			parentNode = node.parentNode.nodeName	
				
			If attrib.name = "NAME" Then 
			
				cmdName = attrib.value
				
				End If
				
			If attrib.name = "ENABLE" Then 
			
				cmdEnable = attrib.value
				
			End If
			
			If attrib.name = "SUBROUTINE" And cmdEnable = 1 Then 
			
				tmpStr = cmdEnable & "," & attrib.value
				cmdType = attrib.name
				
			Elseif attrib.name = "PATH" And cmdEnable = 1 Then
			
				tmpStr = cmdEnable & "," & attrib.value
				cmdType = attrib.name
				
			Elseif attrib.name = "KEYIN" And cmdEnable = 1 Then
			
				tmpStr = cmdEnable & "," & attrib.value	
				cmdType = attrib.name
				
			End If
			
		End If
'-------------------------------------------------------------------------------------------------------
		If node.nodeName = "FORMCOMMAND" And FlipChipMenuType = mainAppName Then ' get main command categories
		
			If attrib.name = "NAME" Then ' attribute name of each category (route, DRC Utilities etc)
			
				sFormButton = attrib.value
								
			End If
			
			If attrib.name = "ENABLE" Then ' is category going to be used?
			
'				If attrib.value = 1 Then
				
					bEnable = attrib.value
										
'				End If
				
			End If
			
			If attrib.name = "TOOLTIP" Then
			
				cmdToolTip = attrib.value
				
			End If
			
			If attrib.name = "SUBROUTINE" Or attrib.name = "PATH" Or attrib.name = "KEYIN" Or attrib.name = "HTML" Then 
			
				cmdType = attrib.name
				sPath = attrib.value
			
			Call AddToDict(dctFormButtons, sFormButton, Array(bEnable, cmdToolTip, cmdType, sPath))
				
			End If
			

		End If
'-------------------------------------------------------------------------------------------------------
		If node.parentNode.nodeName = "CATEGORY" And FlipChipMenuType = mainAppName Then		
			parentNode = node.parentNode.nodeName
			If attrib.name = "NAME" Then 
				cmdName = attrib.value
				End If
			If attrib.name = "ENABLE" Then 
				cmdEnable = attrib.value
			End If
			If attrib.name = "TOOLTIP" Then
				cmdToolTip = attrib.value
			End If
			If attrib.name = "SUBROUTINE" And cmdEnable = 1 And FlipChipMenuType = mainAppName Then 
				tmpKey = catButton & "," & cmdName
				tmpValue = "SUBROUTINE=" & attrib.value
				If cmdToolTip <> "" Then 
					tmpValue = tmpValue & "," & cmdToolTip
				End If
				dctCmds.Add tmpKey, tmpValue
				If attrib.value = "FindDiffPairPins" Then
					tCntr = tCntr + 1
				End If
				tmpKey =""
				tmpValue=""
			Elseif attrib.name = "PATH" And cmdEnable = 1 And FlipChipMenuType = mainAppName Then
				tmpKey = catButton & "," & cmdName
				tmpValue = "PATH=" & attrib.value
				If cmdToolTip <> "" Then 
					tmpValue = tmpValue & "," & cmdToolTip
				End If			
				dctCmds.Add tmpKey, tmpValue
				tmpKey =""
				tmpValue=""
			Elseif attrib.name = "KEYIN" And cmdEnable = 1 And FlipChipMenuType = mainAppName Then
				tmpKey = catButton & "," & cmdName
				tmpValue = "KEYIN=" & attrib.value
				If cmdToolTip <> "" Then 
					tmpValue = tmpValue & "," & cmdToolTip
				End If			
				dctCmds.Add tmpKey, tmpValue
				tmpKey =""
				tmpValue=""
			End If			
		End If
	Next
	
	cmdToolTip=""
	
End Function

Sub addToDict(dict, key, value)
	If dict.Exists(key) Then
		dict(key) = value
	Else
		dict.Add key, value
	End If
End Sub


Dim pcbApp, pcbDoc, Doc

Dim FoT
Dim root
Dim cntr
Dim fCntr
Dim xmlDoc, xmlFile, xmlFileFallback, xmlFileDeveloper
Dim child
Dim indent: indent=0
Dim aatk_env, wdir_env, AATK_TOOLBOX_CONFIG_FILE_env, aatk_flipchip_config_file_env
Dim cmdName, cmdEn, cmdSub, cmdPath, cmdEnable, cat, cmdStr, cmdType, cmdToolTip, catButton
Dim mButtonClick
Dim FlipChipMenuType:FlipChipMenuType = 0
Dim firstCmdTabClick:firstCmdTabClick=True
Dim mainAppName
Dim wscript
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim ActiveCat: ActiveCat=""
Dim firstClick: firstClick = 0
Dim clickCntr:clickCntr=1
Dim catCnt
Dim formCnt
Dim FirstCat: FirstCat = True
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const btnTop = 90, btnHt = 27
aatk_env = Scripting.GetEnvVariable("AATK")
AATK_TOOLBOX_CONFIG_FILE_env = Scripting.GetEnvVariable("AATK_TOOLBOX_CONFIG_FILE")
aatk_flipchip_config_file_env = Scripting.GetEnvVariable("AATK_FLIPCHIP_CONFIG_FILE")
Set pcbApp = Application
Dim cmdListener: Set cmdListener = pcbApp.Gui.CommandListener
Scripting.AttachEvents cmdListener, "cmd"
' Define project integration enums, dll is not registered.
Const eprjintStatusInSynch = 2
Const eprjintStatusRequired = 1
Const eprjintStatusNoCES = 3
Const FlipChipViewer = 2
Const FlipChip = 1
Dim numCols: numCols = 3
Dim dctCats: Set dctCats = CreateObject("Scripting.Dictionary")
Dim dctCmds: Set dctCmds = CreateObject("Scripting.Dictionary")
Dim dctCatButtons: Set dctCatButtons = CreateObject("Scripting.Dictionary")
Dim dctFormButtons: Set dctFormButtons = CreateObject("Scripting.Dictionary")
Dim sShowHideTP: sShowHideTP = "Click to refresh menus. Double+Click to show Lock/Unlock EFM buttons"

Dim slash
If Scripting.IsUnix Then
	slash = "/"
Else
	slash = "\"
End If


If Scripting.IsUnix Then
    
	xmlFileFallback = Scripting.ExpandEnvironmentStrings("%AATK%/FlipChip/FlipChipConfig.xml")
	xmlFileDeveloper = FindFileInWdir("FlipChipConfigDeveloper.xml")
Else
    
	xmlFileFallback = Scripting.ExpandEnvironmentStrings("%AATK%\FlipChip\FlipChipConfig.xml")
	xmlFileDeveloper = FindFileInWdir("FlipChipConfigDeveloper.xml")
End If


'-----------------------------------------
'Added 01/30/2015 by Ian J Gabbitas
'Application version checking
Dim appVer, appVerVec, eeVer  
appVer = Scripting.GetEnvVariable("SDD_VERSION")
appVerVec = split(appVer, ".")

If appVerVec(0) = "7" Then 
	eeVer = "EE"
	If pcbApp.Name = "Discovery PCB - Viewer" Then
		mainAppName = FlipChipViewer
	Else
		mainAppName = FlipChip
	End If
Else
	eeVer = "VX"
	If pcbApp.Name = "Xpedition xPCB Viewer" Or  pcbApp.Name = "Xpedition Layout Viewer" Then
		mainAppName = FlipChipViewer
	Elseif pcbApp.Name = "xPCB Viewer" Then
		mainAppName = FlipChipViewer
	Else
		mainAppName = FlipChip
	End If
End If

'-----------------------------------------

'Set xmlDoc = createXmlDomDocument(xmlDoc)
set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async = False
xmlDoc.validateOnParse=False

initForm

Sub initForm()

	If fso.FileExists(AATK_TOOLBOX_CONFIG_FILE_env) Then
	
		xmlFile = AATK_TOOLBOX_CONFIG_FILE_env ' if file exists here use it
		
	ElseIf fso.FileExists(aatk_flipchip_config_file_env) Then
	
		xmlFile = aatk_flipchip_config_file_env ' (old env var) if file exists then use it
		
	Elseif fso.FileExists(xmlFileFallback) Then
	
		xmlFile = xmlFileFallback ' use this if env var file not found
		
	Else
		msgbox "Error loading FlipChip form." & vbCrlf &_
		"Cannot find AATK FlipChip configuration file " &_
		"using environment variable AATK_TOOLBOX_CONFIG_FILE, or the fallback configuration file: " &_
		vbCrlf & vbCrlf &_
		xmlFileFallback
		Exit Sub
	End If

	'theview.Color = rgb(102,205,170)
	'Bitmap1.Bitmap = aatk_env & "/FlipChip/Icons/bg.bmp"
	dctCmds.RemoveAll
	dctCats.RemoveAll
	cntr=1
	'xmlFileDeveloper = aatk_env & "/FlipChip/FlipChipConfigDeveloper.xml"
	' load the standard XML configuration file
	xmlDoc.load xmlFile
	runXMLParser
	formCnt = 1
	
	loadFormBtns
	'SetFormButtons
	
	catCnt = 1
	loadCatBtns
	
	' Load the "developer" XML to access all the secret scripts
	xmlDoc.load xmlFileDeveloper
	runXMLParser
	
	catCnt = 1
	loadCatBtns
	
	loadCmdBtns
	'If catTab.Tabs.count < 2 Then
	'	catTab.Tabs.Remove(1)
	'End If
	'frmCommandTabs.Text = catTab.SelectedItem
End Sub


Sub loadXMLFile()
	Dim FlipChipMenu: FlipChipMenu = False
	Dim objFileSource, strCurrentLine
	Dim re: Set re = new regexp  'Create the RegExp object
	re.IgnoreCase = True
	'Dim regexpFormStart: regexpFormStart = "^\s*\<FORM NAME\=\" & Chr(34) & "FlipChip" & Chr(34) & ">$"
	Dim regexpFormStart: regexpFormStart = "<FORM NAME=""FlipChip"">"
	Dim regexpFormEnd: regexpFormEnd = "^\s*\</FORM\>$"
    If fso.fileExists(xmlFile) Then
		Set objFileSource = fso.OpenTextFile(xmlFile, ForReading, True, 0)
		Do While objFileSource.AtEndOfStream <> True
			Dim tmpStr, tmpArry, tmpArry2
			strCurrentLine = objFileSource.ReadLine   
			re.pattern = regexpFormStart
			If re.test(strCurrentLine) Then
				FlipChipMenu = True
				AppendOutput "AATK", "Form Start"
 			End If
			re.pattern = regexpFormEnd
			If re.test(strCurrentLine) Then
				FlipChipMenu = False
				AppendOutput "AATK", "Form End"
 			End If
		Loop
		objFileSource.Close
	End If
End Sub


'=========================================================================
' Event Handlers for Application
'=========================================================================
Function cmd_PreOnCommand(name, id)
	' COmmand Exit id=57665
	If id = 57665 Then
		TheView.Cancel ' quit this app
		'pcbApp.quit ' try to quit expedition app
	End If
End Function


'========================================================================
' build main category tabstrip
'========================================================================
Sub loadCatBtns()
	'hideCatBtns
	Dim  i
	For Each cat In dctCats
		If FirstCat = True Then
			ActiveCat = dctCats(cat)
			txt_Cat.Text = ActiveCat
			TidyCats
			c1.BackColor = rgb(0,0,130)
			c1.ForeColor = rgb(255,255,0)
			FirstCat = False
			loadCmdBtns
			'AppendOutput "AATK", "key= " & cat & ", value= " & dctCats(cat)
		End If
		Call displayCatButtons(catCnt, dctCats(cat))
		catCnt = catCnt + 1
	Next
End Sub



'========================================================================
' build command tabs for active category tab
'========================================================================
Sub loadCmdBtns() ' 
	TidyCmds
	Dim i, cmd, tmpArry, cnt, tmpArry2
	Dim row, mess
	' cycle thru all commands and add tabs for the category tab that
	' is currently active
	i = 1
	'On Error Resume Next
	For Each cmd In dctCmds	
		tmpArry = split(cmd, ",")
		tmpArry2 = split(dctCmds(cmd), ",")
			'AppendOutput "AATK", "key= " & cmd & ", value= " & dctCmds(cmd)

		If tmpArry(0) = ActiveCat Then ' active cat tab	
			If i > 20 Then
				editBox_statusbar "Error: Exceeded maximum number of commands allowed in FlipChip form for the " & ActiveCat & " tab."
				'Exit Sub
			Else
				If UBound(tmpArry2) > 0 Then
					Call displayCmdButton(i,tmpArry(1), tmpArry2(1))
				Else
					Call displayCmdButton(i,tmpArry(1),"")
				End If
			End If
			i = i + 1	
		End If
	Next
	txt_Cat.Bottom = 5 + txt_Cat.Top + btnHT*i
	editBox_statusbar.Top = 5 + txt_Cat.Top + btnHT*i
	editBox_statusbar.Bottom = 50 + txt_Cat.Top + btnHT*i
	TheFrame.Height = 75 + txt_Cat.Top + btnHT*i
	TheFrame.Width = 260
End Sub


Sub updateSelected(r,c)
	'catTab.ResetSelection
End Sub

Sub runXMLParser()

	If xmlDoc.parseError.errorcode = 0 Then
		'Walk from the root to each of its child nodes:
		treeWalk(xmlDoc)
	End If
End Sub

' get all the nodes and configure the main menu
Function treeWalk(node)
	Dim nodeName, cmdtype
	cmdName=""
	For Each child In node.childNodes
		If (child.nodeType=1) Then
			If (child.attributes.length>0) Then
				attributeWalk(child)
			End If
		End If
		If (child.hasChildNodes) Then
			treeWalk(child)
		End If
	Next
End Function


Function createXmlDomDocument(xd)
	On Error Resume Next
	Set xd = CreateObject("MSXML2.DOMDocument.6.0")
	If (IsObject(xd) = False) Then
		alert("DOM document not created. Check MSXML version used in createXmlDomDocument.")
	Else
		Set createXmlDomDocument = xd
	End If
End Function

Sub ExecCommand(cmd_script)
	Dim ccmd, ccmd_exe
	Dim aatk_env
	aatk_env = Scripting.GetEnvVariable("AATK")
	
	Set doc = application.ActiveDocument
	If ValidateServer(doc) = 0 Then
		msgbox "Server validation failed."
	End If	
	If application.LockServer Then
		If (Not doc Is Nothing) Then
			ccmd = aatk_env & slash & "FlipChip" & slash & cmd_script
			MsgOut "executing command: " & ccmd
			ccmd_exe = application.ProcessScript(ccmd,True)
		End If
		application.UnlockServer
	End If
End Sub

Sub ExecHtml(sHtmlPath)
	
	Dim helpFile : helpFile = aatk_env &  slash & "FlipChip" & slash & sHtmlPath
	Dim exec
	Set exec = CreateObject("ViewLogic.Exec")
	
	If Scripting.IsUnix Then
		Exec.Run "firefox " & helpFile 'Linux
	Else
		'Call exec.Run("iexplore.exe " & helpFile) 'Windows
		CreateObject("ViewLogic.Exec").Run("iexplore.exe " & helpFile)
	End If
	
	
End Sub

Sub ExecKeyin(cmd_script)
	Dim ccmd, ccmd_exe
	Dim aatk_env
	aatk_env = Scripting.GetEnvVariable("AATK")
	
	Set doc = application.ActiveDocument
	If ValidateServer(doc) = 0 Then
		msgbox "Server validation failed."
	End If	
	If application.LockServer Then
		If (Not doc Is Nothing) Then
			ccmd = "run " & aatk_env & "\FlipChip\" & cmd_script
			'msgbox ccmd
			ccmd_exe = application.Gui.ProcessKeyin(ccmd)
		End If
		application.UnlockServer
	End If
End Sub


Sub ExecCommandNoDoc(cmd_script)
	Dim ccmd, ccmd_exe
	
	If application.LockServer Then
		'	If (Not doc Is Nothing) Then
		ccmd = aatk_env & "\FlipChip\" & cmd_script
		ccmd_exe = application.ProcessScript(ccmd,True)
		'	End If
		application.UnlockServer
	End If
End Sub

' Function to validate document
Private Function ValidateServer(doc)
	dim key, licenseServer, licenseToken
	key = doc.Validate(0)
	Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")
	licenseToken = licenseServer.GetToken(key)
	Set licenseServer = nothing
	'On Error Resume Next
	Err.Clear
	doc.Validate(licenseToken)
	If Err Then
		ValidateServer = 0
	Else
		ValidateServer = 1
	End If
	Scripting.AttachEvents pcbApp, "pcbApp"
	Scripting.AttachEvents doc, "doc"
	'Get CommandListener Handler
End Function

' Update the Layout Driven Netlist sync button based on value of global variable
Sub setLDNButton()
	If pcbApp.IsXtremeClient = False Then
		' if the global scripting variable has a value
		If Scripting.Globals("LayoutDrivenNetlistStatus") <> "" Then
			If Scripting.Globals("LayoutDrivenNetlistStatus") = 1 Then
				ButtonSyncDesign.Text = "       Netlist             Out-of-Sync"
				ButtonSyncDesign.Bitmap = aatk_env & "/FlipChip/icons/org32x32.bmp"
			Else
				ButtonSyncDesign.Text = "Netlist             In-Sync"
				ButtonSyncDesign.Bitmap = aatk_env & "/FlipChip/icons/grn32x32.bmp"
			End If
			ButtonSyncDesign.Enable = True
		Else
			ButtonSyncDesign.Text = "       Netlist             Out-of-Sync"
			ButtonSyncDesign.Bitmap = aatk_env & "/FlipChip/icons/gry32x32.bmp"
			ButtonSyncDesign.Enable = False
		End If
	Else
		ButtonSyncDesign.Text = "       Netlist             Out-of-Sync"
		ButtonSyncDesign.Bitmap = aatk_env & "/FlipChip/icons/gry32x32.bmp"
		ButtonSyncDesign.Enable = False
	End If
	'msgbox "LayoutDrivenNetlistStatus=" & Scripting.Globals("LayoutDrivenNetlistStatus")
End Sub



Sub pcbApp_OnOpenDocument(Flags)
	Select Case Flags
		Case epcbOnOpenDocOpen
		Set pcbDoc = pcbApp.ActiveDocument
		
		Case epcbOnOpenDocReload
		' Skip the custom init & verify steps
		
	End Select
End Sub

'========================================================================
' Subroutines for command buttons. thes should really but moved to their
' corresponding external .vbs scripts
'========================================================================
Sub runConnectionPlanner()
	Gui.ProcessCommand("Smart &Utilities->Connection Planner...")
End Sub
Sub FindDiffPairPins()
	Dim usrPrompt
	usrPrompt = msgbox("Find and mark all diff-pair pins. OK to continue?",Vbokcancel)
	If usrPrompt = vbCancel Then 
    'cancel button was pressed
    Exit Sub
	End If
	Call ExecCommand("Route/Find Diff pairs.vbs")
End Sub
Sub DeleteDuplicateVias()
	Dim usrPrompt
	usrPrompt = msgbox("This command will find and remove duplicate vias. OK to continue?",Vbokcancel)
	If usrPrompt = vbCancel Then 
    'cancel button was pressed
    Exit Sub
	End If			
	Call ExecCommand("Route/DeleteDuplicateVias.vbs")
End Sub
Sub DRC_KillerVia()
	Dim usrPrompt
	usrPrompt = msgbox("Generate a report listing die pins that are connected by a single via." &  vbCrLf & "OK To continue?",Vbokcancel)
	If usrPrompt = vbCancel Then 
    'cancel button was pressed
    Exit Sub
	End If		
	Call ExecCommand("DRC/DRC_KillerVia.vbs")
End Sub
Sub LayerStackupWithMotionGfx()
	Dim usrPrompt
	usrPrompt = msgbox("Place layer stackup table in the design. OK to continue?",Vbokcancel)
	If usrPrompt = vbCancel Then 
    'cancel button was pressed
    Exit Sub
	End If
	Call ExecCommand("MFG/LayerStackupWithMotionGfx_um.vbs")
End Sub
Sub Add_Teardrop_User_Layers()
	Dim usrPrompt
	usrPrompt = msgbox("This command will copy teardrops to user layers (Teardrop 1-n). OK to continue?",Vbokcancel)
	If usrPrompt = vbCancel Then 
    'cancel button was pressed
    Exit Sub
	End If
	Call ExecCommand("MFG/Add_Teardrop_User_Layers.vbs")
End Sub
Sub ExcelCompList()
		Dim cmd_script
		Dim fileName
	
	    cmd_script = "Reports\ExcelCompList.xls"	
		fileName = aatk_env & "\FlipChip\" & cmd_script
	    Dim sh
		Set sh = CreateObject("Shell.Application")
		
		'msgbox filename
		sh.ShellExecute "excel.exe", fileName
		Set sh = Nothing
End Sub
Sub swap_vias
    'Scripting.AddTypeLibrary ("MGCPCBEngines.MaskEngine")
    'Scripting.AddTypeLibrary ("MGCPCB.Application")
    
    ' set the units for the mask engine so that we know the size of the incoming data
    'mskeng.CurrentUnit = emeUnitMils

    ' get expedition document
    'Set expApp = GetObject(, "MGCPCB.Application")
    Dim expApp
    Set expApp = Application
    Set pcbdoc = expApp.ActiveDocument

    If ValidateServer(pcbdoc) = 0 Then
		msgbox "Server validation failed."
	End If

    'docobj.CurrentUnit = epcbUnitMils
    Dim via1name, via2name, NetN
    Dim via 'As via
    Dim vias 'As vias
    Set vias = pcbdoc.vias(epcbSelectSelected)
    'pcbApp.Gui.ProcessCommand ("Partly Selected Nets")
    
    Dim pin1name, pin2name, NetName, PinObjCnt
    Dim pin 'As pin
    Dim pins 'As pins
    Set pins = pcbdoc.pins(epcbSelectSelected)
    'pcbApp.Gui.ProcessCommand ("Partly Selected Nets")
    Dim changeNetLib: Set changeNetLib = Scripting.Globals("ChangeNetLib")

    If pins.Count = 2 Then
		For Each pin In Pins
			PinObjCnt = pin.connectedObjects.count
			If PinObjCnt > 0 Then
				MsgBox "Pin is connected and cannot be swapped"
				Exit Sub
			End If
		Next    
	   
        pin1name = pins.Item(1).Net.Name
        pin2name = pins.Item(2).Net.Name
         
        Set NetName = pcbdoc.FindNet(pin2name)
        Call changeNetLib.SetPinNet(pins.Item(1), NetName, epcbDRCModeDRC)
        Set NetName = pcbdoc.FindNet(pin1name)
        Call changeNetLib.SetPinNet(pins.Item(2), NetName, epcbDRCModeDRC)
		Scripting.Globals("LayoutDrivenNetlistStatus") = 1
		setLDNButton
    Elseif vias.Count = 2 Then
		For Each via In vias
			PinObjCnt = via.connectedObjects.count
			If PinObjCnt > 0 Then
				MsgBox "Via is connected and cannot be swapped"
				Exit Sub
			End If
		Next    
   
        via1name = vias.Item(1).Net.Name
        via2name = vias.Item(2).Net.Name
         
        Set NetN = pcbdoc.FindNet(via2name)
        Set vias.Item(1).Net = NetN
        Set NetN = pcbdoc.FindNet(via1name)
        Set vias.Item(2).Net = NetN  
    Else
        MsgBox "select 2 vias or pins to swap them"
    End If
End Sub

' This library contains a single function for finding a file in the wdir
' The function will search all paths in WDIR for a specific file name and 
' return the first instnace of that file name it finds.
Function FindFileInWdir(fileName)
    FindFileInWdir = ""   
    Dim wdirVar, wdirPaths
    wdirVar = Scripting.GetEnvVariable("WDIR")
    If Scripting.IsUnix Then
    	wdirPaths = split(wdirVar, ":")
    Else
    	wdirPaths = split(wdirVar, ";")
    End If
    Dim i, count, filePathName
    count = UBound(wdirPaths,1)
    For i = 0 To count
        If Scripting.IsUnix Then
            filePathName = wdirPaths(i) & "/" & fileName
        Else
            filePathName = wdirPaths(i) & "\" & fileName
        End If
        If fso.FileExists(filePathName) Then
            FindFileInWdir = filePathName
'msgbox FindFileInWdir
            Exit Function
        End If
    Next
End Function

Sub CompReport()
    If scripting.isunix Then
    	msgbox "This does not run on Linux"
    	Exit Sub
    End if
		Dim cmd_script
		Dim fileName
	    cmd_script = "Reports\ExcelCompList.xls"	
		fileName = aatk_env & "\FlipChip\" & cmd_script
	    Dim sh
		Set sh = CreateObject("Shell.Application")
		
		'msgbox filename
		sh.ShellExecute "excel.exe", fileName
		Set sh = Nothing
End Sub

Sub forceSyncBtnEnable()
	Scripting.Globals("LayoutDrivenNetlistStatus") = 1
	setLDNButton
End Sub

Sub toggleEFMLocks()

	If btn_M1.Visible = True Then
	
		sShowHideTP = "Click to refresh menus. Double+Click to show Lock/Unlock EFM buttons"
'		btn_Unlock.Visible = False
'		btn_Lock.Visible = False
		btn_M1.Visible = False
		
	Else
	
		sShowHideTP = "Click to refresh menus. Double+Click to hide Lock/Unlock EFM buttons"
'		btn_Unlock.Visible = True
'		btn_Lock.Visible = True
		btn_M1.Visible = True
		
	End If
End Sub

'--------------------------------------------------------------------------- 
' output messages to message window
'--------------------------------------------------------------------------- 
Sub MsgOut(str)

	Call AppendOutput("AATK", str)

End Sub

Function AppendOutput(sOutputTab, str)
	If pcbApp.Name = "Discovery PCB - Viewer" or pcbApp.Name = "Xpedition xPCB Viewer"  or pcbApp.Name = "Xpedition Layout Viewer"Then
	Else
		Dim mnu, OutputControl, objTab
		Set mnu = Gui.CommandBars("Document Menu Bar").Controls("&View").Controls("Message Window")		  
		If mnu.Checked = False Then
			Call Gui.ProcessCommand(33125)
		End If
		Set OutputControl = Addins.Item("Message Window").Control
		Set objTab = OutputControl.AddTab(sOutputTab)
		Call objTab.Activate
		Addins("Message Window").Control.AddTab(sOutputTab).AppendText(str & vbCrlf)
	End If
End Function

Function ClearOutputWindow(sOutputTab)
	Addins("Message Window").Control.AddTab(sOutputTab).Clear
End Function

' ===========================================
' Category button location, visibility, label etc
' ===========================================
Sub displayCatButtons(ButtonId, Label)

	Dim btnHT: btnHT = 30
	Dim btnWD: btnWD = 60
	Dim btnLT: btnLT = 10
	Dim btnRT: btnRT = 70
	Select Case ButtonId
		Case 1
			Call ConfigureButtons(c1, 180, 210, btnLT, btnRT, Label, "", True)
		Case 2
			Call ConfigureButtons(c2, c1.Top, c1.Bottom, btnLT+btnWD, btnRT+btnWD, Label, "", True)
		Case 3
			Call ConfigureButtons(c3, c2.Top, c2.Bottom, c2.Left+btnWD, c2.Right+btnWD, Label, "", True)
		Case 4
			Call ConfigureButtons(c4, c3.Top, c3.Bottom, c3.Left+btnWD, c3.Right+btnWD, Label, "", True)
		Case 5
			Call ConfigureButtons(c5, c1.Top+btnHT, c1.Bottom+btnHT, btnLT, btnRT, Label, "", True)
		Case 6
			Call ConfigureButtons(c6, c5.Top, c5.Bottom, c5.Left+btnWD, c5.Right+btnWD, Label, "", True)
		Case 7
			Call ConfigureButtons(c7, c6.Top, c6.Bottom, c6.Left+btnWD, c6.Right+btnWD, Label, "", True)
		Case 8
			Call ConfigureButtons(c8, c7.Top, c7.Bottom, c7.Left+btnWD, c7.Right+btnWD, Label, "", True)
		Case 9
			Call ConfigureButtons(c9, c5.Top+btnHT, c5.Bottom+btnHT, c5.Left, c5.Left+90, Label, "", True)
		Case 10
			Call ConfigureButtons(c10, c9.Top, c9.Bottom, c9.Left+90, c9.Right+150, Label, "", True)
' be careful changing these
		Case 11 
			Call ConfigureButtons(c11, c9.Top+btnHT, c9.Bottom+20, btnLT, btnRT, Label, "", True)
' Updated 06/08/15 Ian Gabbitas - changed width of button #12 for Editor Control category
		Case 12
			Call ConfigureButtons(c12, c11.Top, c11.Bottom, c11.Left+btnWD, c11.Right+btnWD+60, Label, "", True)
		Case 13 ' Updated 06/08/15 Ian Gabbitas - changed left prop of btn #13
			Call ConfigureButtons(c13, c12.Top, c12.Bottom, c12.Left+btnWD+60, c12.Right+btnWD, Label, "", True)
		'Case 14
			'Call ConfigureButtons(c14, c13.Top, c13.Bottom, c13.Left+btnWD, c13.Right+btnWD, Label, "", True)
	End Select
End Sub

Sub displayCmdButton(ButtonId, Label, ToolTipText)
	Dim btnHT: btnHT = 27
	Dim btnLT: btnLT = 20
	Dim btnRT: btnRT = 240
	Select Case ButtonId
		Case 1
			Call ConfigureButtons(e1, 320, 345, 20, 240, Label, ToolTipText, True)
		Case 2
			Call ConfigureButtons(e2, e1.Top+btnHT, e1.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 3
			Call ConfigureButtons(e3, e2.Top+btnHT, e2.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 4                                         
			Call ConfigureButtons(e4, e3.Top+btnHT, e3.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 5                                         
			Call ConfigureButtons(e5, e4.Top+btnHT, e4.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 6                                         
			Call ConfigureButtons(e6, e5.Top+btnHT, e5.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 7                                         
			Call ConfigureButtons(e7, e6.Top+btnHT, e6.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 8                                         
			Call ConfigureButtons(e8, e7.Top+btnHT, e7.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 9                                         
			Call ConfigureButtons(e9, e8.Top+btnHT, e8.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 10
			Call ConfigureButtons(e10, e9.Top+btnHT, e9.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 11
			Call ConfigureButtons(e11, e10.Top+btnHT, e10.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 12                                            
			Call ConfigureButtons(e12, e11.Top+btnHT, e11.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 13                                            
			Call ConfigureButtons(e13, e12.Top+btnHT, e12.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 14                                            
			Call ConfigureButtons(e14, e13.Top+btnHT, e13.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 15                                            
			Call ConfigureButtons(e15, e14.Top+btnHT, e14.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 16                                            
			Call ConfigureButtons(e16, e15.Top+btnHT, e15.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 17                                            
			Call ConfigureButtons(e17, e16.Top+btnHT, e16.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 18                                            
			Call ConfigureButtons(e18, e17.Top+btnHT, e17.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 19                                            
			Call ConfigureButtons(e19, e18.Top+btnHT, e18.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
		Case 20                                            
			Call ConfigureButtons(e20, e19.Top+btnHT, e19.Bottom+btnHT, btnLT, btnRT, Label, ToolTipText, True)
	End Select
End Sub

Sub ConfigureButtons(btn_Id, btn_Top, btn_Bot, btn_left, btn_right, Label, ToolTipText, bEnable)
	With btn_Id
		.Bottom = btn_Bot
		.Top = btn_Top
		.Left = btn_left
		.Right 	= btn_right
		.Visible = True
		.Text = Label
		.ToolTipText = ToolTipText
		.Enable = bEnable
	End With
End Sub

Sub ClearCatBtnColors
	c1.BackColor = -1
	c1.ForeColor = -1
	c2.BackColor = -1
	c2.ForeColor = -1
	c3.BackColor = -1
	c3.ForeColor = -1
	c4.BackColor = -1
	c4.ForeColor = -1
	c5.BackColor = -1
	c5.ForeColor = -1
	c6.BackColor = -1
	c6.ForeColor = -1
	c7.BackColor = -1
	c7.ForeColor = -1
	c8.BackColor = -1
	c8.ForeColor = -1
	c9.BackColor = -1
	c9.ForeColor = -1
	c10.BackColor = -1
	c10.ForeColor = -1
	c11.BackColor = -1
	c11.ForeColor = -1
	c12.BackColor = -1
	c12.ForeColor = -1
	c13.BackColor = -1
	c13.ForeColor = -1
	c14.BackColor = -1
	c14.ForeColor = -1
	c15.BackColor = -1
	c15.ForeColor = -1
	c16.BackColor = -1
	c16.ForeColor = -1
End Sub

Sub TidyFormBtns()

	Call TidyBtns(f1, 30, 320)
	Call TidyBtns(f2, 60, 320)
	Call TidyBtns(f3, 90, 320)
	Call TidyBtns(f4, 120, 320)
	Call TidyBtns(f5, 150, 320)
	Call TidyBtns(f6, 180, 320)

End Sub

Sub TidyCats()
	Call TidyBtns(c1, 30, 320)
	Call TidyBtns(c2, 60, 320)
	Call TidyBtns(c3, 90, 320)
	Call TidyBtns(c4, 120, 320)
	Call TidyBtns(c5, 150, 320)
	Call TidyBtns(c6, 180, 320)
	Call TidyBtns(c7, 210, 320)
	Call TidyBtns(c8, 240, 320)
	Call TidyBtns(c9, 270, 320)
	Call TidyBtns(c10, 300, 320)
	Call TidyBtns(c11, 330, 320)
	Call TidyBtns(c12, 360, 320)
	Call TidyBtns(c13, 390, 320)
	Call TidyBtns(c14, 420, 320)
	Call TidyBtns(c15, 420, 320)
	Call TidyBtns(c16, 420, 320)
End Sub

Sub TidyCmds()
	Call TidyBtns(e1, 30, 260)
	Call TidyBtns(e2, 60, 260)
	Call TidyBtns(e3, 90, 260)
	Call TidyBtns(e4, 120, 260)
	Call TidyBtns(e5, 150, 260)
	Call TidyBtns(e6, 180, 260)
	Call TidyBtns(e7, 210, 260)
	Call TidyBtns(e8, 240, 260)
	Call TidyBtns(e9, 270, 260)
	Call TidyBtns(e10, 300, 260)
	Call TidyBtns(e11, 330, 260)
	Call TidyBtns(e12, 360, 260)
	Call TidyBtns(e13, 390, 260)
	Call TidyBtns(e14, 420, 260)
	Call TidyBtns(e15, 450, 260)
	Call TidyBtns(e16, 480, 260)
	Call TidyBtns(e17, 510, 260)
	Call TidyBtns(e18, 540, 260)
	Call TidyBtns(e19, 570, 260)
	Call TidyBtns(e20, 600, 260)
End Sub

Sub TidyBtns(oBtn, offset, leftside)

	With oBtn
	
		.Visible = False
		.BackColor = -1
		.ForeColor = -1
		.Left = leftside
		.Right = leftside + 40
		.Top = btnTop + offset
		.Bottom = btnTop + offset + 30
		
	End With	
	
End Sub

Sub ResetAllButtons()

	Dim btnHT: btnHT = 40
	Dim ToolTipText: ToolTipText = ""
	Call ConfigureButtons(e1, 260, 300, 120, 150, "e1", ToolTipText, True)
	Call ConfigureButtons(e2, e1.Top+btnHT, e1.Bottom+btnHT, e1.Left, e1.Right, "e2", ToolTipText, True)
	Call ConfigureButtons(e3, e2.Top+btnHT, e2.Bottom+btnHT, e1.Left, e1.Right, "e3", ToolTipText, True)
	Call ConfigureButtons(e4, e3.Top+btnHT, e3.Bottom+btnHT, e1.Left, e1.Right, "e4", ToolTipText, True)
	Call ConfigureButtons(e5, e4.Top+btnHT, e4.Bottom+btnHT, e1.Left, e1.Right, "e5", ToolTipText, True)
	Call ConfigureButtons(e6, e5.Top+btnHT, e5.Bottom+btnHT, e1.Left, e1.Right, "e6", ToolTipText, True)
	Call ConfigureButtons(e7, e6.Top+btnHT, e6.Bottom+btnHT, e1.Left, e1.Right, "e7", ToolTipText, True)
	Call ConfigureButtons(e8, e7.Top+btnHT, e7.Bottom+btnHT, e1.Left, e1.Right, "e8", ToolTipText, True)
	Call ConfigureButtons(e9, e8.Top+btnHT, e8.Bottom+btnHT, e1.Left, e1.Right, "e9", ToolTipText, True)
	Call ConfigureButtons(e10, e9.Top+btnHT, e9.Bottom+btnHT, e1.Left, e1.Right, "e10", ToolTipText, True)
	Call ConfigureButtons(e11, e10.Top+btnHT, e10.Bottom+btnHT, e1.Left, e1.Right, "e11", ToolTipText, True)
	Call ConfigureButtons(e12, e11.Top+btnHT, e11.Bottom+btnHT, e1.Left, e1.Right, "e12", ToolTipText, True)
	Call ConfigureButtons(e13, e12.Top+btnHT, e12.Bottom+btnHT, e1.Left, e1.Right, "e13", ToolTipText, True)
	Call ConfigureButtons(e14, e13.Top+btnHT, e13.Bottom+btnHT, e1.Left, e1.Right, "e14", ToolTipText, True)
	Call ConfigureButtons(e15, e14.Top+btnHT, e14.Bottom+btnHT, e1.Left, e1.Right, "e15", ToolTipText, True)
	Call ConfigureButtons(e16, e15.Top+btnHT, e15.Bottom+btnHT, e1.Left, e1.Right, "e16", ToolTipText, True)
	Call ConfigureButtons(e17, e16.Top+btnHT, e16.Bottom+btnHT, e1.Left, e1.Right, "e17", ToolTipText, True)
	Call ConfigureButtons(e18, e17.Top+btnHT, e17.Bottom+btnHT, e1.Left, e1.Right, "e18", ToolTipText, True)
	Call ConfigureButtons(e19, e18.Top+btnHT, e18.Bottom+btnHT, e1.Left, e1.Right, "e19", ToolTipText, True)
	Call ConfigureButtons(e20, e19.Top+btnHT, e19.Bottom+btnHT, e1.Left, e1.Right, "e20", ToolTipText, True)

	Call ConfigureButtons(c1, 320, 360, 120, 150, "c1", ToolTipText, True)
	Call ConfigureButtons(c2, c1.Top+btnHT, c1.Bottom+btnHT, c1.Left, c1.Right, "c2", ToolTipText, True)
	Call ConfigureButtons(c3, c2.Top+btnHT, c2.Bottom+btnHT, c1.Left, c1.Right, "c3", ToolTipText, True)
	Call ConfigureButtons(c4, c3.Top+btnHT, c3.Bottom+btnHT, c1.Left, c1.Right, "c4", ToolTipText, True)
	Call ConfigureButtons(c5, c4.Top+btnHT, c4.Bottom+btnHT, c1.Left, c1.Right, "c5", ToolTipText, True)
	Call ConfigureButtons(c6, c5.Top+btnHT, c5.Bottom+btnHT, c1.Left, c1.Right, "c6", ToolTipText, True)
	Call ConfigureButtons(c7, c6.Top+btnHT, c6.Bottom+btnHT, c1.Left, c1.Right, "c7", ToolTipText, True)
	Call ConfigureButtons(c8, c7.Top+btnHT, c7.Bottom+btnHT, c1.Left, c1.Right, "c8", ToolTipText, True)
	Call ConfigureButtons(c9, c8.Top+btnHT, c8.Bottom+btnHT, c1.Left, c1.Right, "c9", ToolTipText, True)
	Call ConfigureButtons(c10, c9.Top+btnHT, c9.Bottom+btnHT, c1.Left, c1.Right, "c1", ToolTipText, True)
	Call ConfigureButtons(c11, c10.Top+btnHT, c10.Bottom+btnHT, c1.Left, c1.Right, "c11", ToolTipText, True)
	Call ConfigureButtons(c12, c11.Top+btnHT, c11.Bottom+btnHT, c1.Left, c1.Right, "c12", ToolTipText, True)
	Call ConfigureButtons(c13, c12.Top+btnHT, c12.Bottom+btnHT, c1.Left, c1.Right, "c13", ToolTipText, True)
	Call ConfigureButtons(c14, c13.Top+btnHT, c13.Bottom+btnHT, c1.Left, c1.Right, "c14", ToolTipText, True)
	Call ConfigureButtons(c15, c14.Top+btnHT, c14.Bottom+btnHT, c1.Left, c1.Right, "c15", ToolTipText, True)
	Call ConfigureButtons(c16, c15.Top+btnHT, c15.Bottom+btnHT, c1.Left, c1.Right, "c16", ToolTipText, True)

	Call ConfigureButtons(f1, 380, 430, 120, 150, "f1", ToolTipText, True)
	Call ConfigureButtons(f2, f1.Top+btnHT, f1.Bottom+btnHT, f1.Left, f1.Right, "f2", ToolTipText, True)
	Call ConfigureButtons(f3, f2.Top+btnHT, f2.Bottom+btnHT, f1.Left, f1.Right, "f3", ToolTipText, True)
	Call ConfigureButtons(f4, f3.Top+btnHT, f3.Bottom+btnHT, f1.Left, f1.Right, "f4", ToolTipText, True)
	Call ConfigureButtons(f5, f4.Top+btnHT, f4.Bottom+btnHT, f1.Left, f1.Right, "f5", ToolTipText, True)
	Call ConfigureButtons(f6, f5.Top+btnHT, f5.Bottom+btnHT, f1.Left, f1.Right, "f6", ToolTipText, True)

End Sub

'End of (Declarations)

Sub Bitmap2_EventClick()
Dim This : Set This = Bitmap2
	initForm
	'loadCmdTab
	'Call ExecCommand("/Default/AboutDisclaimer.efm")
End Sub 

Sub Bitmap2_EventDblClick()
Dim This : Set This = Bitmap2
	toggleEFMLocks
End Sub 

Sub btn_Lock_EventClick()
Dim This : Set This = btn_Lock
	Call ExecCommand("Default\setReadOnlyEFM.vbs")
End Sub 

Sub btn_Lock_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = btn_Lock
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = "Set all efm forms to readonly"
	End If
End Sub 

Sub btn_M1_EventClick()
Dim This : Set This = btn_M1
	Dim strFileEdit
	strFileEdit =  "%SystemRoot%/system32/notepad.exe " & xmlFile
	CreateObject("WScript.Shell").Run(strFileEdit)

End Sub 

Sub btn_M1_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = btn_M1
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = "Edit main menu XML file"
	End If
End Sub 

Sub btn_Unlock_EventClick()
Dim This : Set This = btn_Unlock
	Call ExecCommand("Default\unsetReadOnlyEFM.vbs")	
End Sub 

Sub btn_Unlock_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = btn_Unlock
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = "Unlock all efm forms for editing"
	End If
End Sub 

Sub ButtonSyncDesign_EventClick()
Dim This : Set This = ButtonSyncDesign
	Call ExecCommandNoDoc("/Netlist/sync_netlist.vbs")
	
	setLDNButton
End Sub 

Sub c1_EventClick()
Dim This : Set This = c1
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c1_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c1
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c10_EventClick()
Dim This : Set This = c10
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c10_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c10
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c11_EventClick()
Dim This : Set This = c11
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c12_EventClick()
Dim This : Set This = c12
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c13_EventClick()
Dim This : Set This = c13
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c14_EventClick()
Dim This : Set This = c14
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c15_EventClick()
Dim This : Set This = c15
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c16_EventClick()
Dim This : Set This = c16
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c2_EventClick()
Dim This : Set This = c2
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c2_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c2
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c3_EventClick()
Dim This : Set This = c3
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns

End Sub 

Sub c3_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c3
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c4_EventClick()
Dim This : Set This = c4
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c4_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c4
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c5_EventClick()
Dim This : Set This = c5
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c5_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c5
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c6_EventClick()
Dim This : Set This = c6
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c6_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c6
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c7_EventClick()
Dim This : Set This = c7
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c7_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c7
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c8_EventClick()
Dim This : Set This = c8
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c8_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c8
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub c9_EventClick()
Dim This : Set This = c9
	ClearCatBtnColors
	This.BackColor = rgb(0,0,130)
	This.ForeColor = rgb(255,255,0)
	ActiveCat = This.Text
	txt_Cat.Text = ActiveCat
	loadCmdBtns
End Sub 

Sub c9_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = c9
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 5) And (X < .Right - .Left -5) And _
			(Y > 5) And (Y < .Bottom - .Top -5)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		'StatusBar1.Panels(1).Text = "" ' clear field when cursor gets close to edge of button
		editBox_statusbar.Text = ""
	Else
		'StatusBar1.Panels(1).Text = "Delete plane shadow voids for the selected nets."
		editBox_statusbar.Text = ""
	End If
End Sub 

Sub e1_EventClick()
Dim This : Set This = e1
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			If strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			
			Exit Sub
		End If
	Next
End Sub 

Sub e1_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e1
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e10_EventClick()
Dim This : Set This = e10
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e10_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e10
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e11_EventClick()
Dim This : Set This = e11
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e11_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e11
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e12_EventClick()
Dim This : Set This = e12
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e12_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e12
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e13_EventClick()
Dim This : Set This = e13
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e13_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e13
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e14_EventClick()
Dim This : Set This = e14
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e14_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e14
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e15_EventClick()
Dim This : Set This = e15
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e15_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e15
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e16_EventClick()
Dim This : Set This = e16
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e16_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e16
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e17_EventClick()
Dim This : Set This = e17
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e17_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e17
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e18_EventClick()
Dim This : Set This = e18
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e18_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e18
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e19_EventClick()
Dim This : Set This = e19
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e19_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e19
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e2_EventClick()
Dim This : Set This = e2
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e2_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e2
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If

End Sub 

Sub e20_EventClick()
Dim This : Set This = e20
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e20_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e20
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e3_EventClick()
Dim This : Set This = e3
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e3_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e3
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e4_EventClick()
Dim This : Set This = e4
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e4_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e4
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e5_EventClick()
Dim This : Set This = e5
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e5_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e5
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e6_EventClick()
Dim This : Set This = e6
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e6_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e6
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e7_EventClick()
Dim This : Set This = e7
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e7_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e7
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e8_EventClick()
Dim This : Set This = e8
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e8_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e8
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub e9_EventClick()
Dim This : Set This = e9
	Dim dctKey
	Dim strArr1, strArr2, strArr3
	For Each dctKey In dctCmds
		strArr1 = split(dctKey, ",")
		strArr3 = split(dctCmds(dctKey),",") ' split dictionary value (path, tooltip)
		strArr2 = split(strArr3(0),"=") ' get PATH/SUB string value
		If strArr1(0) = ActiveCat And strArr1(1) = This.Text Then		
			If strArr2(0) = "SUBROUTINE" Then ' this button will run a subroutine
				Execute strArr2(1)
			End If
			if strArr2(0) = "PATH" Then ' this button will run another script
				Call ExecCommand(strArr2(1))
			End If
			If strArr2(0) = "KEYIN" Then ' this button will run a keyin
				Call ExecKeyin(strArr2(1))
			End If
			Exit Sub
		End If
	Next

End Sub 

Sub e9_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = e9
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If
End Sub 

Sub f1_EventClick()
Dim This : Set This = f1

	RunCommand(This)
	
End Sub 

Sub f1_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = f1
	
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If

End Sub 

Sub f2_EventClick()
Dim This : Set This = f2

	RunCommand(This)

End Sub 

Sub f2_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = f2
		
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If

End Sub 

Sub f3_EventClick()
Dim This : Set This = f3
	
	RunCommand(This)

End Sub 

Sub f3_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = f3
		
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If

End Sub 

Sub f4_EventClick()
Dim This : Set This = f4

	RunCommand(This)

End Sub 

Sub f4_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = f4
		
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If

End Sub 

Sub f5_EventClick()
Dim This : Set This = f5
	
	RunCommand(This)

End Sub 

Sub f5_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = f5
		
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If

End Sub 

Sub f6_EventClick()
Dim This : Set This = f6
	
	RunCommand(This)

End Sub 

Sub f6_EventMouseMove(Button, Shift, X, Y)
Dim This : Set This = f6
		
	Dim Flag, i
	Flag = False	
	With This
		If ((X > 1) And (X < .Right - .Left -1) And _
			(Y > 1) and (Y < .Bottom - .Top -1)) Then
			Flag = True
		Else
			flag = False
		End If
	End With
	
	If Flag = False Then 
		editBox_statusbar.Text = "" ' clear field when cursor gets close to edge of button
	Else
		editBox_statusbar.Text = This.ToolTipText
	End If

End Sub 

Sub FormOnTop_EventClick()
Dim This : Set This = FormOnTop
	Dim aatk_env
	Dim FormnameVis
	Dim FormName
	Dim Vis
	Dim FoTicon
	Dim prog
	Dim exe
	Dim TransLvl
	
	REM Dim FoT 'Move to General Declarations

	FormName = theframe.Title
	'msgbox formname
	TransLvl = "@" & 255
	'TransLvl = "@" & FormOnTopSlider.Value
	aatk_env = Scripting.GetEnvVariable("AATK")
	prog = aatk_env & "/FlipChip/Utilities/ontop.exe" 
	If FoT = False Then
		FoTicon = aatk_env & "/FlipChip/icons/FormOnTop.bmp"
		this.Bitmap = FoTicon
		this.BackColor = 16777088
		FormnameVis = FormName & "@True" & TransLvl
		'FormOnTopSlider.Visible	= True
		FoT = True
	Else
		FoTicon = aatk_env & "/FlipChip/icons/FormOnBot.bmp"
		this.Bitmap = FoTicon
		this.BackColor = -1
		FormnameVis = FormName & "@False"
		'FormOnTopSlider.Visible	= False
		FoT = False 
	End If
	Call application.ProcessScriptEx(prog, True, FormnameVis, exe, False)
	
End Sub 

Sub Text4_EventClick()
Dim This : Set This = Text4
	initForm
End Sub 

Sub Text4_EventDblClick()
Dim This : Set This = Text4
	toggleEFMLocks
End Sub 

Sub Text5_EventClick()
Dim This : Set This = Text5
	initForm
End Sub 

Sub Text5_EventDblClick()
Dim This : Set This = Text5
	toggleEFMLocks
End Sub 

Sub Text7_EventInitialize()
Dim This : Set This = Text7
	Dim fso
'Open clip file
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim aatk_env, aatk_main, ver_file, inline, parse
	aatk_env = Scripting.GetEnvVariable("AATK")
    Set ver_file = fso.OpenTextFile(aatk_env & slash & "Version.txt", 1)
     Do While Not ver_file.AtEndOfStream
     
        inline = ver_file.ReadLine
        parse = split(inline,"=")
        If parse(0) = "BUILD_NUMBER" Then
        	this.Text = "Build: " & parse(1)
        	'msgbox parse(1)
        End If
        
      Loop
End Sub 

Sub TheView_EventClick()
Dim This : Set This = TheView
	
End Sub 

Sub TheView_EventInitialize()
Dim This : Set This = TheView
	'TheFrame.Height = 770
	TheFrame.Width = 260

End Sub 

Sub TheView_EventTerminate()
Dim This : Set This = TheView
	TidyCats
	TidyCmds
	ResetAllButtons
End Sub 

Sub txt_Cat_EventClick()
Dim This : Set This = txt_Cat
	setLDNButton
End Sub 