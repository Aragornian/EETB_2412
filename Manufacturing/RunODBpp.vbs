Option Explicit

' Add type library
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
Scripting.AddTypeLibrary("MGCPCBEngines.ODBPPOutputEngine")
Scripting.AddTypeLibrary("Scripting.FileSystemObject")

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj, jobName, masterPath
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

' Master path for team layout
jobName = pcbDocObj.FullName
masterPath = pcbDocObj.MasterPath
Dim odbOutputPath
odbOutputPath = masterPath + "Output\ODBpp\"

' Create a FileSystemObject
Dim fileSysObj 
Set fileSysObj = CreateObject("Scripting.FileSystemObject")
' Create folder for ODBpp files
If fileSysObj.FolderExists(odbOutputPath) = False Then
    fileSysObj.CreateFolder(odbOutputPath)
End If
' Running Executables
Dim execObj
Set execObj = CreateObject("viewlogic.Exec")

If pcbAppObj.LockServer = True Then  
    pcbAppObj.Gui.CursorBusy(True)
    
    ' Save file
    pcbAppObj.Gui.SuppressTrivialDialogs = True
    pcbAppObj.Gui.ProcessCommand("File->Save")
    pcbAppObj.Gui.SuppressTrivialDialogs = False
    ' Run ODB++ output engine
    RunODBPP()

    pcbAppObj.Gui.CursorBusy(False)
    pcbAppObj.UnlockServer
End If

'************************************************************************ 
' Run ODB++ output
Sub RunODBPP()
    ' Create ODB++ Engine object 
    Dim oODBPPEngine 
    Set oODBPPEngine = CreateObject("MGCPCBEngines.ODBPPOutputEngine")

    ' Set the design file name 
    oODBPPEngine.DesignFileName = jobName 

    ' Setup files for ODB++ output 
    Call SetupODBPPOutput(oODBPPEngine)

    ' Run the ODB++ Engine 
    On Error Resume Next
    Err.Clear()
    oODBPPEngine.Go()

    ' Check errors 
    Dim oErrors
    Dim oErr 
    Set oErrors = oODBPPEngine.Errors 
    For Each oErr in oErrors 
        Write oError.ErrorString
    Next 
    If oErrors.Count = 0 Then
        ' Open folder after output
        Dim openCommand, visibleInt, waitBool
        visibleInt = 1
        waitBool = True
        ' cmd command, path with space should be quote use ", """" represent "" in vbs
        ' e.g. : start "" "C:\user\desktop\odb output"
        openCommand = "cmd.exe /k start """" " + """" + odbOutputPath + """"
        Call execObj.Run(openCommand, visibleInt, waitBool)
        ' MsgBox("Output ODB++ Successfully!")
        Write "Output ODB++ Successfully! " & Now()        
    End If
    Err.Clear()
End Sub

' Write message in Addin windows
Sub Write(sMsg)
	If Not IsObject(pcbAppObj) Then
		Echo sMsg
	ElseIf Not pcbAppObj.Addins("Message Window") Is Nothing Then
		pcbAppObj.Addins("Message Window").Control.AddTab("RunODBpp").AppendText sMsg & vbCrLf
        pcbAppObj.Addins("Message Window").Control.ActivateTab("RunODBpp")
	End If
End Sub

' Setup ODB++ output
Sub SetupODBPPOutput(oODBPPEngine)
    Dim outputName, pcbName, odbOutputLogFileDir
    pcbName = Split(pcbDocObj.Name, ".pcb")
    outputName = "Designodb_" + pcbName(0)
    odbOutputLogFileDir = masterPath + "\LogFiles"

    With oODBPPEngine
        ' .ClearAdvancedPackagingLayers()
        .ClearExportLayers()
        ' .ClearVariants()
        .OutputPath = odbOutputPath
        .LogFileDirectory = odbOutputLogFileDir
        .OutputJobName = outputName
        .Units = eengOutputUnitsEnglish
        .ODBPPVersion = 8
        .CompressOutput = TRUE
        .AdvancedPackagingData = FALSE
        .AppendToLogFile = FALSE
        .BoardPanelOutline = TRUE
        .ExportEmbeddedPassiveLayers = TRUE
        .ExportPartNumbers = FALSE
        .GenerateSeparateVariantOutputs = FALSE
        .IgnoreComponentLayout = FALSE
        ' .IncludeAllVariantData = TRUE
        .NeutralizeNets = FALSE
        .NonFuncPinPadRemoval = eengNonFuncRemoveNone
        .NonFuncViaPadRemoval = eengNonFuncRemoveNone
        .NPINetTypes = FALSE
        .ODBPPExportMode = eengFullExport
        .PackageOutlineLayer = eengPackageOutlineAssemblyOutline
        .ReadDRCFeatures = FALSE
        .RemoveCadnetNetlist = FALSE
        .RemoveEDAData = FALSE
        .RoundCorners = TRUE
        .UseGeneratedSilkscreenData = FALSE
        ' Export fab and assy data
        .LayersToExport = .AssemblyLayers
        .LayersToExport = .FabricationLayers
    End With
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
