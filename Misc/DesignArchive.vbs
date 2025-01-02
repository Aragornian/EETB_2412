Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj, jobName, mPath
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

mPath = pcbDocObj.MasterPath

' Get the application object
Dim dllApp
Set dllApp = CreateObject("MGCPCBReleaseEnvironmentLib.MGCPCBReleaseEnvServer")

Dim execObj
Set execObj = CreateObject("viewlogic.Exec")

Dim DxArchiverFileName, projectFileName, outputDir
DxArchiverFileName = dllApp.sddHome + "\common\win64\bin\DxArchiver.exe"
projectFileName = pcbDocObj.ProjectIntegration.ProjectFile
' outputDir = ".\"
outputDir = Left(mPath, Len(mPath) - 4)

' Save design before archiving
' FILE_SAVE   57603
pcbAppObj.Gui.SuppressTrivialDialogs = True
pcbAppObj.Gui.ProcessCommand(57603)
pcbAppObj.Gui.SuppressTrivialDialogs = False

' Command usage
' DxArchiver [−p file_name] [−t directory] [−c file_name] [−l log_file] [-zip | -edx] [-createpdf] [-setStatic] [−noGUI]
'   -p project_file_name
'   -t ouput_directory
'   -noGUI
'   -zip compress the contents
' Run this command from a command shell
Dim archiveCmd, visibleInt, waitBool
archiveCmd = "cmd.exe /k" + DxArchiverFileName + " -noGUI -p " + """" + projectFileName + """" + " -t " + """" + outputDir + """" + " -zip"
visibleInt = 1
waitBool = True

' Archive the project
' Compressed file is named as "project_name + time"
Call execObj.Run(archiveCmd, visibleInt, waitBool)

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