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
Dim jobName  
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
    Set app = GetObject(, "MGCPCB.ExpeditionPCBApplication." & sComVersion)
    Set docObj = GetLicensedDoc(app)
    jobName = docObj.FullName
End If

ExportDXF

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
    
    ' Get design Config directory
    Dim pcbPath
    pcbPath = Split(jobName, "PCB")
    If InStr(jobName, ".pnl") Then
        pos = InstrRev(jobName, dirSlash, -1, 1)
        pcbPath = Left(jobName, Len(jobName) - pos)
        outputFile = pcbPath(0) + "/Output/" + "DXFExport.dxf"
    Else 
        pcbPath = Split(jobName, "PCB")
        outputFile = pcbPath(0) + "PCB/Output/" + "DXFExport.dxf"    
    End If

    oDXFExport.DesignFileName = jobName
    oDXFExport.Scheme(0) = "Default"
    oDXFExport.OutputFileName=outputFile
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
    oDXFExport.FillPads = TRUE
    oDXFExport.MirrorOutput = False
    Dim nLayerCount
    nLayerCount = oDXFExport.LayerCount
    Call SetDXFExportElements(oDXFExport)
End Sub

'*************************
' set DXFExportElements
'*************************
Sub SetDXFExportElements(oDXFExport)
    Call oDXFExport.PutElementExport(eengElementTypeConductiveShape, "uniquelayer",0)
    Call oDXFExport.PutElementExport(eengElementTypeAssemblyOutlineBottom,"",1)
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