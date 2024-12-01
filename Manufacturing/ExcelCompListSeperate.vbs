' This script creates excel and loads it with component information.
' The script will also supports cross probing by listening for 
' selections in the PCB tool and making those selections 
' in the Excel spreadsheet.
' Cross probing can be done from Excel to the PCB tool but that
' must be implemented using Excels VBA. (Author: Toby Rimes)

' Add more component information.
' Modify by szg on 2024/11/19
Option Explicit     

' Constants
Const SHEETNAME_COMP_LIST_TOP = "Component List Top"
Const SHEETNAME_COMP_LIST_BOTTOM = "Component List Bottom"
Const SHEETNAME_COMP_LIST_INTERNAL = "Component List Internal"

Const REFDES_COL = "A"
Const PARTNAME_COL = "B"
Const PINS_COL = "C"
Const LAYER_COL = "D"
Const ORIENTATION_COL = "E"
Const LOCATION_X_COL = "F"
Const LOCATION_Y_COL = "G"
Const PARTNUMBER_COL = "H"
Const VALUE_COL = "I"
Const HEIGHT_COL = "J"

Const HEADER_ROW = 1
Const FIRST_COMPONENT_ROW = 2

' Import type library doesn't work for Excel.
' Define used Excel constants here.
Const xlWhole = 1
Const xlPart = 2

' Add any type libraries to be used.
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

' Global variables
Dim pcbAppObj                'Application object
Dim pcbDocObj                'Document object
Dim excelAppObj              'Excel application

' Get the application object.
Set pcbAppObj = Application

' Get the active document
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)
    
' Create the excel applicatoin
Set excelAppObj = CreateObject("Excel.Application")
   
' Load excel
pcbAppObj.Gui.CursorBusy(True)
Call LoadExcel()
pcbAppObj.Gui.CursorBusy(False)

' Make the excel application visible.
excelAppObj.Visible = True

' Attach events to the document object to get selection changes.
' Call Scripting.AttachEvents(pcbDocObj, "pcbDocObj")

' Hang around to listen to events 
' Scripting.DontExit = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Event Handlers

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main Functions

' Loads excel with components and header information
Sub LoadExcel()
	' Create a workbook
	Dim workbookObj 
	Set workbookObj = excelAppObj.Workbooks.Add
	workbookObj.worksheets.Add
	workbookObj.worksheets.Add
	
	' Get the first sheet
	Dim cmpListTopSheetObj, cmpListBotSheetObj, cmpListInSheetObj
	Set cmpListTopSheetObj = workbookObj.Worksheets.Item(1)
	Set cmpListBotSheetObj = workbookObj.worksheets.Item(2)
	Set cmpListInSheetObj = workbookObj.worksheets.Item(3)
	
	' Rename the worksheet.
	cmpListTopSheetObj.Name = SHEETNAME_COMP_LIST_TOP
	cmpListBotSheetObj.Name = SHEETNAME_COMP_LIST_BOTTOM
	cmpListInSheetObj.Name = SHEETNAME_COMP_LIST_INTERNAL
	
	' Set the header information
	Call DefineHeaders(cmpListTopSheetObj)
	Call DefineHeaders(cmpListBotSheetObj)
	Call DefineHeaders(cmpListInSheetObj)
	
	' Get the components
	Dim cmpColl
	Set cmpColl = pcbDocObj.Components

	Dim	cmpListSheetObjArr
	cmpListSheetObjArr = Array(cmpListTopSheetObj, cmpListBotSheetObj, cmpListInSheetObj)
	
	' Sort the component collection
	Call cmpColl.Sort()
	
	' Add the collection
	Call AddComponents(cmpListSheetObjArr, FIRST_COMPONENT_ROW, cmpColl)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Utility functions

' Creates header information.
' sheetObj - Excel Worksheet Object
Sub DefineHeaders(sheetObj)
	Dim unit
	Select Case pcbDocObj.currentUnit
		Case epcbUnitInch
			unit = "inch"
		Case epcbUnitMils
			unit = "mil"
		Case epcbUnitMM
			unit = "mm"
		Case epcbUnitUM
			unit = "um"
	End Select
	sheetObj.Range(REFDES_COL & HEADER_ROW).Value = "Ref Des"
	sheetObj.Range(PARTNAME_COL & HEADER_ROW).Value = "PartName"
	sheetObj.Range(PINS_COL & HEADER_ROW).Value = "Pins"
	sheetObj.Range(LAYER_COL & HEADER_ROW).Value = "Layer"
	sheetObj.Range(ORIENTATION_COL & HEADER_ROW).Value = "Orientation"
	sheetObj.Range(LOCATION_X_COL & HEADER_ROW).Value = "X" & "(" & unit & ")"
	sheetObj.Range(LOCATION_Y_COL & HEADER_ROW).Value = "Y" & "(" & unit & ")"
	sheetObj.Range(PARTNUMBER_COL & HEADER_ROW).Value = "PartNumber"
	sheetObj.Range(VALUE_COL & HEADER_ROW).Value = "Value"
	sheetObj.Range(HEIGHT_COL & HEADER_ROW).Value = "Height" & "(" & unit & ")"
	
	sheetObj.Rows(HEADER_ROW).Font.Bold = True
End Sub

' Add the collection of components
' sheetObjArr - Excel Worksheet Object Array
' startRowInt - Integer
' cmpColl - Component Collection
Sub AddComponents(sheetObjArr, startRowInt, cmpColl)
	Dim topSheetRowInt, botSheetRowInt, inSheetRowInt
	topSheetRowInt = startRowInt
	botSheetRowInt = startRowInt
	inSheetRowInt = startRowInt
	Dim cmpObj, sheetObj
	For Each cmpObj In cmpColl
		If cmpObj.Layer = 1 Then
			Set sheetObj = sheetObjArr(0)
			Call AddComponent(sheetObj, topSheetRowInt, cmpObj)
			topSheetRowInt = topSheetRowInt + 1
		ElseIf cmpObj.Layer = pcbDocObj.LayerCount Then
			Set sheetObj = sheetObjArr(1)
			Call AddComponent(sheetObj, botSheetRowInt, cmpObj)
			botSheetRowInt = botSheetRowInt + 1
		Else
			Set sheetObj = sheetObjArr(2)
			Call AddComponent(sheetObj, inSheetRowInt, cmpObj)
			inSheetRowInt = inSheetRowInt + 1
		End If
	Next
	
	' Adjust the cells size to fit the text
	Dim i
	For i = 0 To UBound(sheetObjArr)
		Call sheetObjArr(i).Columns.AutoFit()
	Next
End Sub

' Add a single component to row rowInt
' sheetObj - Excel Worsheet Object
' rowInt - Integer
' cmpObj - Component Object
Sub AddComponent(sheetObj, rowInt, cmpObj)
	Dim cmpLayerInt, cmpLayerStr, cmpPlacementOutlineColl, cmpPlacementOutlineObj, cmpHeightInt, cmpPropertyObj

	cmpLayerInt = cmpObj.Layer
	cmpHeightInt = 0
	Set cmpPlacementOutlineColl = cmpObj.PlacementOutlines
	Set cmpPropertyObj = cmpObj.FindProperty("Value")

	sheetObj.Range(REFDES_COL & rowInt).Value = cmpObj.Name
	sheetObj.Range(PARTNAME_COL & rowInt).Value = cmpObj.PartName
	sheetObj.Range(PINS_COL & rowInt).Value = cmpObj.Pins.Count
	Select Case cmpLayerInt
		Case 1 
			cmpLayerStr = "TOP"
		Case pcbDocObj.LayerCount
			cmpLayerStr = "BOTTOM"
		Case Else
			cmpLayerStr = "INTERNAL"
	End Select
	sheetObj.Range(LAYER_COL & rowInt).Value = cmpLayerStr
	sheetObj.Range(PARTNUMBER_COL & rowInt).Value = cmpObj.PartNumber

	If cmpPropertyObj Is Nothing Then

	Else
		sheetObj.Range(VALUE_COL & rowInt).Value = cmpObj.FindProperty("Value").Value
	End If

	For Each cmpPlacementOutlineObj In cmpPlacementOutlineColl
		If cmpPlacementOutlineObj.Height > cmpHeightInt Then
			cmpHeightInt = cmpPlacementOutlineObj.Height
		End If
	Next
	If cmpHeightInt = 0 Then
		sheetObj.Range(HEIGHT_COL & rowInt).Value = "None"
	Else 
		sheetObj.Range(HEIGHT_COL & rowInt).Value = cmpHeightInt
	End If
	
	If cmpObj.Placed Then
		sheetObj.Range(LOCATION_X_COL & rowInt).Value = cmpObj.PositionX
		sheetObj.Range(LOCATION_Y_COL & rowInt).Value = cmpObj.PositionY
		sheetObj.Range(ORIENTATION_COL & rowInt).Value = cmpObj.Orientation
	Else
		sheetObj.Range(LOCATION_X_COL & rowInt).Value = "Unplaced"
		sheetObj.Range(LOCATION_Y_COL & rowInt).Value = "Unplaced"
		sheetObj.Range(ORIENTATION_COL & rowInt).Value = "Unplaced"
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Helper Functions

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Miscelaneous Functions

' Validate server function
Private Function ValidateServer(doc)
    
    Dim key, licenseServer, licenseToken

    ' Ask Expedition�s document for the key
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
