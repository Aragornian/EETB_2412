' This script creates excel and loads it with component information.
' The script will also supports cross probing by listening for 
' selections in the PCB tool and making those selections 
' in the Excel spreadsheet.
' Cross probing can be done from Excel to the PCB tool but that
' must be implemented using Excels VBA. (Author: Toby Rimes)
Option Explicit     

' Constants
Const SHEETNAME_COMP_LIST = "Component List"

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

Const HEADER_ROW = 4
Const FIRST_COMPONENT_ROW = 5

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
Call LoadExcel()

' Make the excel application visible.
excelAppObj.Visible = True

' Attach events to the document object to get selection changes.
Call Scripting.AttachEvents(pcbDocObj, "pcbDocObj")

' Hang around to listen to events 
Scripting.DontExit = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Event Handlers

' Document event fired when there is a selection change
' typeEnum - Unused Enumerate
Sub pcbDocObj_OnSelectionChange(typeEnum)
	If ExcelIsRunning(excelAppObj) Then
		Dim cmpsColl
		Set cmpsColl = pcbDocObj.Components(epcbSelectSelected)
		Call SelectComponentsInExcel(cmpsColl)
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main Functions

' Loads excel with components and header information
Sub LoadExcel()
	' Create a workbook
	Dim workbookObj 
	Set workbookObj = excelAppObj.Workbooks.Add
	
	' Get the first sheet
	Dim cmpListSheetObj
	Set cmpListSheetObj = workbookObj.Worksheets.Item(1)
	
	' Rename the worksheet.
	cmpListSheetObj.Name = SHEETNAME_COMP_LIST
	
	' Set the header information
	Call DefineHeaders(cmpListSheetObj)
	
	' Get the components
	Dim cmpColl
	Set cmpColl = pcbDocObj.Components
	
	' Sort the component collection
	Call cmpColl.Sort()
	
	' Add the collection
	Call AddComponents(cmpListSheetObj, FIRST_COMPONENT_ROW, cmpColl)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Utility functions

' Select the components in the spreadsheet.
' cmpsColl - Component Collection
Sub SelectComponentsInExcel(cmpsColl)
	' Get the sheet
	Dim cmpListSheetObj
	Set cmpListSheetObj = excelAppObj.Sheets.Item(1)
	
	' Instantiate a range object to hold multiple cells
	Dim multiRangeObj
	Set multiRangeObj = Nothing
    
    ' Loop through all components to build a Range 
    ' of rows for the components in cmpsColl
    Dim cmpObj
    For Each cmpObj In cmpsColl
        ' Find the ref des in the Ref Des column. Match the whole name.
        Dim foundCellObj
        Set foundCellObj = cmpListSheetObj.Columns(REFDES_COL).Find(cmpObj.Name,,,xlWhole)
        ' If we found something add it to the multi range
        If Not foundCellObj Is Nothing Then   
        	If  Not multiRangeObj Is Nothing Then       
            	Set multiRangeObj = excelAppObj.Union(multiRangeObj, _
            	    cmpListSheetObj.Rows(foundCellObj.Row))
            Else
            	Set multiRangeObj = cmpListSheetObj.Rows(foundCellObj.Row)
            End If     
        End If
    Next
    
    ' Select all the rows in the range.
    Call SelectExcelRange(cmpListSheetObj, multiRangeObj)
End Sub

' Creates header information.
' sheetObj - Excel Worksheet Object
Sub DefineHeaders(sheetObj)
	sheetObj.Range(REFDES_COL & HEADER_ROW).Value = "Ref Des"
	sheetObj.Range(PARTNAME_COL & HEADER_ROW).Value = "PartName"
	sheetObj.Range(PINS_COL & HEADER_ROW).Value = "Pins"
	sheetObj.Range(LAYER_COL & HEADER_ROW).Value = "Layer"
	sheetObj.Range(ORIENTATION_COL & HEADER_ROW).Value = "Orientation"
	sheetObj.Range(LOCATION_X_COL & HEADER_ROW).Value = "X"
	sheetObj.Range(LOCATION_Y_COL & HEADER_ROW).Value = "Y"
	sheetObj.Range(PARTNUMBER_COL & HEADER_ROW).Value = "PartNumber"
	sheetObj.Range(VALUE_COL & HEADER_ROW).Value = "Value"
	sheetObj.Range(HEIGHT_COL & HEADER_ROW).Value = "Height"
	
	sheetObj.Rows(HEADER_ROW).Font.Bold = True
End Sub

' Add the collection of components
' sheetObj - Excel Worksheet Object
' startRowInt - Integer
' cmpColl - Component Collection
Sub AddComponents(sheetObj, startRowInt, cmpColl)
	Dim rowInt
	rowInt = startRowInt
	Dim cmpObj
	For Each cmpObj In cmpColl
		Call AddComponent(sheetObj, rowInt, cmpObj)
		rowInt = rowInt + 1
	Next
	
	' Adjust the cells size to fit the text
	Call sheetObj.Columns.AutoFit()
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

' Returns true if Excel is still running false otherwise
' appObj - Excel Application Object
Function ExcelIsRunning(appObj)
	' Initialize return value
	ExcelIsRunning = True

	' If the variable is nothing return false
	If appObj Is Nothing Then
		ExcelIsRunning = False
		Exit Function
	End If
	
	' Check to see it excel is running by trying to call a method on
	' excel.  If there is an exception assume it has been shut down.
	On Error Resume Next
	Call Err.Clear()
	
	' Make a call that would cause an exception
	Dim sheetsObj
	Set sheetsObj = appObj.Sheets
	
	' Check the error value.
	If Err Then
		Set appObj = Nothing
		ExcelIsRunning = False
	End If
End Function

' Selects a range of objects and colors the range yellow
' sheetObj - Excel Worksheet object
' rangeObj - Excel Range Object
Sub SelectExcelRange(sheetObj, rangeObj)
	' Remove the yellow
	Call RemoveExcelFill(sheetObj)
	If Not rangeObj Is Nothing Then
		' Set the interior to yellow and select
		rangeObj.Interior.ColorIndex = 6
		Call rangeObj.Select()
	Else
		' Cause an unselection by selection first cell
		Call sheetObj.Range("A1").Select()
	End If
End Sub

' Sets the fill to white for all the cells on a sheet
' sheetobj - Excel Sheet Object
Sub RemoveExcelFill(sheetObj)
	' Set all to white
	sheetObj.Cells.Interior.ColorIndex = -4142
End Sub

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
