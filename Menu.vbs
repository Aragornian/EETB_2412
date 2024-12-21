Option Explicit

' Add any type libraries to be used.
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
Scripting.AddTypeLibrary("MGCSDD.CommandBarsEx")

' Get the application object.
Dim pcbApp
Set pcbApp = Application

' Find the document menu bar.
Dim docMenuBarObj
Set docMenuBarObj = pcbApp.Gui.CommandBars("Document Menu Bar")

' Get the collection of controls for the menu 
'(i.e. menu popup buttons, File, Edit, View, etc...)
Dim docMenuBarCtrlColl
Set docMenuBarCtrlColl = docMenuBarObj.Controls

Dim eetbMenuObj, levelOneMenuObj, levelOneControlsColl, levelTwoMenuObj

'''''''''''''''''' Add the EETB menu '''''''''''''''''''''''''''''''''''''

' Check to see if the eetb menu is already there
Set eetbMenuObj = FindMenu("EETB", docMenuBarObj)

' If there is already a eetb menu then delete it
If Not eetbMenuObj Is Nothing Then
    eetbMenuObj.Delete()
End If
' Create the new button by adding to the control collection
Set eetbMenuObj = docMenuBarCtrlColl.Add(cmdControlPopup,,,-1)
' Configure the menu control
eetbMenuObj.Caption = "EETB"

'Get the control collection for the new EETB menu
Dim eetbControlsColl
Set eetbControlsColl = eetbMenuObj.Controls

''''''''''''''Add toolbox dialog menu under EETB ''''''''''''''''''''''''''

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "EE Auto Tool Box"
    levelOneMenuObj.OnAction = "run %EETB%\XpeditionAutoToolBox.efm"
Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButtonSeparator,,,-1)

'''''''''''''''''' Add "Display" menu under EETB ''''''''''''''''''''''''''

' Create the new button by adding to the control collection
Set levelOneMenuObj = eetbControlsColl.Add(cmdControlPopup,,,-1)
    levelOneMenuObj.Caption = "Display Schemes"
    Set levelOneControlsColl = levelOneMenuObj.Controls
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButton,,,-1)
        levelTwoMenuObj.Caption = "Route"
        levelTwoMenuObj.OnAction = "run %EETB%\Display\DisplaySchemeRoute.vbs"
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButtonSeparator,,,-1)
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButton,,,-1)
        levelTwoMenuObj.Caption = "AssemblyTop"
        levelTwoMenuObj.OnAction = "run %EETB%\Display\DisplaySchemeAssemblyTop.vbs"
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButton,,,-1)
        levelTwoMenuObj.Caption = "AssemblyBottom"
        levelTwoMenuObj.OnAction = "run %EETB%\Display\DisplaySchemeAssemblyBottom.vbs"
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButtonSeparator,,,-1)
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButton,,,-1)
        levelTwoMenuObj.Caption = "Color Power/GND"
        levelTwoMenuObj.OnAction = "run %EETB%\Display\SetGndPwrNetsColor.vbs"
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButton,,,-1)
        levelTwoMenuObj.Caption = "Color Impedance NetClasses"
        levelTwoMenuObj.OnAction = "run %EETB%\Display\SetNetClassesColor.vbs"

'''''''''''''''''' Add another menu under EETB ''''''''''''''''''''''''''''

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "Set PadEntry"
    levelOneMenuObj.OnAction = "run %EETB%\Constraint\SetPadEntry.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "Unfix & Unlock"
    levelOneMenuObj.OnAction = "run %EETB%\Route\UnfixUnlockObject.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "Get Polygon Area"
    levelOneMenuObj.OnAction = "run %EETB%\Route\GetPolyArea.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "Adjust Silkscreen RefDes"
    levelOneMenuObj.OnAction = "run %EETB%\Manufacturing\RefDesAdjust.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlPopup,,,-1)
    levelOneMenuObj.Caption = "DXF Output"
    Set levelOneControlsColl = levelOneMenuObj.Controls
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButton,,,-1)
    levelTwoMenuObj.Caption = "DXF Top"
    levelTwoMenuObj.OnAction = "run %EETB%\Manufacturing\RunDXFExportTop.vbs"
    Set levelTwoMenuObj = levelOneControlsColl.Add(cmdControlButton,,,-1)
    levelTwoMenuObj.Caption = "DXF Bottom"
    levelTwoMenuObj.OnAction = "run %EETB%\Manufacturing\RunDXFExportBottom.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "Gerber Output"
    levelOneMenuObj.OnAction = "run %EETB%\Manufacturing\RunGerber.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "NCDrill Output"
    levelOneMenuObj.OnAction = "run %EETB%\Manufacturing\RunNCDrill.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "ODB++ Output"
    levelOneMenuObj.OnAction = "run %EETB%\Manufacturing\ODBPPOutput.vbs"

Set levelOneMenuObj = eetbControlsColl.Add(cmdControlButton,,,-1)
    levelOneMenuObj.Caption = "BOM"
    levelOneMenuObj.OnAction = "run %EETB%\Manufacturing\ExcelComplist.vbs"

' Keep this script running so that the handler can be executed.
' Scripting.DontExit = True

''''''''''''''''''' Local Functions ''''''''''''''''''''''''''''''''''''''''

Function FindMenu(menuToFind, menuBar)
    Dim ctrls : Set ctrls = menuBar.Controls
    Dim ctrl
    
    Set FindMenu = Nothing
    
    For Each ctrl In ctrls
       Dim capt: capt = ctrl.Caption
       capt = Replace(capt, "&", "")
       If capt = menuToFind Then
           Set FindMenu = ctrl
           Exit For
       End If
    Next 
End Function

' Server validation function
Function ValidateServer(docObj)
    
    Dim keyInt
    Dim licenseTokenInt
    Dim licenseServerObj

    ' Ask Expedition’s document for the key
    keyInt = docObj.Validate(0)

    ' Get license server
    Set licenseServerObj = CreateObject("MGCPCBAutomationLicensing.Application")

    ' Ask the license server for the license token
    licenseTokenInt = licenseServerObj.GetToken(keyInt)

    ' Release license server
    Set licenseServerObj = nothing

    ' Turn off error messages.  
    On Error Resume Next
    Err.Clear

    ' Ask the document to validate the license token
    docObj.Validate(licenseTokenInt)
    If Err Then
        ValidateServer = 0    
    Else
        ValidateServer = 1
    End If

End Function