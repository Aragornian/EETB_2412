Option Explicit

' Add type library so that we can use BindCommand and BindAccelerator constants.
Scripting.AddTypeLibrary("MGCSDD.KeyBindings")
' Get the application object
Dim pcbApp
Set pcbApp = Application
' Get the document key bind table
Dim keyBindTableColl
Set keyBindTableColl = pcbApp.Gui.Bindings("Document")

' Add key to bind table
' Change layer shortkeys
keyBindTableColl.AddKeyBinding "1","cl 1",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "2","cl 2",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "3","cl 3",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "4","cl 4",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "5","cl 5",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "6","cl 6",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "7","cl 7",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "8","cl 8",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "9","cl 9",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "0","cl 10",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "-","cl 11",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "=","cl 12",BindCommand,BindAccelerator

keyBindTableColl.AddKeyBinding "num 1","cl 1",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 2","cl 2",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 3","cl 3",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 4","cl 4",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 5","cl 5",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 6","cl 6",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 7","cl 7",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 8","cl 8",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 9","cl 9",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num 0","cl 10",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num -","cl 11",BindCommand,BindAccelerator
keyBindTableColl.AddKeyBinding "num +","cl 12",BindCommand,BindAccelerator

' Snap
Call ExecuteKeyBindFunction("a","RunSnapToggleHover")

' Edit Fix & Lock
Call ExecuteKeyBindFunction("s","RunUnfix")
Call ExecuteKeyBindFunction("d","RunSemiFix")
Call ExecuteKeyBindFunction("f","RunFix")

' Move rich graph
Call ExecuteKeyBindFunction("e","RunMoveWithRichGraphic")

' Toggle DRC
keyBindTableColl.AddKeyBinding "g","run %EETB_2412%\Misc\ToggleDRC.vbs",BindCommand,BindAccelerator

' Toggle display patterns
keyBindTableColl.AddKeyBinding "i","run %EETB_2412%\Display\ToggleDisplayPatterns.vbs",BindCommand,BindAccelerator

' Toggle display planes
keyBindTableColl.AddKeyBinding "j","run %EETB_2412%\Display\ToggleDisplayPlanes.vbs",BindCommand,BindAccelerator

' Measure distance
Call ExecuteKeyBindFunction("l","RunMeasureDistanceCenter")

' Mirror view
keyBindTableColl.AddKeyBinding "m","run %EETB_2412%\Display\ToggleMirrorView.vbs",BindCommand,BindAccelerator

' Toggle display netlines
keyBindTableColl.AddKeyBinding "n","run %EETB_2412%\Route\ToggleDisplayNetlines.vbs",BindCommand,BindAccelerator

' Toggle mode
keyBindTableColl.AddKeyBinding "q","run %EETB_2412%\Misc\ToggleMode.vbs",BindCommand,BindAccelerator
Call ExecuteStrokeBindFunction("258","RunDrawMode")
Call ExecuteStrokeBindFunction("654","RunPlaceMode")
Call ExecuteStrokeBindFunction("456","RunRouteMode")

' Swap parts and nets
Call ExecuteKeyBindFunction("r","RunSwapParts")
Call ExecuteKeyBindFunction("t","RunSwapNets")

' Find next open net
keyBindTableColl.AddKeyBinding "v","fnl",BindCommand,BindAccelerator

' Edit Shape
Call ExecuteKeyBindFunction("w","RunModifyShape")

' Cut trace
Call ExecuteKeyBindFunction("`","RunCutTrace")

' Draw shape
Call ExecuteKeyBindFunction("o","RunDrawPlaneShape")

' Show whole board
keyBindTableColl.AddKeyBinding "z","zb",BindCommand,BindAccelerator

' Assign net name
keyBindTableColl.AddKeyBinding "Alt+a","run %EETB_2412%\Route\AssignNetName.vbs",BindCommand,BindAccelerator

' Align object
Call ExecuteKeyBindFunction("Alt+e","RunAlignTop")
Call ExecuteKeyBindFunction("Alt+d","RunAlignBottom")
Call ExecuteKeyBindFunction("Alt+s","RunAlignLeft")
Call ExecuteKeyBindFunction("Alt+f","RunAlignRight")

' Rotation
keyBindTableColl.AddKeyBinding "Shift+q","rs 45",BindCommand,BindAccelerator

' Change plane to conductiveshape
keyBindTableColl.AddKeyBinding "ctrl+\","run %EETB_2412%\Route\ChangePlaneToConductiveShape.vbs",BindCommand,BindAccelerator

' Change conductiveshape to plane
keyBindTableColl.AddKeyBinding "ctrl+/","run %EETB_2412%\Route\ChangeConductiveShapeToPlane.vbs",BindCommand,BindAccelerator

' Keep this script running so that the handler can be executed 
Scripting.DontExit = True

Sub ExecuteKeyBindFunction(shortcutkey,usrfunction)
    Dim bindObj
    Set bindObj = keyBindTableColl.AddKeyBinding(shortcutkey, usrfunction, BindFunction, BindAccelerator)
    ' Associate the current script engine with the key binding  
    bindObj.Target = ScriptEngine
    ' Call method below with this name 
    bindObj.ExecuteMethod = usrfunction 
End Sub

Sub ExecuteStrokeBindFunction(stroke,usrfunction)
    Dim bindObj
    Set bindObj = keyBindTableColl.AddStrokeBinding(stroke, usrfunction, BindFunction, BindAccelerator)
    ' Associate the current script engine with the key binding  
    bindObj.Target = ScriptEngine
    ' Call method below with this name 
    bindObj.ExecuteMethod = usrfunction 
End Sub

Sub RunCutTrace()
    ' "Smart Utilities->Editing->Cut Trace By Area"
    Gui.ProcessCommand(55026)
End Sub

Sub RunMeasureDistanceCenter()
    ' "Edit->Measure->Measure Minimum Distance"
    Gui.ProcessCommand(60853)
End Sub

Sub RunDrawPlaneShape()
    ' "Planes->General->Draw Plane Shape"
    Gui.ProcessCommand(32888)
End Sub

Sub RunModifyShape()
    ' "Draw->Edit->Modify Shape"
    Gui.ProcessCommand(53308)
End Sub

Sub RunUnfix()
    ' "Edit->Fix Lock->Unfix"
    Gui.ProcessCommand(32916)
End Sub

Sub RunSemiFix()
    ' "Edit->Fix Lock->Semi Fix"
    Gui.ProcessCommand(33435)
End Sub

Sub RunFix()
    ' "Edit->Fix Lock->Fix"
    Gui.ProcessCommand(32866)
End Sub

Sub RunSwapParts()
    ' "Place->Swap Parts"
    Gui.ProcessCommand(33045)
End Sub

Sub RunSwapNets()
    ' "Place->Swap Nets"
    Gui.ProcessCommand(62301)
End Sub

Sub RunAlignTop()
    ' "Edit->Align->Align Top"
    Gui.ProcessCommand(33215)
End Sub

Sub RunAlignBottom()
    ' "Edit->Align->Align Bottom"
    Gui.ProcessCommand(33216)
End Sub

Sub RunAlignLeft()
    ' "Edit->Align->Align Left"
    Gui.ProcessCommand(33213)
End Sub

Sub RunAlignRight()
    ' "Edit->Align->Align Right"
    Gui.ProcessCommand(33214)
End Sub

Sub RunMoveWithRichGraphic()
    ' "Smart Utilities->Design Editing Aids->Move /w Rich Graphics"
    Gui.ProcessCommand(55055)
End Sub

Sub RunSnapToggleHover()
    ' "Smart Utilities->Design Editing Aids->Move /w Rich Graphics"
    Gui.ProcessCommand(59417)
End Sub

Sub RunPlaceMode()
    ' CMD_MODE_PLACE
    Gui.ProcessCommand(32813)
End Sub

Sub RunRouteMode()
    ' CMD_MODE_ROUTE
    Gui.ProcessCommand(32814)
End Sub

Sub RunDrawMode()
    ' CMD_MODE_DRAW
    Gui.ProcessCommand(32835)
End Sub