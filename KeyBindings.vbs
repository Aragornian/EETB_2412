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
keyBindTableColl.AddKeyBinding "a","Edit->Snap->Toggle Hover Snap",BindMenu,BindAccelerator

' Edit Fix & Lock
keyBindTableColl.AddKeyBinding "s","Edit->Fix Lock->Unfix",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "d","Edit->Fix Lock->Semi Fix",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "f","Edit->Fix Lock->Fix",BindMenu,BindAccelerator

' Move rich graph
keyBindTableColl.AddKeyBinding "e","Smart Utilities->Design Editing Aids->Move /w Rich Graphics",BindMenu,BindAccelerator

' Toggle DRC
keyBindTableColl.AddKeyBinding "g","run ToggleDRC.vbs",BindCommand,BindAccelerator

' Toggle display patterns
keyBindTableColl.AddKeyBinding "i","run ToggleDisplayPatterns.vbs",BindCommand,BindAccelerator

' Toggle display planes
keyBindTableColl.AddKeyBinding "j","run ToggleDisplayPlanes.vbs",BindCommand,BindAccelerator

' Measure distance
keyBindTableColl.AddKeyBinding "l","Edit->Measure->Minimum Distance",BindMenu,BindAccelerator

' Mirror view
keyBindTableColl.AddKeyBinding "m","run ToggleMirrorView.vbs",BindCommand,BindAccelerator

' Toggle display netlines
keyBindTableColl.AddKeyBinding "n","run ToggleDisplayNetlines.vbs",BindCommand,BindAccelerator

' Toggle mode
keyBindTableColl.AddKeyBinding "q","run ToggleMode.vbs",BindCommand,BindAccelerator

' Swap parts and nets
keyBindTableColl.AddKeyBinding "r","Place->Swap Parts",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "t","Route->Swap->Nets",BindMenu,BindAccelerator

' Find next open net
keyBindTableColl.AddKeyBinding "v","fnl",BindCommand,BindAccelerator

' Edit Shape
keyBindTableColl.AddKeyBinding "w","Draw->Edit->Modify Shape",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "`","Smart Utilities->Design Editing Aids->Area Cut Trace",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "o","Planes->Plane Shape",BindMenu,BindAccelerator

' Show whole board
keyBindTableColl.AddKeyBinding "z","zb",BindCommand,BindAccelerator

' Assign net name
keyBindTableColl.AddKeyBinding "Alt+a","run AssignNetName.vbs",BindCommand,BindAccelerator

' Align object
keyBindTableColl.AddKeyBinding "Alt+e","Edit->Align->Align Top",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "Alt+d","Edit->Align->Align Bottom",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "Alt+s","Edit->Align->Align Left",BindMenu,BindAccelerator
keyBindTableColl.AddKeyBinding "Alt+f","Edit->Align->Align Right",BindMenu,BindAccelerator

' Rotation
keyBindTableColl.AddKeyBinding "Shift+q","rs 45",BindCommand,BindAccelerator

' Change plane to conductiveshape
keyBindTableColl.AddKeyBinding "ctrl+\","run ChangePlaneToConductiveShape.vbs",BindCommand,BindAccelerator

' Change conductiveshape to plane
keyBindTableColl.AddKeyBinding "ctrl+/","run ChangeConductiveShapeToPlane.vbs",BindCommand,BindAccelerator
