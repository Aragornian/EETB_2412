Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Toggle design mode
' CND_MODE_DRAW         32835
' CMD_MODE_ROUTE        32814
' CMD_MODE_PLACE        32813
Dim CurrentModeEnum
CurrentModeEnum = pcbAppObj.Gui.ActiveMode

Select Case CurrentModeEnum
    Case epcbModePlace
        pcbAppObj.Gui.Processcommand(32814)
    Case epcbModeRoute
        pcbAppObj.Gui.Processcommand(32835)
    Case epcbModeDrawing
        pcbAppObj.Gui.Processcommand(32813)
    Case Else
        MsgBox("Unsupported")
End Select
