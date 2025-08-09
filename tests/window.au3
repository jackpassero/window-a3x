#include-once ;УТФ-8
If $CmdLine[0] < 1 Then Exit 1
Main($CmdLine[1])

Func Main($title)
    Local $x = 10
    Local $w = 300
    Local $hGUI = GUICreate($CmdLine[1], $w, 60, 0, 0)
    Local $pid = WinGetProcess($hGUI)

    GUICtrlCreateLabel('pid: ' & $pid, $x, 10, $w - $x - 10, 20)
    GUICtrlCreateLabel('handle: ' & $hGUI, $x, 30, $w - $x - 10, 20)
    GUISetState(@SW_SHOW, $hGUI)
    ConsoleWrite('ready')
    While 1
        Switch GUIGetMsg()
            Case -3
                ExitLoop
        EndSwitch
    WEnd
    GUIDelete($hGUI)
EndFunc
