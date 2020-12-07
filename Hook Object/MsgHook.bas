Attribute VB_Name = "MsgHookLib"
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public hooker As CHookManager
    Public Function Hook(ByVal hWnd As Long) As Long
        Hook = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    End Function
    Public Sub UnHook(ByVal hWnd As Long, ByVal lpPrevWndProc As Long)
        Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevWndProc)
    End Sub
    Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
            Dim Proc As Long
            Proc = hooker.HookMe(hw, uMsg, wParam, lParam)
            If Proc <> 0 Then
                WindowProc = CallWindowProc(Proc, hw, uMsg, wParam, lParam)
            End If
    End Function

