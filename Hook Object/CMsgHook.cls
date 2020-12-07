VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMsgHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   Message Hooker
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 5/31/98
'
'       Internal Dependencies:
'           MsgHook.bas
'           CMsgHook.cls
'
'       External Dependencies:
'           -none-
'
'**************************************************************


'local variable(s) to hold property value(s)
Private mvarhWnd As Long 'local copy
Private mvaruMsg As Long 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent onHook[(arg1, arg2, ... , argn)]
Public Event onHook(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wp As Long, ByVal lp As Long)

Public Sub StopHook(ByVal hand As Long)
    UnHook hand
    'WM_HOOKMESSAGE = 0
End Sub

Public Sub StartHook()
    Set hooker = Me
    WM_HOOKMESSAGE = mvaruMsg
    Hook mvarhWnd
End Sub

Friend Sub HookMe(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    RaiseEvent onHook(hw, uMsg, wParam, lParam)
End Sub

Public Property Let uMsg(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.uMsg = 5
    mvaruMsg = vData
End Property


Public Property Get uMsg() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.uMsg
    uMsg = mvaruMsg
End Property



Public Property Let hWnd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hWnd = 5
    mvarhWnd = vData
End Property


Public Property Get hWnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hWnd
    hWnd = mvarhWnd
End Property


Public Function GetVersion() As String
    GetVersion = "[" & App.Title & "] " & App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Sub Class_Initialize()

End Sub
