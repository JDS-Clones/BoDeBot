VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHookManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   Message Hooker Manager
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 07/18/98
'
'       Internal Dependencies:
'           CHookManager.cls
'
'       External Dependencies:
'           MsgHook.dll
'
'**************************************************************

'This collection will contain references to
'everything we are currently hooking. Every
'object we add to this class must have a
'public hWnd property and a public WndProc property. Each object must also
'have a Public Hook() method.
Private BagOfHooks As New Collection
'This is our actual hooker
'Private WithEvents hooker As CMsgHook

Public Sub AddHook(Ob As Object, ByVal uMsg As Long)
    Dim szKey As String
    szKey = CStr(Ob.hWnd)
    'add object to collection using its hWnd as the key
    'If Not (BagOfHooks.Item(szKey) Is Nothing) Then
        BagOfHooks.Add Ob, szKey
        'start the new hook
        Ob.WndProc = Hook(Ob.hWnd)
    'End If
End Sub
Public Sub DelHook(Ob As Object, ByVal uMsg As Long)
    Dim szKey As String
    szKey = CStr(Ob.hWnd)
    'delete all matcing hWnds
    BagOfHooks.Remove szKey
    'stop the hook
    UnHook Ob.hWnd, Ob.WndProc
End Sub

Private Sub Class_Initialize()
    Set hooker = Me
End Sub

Private Sub Class_Terminate()
    'We need to first unhook everything
    Dim i As Integer
    For i = 1 To BagOfHooks.Count
        UnHook BagOfHooks.Item(i).hWnd, BagOfHooks.Item(i).WndProc
    Next i
    'then kill the hooker object
End Sub

Friend Function HookMe(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    Dim szKey As String
    szKey = CStr(hWnd)
    'send the hook to the write place
    If BagOfHooks.Count > 0 Then
        HookMe = BagOfHooks.Item(szKey).Hook(uMsg, wp, lp)
    End If
End Function

Public Function GetVersion() As String
    GetVersion = "[" & App.Title & "] " & App.Major & "." & App.Minor & "." & App.Revision
End Function

