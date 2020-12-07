VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CScriptManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
'   Scripting Manager
'       Author: Nicholas J. Felmlee
'       Email: Nick@Felmlee.Com
'       Last Revision: 7/15/98
'Windows95 users or NT users without IIS need to download:
'http://www.microsoft.com/msdownload/vbscript/scripting.asp
'
'**************************************************************


Private WithEvents ScriptEngine As ScriptControl
Attribute ScriptEngine.VB_VarHelpID = -1
Private ScObjects As Collection
Private ScKeys(1 To 255) As String
Private lastkey As Integer
'local variable(s) to hold property value(s)
Private mvarEnabled As Boolean 'local copy
Public Sub Reset()
    On Error Resume Next
    Dim i As Integer
    'disable any calls to the script
    mvarEnabled = False
    'reset the object
    ScriptEngine.Reset
    'load the new script
    For i = 1 To ScObjects.count
        ScriptEngine.AddObject ScKeys(i), ScObjects.Item(i), True
    Next i
    Open App.Path & "\required.scp" For Input As #1
        ScriptEngine.AddCode Input(LOF(1), 1)
    Close #1
    Open App.Path & "\user.scp" For Input As #1
        ScriptEngine.AddCode Input(LOF(1), 1)
    Close #1
    'enable calls to script
    mvarEnabled = True
End Sub

Public Sub AddObject(Ob As Object, ByVal szName As String)
    ScObjects.Add Ob, szName
    ScKeys(lastkey) = szName
    lastkey = lastkey + 1
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
End Property


Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Sub onBotCommands(UserOb As Object, ByVal szCommand As String, ByVal szParams As String)
'FIRES WHEN SOMEONE PERFORMS A BOT COMMAND
    If mvarEnabled = False Then Exit Sub
    ScriptEngine.Run "onBotCommandEvent", UserOb, szCommand, szParams
End Sub

Public Sub onNotifyList(ByVal szNotifyList As String)
'FIRES WHEN WE RECIEVE A NOTIFY ALERT
    If mvarEnabled = False Then Exit Sub
    ScriptEngine.Run "onNotifyListEvent", szNotifyList
End Sub
Public Sub onChannelJoin(ByVal szChannel As String, ByVal szNick As String)
'FIRES WHEN WE OR SOMEONE ELSE JOINS A CHANNEL
    If mvarEnabled = False Then Exit Sub
    ScriptEngine.Run "onChannelJoinEvent", szChannel, szNick
End Sub

Public Sub onChannelPart(ByVal szChannel As String, ByVal szNick As String)
'FIRES WHEN WE OR SOMEONE ELSE PARTS A CHANNEL
    If mvarEnabled = False Then Exit Sub
    ScriptEngine.Run "onChannelPartEvent", szChannel, szNick
End Sub

Public Sub onChanCommands(ChannelWindow As Object, ByVal szCommand As String, ByVal szParams As String)
'FIRES WHEN A USER PERFORMS A NON-STANDARD COMMAND IN A CHANNEL WINDOW (/COMMAND)
    If mvarEnabled = False Then Exit Sub
    On Error Resume Next
    ScriptEngine.Run "onChanCommandEvent", ChannelWindow, szCommand, szParams
End Sub

Public Sub onMsgCommands(MessageWindow As Object, ByVal szCommand As String, ByVal szParams As String)
'FIRES WHEN A USER PERFORMS A NON-STANDARD COMMAND IN A MESSAGE WINDOW (/COMMAND)
    If mvarEnabled = False Then Exit Sub
    On Error Resume Next
    ScriptEngine.Run "onMsgCommandEvent", MessageWindow, szCommand, szParams
End Sub

Public Sub onConsoleCommands(ConsoleWindow As Object, ByVal szCommand As String, ByVal szParams As String)
'FIRES WHEN A USER PERFORMS A NON-STANDARD COMMAND IN A CONSOLE WINDOW (/COMMAND)
    If mvarEnabled = False Then Exit Sub
    On Error Resume Next
    ScriptEngine.Run "onConsoleCommandEvent", ConsoleWindow, szCommand, szParams
End Sub

Public Sub onCTCPCommand(ByVal szSender As String, ByVal szCommand As String, szParams As String)
'FIRES WHEN WE GET A NON-STANDARD INCOMING CTCP COMMAND
    If mvarEnabled = False Then Exit Sub
    On Error Resume Next
    ScriptEngine.Run "onCTCPCommandEvent", szSender, szCommand, szParams
End Sub
Public Sub onChannelMessage(ByVal szSender As String, ByVal szChannel As String, ByVal szLine As String)
'FIRES WHEN WE GET A NON-STANDARD INCOMING CTCP COMMAND
    If mvarEnabled = False Then Exit Sub
    On Error Resume Next
    ScriptEngine.Run "onChannelMessageEvent", szSender, szChannel, szLine
End Sub
Public Sub onConnectionStatusChange(ByVal Status As Boolean)
'FIRES WHEN WE CONNECT OR DISCONNECT FROM IRC
'CONNECTED -> Status=True
'DISCONNECTED -> Status=False
'TO DO: not implemented yet
    If mvarEnabled = False Then Exit Sub
    On Error Resume Next
    ScriptEngine.Run "onConnectionStatusChangeEvent", Status
End Sub

Private Sub Class_Initialize()
    Set ScriptEngine = MDI.ScriptControl1
    'ScriptEngine.Language = vbNull 'set to null as a hack around ocx bug
    ScriptEngine.Language = "VBScript"
    ScriptEngine.AllowUI = True
    Set ScObjects = New Collection
    Me.Enabled = False
    lastkey = 1
End Sub

Private Sub Class_Terminate()
    ScriptEngine.Reset
    Set ScriptEngine = Nothing
End Sub

Private Sub ScriptEngine_Error()
    On Error Resume Next
    MsgBox "There was an error in your script! " & ScriptEngine.Error.Text, vbExclamation, "Script Error"
End Sub
