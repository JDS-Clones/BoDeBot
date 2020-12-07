VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.MDIForm MDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000001&
   Caption         =   "- B o D e B o T -"
   ClientHeight    =   6690
   ClientLeft      =   3075
   ClientTop       =   2340
   ClientWidth     =   9885
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin ComctlLib.Toolbar Mainbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ToolbarImageList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "connect"
            Description     =   "connect"
            Object.ToolTipText     =   "Connect to IRC"
            Object.Tag             =   "connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "disconnect"
            Description     =   "Disconnect"
            Object.ToolTipText     =   "Disconnect from IRC"
            Object.Tag             =   "disconnect"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "setup"
            Description     =   "Setup"
            Object.ToolTipText     =   "Setup"
            Object.Tag             =   "setup"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "dcc send"
            Description     =   "DCC Send"
            Object.ToolTipText     =   "DCC Send"
            Object.Tag             =   "dcc send"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "dcc chat"
            Description     =   "DCC Chat"
            Object.ToolTipText     =   "DCC Chat"
            Object.Tag             =   "dcc chat"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "script editor"
            Description     =   "script editor"
            Object.ToolTipText     =   "Script Editor"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "bot setup"
            Description     =   "Bot Setup"
            Object.ToolTipText     =   "Bot Setup"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "notify"
            Description     =   "notify list"
            Object.ToolTipText     =   "Notify List"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "reconnect"
            Description     =   "Auto Reconnect"
            Object.ToolTipText     =   "Automatically Reconnect on Disconnect"
            Object.Tag             =   "reconnect"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "scripton"
            Object.ToolTipText     =   "Turn on the scripting engine"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   3240
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6435
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "connected_status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bot_user_count"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Key             =   "script"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Script Engine Status"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   9234
            Key             =   "my_time"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer NotifyTimer 
      Interval        =   60000
      Left            =   5190
      Top             =   2880
   End
   Begin VB.PictureBox TabBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   9885
      TabIndex        =   0
      Top             =   6105
      Width           =   9885
      Begin ComctlLib.TabStrip Tab1 
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   661
         Style           =   1
         ImageList       =   "TabImage"
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Console"
               Key             =   "Console"
               Object.Tag             =   "Console"
               Object.ToolTipText     =   "Console Window"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer ReConnectTimer 
      Enabled         =   0   'False
      Interval        =   35000
      Left            =   4710
      Top             =   2880
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   3870
      Top             =   3990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin ComctlLib.ImageList ToolbarImageList 
      Left            =   90
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":116E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":14C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":1A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":1B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2076
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2188
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":229A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":25EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList MainBarImage 
      Left            =   2700
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":293E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":2F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":328C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":35A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":38C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":3BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":3EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":420E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList TabImage 
      Left            =   3330
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":4528
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":4842
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":4B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI.frx":4E76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu File_mnu 
      Caption         =   "&File"
      Begin VB.Menu connect_mnu 
         Caption         =   "&Connect"
      End
      Begin VB.Menu disconnect_mnu 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu setup_mnu 
         Caption         =   "&Setup..."
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu exit_mnu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu scripts_mnu 
         Caption         =   "&Scripts"
         Begin VB.Menu scnew_mnu 
            Caption         =   "New"
            Visible         =   0   'False
         End
         Begin VB.Menu scedit_mnu 
            Caption         =   "Edit"
         End
         Begin VB.Menu scload_mnu 
            Caption         =   "Load"
            Visible         =   0   'False
         End
         Begin VB.Menu screset_mnu 
            Caption         =   "Reset"
         End
      End
      Begin VB.Menu mnuReconnect 
         Caption         =   "&Reconnect"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileH 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileV 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowCloseAll 
         Caption         =   "Close &All"
      End
   End
   Begin VB.Menu help_mnu 
      Caption         =   "&Help"
      Begin VB.Menu contents_mnu 
         Caption         =   "Contents"
      End
      Begin VB.Menu about_mnu 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5
Private Const TBSTYLE_FLAT = &H800
Private Const GWL_STYLE = (-16)
Public WithEvents NetClient As CClientManager
Attribute NetClient.VB_VarHelpID = -1

Public Sub about_mnu_Click()
    frmSplash.Show
End Sub


Public Sub connect_mnu_Click()
  NetClient.ConnectToIrc
End Sub

Public Sub disconnect_mnu_Click()
  NetClient.DisconnectFromIrc
End Sub

Public Sub exit_mnu_Click()
    If NetClient.Connected = True Then
        NetClient.DisconnectFromIrc
    End If
    Unload Me
    'End
End Sub

Private Sub Mainbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case Is = "connect"
            connect_mnu_Click
        Case Is = "disconnect"
            disconnect_mnu_Click
        Case Is = "setup"
            setup_mnu_Click
        Case Is = "bot setup"
            BotSetup.Show
        Case Is = "notify"
            If Button.Value = 1 Then
                NotifyFrm.ShowNotify
            Else
                NotifyFrm.Hide
            End If
        Case Is = "script editor"
            ScriptEditor.Show
        Case Is = "scripton"
            If NetClient.ToggleScriptStatus = True Then
                StatusBar1.Panels(3).Text = " Script Enabled "
            Else
                StatusBar1.Panels(3).Text = " Script Disabled "
            End If
        Case Is = "dcc send"
            startDCCSend InputBox("DCC Send" & vbCrLf & "Enter Nick to send to.", "DCC Send")
        Case Is = "dcc chat"
            startDCCChat InputBox("DCC Chat" & vbCrLf & "Enter Nick to Chat with.", "DCC Chat")
    End Select
End Sub

Private Sub MDIForm_Load()
Dim h As Long
  'Start a new Client Manager
  Set NetClient = New CClientManager
  NetClient.hWnd = Me.hWnd
  ' also hook events from the console
  'Set WConsole.NetClient = NetClient
  '
  StatusBar1.Panels(1).Text = "Disconnected"
  'make flat toolbar
  h = GetWindow(Mainbar.hWnd, GW_CHILD)
  SetWindowLong h, GWL_STYLE, GetWindowLong(h, GWL_STYLE) Or TBSTYLE_FLAT
  WConsole.Show
  Tab1.Left = TabBar.ScaleLeft
  Tab1.Top = TabBar.ScaleTop
  Tab1.Width = TabBar.ScaleWidth
  Tab1.Height = TabBar.ScaleHeight
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If NetClient.Connected = True Then
        NetClient.DisconnectFromIrc
    End If
    Unload WConsole
    Unload BotSetup
    Unload IDSRV
    Unload NotifyFrm
    Unload ScriptEditor
    Unload Servers
    Unload frmSplash
    Unload AddHost_dlg
    NetClient.CleanHouse
'    'unload any other open forms
'    While Forms.count > 1
'        If Not Forms.Item(0).hWnd = Me.hWnd Then
'            Unload Forms.Item(0)
'        End If
'    Wend
    Set NetClient = Nothing
End Sub

Private Sub MDIForm_Resize()
    Tab1.Left = TabBar.ScaleLeft
    Tab1.Top = TabBar.ScaleTop
    Tab1.Width = TabBar.ScaleWidth
    Tab1.Height = TabBar.ScaleHeight
End Sub

Private Sub NetClient_onBotConnection(ByVal szNick As String, ByVal dwCount As Integer)
    StatusBar1.Panels(2).Text = " " & dwCount & " Bot Connections "
End Sub

Private Sub NetClient_onBotDisconnection(ByVal szNick As String, ByVal dwCount As Integer)
    StatusBar1.Panels(2).Text = " " & dwCount & " Bot Connections "
End Sub

Private Sub NetClient_onConnect(ByVal szHost As String)
    StatusBar1.Panels(1).Text = " Connected to " & szHost & " "
End Sub

Private Sub NetClient_onDisconnect(ByVal szHost As String)
    StatusBar1.Panels(1).Text = " Disconnected "
    If NetClient.Retries > 0 And NetClient.Connected = False Then
        ReConnectTimer.Enabled = True
    End If
End Sub

Private Sub NetClient_onHighLightEvent(ByVal szChannel As String)
    Dim x As Integer
    x = Tab1.Tabs.Item(szChannel).Index - 1
    Call SendMessage(Tab1.hWnd, &H1300 + 51, x, &H1)
End Sub

Private Sub NetClient_onKillChannelWindow(ByVal szChan As String)
    Tab1.Tabs.Remove szChan
End Sub

Private Sub NetClient_onNewChannelWindow(ByVal szChan As String)
    Tab1.Tabs.Add , szChan, szChan, 1
End Sub

Private Sub NetClient_onNewMessageWindow(ByVal szNick As String)
    Tab1.Tabs.Add , szNick, szNick, 2
End Sub

Private Sub NetClient_onQueryAppVersion(ByVal szInfo As String)
    UpdateRTB WConsole.RTB, szInfo, NetClient.CLR_NOTICE
End Sub

Private Sub NetClient_onSelectFileForDCCSend(ByVal szNick As String)
    startDCCSend szNick
End Sub

Private Sub NetClient_onUpdateNotifyWindow(ByVal szNicks As String)
    Dim temp As String
    Dim pos As Integer
    'get first nick
    szNicks = Trim(szNicks)
    NotifyFrm.NotifyList.Clear
    While (Not Trim(szNicks) = "")
        pos = InStr(1, szNicks, " ")
        If pos <> 0 Then
            temp = Mid(szNicks, 1, pos - 1)
            szNicks = Mid(szNicks, pos + 1)
        Else
            temp = Mid(szNicks, 1)
            szNicks = ""
        End If
        If (InStr(1, NotifyFrm.NotifyList.Text, temp) = 0) Then
            NotifyFrm.NotifyList.AddItem temp
        End If
        'loop
    Wend
End Sub

Private Sub NotifyTimer_Timer()
    If (NetClient.Connected = True) And (Not Trim(MDI.NetClient.NOTIFYNICKS) = "") Then
        NetClient.Send "ISON " & Trim(MDI.NetClient.NOTIFYNICKS) & vbCrLf
    End If
End Sub

Private Sub ReConnectTimer_Timer()
    If NetClient.Retries > 0 And NetClient.Connected = False Then
        NetClient.ConnectToIrc
        ReConnectTimer.Enabled = False
    End If
End Sub

Public Sub scedit_mnu_Click()
    ScriptEditor.Show
End Sub

Public Sub scload_mnu_Click()
'    CMD.Filter = "TXT (*.txt)|*.txt|SCP (*.scp)|*.scp"
'    CMD.ShowOpen
'    OFILE = CMD.filename
End Sub


Private Sub ScriptManager_Error()

End Sub

Public Sub screset_mnu_Click()
    NetClient.ResetScriptEngine
    StatusBar1.Panels(3).Text = " Script Enabled "
End Sub

Public Sub setup_mnu_Click()
    Servers.Show
End Sub


Private Sub Tab1_Click()
    'Call SendMessage(Tab1.hWnd, &H1300 + 51, Tab1.SelectedItem.Index - 1, ByVal 0)
    If Mid(Tab1.SelectedItem, 1, 1) = "#" And Not WINMIN Then
        NetClient.ChanWins.Item(Tab1.SelectedItem).WindowState = WINDOW_STATE
        NetClient.ChanWins.Item(Tab1.SelectedItem).Visible = True
        NetClient.ChanWins.Item(Tab1.SelectedItem).SetFocus
    ElseIf InStr(1, Tab1.SelectedItem, "Console") = 0 And Not WINMIN Then
        NetClient.MessWins.Item(Tab1.SelectedItem).WindowState = WINDOW_STATE
        NetClient.MessWins.Item(Tab1.SelectedItem).Visible = True
        NetClient.MessWins.Item(Tab1.SelectedItem).SetFocus
    ElseIf InStr(1, Tab1.SelectedItem, "Console") <> 0 And Not WINMIN Then
        WConsole.WindowState = WINDOW_STATE
        WConsole.Visible = True
        WConsole.SetFocus
    ElseIf NetClient.DccWins.count > 0 And Not WINMIN Then
        NetClient.DccWins.Item(Tab1.SelectedItem.key).WindowState = WINDOW_STATE
        NetClient.DccWins.Item(Tab1.SelectedItem.key).Visible = True
        NetClient.DccWins.Item(Tab1.SelectedItem.key).SetFocus
    End If
    WINMIN = False
End Sub

