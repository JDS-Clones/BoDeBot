VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form chan 
   Caption         =   "channel window"
   ClientHeight    =   2910
   ClientLeft      =   4005
   ClientTop       =   9945
   ClientWidth     =   7785
   Icon            =   "chan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   7785
   Begin ComctlLib.ListView nicklist 
      Height          =   2175
      Left            =   5400
      TabIndex        =   2
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3836
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16744448
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1835
      EndProperty
   End
   Begin RichTextLib.RichTextBox sendrtb 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   327681
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"chan.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox chanrtb 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3201
      _Version        =   327681
      BackColor       =   16776960
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   5400
      TextRTF         =   $"chan.frx":0548
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu file_mnu 
      Caption         =   "&File"
      Begin VB.Menu connect_mnu 
         Caption         =   "&Connect"
      End
      Begin VB.Menu disconect_mnu 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu setup_mnu 
         Caption         =   "&Setup"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu exit_mnu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu msnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu scripts_mnu 
         Caption         =   "&Scripts"
         Begin VB.Menu edit_mnu 
            Caption         =   "Edit"
         End
         Begin VB.Menu reset_mnu 
            Caption         =   "Reset"
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close &All"
      End
   End
   Begin VB.Menu mnu_nicklist 
      Caption         =   "NickNames"
      Visible         =   0   'False
      Begin VB.Menu opstuff_sbmnu 
         Caption         =   "OP Stuff"
         Begin VB.Menu op_sbmnu 
            Caption         =   "OP"
         End
         Begin VB.Menu deop_sbmnu 
            Caption         =   "Deop"
         End
         Begin VB.Menu kick_sbmnu 
            Caption         =   "Kick!"
         End
         Begin VB.Menu ban_sbmnu 
            Caption         =   "Ban!"
         End
      End
      Begin VB.Menu whois_sbmnu 
         Caption         =   "Whois"
      End
      Begin VB.Menu away_sbmnu 
         Caption         =   "Away"
      End
      Begin VB.Menu dcc_sbmnu 
         Caption         =   "DCC Stuff"
         Begin VB.Menu dccsnd_mnu 
            Caption         =   "DCC send"
         End
         Begin VB.Menu dcccht_sbmnu 
            Caption         =   "DCC chat"
         End
      End
   End
   Begin VB.Menu help_mnu 
      Caption         =   "Help"
      Begin VB.Menu contents_mnu 
         Caption         =   "Contents"
      End
      Begin VB.Menu about_mnu 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "chan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Stack As CStringStack
Private Sub about_mnu_Click()
    frmSplash.Show
End Sub

Private Sub ban_sbmnu_Click()
    Dim mnuline As String
    Dim wnick As String
    
    wnick = nicklist.SelectedItem
    If Mid(nicklist.SelectedItem, 1, 1) = "@" Then
        wnick = Mid(nicklist.SelectedItem, 2, Len(nicklist.SelectedItem) - 1)
    End If
    mnuline = "/mode " & Me.Tag & " +b " & wnick
    
    MDI.NetClient.ProcessUserCommands Me.chanrtb, mnuline, ""

End Sub


Private Sub chanrtb_GotFocus()
    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
    Dim x As Integer
    x = MDI.Tab1.Tabs.Item(Me.Tag).Index - 1
    Call SendMessage(MDI.Tab1.hWnd, &H1300 + 51, x, ByVal 0)
End Sub


Private Sub connect_mnu_Click()
    MDI.connect_mnu_Click
End Sub

Private Sub dcccht_sbmnu_Click()
    Dim mnuline As String
    Dim wnick As String
    
    wnick = nicklist.SelectedItem
    If (Mid(nicklist.SelectedItem, 1, 1) = "@") Or (Mid(nicklist.SelectedItem, 1, 1) = "+") Then
    wnick = Mid(nicklist.SelectedItem, 2, Len(nicklist.SelectedItem) - 1)
    End If
    mnuline = "/chat " & wnick
    
    MDI.NetClient.ProcessUserCommands Me.chanrtb, mnuline, ""
End Sub

Private Sub dccsnd_mnu_Click()
    Dim mnuline As String
    Dim wnick As String
    
    wnick = nicklist.SelectedItem
    If (Mid(nicklist.SelectedItem, 1, 1) = "@") Or (Mid(nicklist.SelectedItem, 1, 1) = "+") Then
    wnick = Mid(nicklist.SelectedItem, 2, Len(nicklist.SelectedItem) - 1)
    End If
    mnuline = "/dcc " & wnick
    
    MDI.NetClient.ProcessUserCommands Me.chanrtb, mnuline, ""
End Sub

Private Sub deop_sbmnu_Click()
    Dim mnuline As String
    Dim wnick As String
    wnick = nicklist.SelectedItem
    If (Mid(nicklist.SelectedItem, 1, 1) = "@") Or (Mid(nicklist.SelectedItem, 1, 1) = "+") Then
        wnick = Mid(nicklist.SelectedItem, 2, Len(nicklist.SelectedItem) - 1)
    End If
    mnuline = "/mode " & Me.Tag & " -o " & wnick
    MDI.NetClient.ProcessUserCommands Me.chanrtb, mnuline, ""

End Sub

Private Sub disconect_mnu_Click()
    MDI.disconnect_mnu_Click
End Sub

Private Sub edit_mnu_Click()
    MDI.scedit_mnu_Click
End Sub

Private Sub exit_mnu_Click()
    MDI.exit_mnu_Click
End Sub

Private Sub Form_GotFocus()
    Dim x As Integer
    x = MDI.Tab1.Tabs.Item(Me.Tag).Index - 1
    Call SendMessage(MDI.Tab1.hWnd, &H1300 + 51, x, ByVal 0)
End Sub

Private Sub Form_Load()
''========================================================
''
''  Code Copyright (c) 1998 Nicholas J. Felmlee
''  http://www.felmlee.com/bodebot
''========================================================
    Set Stack = New CStringStack
    Stack.MallocStack 10
    Dim RTBCODE$
    RTBCODE$ = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman Times New Roman;}}{\colortbl\red0\green0\blue0;\red255\green0\blue0;}"
    chanrtb.TextRTF = RTBCODE$ & "\par }"
    chanrtb.Top = ScaleTop - 50
    chanrtb.Height = ScaleHeight - 275
    chanrtb.Width = ScaleWidth - 1635
    
    sendrtb.Top = chanrtb.Top + chanrtb.Height
    sendrtb.Width = chanrtb.Width
    sendrtb.Height = 375
    nicklist.Left = chanrtb.Left + chanrtb.Width
    nicklist.Top = ScaleTop - 50
    nicklist.Height = chanrtb.Height + sendrtb.Height + 40
    chanrtb.BackColor = QBColor(MDI.NetClient.CLR_TEXTBACK)
    nicklist.BackColor = QBColor(MDI.NetClient.CLR_NICKBACK)
    nicklist.ForeColor = QBColor(MDI.NetClient.CLR_NICK)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDI.NetClient.Send "PART :" & Me.Tag & vbCrLf
    MDI.NetClient.ChanWins.Remove Me.Tag
    MDI.Tab1.Tabs.Remove Me.Tag
End Sub

Private Sub Form_Resize()
If WindowState <> 1 And ScaleHeight <> 0 Then
    If ScaleTop - 50 > 0 Then
        chanrtb.Top = ScaleTop - 50
    Else
        chanrtb.Top = 0
    End If
    
    If ScaleHeight - 275 > 0 Then
        chanrtb.Height = ScaleHeight - 275
    Else
        chanrtb.Height = 0
    End If
    
    If ScaleWidth - 1635 > 0 Then
        chanrtb.Width = ScaleWidth - 1635
    Else
        chanrtb.Width = 0
    End If
    
    sendrtb.Top = chanrtb.Top + chanrtb.Height
    sendrtb.Width = chanrtb.Width
    sendrtb.Height = 375
    nicklist.Left = chanrtb.Left + chanrtb.Width
    
    If ScaleTop - 50 > 0 Then
        nicklist.Top = ScaleTop - 50
    Else
        nicklist.Top = 0
    End If
    
    nicklist.Height = chanrtb.Height + sendrtb.Height + 40
    
    If chanrtb.Width - 200 > 0 Then
        chanrtb.RightMargin = chanrtb.Width - 200
    Else
        chanrtb.RightMargin = 0
    End If
    WINDOW_STATE = WindowState
    
ElseIf WindowState = 1 Then
    WINMIN = True
    Me.Visible = False
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'SOCKET.Send "PART :" & Me.Tag & vbCrLf
End Sub

Private Sub kick_sbmnu_Click()
    Dim temp As String
    temp = Me.nicklist.SelectedItem
    If (InStr(1, Me.nicklist.SelectedItem, "@") <> 0 Or InStr(1, Me.nicklist.SelectedItem, "+") <> 0) Then
        temp = Mid(Me.nicklist.SelectedItem, 2)
    End If
    MDI.NetClient.Send "KICK " & Me.Tag & " " & temp & " :The lunatic is in my head!" & vbCrLf
End Sub

Private Sub nicklist_GotFocus()
    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
    Dim x As Integer
    x = MDI.Tab1.Tabs.Item(Me.Tag).Index - 1
    Call SendMessage(MDI.Tab1.hWnd, &H1300 + 51, x, ByVal 0)
End Sub

Private Sub nicklist_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnu_nicklist
    End If

End Sub

Private Sub op_sbmnu_Click()
    Dim mnuline As String
    Dim wnick As String
    
    wnick = nicklist.SelectedItem
    If (Mid(nicklist.SelectedItem, 1, 1) = "@") Or (Mid(nicklist.SelectedItem, 1, 1) = "+") Then
    wnick = Mid(nicklist.SelectedItem, 2, Len(nicklist.SelectedItem) - 1)
    End If
    mnuline = "/mode " & Me.Tag & " +o " & wnick
    
    MDI.NetClient.ProcessUserCommands Me.chanrtb, mnuline, ""

End Sub

Private Sub reset_mnu_Click()
    MDI.screset_mnu_Click
End Sub

Private Sub sendrtb_GotFocus()
    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
    Dim x As Integer
    x = MDI.Tab1.Tabs.Item(Me.Tag).Index - 1
    Call SendMessage(MDI.Tab1.hWnd, &H1300 + 51, x, ByVal 0)
End Sub

Private Sub sendrtb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
            sendrtb.Text = Stack.Pop
            sendrtb.SelStart = Len(sendrtb.Text)
    ElseIf KeyCode = vbKeyUp Then
            sendrtb.Text = Stack.UnPop
            sendrtb.SelStart = Len(sendrtb.Text)
    End If
End Sub

Private Sub sendrtb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Mid(sendrtb.Text, 1, 1) = "/" Then
            UpdateRTB chanrtb, "[[" & MDI.NetClient.NickName & "]]   " & sendrtb.Text, MDI.NetClient.CLR_MESSAGE
        End If
        MDI.NetClient.ProcessUserCommands Me.chanrtb, sendrtb.Text, Me.Tag
        Stack.Push sendrtb.Text
        sendrtb.Text = ""
    End If
End Sub



Private Sub setup_mnu_Click()
    Servers.Show
End Sub

Private Sub whois_sbmnu_Click()
    Dim wnick As String
    wnick = nicklist.SelectedItem
    If Mid(nicklist.SelectedItem, 1, 1) = "@" Or Mid(nicklist.SelectedItem, 1, 1) = "+" Then
        wnick = Mid(nicklist.SelectedItem, 2, Len(nicklist.SelectedItem) - 1)
    End If
    MDI.NetClient.Send "WHOIS :" & wnick & vbCrLf
    
End Sub

Public Sub UpdateMeFromScript(ByVal szLine As String)
    UpdateRTB Me.chanrtb, szLine, MDI.NetClient.CLR_CTCP
End Sub
