VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form DCCChat 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   5835
   ClientTop       =   7290
   ClientWidth     =   6720
   Icon            =   "DCCChat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   6720
   Begin RichTextLib.RichTextBox msendrtb 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      _Version        =   327681
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"DCCChat.frx":0442
   End
   Begin RichTextLib.RichTextBox messgrtb 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4048
      _Version        =   327681
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"DCCChat.frx":0537
   End
End
Attribute VB_Name = "DCCChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Stack As CStringStack

Private Sub Form_GotFocus()
'    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(CStr(Me.Tag))
End Sub

Private Sub Form_Load()
''========================================================
''
''  Code Copyright (c) 1997 Nicholas J. Felmlee
''  http://www.felmlee.com/grave
''========================================================
    Set Stack = New CStringStack
    Stack.MallocStack 10
    messgrtb.Top = ScaleTop - 50
    messgrtb.Height = ScaleHeight - 255
    messgrtb.Width = ScaleWidth
    
    msendrtb.Top = messgrtb.Top + messgrtb.Height
    msendrtb.Width = messgrtb.Width
    msendrtb.Height = 375
    messgrtb.BackColor = QBColor(MDI.NetClient.CLR_TEXTBACK)
    'index these in the tab list by the hWnd
    Dim szKey As String
    szKey = CStr(Me.hWnd)
    Me.Tag = szKey
'   MDI.Tab1.Tabs.Add , CStr(Me.Tag), Me.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call closesocket(CHATSOCKET)
    MDI.NetClient.KillDcc Me.hWnd
    'MDI.Tab1.Tabs.Remove Me.Tag
    MDI.NetClient.DccWins.Remove Me.Tag
End Sub

Private Sub Form_Resize()
If WindowState <> 1 And ScaleHeight <> 0 Then
    messgrtb.Top = ScaleTop - 50
    messgrtb.Height = ScaleHeight - 255
    messgrtb.Width = ScaleWidth

    msendrtb.Top = messgrtb.Top + messgrtb.Height
    msendrtb.Width = messgrtb.Width
    msendrtb.Height = 375
    messgrtb.RightMargin = messgrtb.Width - 200
    WINDOW_STATE = WindowState
'ElseIf WindowState = 1 Then
'    WINMIN = True
'    Me.Visible = False
End If
End Sub


Private Sub messgrtb_GotFocus()
    'MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
End Sub

Private Sub msendrtb_GotFocus()
    'MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
End Sub

Private Sub msendrtb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
            msendrtb.Text = Stack.Pop
            msendrtb.SelStart = Len(msendrtb.Text)
    ElseIf KeyCode = vbKeyUp Then
            msendrtb.Text = Stack.UnPop
            msendrtb.SelStart = Len(msendrtb.Text)
    End If
End Sub

Private Sub msendrtb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Mid(msendrtb.Text, 1, 1) = "/" Then
            UpdateRTB Me.messgrtb, "[[" & MDI.NetClient.NickName & "]]   " & msendrtb.Text, MDI.NetClient.CLR_MESSAGE
            MDI.NetClient.SendChatText msendrtb.Text & vbCrLf, Me.hWnd
        Else
            MDI.NetClient.ProcessUserCommands messgrtb, msendrtb.Text
        End If
        Stack.Push msendrtb.Text
        msendrtb.Text = ""
    End If
End Sub

