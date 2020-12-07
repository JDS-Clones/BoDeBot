VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form DCCChat 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   3060
   ClientTop       =   2760
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TextRTF         =   $"DCCChat.frx":0000
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
      TextRTF         =   $"DCCChat.frx":00F5
   End
End
Attribute VB_Name = "DCCChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Stack As CStringStack
Public Cdcc As CDCCFactory


Private Sub Form_GotFocus()
    'MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
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
    messgrtb.BackColor = QBColor(TEXTBACKCOLOR)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call closesocket(CHATSOCKET)
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
ElseIf WindowState = 1 Then
    WINMIN = True
    Me.Visible = False
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
    ElseIf KeyCode = vbKeyUp Then
            msendrtb.Text = Stack.UnPop
    End If
End Sub

Private Sub msendrtb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Mid(msendrtb.Text, 1, 1) = "/" Then
            Cdcc.RaiseBoxInput Me.messgrtb, "[[" & NickName & "]]   " & msendrtb.Text
            Call SendData(Cdcc.lpDCC_SOCKET, msendrtb.Text & vbCrLf)
        Else
            Cdcc.RaiseUserCommand messgrtb, msendrtb.Text
        End If
        Stack.Push msendrtb.Text
        msendrtb.Text = ""
    End If
End Sub

