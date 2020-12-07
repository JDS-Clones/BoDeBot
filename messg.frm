VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form messg 
   Caption         =   "message"
   ClientHeight    =   3375
   ClientLeft      =   10500
   ClientTop       =   2730
   ClientWidth     =   6540
   Icon            =   "messg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3375
   ScaleWidth      =   6540
   Begin RichTextLib.RichTextBox msendrtb 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   327681
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"messg.frx":0442
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
   Begin RichTextLib.RichTextBox messgrtb 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3413
      _Version        =   327681
      BackColor       =   16776960
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   5000.28
      TextRTF         =   $"messg.frx":0548
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
End
Attribute VB_Name = "messg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Stack As CStringStack

Private Sub Form_GotFocus()
    Dim x As Integer
    x = MDI.Tab1.Tabs.Item(Me.Tag).Index - 1
    Call SendMessage(MDI.Tab1.hWnd, &H1300 + 51, x, ByVal 0)
End Sub

Private Sub Form_Load()
''========================================================
''
''  Code Copyright (c) 1997 Nicholas J. Felmlee
''  http://www.felmlee.com/bodebot
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

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDI.NetClient.MessWins.Remove Me.Caption
    MDI.Tab1.Tabs.Remove Me.Caption
End Sub


Private Sub Form_Resize()
If WindowState <> 1 And ScaleHeight <> 0 Then
    If ScaleTop - 50 > 0 Then
        messgrtb.Top = ScaleTop - 50
    Else
        messgrtb.Top = 0
    End If
    If ScaleHeight - 255 > 0 Then
        messgrtb.Height = ScaleHeight - 255
    Else
        messgrtb.Height = 0
    End If
    messgrtb.Width = ScaleWidth

    msendrtb.Top = messgrtb.Top + messgrtb.Height
    msendrtb.Width = messgrtb.Width
    msendrtb.Height = 375
    If messgrtb.RightMargin = messgrtb.Width - 200 > 0 Then
        messgrtb.RightMargin = messgrtb.Width - 200
    Else
        messgrtb.RightMargin = 0
    End If
    WINDOW_STATE = WindowState
ElseIf WindowState = 1 Then
    WINMIN = True
    Me.Visible = False
End If
End Sub


Private Sub messgrtb_GotFocus()
    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
    Dim x As Integer
    x = MDI.Tab1.Tabs.Item(Me.Tag).Index - 1
    Call SendMessage(MDI.Tab1.hWnd, &H1300 + 51, x, ByVal 0)
End Sub

Private Sub msendrtb_GotFocus()
    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
    Dim x As Integer
    x = MDI.Tab1.Tabs.Item(Me.Tag).Index - 1
    Call SendMessage(MDI.Tab1.hWnd, &H1300 + 51, x, ByVal 0)
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
        UpdateRTB Me.messgrtb, "[[" & MDI.NetClient.NickName & "]]   " & msendrtb.Text, MDI.NetClient.CLR_MESSAGE
        MDI.NetClient.ProcessUserCommands Me.messgrtb, msendrtb.Text, Me.Caption
        Stack.Push msendrtb.Text
        msendrtb.Text = ""
    End If
End Sub


