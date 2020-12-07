VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form WConsole 
   Caption         =   "Console"
   ClientHeight    =   4500
   ClientLeft      =   3225
   ClientTop       =   3900
   ClientWidth     =   6390
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   6390
   Begin RichTextLib.RichTextBox msendrtb 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   327681
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0442
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5953
      _Version        =   327681
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0537
   End
End
Attribute VB_Name = "WConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Stack As CStringStack
'formload
'keydown
Private Sub Form_Load()
    Set Stack = New CStringStack
    Stack.MallocStack 10
    RTB.Top = ScaleTop - 50
    RTB.Height = ScaleHeight - 255
    RTB.Width = ScaleWidth
    
    msendrtb.Top = RTB.Top + RTB.Height
    msendrtb.Width = RTB.Width
    msendrtb.Height = 375
    RTB.BackColor = QBColor(MDI.NetClient.CLR_TEXTBACK)
    Me.Tag = "Console"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set Stack = Nothing
End Sub

Private Sub Form_Resize()
    If WindowState <> 1 And ScaleHeight <> 0 Then
        RTB.Top = ScaleTop - 50
        RTB.Height = ScaleHeight - 255
        RTB.Width = ScaleWidth
        msendrtb.Top = RTB.Top + RTB.Height
        msendrtb.Width = RTB.Width
        msendrtb.Height = 375
        RTB.RightMargin = RTB.Width - 200
        WINDOW_STATE = WindowState
    ElseIf WindowState = 1 Then
        WINMIN = True
        Me.Visible = False
    End If

End Sub

Private Sub msendrtb_GotFocus()
    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
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
        'BBotCommand Me.RTB, cmline, msendrtb.Text, tot, res
        MDI.NetClient.ProcessUserCommands Me.RTB, msendrtb.Text, ""
        Stack.Push msendrtb.Text
        msendrtb.Text = ""
    End If
End Sub

Private Sub RTB_GotFocus()
    MDI.Tab1.SelectedItem = MDI.Tab1.Tabs(Me.Tag)
End Sub
