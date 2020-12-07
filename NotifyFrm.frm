VERSION 5.00
Begin VB.Form NotifyFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Notify"
   ClientHeight    =   3690
   ClientLeft      =   12585
   ClientTop       =   2415
   ClientWidth     =   2175
   Icon            =   "NotifyFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox NotifyList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      ItemData        =   "NotifyFrm.frx":0442
      Left            =   120
      List            =   "NotifyFrm.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "NotifyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private setup As Boolean
Private Sub Form_Load()
    setup = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDI.Mainbar.Buttons(12).Value = tbrUnpressed
End Sub

Private Sub Form_Resize()
    If WindowState <> 1 And ScaleHeight <> 0 Then
        NotifyList.Top = Me.ScaleTop
        NotifyList.Left = Me.ScaleLeft
        NotifyList.Width = Me.ScaleWidth
        NotifyList.Height = Me.ScaleHeight
    End If
End Sub

Public Sub ShowNotify()
    Me.Show
    If setup = True Then
        Me.Height = 1440
        Me.Width = 1440
        NotifyList.Top = Me.ScaleTop
        NotifyList.Left = Me.ScaleLeft
        NotifyList.Width = Me.ScaleWidth
        NotifyList.Height = Me.ScaleHeight
    End If
    setup = False
End Sub
