VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form DCCSend 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DCC Send"
   ClientHeight    =   1455
   ClientLeft      =   8250
   ClientTop       =   3045
   ClientWidth     =   4890
   Icon            =   "DCCSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1455
   ScaleWidth      =   4890
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3600
      Picture         =   "DCCSend.frx":0442
      ScaleHeight     =   975
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   1320
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1e18
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "DCCSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Alive As Boolean

Private Sub Form_Load()
    Alive = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Alive = True Then
        MDI.NetClient.KillDcc Me.hWnd
        Alive = False
    End If
End Sub

Private Sub Form_Terminate()
    If Alive = True Then
        MDI.NetClient.KillDcc Me.hWnd
        Alive = False
    End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()
    MDI.NetClient.KillDcc Me.hWnd
    Me.Caption = Me.Caption & " [DCC Timed Out!]"
    Alive = False
    Timer1.Enabled = False
End Sub
