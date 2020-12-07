VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form DCCGet 
   Caption         =   "DCC Get"
   ClientHeight    =   1620
   ClientLeft      =   2340
   ClientTop       =   3255
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   5760
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   840
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1245
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "PATH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "HOST"
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
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "DCCGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cdcc As CDCCFactory
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call closesocket(Cdcc.lpDCC_SOCKET)
End Sub




Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
