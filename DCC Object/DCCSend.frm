VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form DCCSend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DCC Send"
   ClientHeight    =   1785
   ClientLeft      =   8250
   ClientTop       =   3045
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   1320
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1410
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "FILE"
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
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4455
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
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "DCCSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cdcc As CDCCFactory

