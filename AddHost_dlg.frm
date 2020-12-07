VERSION 5.00
Begin VB.Form AddHost_dlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Server"
   ClientHeight    =   1290
   ClientLeft      =   8025
   ClientTop       =   3135
   ClientWidth     =   5010
   Icon            =   "AddHost_dlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "AddHost_dlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open "servers.ini" For Append As #1
        Print #1, Trim(Text1.Text) & " " & Trim(Text2.Text)
    Close #1
    Servers.Combo1.AddItem Trim(Text1.Text)
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub
