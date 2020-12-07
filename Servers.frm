VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form Servers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   5490
   ClientLeft      =   2520
   ClientTop       =   1605
   ClientWidth     =   5490
   Icon            =   "Servers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   51
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3120
      TabIndex        =   50
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame misc_fr 
      Height          =   1335
      Left            =   5040
      TabIndex        =   66
      Top             =   8400
      Width           =   4695
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   1200
         TabIndex        =   70
         Text            =   "Text27"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   960
         TabIndex        =   68
         Text            =   "Text26"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "Highlight Text"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Connection Retries"
         Height          =   495
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame reply_fr 
      Height          =   2295
      Left            =   5040
      TabIndex        =   59
      Top             =   5760
      Width           =   4695
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   1080
         TabIndex        =   104
         Text            =   "http://www.felmlee.com/bodebot"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   1080
         TabIndex        =   65
         Text            =   "Text25"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   1080
         TabIndex        =   63
         Text            =   "Text24"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   1080
         TabIndex        =   61
         Text            =   "Text23"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label23 
         Caption         =   "QUIT"
         Height          =   255
         Left            =   360
         TabIndex        =   103
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "PING"
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "FINGER"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "USERINFO"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame color_fr 
      Height          =   4695
      Left            =   6720
      TabIndex        =   20
      Top             =   360
      Width           =   4095
      Begin VB.TextBox Text43 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   102
         Text            =   "Text29"
         Top             =   1680
         Width           =   150
      End
      Begin VB.TextBox Text42 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "Text29"
         Top             =   1440
         Width           =   150
      End
      Begin VB.TextBox Text41 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "Text29"
         Top             =   1200
         Width           =   150
      End
      Begin VB.TextBox Text40 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "Text29"
         Top             =   960
         Width           =   150
      End
      Begin VB.TextBox Text39 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "Text29"
         Top             =   720
         Width           =   150
      End
      Begin VB.TextBox Text38 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "Text29"
         Top             =   480
         Width           =   150
      End
      Begin VB.TextBox Text37 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "Text29"
         Top             =   2640
         Width           =   150
      End
      Begin VB.TextBox Text36 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "Text29"
         Top             =   2400
         Width           =   150
      End
      Begin VB.TextBox Text35 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text29"
         Top             =   2160
         Width           =   150
      End
      Begin VB.TextBox Text34 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text29"
         Top             =   1920
         Width           =   150
      End
      Begin VB.TextBox Text33 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "Text29"
         Top             =   1680
         Width           =   150
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Text29"
         Top             =   1440
         Width           =   150
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text29"
         Top             =   1200
         Width           =   150
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text29"
         Top             =   960
         Width           =   150
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text29"
         Top             =   720
         Width           =   150
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Text28"
         Top             =   480
         Width           =   150
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1320
         TabIndex        =   44
         Text            =   "Text17"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   1320
         TabIndex        =   43
         Text            =   "Text16"
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1320
         TabIndex        =   42
         Text            =   "Text15"
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1320
         TabIndex        =   41
         Text            =   "Text14"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Text            =   "Text13"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         Text            =   "Text12"
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Text            =   "Text11"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Text            =   "Text10"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Text            =   "Text9"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1320
         TabIndex        =   35
         Text            =   "Text8"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Text            =   "Text7"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Text            =   "Text6"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label46 
         Caption         =   "15"
         Height          =   255
         Left            =   2400
         TabIndex        =   96
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label45 
         Caption         =   "14"
         Height          =   255
         Left            =   2400
         TabIndex        =   95
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label44 
         Caption         =   "13"
         Height          =   255
         Left            =   2400
         TabIndex        =   94
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label43 
         Caption         =   "12"
         Height          =   255
         Left            =   2400
         TabIndex        =   93
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label42 
         Caption         =   "11"
         Height          =   255
         Left            =   2400
         TabIndex        =   92
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label41 
         Caption         =   "10"
         Height          =   255
         Left            =   2400
         TabIndex        =   91
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label40 
         Caption         =   "9"
         Height          =   255
         Left            =   2040
         TabIndex        =   80
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label39 
         Caption         =   "8"
         Height          =   255
         Left            =   2040
         TabIndex        =   79
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "7"
         Height          =   255
         Left            =   2040
         TabIndex        =   78
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label37 
         Caption         =   "6"
         Height          =   255
         Left            =   2040
         TabIndex        =   77
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label36 
         Caption         =   "5"
         Height          =   255
         Left            =   2040
         TabIndex        =   76
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label35 
         Caption         =   "4"
         Height          =   255
         Left            =   2040
         TabIndex        =   75
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label34 
         Caption         =   "3"
         Height          =   255
         Left            =   2040
         TabIndex        =   74
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label33 
         Caption         =   "2"
         Height          =   255
         Left            =   2040
         TabIndex        =   73
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label32 
         Caption         =   "1"
         Height          =   255
         Left            =   2040
         TabIndex        =   72
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label31 
         Caption         =   "0"
         Height          =   255
         Left            =   2040
         TabIndex        =   71
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label20 
         Caption         =   "COLOR CODES"
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "NoticeColor"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "HighLightColor"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "MessageColor"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "TextBackColor"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "NickBackColor"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "NickForeColor"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "ActionColor"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "PartColor"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "JoinColor"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "CTCPColor"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "KickColor"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "ModeColor"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame dcc_fr 
      Height          =   2775
      Left            =   4320
      TabIndex        =   15
      Top             =   9840
      Width           =   4215
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   2520
         TabIndex        =   56
         Text            =   "Text22"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   1080
         TabIndex        =   54
         Text            =   "Text21"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto-Get"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   765
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label25 
         Caption         =   "Time Out"
         Height          =   255
         Left            =   1800
         TabIndex        =   55
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Packet Size"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Recieve Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame ident_fr 
      Height          =   1455
      Left            =   0
      TabIndex        =   10
      Top             =   9240
      Width           =   4095
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "System"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame gen_fr 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   4695
      Begin VB.CommandButton Command4 
         Caption         =   "Remove Server"
         Height          =   375
         Left            =   2160
         TabIndex        =   58
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Invisible"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   57
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Server"
         Height          =   375
         Left            =   840
         TabIndex        =   52
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   1200
         TabIndex        =   49
         Text            =   "Text19"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   1200
         TabIndex        =   47
         Text            =   "Text18"
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label22 
         Caption         =   "Alt NickName"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "NickName"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Email"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Real Name"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Username"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Server"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin ComctlLib.TabStrip Tab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9340
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   7
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "IDENTD"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "DCC"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Colors"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fonts"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Replies"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Servers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub fillFields()
    Text1.Text = MDI.NetClient.Username
    Text2.Text = MDI.NetClient.RealName
    Text3.Text = MDI.NetClient.Email
    Text18.Text = MDI.NetClient.NickName
    Text19.Text = MDI.NetClient.AltNickName
    Text4.Text = MDI.NetClient.IdentdName
    Text5.Text = MDI.NetClient.IdentdSystem
    Text6.Text = MDI.NetClient.CLR_MODE
    Text7.Text = MDI.NetClient.CLR_KICK
    Text8.Text = MDI.NetClient.CLR_CTCP
    Text9.Text = MDI.NetClient.CLR_JOIN
    Text10.Text = MDI.NetClient.CLR_PART
    Text11.Text = MDI.NetClient.CLR_ACTION
    Text12.Text = MDI.NetClient.CLR_NICK
    Text13.Text = MDI.NetClient.CLR_NICKBACK
    Text14.Text = MDI.NetClient.CLR_TEXTBACK
    Text15.Text = MDI.NetClient.CLR_MESSAGE
    Text16.Text = MDI.NetClient.CLR_HIGHLIGHT
    Text17.Text = MDI.NetClient.CLR_NOTICE
    'Text20.Text = DCCLISTENPORT
    Text21.Text = MDI.NetClient.DCC_ChunkSize
    Text22.Text = MDI.NetClient.DCC_TimeOut
    Check1.Value = MDI.NetClient.DCC_AutoGet
    Text23.Text = MDI.NetClient.UserInfo
    Text24.Text = MDI.NetClient.Finger
    Text25.Text = MDI.NetClient.PingReply
    Text20.Text = MDI.NetClient.QuitMsg
    Text26.Text = MDI.NetClient.Retries
    Text27.Text = MDI.NetClient.HighLight
    
    Text28.BackColor = QBColor(0)
    Text29.BackColor = QBColor(1)
    Text30.BackColor = QBColor(2)
    Text31.BackColor = QBColor(3)
    Text32.BackColor = QBColor(4)
    Text33.BackColor = QBColor(5)
    Text34.BackColor = QBColor(6)
    Text35.BackColor = QBColor(7)
    Text36.BackColor = QBColor(8)
    Text37.BackColor = QBColor(9)
    Text38.BackColor = QBColor(10)
    Text39.BackColor = QBColor(11)
    Text40.BackColor = QBColor(12)
    Text41.BackColor = QBColor(13)
    Text42.BackColor = QBColor(14)
    Text43.BackColor = QBColor(15)
    
    Combo1.Clear
    Dim TextLine As String
    Open App.Path & "\servers.ini" For Input As #1 ' Open file.
        Do While Not EOF(1) ' Loop until end of file.
            Line Input #1, TextLine ' Read line into variable.
            Combo1.AddItem getNextToken(TextLine, " ")
        Loop
    Close #1    ' Close file.
    Combo1.Text = Combo1.List(1)
End Sub

Private Sub applyChanges()
    'dump to ini from here
    'WritePrivateProfileString
    Dim retval As Long
    retval = WritePrivateProfileString("CLIENT", "USERNAME", Text1.Text, App.Path & "\bodebot2.ini")
'= REALNAME
    retval = WritePrivateProfileString("CLIENT", "REALNAME", Text2.Text, App.Path & "\bodebot2.ini")
'= EMAIL
    retval = WritePrivateProfileString("CLIENT", "EMAIL", Text3.Text, App.Path & "\bodebot2.ini")
'= NICKNAME
    retval = WritePrivateProfileString("CLIENT", "NICKNAME", Text18.Text, App.Path & "\bodebot2.ini")
'= ALTNICKNAME
    retval = WritePrivateProfileString("CLIENT", "ALTNICKNAME", Text19.Text, App.Path & "\bodebot2.ini")
'= USERINFO
    retval = WritePrivateProfileString("CLIENT", "USERINFO", Text23.Text, App.Path & "\bodebot2.ini")
    retval = WritePrivateProfileString("CLIENT", "QUIT", Text20.Text, App.Path & "\bodebot2.ini")
'= FINGER
    retval = WritePrivateProfileString("CLIENT", "FINGER", Text24.Text, App.Path & "\bodebot2.ini")
'= PING
    retval = WritePrivateProfileString("CLIENT", "PING", Text25.Text, App.Path & "\bodebot2.ini")
'= RETRIES
    retval = WritePrivateProfileString("CLIENT", "RETRIES", Text26.Text, App.Path & "\bodebot2.ini")
'= HIGHLIGHT
    retval = WritePrivateProfileString("CLIENT", "HIGHLIGHT", Text27.Text, App.Path & "\bodebot2.ini")
    
    retval = WritePrivateProfileString("CLIENT", "SERVERNAME", Combo1.Text, App.Path & "\bodebot2.ini")

'= IDENTNAME
    retval = WritePrivateProfileString("IDENTD", "IDENTNAME", Text4.Text, App.Path & "\bodebot2.ini")
'= IDENTSYS
    retval = WritePrivateProfileString("IDENTD", "IDENTSYS", Text5.Text, App.Path & "\bodebot2.ini")

'= MODECOLOR
    retval = WritePrivateProfileString("COLOR", "MODECOLOR", Text6.Text, App.Path & "\bodebot2.ini")
'= KICKCOLOR
    retval = WritePrivateProfileString("COLOR", "KICKCOLOR", Text7.Text, App.Path & "\bodebot2.ini")
'= CTCPCOLOR
    retval = WritePrivateProfileString("COLOR", "CTCPCOLOR", Text8.Text, App.Path & "\bodebot2.ini")
'= JOINCOLOR
    retval = WritePrivateProfileString("COLOR", "JOINCOLOR", Text9.Text, App.Path & "\bodebot2.ini")
'= PARTCOLOR
    retval = WritePrivateProfileString("COLOR", "PARTCOLOR", Text10.Text, App.Path & "\bodebot2.ini")
'= ACTIONCOLOR
    retval = WritePrivateProfileString("COLOR", "ACTIONCOLOR", Text11.Text, App.Path & "\bodebot2.ini")
'= NICKCOLOR
    retval = WritePrivateProfileString("COLOR", "NICKCOLOR", Text12.Text, App.Path & "\bodebot2.ini")
'= NICKBACKCOLOR
    retval = WritePrivateProfileString("COLOR", "NICKBACKCOLOR", Text13.Text, App.Path & "\bodebot2.ini")
'= TEXTBACKCOLOR
    retval = WritePrivateProfileString("COLOR", "TEXTBACKCOLOR", Text14.Text, App.Path & "\bodebot2.ini")
'= MESSAGECOLOR
    retval = WritePrivateProfileString("COLOR", "MESSAGECOLOR", Text15.Text, App.Path & "\bodebot2.ini")
'= HIGHLIGHTCOLOR
    retval = WritePrivateProfileString("COLOR", "HIGHLIGHTCOLOR", Text16.Text, App.Path & "\bodebot2.ini")
'= NOTICECOLOR
    retval = WritePrivateProfileString("COLOR", "NOTICECOLOR", Text17.Text, App.Path & "\bodebot2.ini")
    
'= DCCPACKETSIZE
    retval = WritePrivateProfileString("DCC", "DCCPACKETSIZE", Text21.Text, App.Path & "\bodebot2.ini")
'= DCCTIMEOUT
    retval = WritePrivateProfileString("DCC", "DCCTIMEOUT", Text22.Text, App.Path & "\bodebot2.ini")
'= DCCAUTOGET
    retval = WritePrivateProfileString("DCC", "DCCAUTOGET", CStr(CBool(Check1.Value)), App.Path & "\bodebot2.ini")
End Sub

Private Sub Command1_Click()
    applyChanges
    Me.Hide
    'GetStoredConfig
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Command3_Click()
    AddHost_dlg.Show
    fillFields
End Sub

Private Sub Command4_Click()
    MsgBox "Not Implemented Yet!"
End Sub

Private Sub Form_Load()
    Tab1.Top = Me.ScaleTop
    Tab1.Left = Me.ScaleLeft
    Tab1.Width = Me.ScaleWidth
    Tab1.Height = Me.ScaleHeight
    fillFields
    color_fr.Visible = False
    dcc_fr.Visible = False
    ident_fr.Visible = False
    reply_fr.Visible = False
    misc_fr.Visible = False
    gen_fr.Visible = True
    gen_fr.Top = Tab1.ClientTop
    gen_fr.Left = Tab1.ClientLeft
    gen_fr.Width = Tab1.ClientWidth
    gen_fr.Height = Tab1.ClientHeight
    
    
End Sub

Private Sub Form_Resize()
    If WindowState <> 1 Then
        Tab1.Top = Me.ScaleTop
        Tab1.Left = Me.ScaleLeft
        Tab1.Width = Me.ScaleWidth
        Tab1.Height = Me.ScaleHeight
    End If
End Sub

Private Sub Tab1_Click()
    If Tab1.SelectedItem = "General" Then
        color_fr.Visible = False
        dcc_fr.Visible = False
        ident_fr.Visible = False
        misc_fr.Visible = False
        reply_fr.Visible = False
        gen_fr.Visible = True
        gen_fr.Top = Tab1.ClientTop
        gen_fr.Left = Tab1.ClientLeft
        gen_fr.Width = Tab1.ClientWidth
        gen_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "IDENTD" Then
        color_fr.Visible = False
        dcc_fr.Visible = False
        gen_fr.Visible = False
        misc_fr.Visible = False
        reply_fr.Visible = False
        ident_fr.Visible = True
        ident_fr.Top = Tab1.ClientTop
        ident_fr.Left = Tab1.ClientLeft
        ident_fr.Width = Tab1.ClientWidth
        ident_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "DCC" Then
        color_fr.Visible = False
        gen_fr.Visible = False
        ident_fr.Visible = False
        misc_fr.Visible = False
        reply_fr.Visible = False
        dcc_fr.Visible = True
        dcc_fr.Top = Tab1.ClientTop
        dcc_fr.Left = Tab1.ClientLeft
        dcc_fr.Width = Tab1.ClientWidth
        dcc_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "Colors" Then
        gen_fr.Visible = False
        dcc_fr.Visible = False
        ident_fr.Visible = False
        misc_fr.Visible = False
        reply_fr.Visible = False
        color_fr.Visible = True
        color_fr.Top = Tab1.ClientTop
        color_fr.Left = Tab1.ClientLeft
        color_fr.Width = Tab1.ClientWidth
        color_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "Fonts" Then
        color_fr.Visible = False
        dcc_fr.Visible = False
        ident_fr.Visible = False
        gen_fr.Visible = False
        misc_fr.Visible = False
        reply_fr.Visible = False
'        gen_fr.Visible = True
'        gen_fr.Top = Tab1.ClientTop
'        gen_fr.Left = Tab1.ClientLeft
'        gen_fr.Width = Tab1.ClientWidth
'        gen_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "Replies" Then
        color_fr.Visible = False
        dcc_fr.Visible = False
        ident_fr.Visible = False
        gen_fr.Visible = False
        misc_fr.Visible = False
        reply_fr.Visible = True
        reply_fr.Top = Tab1.ClientTop
        reply_fr.Left = Tab1.ClientLeft
        reply_fr.Width = Tab1.ClientWidth
        reply_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "Misc" Then
        color_fr.Visible = False
        dcc_fr.Visible = False
        ident_fr.Visible = False
        gen_fr.Visible = False
        reply_fr.Visible = False
        misc_fr.Visible = True
        misc_fr.Top = Tab1.ClientTop
        misc_fr.Left = Tab1.ClientLeft
        misc_fr.Width = Tab1.ClientWidth
        misc_fr.Height = Tab1.ClientHeight
    End If
End Sub
