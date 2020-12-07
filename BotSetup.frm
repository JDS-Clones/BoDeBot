VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form BotSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bot Setup"
   ClientHeight    =   5910
   ClientLeft      =   2910
   ClientTop       =   1965
   ClientWidth     =   5310
   Icon            =   "BotSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   40
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame bans_fr 
      Height          =   2415
      Left            =   5640
      TabIndex        =   35
      Top             =   4680
      Width           =   5175
      Begin VB.CommandButton Command13 
         Caption         =   "Edit"
         Height          =   375
         Left            =   2160
         TabIndex        =   39
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1200
         TabIndex        =   38
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Add"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame chanprof_fr 
      Height          =   2895
      Left            =   5400
      TabIndex        =   30
      Top             =   7440
      Width           =   4935
      Begin VB.CommandButton Command10 
         Caption         =   "Edit"
         Height          =   375
         Left            =   2280
         TabIndex        =   34
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   2280
         Width           =   855
      End
      Begin ComctlLib.ListView ListView2 
         Height          =   1815
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "channel"
            Object.Tag             =   ""
            Text            =   "#Channel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   "mode"
            Object.Tag             =   ""
            Text            =   "Mode"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame vocabman_fr 
      Height          =   1695
      Left            =   4440
      TabIndex        =   21
      Top             =   10440
      Width           =   4935
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2760
         TabIndex        =   29
         Text            =   "#channel"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Relay to Channel"
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3960
         TabIndex        =   27
         Text            =   "1200"
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "User Defined (vocab.txt)"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "WebCrawler"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TAD"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Delay (seconds)"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Vocab Sources:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame userman_fr 
      Height          =   3375
      Left            =   240
      TabIndex        =   10
      Top             =   9240
      Width           =   3855
      Begin VB.CommandButton Command7 
         Caption         =   "Apply Change"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   855
      End
      Begin VB.ListBox UList 
         Height          =   2985
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Access Level"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Pass"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame status_fr 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   4935
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Text            =   "7788"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "100"
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Alert"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Kill"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Bot"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "user"
            Object.Tag             =   ""
            Text            =   "User"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   "action"
            Object.Tag             =   ""
            Text            =   "Last Action Performed"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Max Connections"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1335
      End
   End
   Begin ComctlLib.TabStrip Tab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   10398
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Status"
            Key             =   "status"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "User Manager"
            Key             =   "usermanager"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Vocab Manager"
            Key             =   "vocabmanager"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Channel Profiles"
            Key             =   "channelprofiles"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Bans"
            Key             =   "bans"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "BotSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Tab1.Top = Me.ScaleTop
    Tab1.Left = Me.ScaleLeft
    Tab1.Width = Me.ScaleWidth
    Tab1.Height = Me.ScaleHeight
'    fillFields
    userman_fr.Visible = False
    vocabman_fr.Visible = False
    chanprof_fr.Visible = False
    bans_fr.Visible = False
'    misc_fr.Visible = False
    status_fr.Visible = True
    status_fr.Top = Tab1.ClientTop
    status_fr.Left = Tab1.ClientLeft
    status_fr.Width = Tab1.ClientWidth
    status_fr.Height = Tab1.ClientHeight
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
    If Tab1.SelectedItem = "Status" Then
        userman_fr.Visible = False
        vocabman_fr.Visible = False
        chanprof_fr.Visible = False
        bans_fr.Visible = False
'        reply_fr.Visible = False
        status_fr.Visible = True
        status_fr.Top = Tab1.ClientTop
        status_fr.Left = Tab1.ClientLeft
        status_fr.Width = Tab1.ClientWidth
        status_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "Channel Profiles" Then
        status_fr.Visible = False
        userman_fr.Visible = False
        vocabman_fr.Visible = False
        bans_fr.Visible = False
'        reply_fr.Visible = False
        chanprof_fr.Visible = True
        chanprof_fr.Top = Tab1.ClientTop
        chanprof_fr.Left = Tab1.ClientLeft
        chanprof_fr.Width = Tab1.ClientWidth
        chanprof_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "Vocab Manager" Then
        userman_fr.Visible = False
        status_fr.Visible = False
        chanprof_fr.Visible = False
        bans_fr.Visible = False
'        reply_fr.Visible = False
        vocabman_fr.Visible = True
        vocabman_fr.Top = Tab1.ClientTop
        vocabman_fr.Left = Tab1.ClientLeft
        vocabman_fr.Width = Tab1.ClientWidth
        vocabman_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "User Manager" Then
        status_fr.Visible = False
        vocabman_fr.Visible = False
        chanprof_fr.Visible = False
        bans_fr.Visible = False
'        reply_fr.Visible = False
        userman_fr.Visible = True
        userman_fr.Top = Tab1.ClientTop
        userman_fr.Left = Tab1.ClientLeft
        userman_fr.Width = Tab1.ClientWidth
        userman_fr.Height = Tab1.ClientHeight
'    ElseIf Tab1.SelectedItem = "Fonts" Then
'        color_fr.Visible = False
'        dcc_fr.Visible = False
'        ident_fr.Visible = False
'        gen_fr.Visible = False
'        misc_fr.Visible = False
'        reply_fr.Visible = False
''        gen_fr.Visible = True
''        gen_fr.Top = Tab1.ClientTop
''        gen_fr.Left = Tab1.ClientLeft
''        gen_fr.Width = Tab1.ClientWidth
''        gen_fr.Height = Tab1.ClientHeight
'    ElseIf Tab1.SelectedItem = "Replies" Then
'        color_fr.Visible = False
'        dcc_fr.Visible = False
'        ident_fr.Visible = False
'        gen_fr.Visible = False
'        misc_fr.Visible = False
'        reply_fr.Visible = True
'        reply_fr.Top = Tab1.ClientTop
'        reply_fr.Left = Tab1.ClientLeft
'        reply_fr.Width = Tab1.ClientWidth
'        reply_fr.Height = Tab1.ClientHeight
    ElseIf Tab1.SelectedItem = "Bans" Then
        status_fr.Visible = False
        userman_fr.Visible = False
        vocabman_fr.Visible = False
        chanprof_fr.Visible = False
        'reply_fr.Visible = False
        bans_fr.Visible = True
        bans_fr.Top = Tab1.ClientTop
        bans_fr.Left = Tab1.ClientLeft
        bans_fr.Width = Tab1.ClientWidth
        bans_fr.Height = Tab1.ClientHeight
    End If
End Sub



