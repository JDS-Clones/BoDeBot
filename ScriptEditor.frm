VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form ScriptEditor 
   Caption         =   "Script Editor"
   ClientHeight    =   6750
   ClientLeft      =   2895
   ClientTop       =   2535
   ClientWidth     =   7770
   Icon            =   "ScriptEditor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   7770
   Begin RichTextLib.RichTextBox EditBox 
      Height          =   4695
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8281
      _Version        =   327681
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"ScriptEditor.frx":0442
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ScriptEditor.frx":053F
      Left            =   120
      List            =   "ScriptEditor.frx":0541
      TabIndex        =   1
      ToolTipText     =   "Select Function or Event to Edit"
      Top             =   600
      Width           =   7095
   End
   Begin ComctlLib.TabStrip Tab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11456
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Events"
            Key             =   "events"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Events"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "User-Defined Code"
            Key             =   "code"
            Object.Tag             =   ""
            Object.ToolTipText     =   "User-Defined Code"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu scr_edit 
      Caption         =   "Script Editor"
      NegotiatePosition=   3  'Right
      Begin VB.Menu scr_save 
         Caption         =   "Save"
      End
      Begin VB.Menu scr_close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "ScriptEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EventSub(16) As String
Private EventAct(16) As Boolean

Private CurrentIndex As Integer

Private GeneralSubs As String
Private GeneralAct As Boolean

Private PreviousTab As String

Private InitDone As Boolean
Private Sub Save()
    Dim i As Integer
    If GeneralAct = True Then
        GeneralSubs = EditBox.Text
    Else
        EventSub(CurrentIndex) = EditBox.Text
    End If
    
    Open "required.scp" For Output As #1
        For i = 0 To Combo1.ListCount - 1
            Print #1, Combo1.List(i) & vbCrLf
            Print #1, EventSub(i + 1)
            Print #1, "End Sub" & vbCrLf
        Next i
    Close #1
    
    Open "user.scp" For Output As #1
        Print #1, GeneralSubs
    Close #1
End Sub
Private Function ExamineLine(ByVal szLine As String) As String
    Dim param1$, param2$, buff$
    ExamineLine = ""
    buff$ = szLine
    param1$ = getNextToken(buff$, " ")
    param2$ = getNextToken(buff$, " ")
    Debug.Print "index=" & CurrentIndex & " -> " & szLine
    If LCase$(param1$) = "sub" And InStr(1, szLine, "(") > 0 Then
        Combo1.AddItem szLine
        CurrentIndex = CurrentIndex + 1
        EventAct(CurrentIndex) = False
    ElseIf LCase$(param1$) = "end" And LCase$(param2$) = "sub" Then
        'do nothing
    ElseIf Trim(szLine) = "" Or Trim(szLine) = vbCrLf Then
        'do nothing
    Else
        ExamineLine = szLine
        EventSub(CurrentIndex) = EventSub(CurrentIndex) & szLine & vbCrLf
    End If
End Function

Private Sub Combo1_Change()
    If InitDone = False Then Exit Sub
End Sub

Private Sub Combo1_Click()
    If InitDone = False Then Exit Sub
    EventAct(CurrentIndex) = False
    EventSub(CurrentIndex) = EditBox.Text
    EditBox.Text = EventSub(Combo1.ListIndex + 1)
    EventAct(Combo1.ListIndex + 1) = True
    CurrentIndex = Combo1.ListIndex + 1
End Sub

Private Sub EditBox_GotFocus()
    On Error Resume Next
    For Each Control In Controls
       Control.TabStop = False
    Next Control
End Sub

Private Sub EditBox_LostFocus()
    On Error Resume Next
    For Each Control In Controls
       Control.TabStop = True
    Next Control
End Sub

Private Sub Form_Load()
    InitDone = False
    CurrentIndex = 0
    Tab1.Top = Me.ScaleTop
    Tab1.Left = Me.ScaleLeft
    Tab1.Width = Me.ScaleWidth
    Tab1.Height = Me.ScaleHeight
    
    Combo1.Top = Tab1.ClientTop
    Combo1.Left = Tab1.ClientLeft
    Combo1.Width = Tab1.ClientWidth
    EditBox.Top = Tab1.ClientTop + Combo1.Top
    EditBox.Height = Tab1.ClientHeight - Combo1.Height
    EditBox.Left = Tab1.ClientLeft
    EditBox.Width = Tab1.ClientWidth
    
    Combo1.Visible = True
    Dim temp$
    EditBox.Text = ""
    Combo1.Clear
    Open App.Path & "\required.scp" For Input As #1
        Do While Not EOF(1)
            Line Input #1, temp$
            Call ExamineLine(temp$)
        Loop
    Close #1
    Combo1.Text = Combo1.List(0)
    EventAct(1) = True
    EditBox.Text = EventSub(1)
    CurrentIndex = 1
    Open App.Path & "\user.scp" For Input As #1
        Do While Not EOF(1)
            GeneralSubs = Input(LOF(1), 1)
        Loop
    Close #1
    PreviousTab = "events"
    InitDone = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Would you like to save any changes?", vbYesNo, "BoDeBoT: Script Editor") = 6 Then
        Save
    End If
    Erase EventSub
    Erase EventAct
End Sub

Private Sub Form_Resize()
    If WindowState <> 1 Then
        Tab1.Top = Me.ScaleTop
        Tab1.Left = Me.ScaleLeft
        Tab1.Width = Me.ScaleWidth
        Tab1.Height = Me.ScaleHeight
        Combo1.Left = Tab1.ClientLeft
        Combo1.Width = Tab1.ClientWidth
        Combo1.Top = Tab1.ClientTop
        EditBox.Top = Tab1.ClientTop + Combo1.Top
        EditBox.Height = Tab1.ClientHeight - Combo1.Height
        EditBox.Left = Tab1.ClientLeft
        EditBox.Width = Tab1.ClientWidth
        WINDOW_STATE = WindowState
    End If
End Sub

Private Sub scr_save_Click()
    Save
End Sub

Private Sub Tab1_Click()
    If Tab1.SelectedItem.key = "events" And PreviousTab = "code" Then
        Combo1.Visible = True
        GeneralSubs = EditBox.Text
        EditBox.Text = EventSub(CurrentIndex)
        GeneralAct = False
        PreviousTab = "events"
    ElseIf Tab1.SelectedItem.key = "code" And PreviousTab = "events" Then
        Combo1.Visible = False
        EventSub(CurrentIndex) = EditBox.Text
        EditBox.Text = GeneralSubs
        GeneralAct = True
        PreviousTab = "code"
    End If
End Sub
