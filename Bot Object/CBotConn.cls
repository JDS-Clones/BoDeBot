VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBotConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarUser As String
Private mvarStatus As Integer
Private mvarLevel As String
Private mvarLastCommand As String
Private mvarConnID As Long
'normal vars
Private Perms() As String
Private PermSize As Integer
Public Property Let ConnID(ByVal vData As Long)
    mvarConnID = vData
End Property


Public Property Get ConnID() As Long
    ConnID = mvarConnID
End Property

Public Property Let LastCommand(ByVal vData As String)
    mvarLastCommand = vData
End Property


Public Property Get LastCommand() As String
    LastCommand = mvarLastCommand
End Property

Public Property Let Level(vData As String)
    mvarLevel = vData
End Property

Public Property Get Level() As String
    Level = mvarLevel
End Property

Public Property Let Status(ByVal vData As Integer)
    mvarStatus = vData
End Property

Public Property Get Status() As Integer
    Status = mvarStatus
End Property

Public Property Let User(ByVal vData As String)
    mvarUser = vData
End Property

Public Property Get User() As String
    User = mvarUser
End Property


Private Sub Push(ByVal szPerm As String)
'ADDS A PERMS TO OUR LIST OF PERMS FOR THIS
'USER.
    PermSize = PermSize + 1
    ReDim Preserve Perms(1 To PermSize)
    Perms(PermSize) = szPerm
End Sub

Private Function Pop() As String
'REMOVES A PERM FROM OUR PERM LIST
    Pop = Perms(PermSize)
    PermSize = PermSize - 1
    ReDim Preserve Perms(1 To PermSize)
End Function

Public Function CheckChanPerms(ByVal szChan As String, ByVal Level As Integer) As Boolean
'TAKES A CHANNEL NAME AS INPUT AND THE REQUIRED LEVEL
'FOR THE ACTION. IF THE USER MEETS THE REQUIRED CHANNEL
'AND LEVEL REQUIREMENT, THIS FUNCTION WILL RETURN FALSE.
'OTHERWISE WE RETURN TRUE.
    Dim temp As String, i As Integer
    CheckChanPerms = True
    For i = 1 To PermSize
        temp = Mid(Perms(i), 3)
        If (StrComp(LCase$(temp), LCase$(szChan), vbBinaryCompare) = 0) And (CInt(Mid(Perms(i), 1, 1)) >= Level) Then
            'okie everything checks out break and return true
            CheckChanPerms = False
            Exit Function
        End If
    Next i
End Function

Public Function CheckBotPerms(ByVal Level As Integer) As Boolean
'TAKES THE REQUIRED LEVEL AS INPUT. RETURNS FALSE IF THE
'USER HAS PERMISSIONS ON THIS GENERAL BOT FUNCTION.
    Dim temp As String, i As Integer
    CheckBotPerms = True
    For i = 1 To PermSize
        temp = Mid(Perms(i), 3)
        If (StrComp(LCase$(temp), "bot", vbBinaryCompare) = 0) And (CInt(Mid(Perms(i), 1, 1)) >= Level) Then
            'okie everything checks out break and return true
            CheckBotPerms = False
            Exit Function
        End If
    Next i
End Function

Public Sub SetPerms(ByVal szPerms As String)
'SETS THE PERMISSIONS FOR THIS USER.
    Dim i As Integer, x As Integer, temp As String
    i = CountTokens(szPerms, vbNullChar) - 1
    'first is the password line so trash it
    Call getNextToken(szPerms, vbNullChar)
    'now loop thru and add all the perms to our list
    For x = 1 To i - 1
        temp = getNextToken(szPerms, vbNullChar)
        Push Mid(temp, InStr(1, temp, "=") + 1)
        Debug.Print "Pushed: " & Mid(temp, InStr(1, temp, "=") + 1)
    Next x
End Sub

Private Sub Class_Terminate()
    Erase Perms
End Sub
