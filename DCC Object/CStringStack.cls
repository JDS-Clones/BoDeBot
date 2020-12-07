VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStringStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Stack() As String
Private StackSize As Integer
Private LastItem As Integer
Const WANT_FREE_PERCENT = 0.1
Const MIN_FREE = 10
Public Function Pop() As String
    If LastItem > 0 Then
        Pop = Stack(LastItem)
    Else
        Pop = Stack(1)
    End If
    If LastItem > 1 Then
        LastItem = LastItem - 1
    End If
End Function

Public Function UnPop() As String
    If LastItem < StackSize Then
        LastItem = LastItem + 1
        If Not Stack(LastItem) = "" Then
            UnPop = Stack(LastItem)
        Else
            LastItem = LastItem - 1
            UnPop = Stack(LastItem)
        End If
    Else
        UnPop = Stack(LastItem)
    End If
End Function


Public Sub Push(ByVal szItem As String)
    LastItem = LastItem + 1
    If LastItem > StackSize Then
        ShiftStack szItem
        LastItem = LastItem - 1
    Else
        Stack(LastItem) = szItem
    End If
End Sub


Private Sub ShiftStack(ByVal szItem As String)
    ReDim temp(1 To StackSize) As String
    Dim x As Integer
    For x = 1 To StackSize - 1
        temp(x) = Stack(x + 1)
    Next x
    For x = 1 To StackSize - 1
        Stack(x) = temp(x)
    Next x
    Stack(StackSize) = szItem
    Erase temp
End Sub

Public Sub MallocStack(dwSize As Integer)
    StackSize = dwSize
    ReDim Stack(1 To StackSize)
End Sub

Private Sub Class_Terminate()
    StackSize = 0
    LastItem = 0
    Erase Stack
End Sub
