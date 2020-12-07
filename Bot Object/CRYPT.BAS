Attribute VB_Name = "Module1"

Const Pass$ = "KJLHDKOI@#UR(*U$#()R()$R(*@!#(U!( *#@JKDJ**&^"

Public Function Encrypt(ByVal szString As String) As String
    Dim H$, J$
    szString = Crypt(Pass$, szString)
    Debug.Print "Encrypted = " & szString
    'When writing an encrypted password to a sequential access file like the
    'INI files, you need to convert the resultant encrypted file to hex data.
    'This is because you can end up with an encrypted password that contains
    'characters which cannot be properly read using sequential access.  So,
    'before saving your encrypted password, use this routine:
    H$ = ""
    For i = 1 To Len(szString)
        J$ = Hex$(Asc(Mid$(szString, i, 1)))
        If Len(J$) = 1 Then J$ = "0" + J$
        H$ = H$ + J$
    Next
   Encrypt = Format$(Len(H$), "00") + H$
   Debug.Print "Encrypted Hex = " & Encrypt
End Function

Public Function Decrypt(ByVal szString As String) As String
    Dim H$, Strg$, J$
    H$ = Mid$(szString, 3, Val(Left$(szString, 2)))

    Strg$ = ""
    For i = 1 To Len(H$) Step 2
        J$ = Mid$(H$, i, 2)
        Strg$ = Strg$ + Chr$(Val("&H" + J$))
    Next
    Decrypt = Crypt(Pass$, Strg$)
    Debug.Print "Decrypted = " & Decrypt
End Function

Private Function Crypt(ByVal Pass$, ByVal Strg$) As String
  Dim a As Integer, B
  a = 1
  For i = 1 To Len(Strg$)
     B = Asc(Mid$(Pass$, a, 1)): a = a + 1: If a > Len(Pass$) Then a = 1
     Mid$(Strg$, i, 1) = Chr$(Asc(Mid$(Strg$, i, 1)) Xor B)
  Next
  Crypt = Strg$
End Function
