Attribute VB_Name = "modWeb"
Option Explicit

Public header As String
Public menu As String
Public bodyend As String

Private ServerStats As String
Private Index As String

Public Function Load_Set_Data()
    'load files
    header = LoadFile(App.Path & "\web\head.tpl")
    menu = LoadFile(App.Path & "\web\menu.tpl")
    bodyend = LoadFile(App.Path & "\web\end.tpl")
End Function

Public Function IndexData(ByVal playersonline As Long, ByVal Accounts As Long, ByVal Guilds As Long) As String
    ServerStats = "<font color=#000000><BR>Status: &quot;</font><b><font color=#00FF00>UP</font></b><font color=#000000>&quot;<BR>Players Online: " & playersonline & "<BR>Accounts: " & Accounts & "<BR></font>Guilds: " & Guilds & "<font color=#000000><br></font>"
    Index = LoadFile(App.Path & "\web\index.tpl")
    IndexData = header & menu & ServerStats & Index & bodyend
End Function

Public Function Welcome(ByVal playersonline As Long, ByVal Accounts As Long, ByVal Guilds As Long) As String
    ServerStats = "<font color=#000000><BR>Status: &quot;</font><b><font color=#00FF00>UP</font></b><font color=#000000>&quot;<BR>Players Online: " & playersonline & "<BR>Accounts: " & Accounts & "<BR></font>Guilds: " & Guilds & "<font color=#000000><br></font>"
    Index = LoadFile(App.Path & "\web\welcome.tpl")
    Welcome = header & menu & ServerStats & Index & bodyend
End Function

Public Function LoadFile(ByVal filename1 As String) As String
    Open filename1 For Binary As #1
    LoadFile = Input(FileLen(filename1), #1)
    Close #1
End Function

Public Function AccountData(ByVal playersonline As Long, ByVal Accounts As Long, ByVal Guilds As Long) As String
    ServerStats = "<font color=#000000><BR>Status: &quot;</font><b><font color=#00FF00>UP</font></b><font color=#000000>&quot;<BR>Players Online: " & playersonline & "<BR>Accounts: " & Accounts & "<BR></font>Guilds: " & Guilds & "<font color=#000000><br></font>"
    Index = LoadFile(App.Path & "\web\Account.tpl")
    AccountData = header & menu & ServerStats & Index & bodyend
End Function

Public Function Error1Data(ByVal playersonline As Long, ByVal Accounts As Long, ByVal Guilds As Long) As String
    ServerStats = "<font color=#000000><BR>Status: &quot;</font><b><font color=#00FF00>UP</font></b><font color=#000000>&quot;<BR>Players Online: " & playersonline & "<BR>Accounts: " & Accounts & "<BR></font>Guilds: " & Guilds & "<font color=#000000><br></font>"
    Index = LoadFile(App.Path & "\web\error1.tpl")
    Error1Data = header & menu & ServerStats & Index & bodyend
End Function

Public Function Error2Data(ByVal playersonline As Long, ByVal Accounts As Long, ByVal Guilds As Long) As String
    ServerStats = "<font color=#000000><BR>Status: &quot;</font><b><font color=#00FF00>UP</font></b><font color=#000000>&quot;<BR>Players Online: " & playersonline & "<BR>Accounts: " & Accounts & "<BR></font>Guilds: " & Guilds & "<font color=#000000><br></font>"
    Index = LoadFile(App.Path & "\web\error2.tpl")
    Error2Data = header & menu & ServerStats & Index & bodyend
End Function

Public Function Web_Field_Read(ByVal field_pos As Long, ByVal text As String) As String
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
    Dim lastfield As Long
    Dim d As String
    Dim Count As Long
    
    On Error Resume Next
    
    LastPos = 0
    FieldNum = 0
    lastfield = 1
    
    For i = 1 To Len(text)
        d = Mid(text, LastPos, 1)
        If d = "&" Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                Web_Field_Read = Mid(text, lastfield, Count - 1)
                Exit Function
            End If
            lastfield = LastPos + 1
            Count = 0
        End If
        Count = Count + 1
        LastPos = LastPos + 1
    Next i
End Function

Public Function Web_Field_Read_Values(ByVal text As String) As String
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
    Dim lastfield As Long
    Dim d As String
    Dim Count As Long
    Dim field_pos As Long
    
    On Error Resume Next
    
    LastPos = 0
    FieldNum = 0
    lastfield = 1
    field_pos = 1
    
    For i = 1 To Len(text)
        d = Mid(text, LastPos, 1)
        If d = "=" Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                Web_Field_Read_Values = Mid(text, lastfield + Count)
                Exit Function
            End If
            lastfield = LastPos + 1
            Count = 0
        End If
        Count = Count + 1
        LastPos = LastPos + 1
    Next i
End Function

Public Function ASCIItoUTF(ByVal data As String) As String
    Dim c As Long
    Dim LoopC As Long
    Dim a As String
    
    For LoopC = 1 To Len(data)
        a = Mid(data, LoopC, 1)
        c = Asc(a)
        If (c < &H80) Then
            ASCIItoUTF = ASCIItoUTF & Chr(c)
        ElseIf (c < &H800) Then
            ASCIItoUTF = ASCIItoUTF & Chr(&HC0 Or (c / (2 ^ 6)))
            ASCIItoUTF = ASCIItoUTF & Chr(&H80 Or (c And &H3F))
        End If
    Next LoopC
End Function

Public Function ConvertUTF8toASCII(ByVal strData As String) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 9/03/2004
'
'*****************************************************************
    Dim pos As Long
    Dim LastPos As Long
    Dim tempStr As String
    
    'First we replace all "+" with spaces
    pos = InStr(strData, "+")
    
    LastPos = 1
    While pos
        tempStr = tempStr & Mid(strData, LastPos, pos - LastPos) & " "
        LastPos = pos + 1
        pos = InStr(LastPos, strData, "+")
    Wend
    
    tempStr = tempStr & Right$(strData, Len(strData) - LastPos + 1)
    
    If LastPos = 1 Then tempStr = strData
    
    'Search for UTF-8 values
    pos = InStr(tempStr, "%")
    
    If pos = 0 Then
        ConvertUTF8toASCII = tempStr
        Exit Function
    End If
    
    LastPos = 1
    While pos
        ConvertUTF8toASCII = ConvertUTF8toASCII & Mid(tempStr, LastPos, pos - LastPos) & Chr(Val("&H " & Mid(tempStr, pos + 1, 2)))
        LastPos = pos + 3
        pos = InStr(LastPos, tempStr, "%")
    Wend
    
    ConvertUTF8toASCII = ConvertUTF8toASCII & Right$(tempStr, Len(tempStr) - LastPos + 1)
End Function
