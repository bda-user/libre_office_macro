REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit

Function Format_Num(ByVal lineNum As Integer, Optional Size As Integer) As String
    Dim sz As Integer : sz = "4"
    If Not IsMissing(Size) Then sz = Size
    Dim s As String : s = "" + lineNum
    s = s + String(sz - Len(s), " ")
    Format_Num = s
End Function

Function Escape_Characters(ByRef s As String) As String
    Dim sz As Integer, c As String,  i As Integer, r As String : r = ""
    sz = Len(s) 
    i = 1
    While i <= sz
        c = Mid(s, i, 1)
        Select Case c
            Case "<"
                r = r + "&lt;"
            Case ">"
                r = r + "&gt;"
            Case Else
                r = r + c
        End Select
        i = i + 1
    Wend
    Escape_Characters = r
End Function

Sub Main
    Dim s As String : s = "'" + Format_Num(43, 12) + "'"
    s = "<?xml version=""1.0"" encoding=""UTF-8""?>"
    s = Escape_Characters(s)
End Sub
