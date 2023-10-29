REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit

Function Format_Num(ByVal lineNum As Integer, Optional Size As Integer) As String
    Dim sz As Integer : sz = "4"
    If Not IsMissing(Size) Then sz = Size
    Dim s As String : s = "" + lineNum
    s = s + String(sz - Len(s), " ")
    Format_Num = s
End Function

Sub Main
    Dim s As String : s = Format_Num(43, 12)
End Sub
