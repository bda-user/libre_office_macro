REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit
Option Compatible
Option ClassModule

Const SEC_HEADING = "Оглавление"
Const STYLE_HEADING = "Contents"
Const SHIFT_CNT = 4

Public docView
Public fontStyles As New Collection

Private Sub Class_Initialize()
    fontStyles.Add(Array("**", "**"), "Bold")
    fontStyles.Add(Array("*", "*"), "Italic")
    fontStyles.Add(Array("<u>", "</u>"), "Underline")
    fontStyles.Add(Array("~~", "~~"), "Strikeout")
End Sub

Function Head(ByRef node)
    Dim i : i = Split(node.name_, " ")(1)
    Head = CHR$(10) & String(i, "#") & " " & _
        node.children(1).value.String & CHR$(10)
End Function

Function Quote(ByRef node)
    Quote = "> " & docView.PrintTree(node) & CHR$(10)
End Function

Function Code(ByRef node)
    If docView.props("CodeLineNum") Then
        Dim props As New Collection
        props.Add(True, "CodeLineNum")
        Code = docView.PrintTree(node, props)
    Else
        Code = docView.PrintTree(node)
    End If
    Code = "```" & Split(node.name_, "_")(1) & CHR$(10) & _
        Code & "```" & CHR$(10)
End Function

Function ParaStyle(ByRef node)
    ParaStyle = docView.PrintTree(node)
End Function

Function List(ByRef node)
    List = docView.PrintTree(node)
End Function

Function Image(ByRef lo)
    Image = "![" & lo.Title & "](" & lo.Graphic.OriginURL & _
        " """ & lo.Description & """)" & CHR$(10)
End Function

Function InlineImage(ByRef lo)
    InlineImage = "<img inline=""true"" src=""" & _
        lo.Graphic.OriginURL & """ />"
End Function

Function Link(ByRef node)
    Dim t As String : t = node.String
    If Left(node.ParaStyleName, 8) = STYLE_HEADING Then
        t = Left(t, Len(t) - 2)
    End If
    Link = "[" & t & "](" & node.HyperLinkURL & ")"
End Function

Function Anchor(ByRef lo)
    Anchor = IIf(lo.IsStart, "<anchor>" & lo.Bookmark.Name & "</anchor>", "")
End Function

Function FontDecorate(ByRef node, style As String)
    Dim s : s = fontStyles(style)
    FontDecorate = s(0) & node.String & s(1)
End Function

Function FormatCell(ByRef txt, level As Long, index As Long, idxRow As Long)
    FormatCell = IIf(index = 0, txt, "|" & txt)
End Function

Function FormatRow(ByRef txt, level As Long, index As Long, Colls As Long)
    Dim i AS Long, r As String : r = ""
    r = "|" & txt & "|" & CHR$(10)
    If index = 0 Then
        r = r & "|"
        For i = 0 To Colls
            r = r & " --- |"
        Next
        r = r & CHR$(10)
    End If
    FormatRow = r
End Function

Function FormatTable(ByRef txt, level As Long)
    FormatTable = txt
End Function

Function FormatList(ByRef list, ByRef txt, level As Long)
    Dim shift As String : shift = String(list.NumberingLevel * SHIFT_CNT, " ")
    Dim lbl As String : lbl = list.ListLabelString
    FormatList = shift & IIf(lbl = "", "-", lbl) & " " & txt & CHR$(10)
End Function

Function FormatPara(ByRef txt, level As Long, extra As Long)
    FormatPara = txt & IIf(extra = 0, "", CHR$(10))
End Function

Function Formula(ByRef txt As String)
    Dim m As New mMath
    m.Set_Formula(txt)
    m.vAdapter = New vLatex
    m.vAdapter.mMath = m
    Formula = "$$" & CHR$(10) & m.Get_Formula() & CHR$(10) &  "$$" & CHR$(10)
End Function

Function GetSectionTitle(ByRef nodeSec)
    GetSectionTitle = ""
    If nodeSec.children.Count > 0 And _
        nodeSec.children(1).type_ =  NodeType.Style Then
        Dim nodeStyle : nodeStyle = nodeSec.children(1)
        If nodeStyle.children.Count > 0 And _
            nodeStyle.children(1).type_ =  NodeType.Paragraph Then
            GetSectionTitle = nodeStyle.children(1).value.getString()
            nodeStyle.children.Remove(1)
        End If
    End If
End Function

Function Section(ByRef nodeSec)
    If nodeSec.level <> 1 Then
        Section = docView.PrintTree(nodeSec)
        Exit Function
    End If
    
    Dim secTitle : secTitle = GetSectionTitle(nodeSec)
    If secTitle = SEC_HEADING Then
        Section = "# " & secTitle & CHR$(10) & docView.PrintTree(nodeSec)
        Exit Function
    End If

    Section = "<spoiler title=""" & _
        secTitle & """>" & CHR$(10) & CHR$(10) & _
        docView.PrintTree(nodeSec) & "</spoiler>" & CHR$(10)
End Function

