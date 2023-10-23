REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit
Option Compatible
Option ClassModule

Const STYLE_HEADING = "Contents"
Const SEC_CLASS = "section"
Const SHIFT_CNT = 4

Public docView
Public fontStyles As New Collection
Public headHtml As New Collection

Private Sub Class_Initialize()
    fontStyles.Add(Array("<b>", "</b>"), "Bold")
    fontStyles.Add(Array("<i>", "</i>"), "Italic")
    fontStyles.Add(Array("<u>", "</u>"), "Underline")
    fontStyles.Add(Array("<del>", "</del>"), "Strikeout")
    
    headHtml.Add(Array("<h1>", "</h1>"), "1")
    headHtml.Add(Array("<h2>", "</h2>"), "2")
    headHtml.Add(Array("<h3>", "</h3>"), "3")
End Sub

Function Head(ByRef node)
   Dim shift As String : shift = String(node.level * SHIFT_CNT, " ")
    Dim s : s = Split(node.name_, " ")(1)
    Dim tag : tag = headHtml(s)
    Head = CHR$(10) & shift & tag(0) & node.children(1).value.String & tag(1) & CHR$(10)
End Function

Function Quote(ByRef node)
    Dim shift As String : shift = String(node.level * SHIFT_CNT, " ")
    Quote = shift & "<blockquote>" & CHR$(10) & _
        docView.PrintTree(node) & shift & "</blockquote>" & CHR$(10)
End Function

Function Code(ByRef node)
    If docView.props("CodeLineNum") Then
        Dim props As New Collection
        props.Add(True, "CodeLineNum")
        Code = docView.PrintTree(node, props)
    Else
        Code = docView.PrintTree(node)
    End If
    Dim shift As String : shift = String(node.level * SHIFT_CNT, " ")
    Code = shift & "<pre class=""code " & Split(node.name_, "_")(1) & """>" & _
        CHR$(10) & Code & shift & "</pre>" & CHR$(10)
End Function

Function ParaStyle(ByRef node)
    Dim shift As String : shift = String(node.level * SHIFT_CNT, " ")
    ParaStyle = shift & "<div>" & CHR$(10) & _
        docView.PrintTree(node) & shift & "</div>" & CHR$(10)
End Function

Function List(ByRef node)
    List = docView.PrintTree(node)
    Dim shift As String : shift = String(node.level * SHIFT_CNT, " ")
    If node.name_ = "Marked" Then
        List = shift & "<ul>" & CHR$(10) & _
            List & shift & "</ul>" & CHR$(10)
    Else
        List = shift & "<ol>" & CHR$(10) & _
            List & shift & "</ol>" & CHR$(10)    
    End If
End Function

Function InlineImage(ByRef node)
    InlineImage = "<img alt=""" & node.Title & """ src=""" & _
        node.Graphic.OriginURL & """ />"
End Function

Function Image(ByRef node)
    Image =  InlineImage(node) & "<br />" & CHR$(10)
End Function

Function Link(ByRef node)
    Dim t As String : t = node.String
    If Left(node.ParaStyleName, 8) = STYLE_HEADING Then
        t = Left(t, Len(t) - 2)
    End If
    Link = "<a href=""" & node.HyperLinkURL  & """ >" & t & "</a>"
End Function

Function Anchor(ByRef lo)
    Anchor = IIf(lo.IsStart, _
        "<a name =""" & lo.Bookmark.Name & """></a>", "")
End Function

Function FontDecorate(ByRef node, style As String)
    Dim s : s = fontStyles(style)
    FontDecorate = s(0) & node.String & s(1)
End Function

Function FormatCell(ByRef txt, level As Long, index As Long, idxRow As Long)
    Dim shift As String : shift = String(level * SHIFT_CNT, " ")
    If idxRow = 0 Then
        FormatCell = shift & "<th>" & CHR$(10) & txt & shift &"</th>" & CHR$(10)
    Else
        FormatCell = shift & "<td>" & CHR$(10) & txt & shift & "</td>" & CHR$(10)
    End If
End Function

Function FormatRow(ByRef txt, level As Long, index As Long, Colls As Long)
    Dim shift As String : shift = String(level * SHIFT_CNT, " ")
    FormatRow = shift & "<tr>" & CHR$(10) & txt & shift & "</tr>" & CHR$(10)
End Function

Function FormatTable(ByRef txt, level As Long)
    Dim shift As String : shift = String(level * SHIFT_CNT, " ")
    FormatTable = shift & "<table>" & CHR$(10) & txt & shift & "</table>" & CHR$(10)
End Function

Function FormatList(ByRef list, ByRef txt, level As Long)
    Dim shift As String : shift = String(level * SHIFT_CNT, " ")
    FormatList = shift & "<li>" & txt & "</li>" & CHR$(10)
End Function

Function FormatPara(ByRef txt, level As Long, extra As Long)
    Dim shift As String : shift = String(level * SHIFT_CNT, " ")
    FormatPara = shift & "<p>" & txt & "</p>" & CHR$(10)
End Function

Function Formula(ByRef txt As String)
    Formula = " " & txt & " "
End Function

Function Section(ByRef nodeSec)
    Dim shift As String : shift = String(nodeSec.level * SHIFT_CNT, " ")
    Section = shift & "<div class=""" & SEC_CLASS & """>" & _
        "<!-- section begin -->" & CHR$(10) & _
        docView.PrintTree(nodeSec) & _
        shift & "</div><!-- section end -->" & CHR$(10)
End Function

