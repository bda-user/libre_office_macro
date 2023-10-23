REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit
Option Compatible
Option ClassModule

Const STYLE_QUOT = "Quotations"
Const STYLE_CODE = "code_"
Const STYLE_HEAD = "Heading"

Public docTree
Public viewAdapter
Public props

Function PrintNodeStyle(ByRef node)
    Dim s As String : s = ""
    If node.name_ = STYLE_QUOT Then
        s = viewAdapter.Quote(node)
    ElseIf Left(node.name_, 5) = STYLE_CODE Then
        s = viewAdapter.Code(node)
    ElseIf Left(node.name_, 7) = STYLE_HEAD Then
        Dim textPortion, enumPortion
        enumPortion = node.children(1).value.createEnumeration()
        Do While enumPortion.hasMoreElements()
            textPortion = enumPortion.nextElement()
            If textPortion.TextPortionType = "Text" Then
                s = s & viewAdapter.Head(node)
            ElseIf textPortion.TextPortionType = "Bookmark" Then
                s = s & viewAdapter.Anchor(textPortion)
            End If
        Loop
    Else
        s = viewAdapter.ParaStyle(node)
    End If
    PrintNodeStyle = s
End Function

Function PrintNodeParaLO(ByRef oPara, level As Long, Optional lineNum As Integer)
    Dim textGraphObj$ : textGraphObj$ = "com.sun.star.text.TextGraphicObject"
    Dim drawShape$ : drawShape$ = "com.sun.star.drawing.Shape"
    Dim textEmbObj$ : textEmbObj$ = "com.sun.star.text.TextEmbeddedObject"
    ' Graphics are anchored to a paragraph are enumerate as TextContent
    Dim contEnum : contEnum = _
        oPara.createContentEnumeration("com.sun.star.text.TextContent")
    Dim curContent, s As String : s = ""
    
    If Not IsMissing (lineNum) AND lineNum > 0 Then s = s & lineNum & " "
    
    Do While contEnum.hasMoreElements()
        curContent = contEnum.nextElement()           
        If curContent.supportsService(textGraphObj$) Then
            ' Image anchored to a paragraph
            s = s & viewAdapter.Image(curContent)
        ElseIf curContent.supportsService(drawShape$) Then
            ' Drawing shape anchored to a paragraph
        End If
    Loop
    
    ' Graphics are anchored a character, or as a character into paragraph
    ' are enumerate as TextPortionType will process when parse paragraph
    Dim textPortion, enumPortion : enumPortion = oPara.createEnumeration()
    Do While enumPortion.hasMoreElements()
        textPortion = enumPortion.nextElement()
        If textPortion.TextPortionType = "Text" Then
            ' Simply text object is here
            If Not IsEmpty(textPortion.HyperLinkURL) And _
                textPortion.HyperLinkURL <> "" Then
			    s = s & viewAdapter.Link(textPortion)
      	    ElseIf textPortion.CharWeight = com.sun.star.awt.FontWeight.BOLD Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Bold")
      	    ElseIf textPortion.CharPosture = com.sun.star.awt.FontSlant.ITALIC Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Italic")
      	    ElseIf textPortion.CharUnderline = com.sun.star.awt.FontUnderline.SINGLE Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Underline")
      	    ElseIf textPortion.CharStrikeout = com.sun.star.awt.FontStrikeout.SINGLE Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Strikeout")
            Else
			    s = s & textPortion.String
			End If
        ElseIf textPortion.TextPortionType = "Frame" Then
            ' Check inline textGraphic & drawing.Shape
            Dim framePortion, enumFrame
            enumFrame = textPortion.createContentEnumeration(textGraphObj$)
            Do While enumFrame.hasMoreElements()
                framePortion = enumFrame.nextElement()
                If framePortion.supportsService(textGraphObj$) Then
                    ' inline IMG is here
                    s = s & viewAdapter.InlineImage(framePortion)
                ElseIf framePortion.supportsService(drawShape$) Then
                    ' inline Shape here
                ElseIf framePortion.supportsService(textEmbObj$) And _
                    framePortion.FrameStyleName = "Formula" Then
                    ' inline Formula here
                    s = s & viewAdapter.Formula(framePortion.Component.Formula)
                End If
            Loop
        ElseIf textPortion.TextPortionType = "Bookmark" Then
            s = s & viewAdapter.Anchor(textPortion)
        End If
    Loop
    If oPara.NumberingIsNumber Then
        PrintNodeParaLO = viewAdapter.FormatList(oPara, s, level)
    ElseIf Not IsMissing (lineNum) AND lineNum = 0 Then
        PrintNodeParaLO = viewAdapter.FormatPara(s, level, 0)       
    ElseIf Not IsMissing (lineNum) AND lineNum > 0 Then
        PrintNodeParaLO = s & CHR$(10)
    Else
        PrintNodeParaLO = viewAdapter.FormatPara(s, level, 1)
    End If
End Function

Function PrintNodePara(ByRef nodePara, Optional lineNum As Integer)
    If IsMissing (lineNum) Then
        PrintNodePara = PrintNodeParaLO(nodePara.value, nodePara.level)
        Exit Function
    End If
    PrintNodePara = PrintNodeParaLO(nodePara.value, nodePara.level, lineNum)
End Function

Function PrintNodeTable(ByRef nodeTable)
    Dim oTable, oCell, oText, oEnum, oPar, t, r, c
    Dim nRow As Long, nCol As Long, Rows As Long, Colls As Long
    oTable = nodeTable.value : t = ""
    Rows = oTable.getRows().getCount() - 1
    Colls = oTable.getColumns().getCount() - 1
    For nRow = 0 To Rows
        r = ""
        For nCol = 0 To Colls
            oCell = oTable.getCellByPosition(nCol, nRow)
            oText = oCell.getText()
            c = ""
            oEnum = oText.createEnumeration()
            Do While oEnum.hasMoreElements()
                oPar = oEnum.nextElement()
                If oPar.supportsService("com.sun.star.text.Paragraph") Then
                    c = c & PrintNodeParaLO(oPar, nodeTable.level + 3, 0)
                End If
            Loop
            r = r & viewAdapter.FormatCell(c, nodeTable.level + 2, nCol, nRow)
        Next
        t = t & viewAdapter.FormatRow(r, nodeTable.level + 1, nRow, Colls)
    Next
    PrintNodeTable = viewAdapter.FormatTable(t, nodeTable.level)
End Function

Function PrintTree(ByRef node, Optional ByRef props As Collection)
    Dim child, lineNum : lineNum = 0
    Dim s : s = ""
    If Not IsMissing(props) And props("CodeLineNum") Then lineNum = 1
    For Each child In node.children
        If child.type_ = NodeType.Section Then
            s = s & viewAdapter.Section(child)
        ElseIf child.type_ = NodeType.Style Then
            s = s & PrintNodeStyle(child)
        ElseIf child.type_ = NodeType.List Then
            s = s & viewAdapter.List(child)
        ElseIf child.type_ = NodeType.Paragraph Then
            If lineNum > 0 Then
                s = s & PrintNodePara(child, lineNum)
                lineNum = lineNum + 1
            Else
                s = s & PrintNodePara(child)
            End If
        ElseIf child.type_ = NodeType.Table Then
            s = s & PrintNodeTable(child)
        End If
    Next
    PrintTree = s
End Function

Public Function MakeView() As String
    MakeView = PrintTree(docTree)
End Function

