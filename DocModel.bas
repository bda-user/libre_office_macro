REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option VBASupport 1

Const STYLE_HEAD = "Heading"

Enum NodeType
    Section = 1
    Style = 2
    List = 3
    Table = 4
    Paragraph = 5
End Enum

Type Node ' node tree
    type_ As NodeType
    value As Variant ' NodeType: Paragraph = 3, Table = 4, Image = 5. LibreOffice Object
    name_ As String ' NodeType: Section = 1,	Style = 2
    children As Variant ' NodeType: Section = 1, Style = 2
    level As Integer ' NodeType: All
End Type

Enum SectionState
    End_ = 1
    New_ = 2
    Continue = 3
End Enum

Function SectionTest(ByRef curNode, ByRef curPara, ByRef sectionNames) As SectionState
    If IsEmpty(curPara.TextSection) Then
        If curNode.level > 0 Then
            SectionTest = SectionState.End_
            Exit Function
        EndIf
    ElseIf IsEmpty(curNode.name_) Or curPara.TextSection.Name <> curNode.name_ Then
        Dim sectionName : sectionName = Null
        If Not IsEmpty(curNode.name_) Then
            On Error Resume Next
            sectionName = sectionNames.Item(curPara.TextSection.Name)
            If Not IsNull(sectionName) Then
                SectionTest = SectionState.End_
                Exit Function
            End If
        End If
        SectionTest = SectionState.New_
        Exit Function
    End If
    SectionTest = SectionState.Continue
End Function

Function MakeNewSection(ByRef curNode, ByRef curPara, ByRef sectionNames) As Node
    sectionNames.Add(True, curPara.TextSection.Name)
    Dim newSec As Node
    With newSec
        .type_ = NodeType.Section
        .name_ = curPara.TextSection.Name
        .level = curNode.level + 1
        .children = New Collection
    End With
    curNode.children.Add(newSec)
    MakeNewSection = newSec
End Function

Function GetNodeStyle(ByRef curNode, ByRef curPara) As Node
    Dim i As Integer : i = curNode.children.Count
    If i > 0 Then
        Dim lastItem : lastItem = curNode.children.Item(i)
        If lastItem.type_ = NodeType.Style And _
            lastItem.name_ = curPara.ParaStyleName Then
            GetNodeStyle = lastItem
            Exit Function
        End If
    End If
    Dim nodeStyle As Node
    With nodeStyle
        .type_ = NodeType.Style
        .name_ = curPara.ParaStyleName
        .level = curNode.level + 1
        .children = New Collection
    End With
    curNode.children.Add(nodeStyle)
    GetNodeStyle = nodeStyle
End Function

Function GetNodeList(ByRef curNode, ByRef curPara) As Node
    Dim i As Integer : i = curNode.children.Count
    If i > 0 Then
        Dim lastItem : lastItem = curNode.children.Item(i)
        If lastItem.type_ = NodeType.List Then
            i = lastItem.children.Count
            If i > 0 Then
                Dim listItem
                listItem = lastItem.children.Item(i)
                If listItem.type_  = NodeType.Paragraph Then
                    If listItem.value.NumberingLevel = curPara.NumberingLevel Then
                        GetNodeList = lastItem
                        Exit Function
                    End If
                Else
                    GetNodeList = GetNodeList(lastItem, curPara)
                    Exit Function
                End If
            End If
            GetNodeList = GetNodeList(lastItem, curPara)
            Exit Function
        End If
    End If
    Dim nodeList As Node
    With nodeList
        .type_ = NodeType.List
        .name_ = IIf(curPara.ListLabelString = "", "Marked", "Numbered")
        .level = curNode.level + 1
        .children = New Collection
    End With
    curNode.children.Add(nodeList)
    GetNodeList = nodeList
End Function

Sub SectionParse(ByRef paraEnum, ByRef curPara, ByRef curNode, ByRef sectionNames)
    ' Enumerate paragraphs, include tables
    Do
' emulate key word "continue" in C++
continue:
		    
        ' Process the tables
        If curPara.supportsService("com.sun.star.text.TextTable") Then
    	    Dim nodeTable As Node
            With nodeTable
                .type_ = NodeType.Table
                .level = curNode.level + 1
                .value = curPara
            End With
            curNode.children.Add(nodeTable)  
    	    
        ' Process the paragrath
        Elseif curPara.supportsService("com.sun.star.text.Paragraph") Then
            Dim secState As SectionState
            secState = SectionTest(curNode, curPara, sectionNames)            
            Select Case secState 
                Case SectionState.End_
                    Exit Sub
                Case SectionState.New_
                    Dim newSec As Node
                    newSec = MakeNewSection(curNode, curPara, sectionNames)
                    SectionParse paraEnum, curPara, newSec, sectionNames
                    GoTo continue
            End Select
            
            ' Process the Style
            Dim nodeStyle As Node
            nodeStyle = GetNodeStyle(curNode, curPara)           
            Dim nodePara As Node
            With nodePara
                .type_ = NodeType.Paragraph
                .level = nodeStyle.level + 1
                .value = curPara
            End With
            If curPara.NumberingIsNumber And _
                Left(curPara.ParaStyleName, 7) <> STYLE_HEAD Then
                Dim nodeList As Node
                nodeList = GetNodeList(nodeStyle, curPara)
                nodePara.level = nodeList.level + 1
                nodeList.children.Add(nodePara)
            Else
                nodeStyle.children.Add(nodePara)
            End If
      
        End If
        If Not paraEnum.hasMoreElements() Then Exit Do
        curPara = paraEnum.nextElement()
    Loop
End Sub

Sub ExportToFile (ByRef text_ As String, Comp As Object, Optional suffix = "_export.txt")
    Dim FileNo As Integer, Filename As String
    Filename = convertToURL(replace(convertFromURL(Comp.URL), ".odt", suffix))
	FileNo = Freefile
	Open Filename For Output As #FileNo
	Print #FileNo, text_
End Sub

Function MakeModel(ByRef Comp As Object) As Node
    Dim sectionNames As New Collection
    Dim docTree As Node
    With docTree
        .type_ = NodeType.Section
        .level = 0
        .children = New Collection
    End With
 
    ' Enumerate paragraphs, include tables
    Dim paraEnum : paraEnum = Comp.getText().createEnumeration()
    Dim curPara
    If paraEnum.hasMoreElements() Then
        curPara = paraEnum.nextElement()
        SectionParse paraEnum, curPara, docTree, sectionNames
    End If
    MakeModel = docTree
End Function

Sub MakeDocHtmlView(Optional Comp As Object)
    Dim doc As Object : doc = ThisComponent
    If Comp Then doc = Comp

    Dim dView As New DocView : dView = New DocView
    Dim vHtml As New ViewHtml : vHtml = New ViewHtml
    vHtml.docView = dView
    dView.docTree = MakeModel(doc)
    dView.viewAdapter = vHtml
    dView.props = New Collection
    With dView.props
        .Add(True, "CodeLineNum") ' Enumerate code lines 1, 2, 3 ... n
    End With
    ExportToFile dView.MakeView(), doc, "_export.html"
End Sub

Sub MakeDocHfmView(Optional Comp As Object)
    Dim doc As Object : doc = ThisComponent
    If Comp Then doc = Comp

    Dim dView As New DocView : dView = New DocView
    Dim vHfm As New ViewHfm : vHfm = New ViewHfm
    vHfm.docView = dView
    dView.docTree = MakeModel(doc)
    dView.viewAdapter = vHfm
    dView.props = New Collection
    With dView.props
        .Add(True, "CodeLineNum") ' Enumerate code lines 1, 2, 3 ... n
    End With
    ExportToFile dView.MakeView(), doc, "_export_hfm.txt"
End Sub

' "C:\Program Files\LibreOffice\program\soffice.exe"  --invisible --nofirststartwizard --headless --norestore macro:///DocExport.DocModel.ExportDir("D:\cpp\habr\002-hfm",0)
Sub ExportDir(Folder As String, Optional Hfm As Boolean = True)  
    Dim Props(0) as New com.sun.star.beans.PropertyValue
    Props(0).NAME = "Hidden" 
    Props(0).Value = True 
    Dim Comp As Object
    Dim url, fname As String : fname = Dir$(Folder + "\" + "*.odt", 0)    
    Do
        url = ConvertToUrl(Folder + "\" + fname)
        Comp = StarDesktop.loadComponentFromURL(url, "_blank", 0, Props)
        If Hfm Then
            MakeDocHfmView Comp
        Else
            MakeDocHtmlView Comp
        End If
        fname = Dir$
        call Comp.close(True)
    Loop Until fname = ""
End Sub

