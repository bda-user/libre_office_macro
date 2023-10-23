REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit
Option Compatible
Option ClassModule

const GROUP_ON = 0
const GROUP_OFF = 1

Type fExtra ' fNode.extra
    name_ As String
    lname As String
    rname As String
End Type

Type fNode ' node tree
    type_ As Integer
    name_ As String
    children As Variant
    extra As fExtra
End Type

Public NodeType As New Collection
Public NodeTune As Variant
Public NodeTune2 As Variant
Public NodeTune3 As Variant
Public vAdapter As Object

' s_ - formula, sz_ - len, i_ - cur pos, c_ - cur char, w_ - cur word
Private s_ As String,  sz_ As Integer,  i_ As Integer, c_ As String, w_ As String
Private t_ As fNode ' root
Private keys As New Collection

Private Sub Class_Initialize()
    t_.children = New Collection
    With NodeType
        .Add(1, "Group") ' {...}, (...), lbrace ... rbrace, etñ.
        .Add(2, "Key")
        .Add(3, "Word")
        .Add(4, "Text") ' "a ... b"
    End With
    With keys
        .Add(0, "left") ' 0 - Group name, GROUP_ON
        .Add(0, "lbrace")
        .Add(0, "langle")
'        .Add(0, "right")
        .Add(1, "rbrace") ' 1 - Group rname, GROUP_OFF
        .Add(1, "rangle")
        .Add(1, "none")
        .Add(2, "overbrace") ' 4 - Key Add sym "\" before Keyword
        .Add(2, "underbrace")
        .Add(2, "newline")
        .Add(2, "matrix")
        .Add(2, "over")
        .Add(2, "vec")
        .Add(2, "times")
        .Add(2, "sqrt")
        .Add(2, "nroot")
        .Add(2, "int")
        .Add(2, "iint")
        .Add(2, "from")
        .Add(2, "to")
        .Add(2, "infty")
        .Add(2, "sum")
    End With
    
    NodeTune = Array("over", "overbrace", "underbrace")
    NodeTune2 = Array("nroot")
    NodeTune3 = Array("matrix")
    GlobalScope.BasicLibraries.loadLibrary("ScriptForge")
End Sub

Sub Set_Formula(ByRef s As String)
    i_ = 1
    s_ = s & "  "
    sz_ = Len(s_)
End Sub

Private Function Make_Node(ByRef node As fNode)  As fNode
    Dim n As fNode
    w_ = Trim(w_)
    If w_ = "" Or w_ = "right" Then
        n.type_ = 0
        w_ = ""
        Make_Node = n
        Exit Function
    End If
    n.type_ = NodeType("Word")
    On Local Error Goto ErrKey
    Dim i As Integer : i = keys(w_) 'if w not found -> ErrKey:
    If i = GROUP_ON Then
        Make_Node = Make_Group(node, w_)
        w_ = ""
        Exit Function
    Elseif i = GROUP_OFF Then
        n.type_ = 0
        node.extra.name_ = w_
        If node.name_ = "left" Then
            node.extra.name_ = "right"
            node.extra.rname = w_
        End If
        w_ = ""
        Make_Node = n
        Exit Function
    Else
        n.type_ = NodeType("Key")
    End If
ErrKey:
    n.name_ = w_
    n.children = New Collection
    node.children.Add(n)
    w_ = ""
    Make_Node = n
End Function

Private Function Make_Group(ByRef node As fNode, key_group As String)  As fNode
    Dim n As fNode
    If node.name_ = "left" And node.extra.lname = "" Then
        node.extra.lname = c_
        If key_group <> "" Then node.extra.lname = key_group
        n.type_ = 0
    Else
        n.type_ = NodeType("Group")
        n.name_ = c_
        If key_group <> "" Then n.name_ = key_group
        n.children = New Collection
        node.children.Add(n)
        i_ = i_ + 1
    End If
    Make_Group = n
End Function

Private Function Make_Op(ByRef node As fNode)  As fNode
    Make_Node(node)
    w_ = c_
    Make_Node(node)
End Function

Private Function Make_Text(ByRef node As fNode)  As fNode
    Make_Node(node)
    i_ = i_ + 1
    While i_ < sz_
        Dim n As fNode
        c_ = Mid(s_, i_, 1)
        If c_ = """" Then
            n.type_ = NodeType("Text")
            n.name_ = w_
            n.children = New Collection
            node.children.Add(n)
            Make_Text = n
            w_ = ""
            Exit Function 
        End If 
        w_ = w_ & c_
        i_ = i_ + 1
    Wend
    w_ = c_
    Make_Node(node)
End Function

Sub Parse(ByRef node As fNode)
    w_ = ""
    While i_ < sz_
        Dim n As fNode
        c_ = Mid(s_, i_, 1)
        Select Case c_
            Case "{", "(", "["
                n = Make_Node(node)
                n = Make_Group(node, "")
                If n.type_ = NodeType("Group") Then Parse n
            Case "}", ")", "]"
                n = Make_Node(node)
                node.extra.name_ = c_
                If node.name_ = "left" Then
                    node.extra.name_ = "right"
                    node.extra.rname = c_
                End If
                Exit Sub
            Case "^", "_"
                n = Make_Op(node)
            Case """"
                n = Make_Text(node)
            Case " ", CHR$(10)
                n = Make_Node(node)
                If node.name_ = "left" Then
                    If node.extra.rname <> "" Then Exit Sub
                ElseIf node.extra.name_ <> "" Then
                    Exit Sub
                End If
                If n.type_ = NodeType("Group") Then Parse n
            Case Else
                w_ = w_ & c_
        End Select
        i_ = i_ + 1
    Wend
End Sub

Sub Tune(ByRef node As fNode)
    Dim i As Integer, n As fNode
    For i = 1 To node.children.Count
        n = node.children.Item(i)
        
        If SF_Array.IndexOf(NodeTune, n.name_) > -1 And _
           n.children.Count = 0 Then
            n.children.Add(node.children.Item(i - 1))
            n.children.Add(node.children.Item(i + 1))
            node.children.Remove(i + 1)
            node.children.Remove(i - 1)
        Elseif SF_Array.IndexOf(NodeTune2, n.name_) > -1 And _
           n.children.Count = 0 Then
            n.children.Add(node.children.Item(i + 1))
            n.children.Add(node.children.Item(i + 2))
            node.children.Remove(i + 2)
            node.children.Remove(i + 1)
        Elseif SF_Array.IndexOf(NodeTune3, n.name_) > -1 And _
           n.children.Count = 0 Then
            n.children.Add(node.children.Item(i + 1))
            node.children.Remove(i + 1)
        End If
        
        Tune n
    Next
End Sub

Function Get_Formula() As String
    Parse t_
    Tune t_
    Get_Formula = vAdapter.PrintMe(t_)
End Function
