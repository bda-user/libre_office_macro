REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit
Option Compatible
Option ClassModule

Public mMath

Private groupChar As Variant
Private rename As New Collection ' replace keywords
Private extraChar As New Collection

Private Sub Class_Initialize()
    GlobalScope.BasicLibraries.loadLibrary("ScriptForge")
    groupChar = Array("{", "}", "(", ")", "[", "]")
    With rename
        .Add("frac", "over")
        .Add("sqrt", "nroot")
        .Add("_", "from")
        .Add("^", "to")
        .Add("\ " & CHR$(10), "newline")
        .Add("& ", "#")
        .Add("\\ ", "##")
    End With
    With extraChar
        .Add("^", "overbrace")
        .Add("_", "underbrace")
    End With
End Sub

Public Function ReplaceKey(ByRef node As fNode) As String
    ReplaceKey = "\" & node.name_ & " "
    On Local Error Goto ErrRename
    Dim s As String : s = rename(node.name_) 'if not found -> ErrRename:
    If Len(s) = 1 Then
        ReplaceKey = s
    Else
        ReplaceKey = "\" & s
    End If
    ErrRename:
End Function

Public Function BeginNode(ByRef node As fNode) As String
    Dim s As String : s = ""
    Select Case node.type_
        Case mMath.NodeType("Text")
            s = "\text{" & node.name_ & "} "
        Case mMath.NodeType("Key")
            s = ReplaceKey(node)
        Case mMath.NodeType("Group")
            s = node.name_ & " "
            If SF_Array.IndexOf(groupChar, node.name_) = -1 Then s = "\" & s
            If node.name_ = "left" Then
                Dim c As String : c = ""
                If SF_Array.IndexOf(groupChar, node.extra.lname) = -1 Then c = "\"
                s = s & c & node.extra.lname & " "
            End If
        Case Else ' Word
            s = node.name_ & " "
            On Local Error Goto ErrRename
                s = rename(node.name_) & " "
            ErrRename:
            If Mid(s, 1, 1) = "%" Then s = "\" & Mid(s, 2, Len(s) - 1) & " " ' greece sym
    End Select
    BeginNode = s
End Function

Public Function EndNode(ByRef node As fNode) As String
    Dim s As String : s = ""
    If node.type_ = mMath.NodeType("Group") Then
        s = node.extra.name_ & " "
        If SF_Array.IndexOf(groupChar, node.name_) = -1 Then s = "\" & s
        If node.extra.name_ = "right" Then
            Dim c As String : c = ""
            If SF_Array.IndexOf(groupChar, node.extra.rname) = -1 Then c = "\"
            s = s & c & node.extra.rname & " "
        End If
    End If
    EndNode = s
End Function

Public Function NodeTune(ByRef node As fNode) As String
    Dim s As String : s = ""
    Dim n As fNode
    s = "\" & node.name_

On Local Error Goto ErrRename
    Dim b1 As String : b1 = ""
    Dim b2 As String : b2 = ""
    s = "\" & rename(node.name_) 'if not found -> ErrRename:
    b1 = "{"
    b2 = "}"
ErrRename:

On Local Error Goto ErrExtraChar
    Dim e As String : e = ""
    e = extraChar(node.name_)
ErrExtraChar:

    Dim b As Boolean : b = True
    For Each n In node.children
        s = s & b1 & BeginNode(n)
        s = s & Print_(n)
        s = s & EndNode(n) & b2
        If b Then
            s = s & e
            b = False
        End If
    Next
    NodeTune = s & " "
End Function

Public Function NodeTune2(ByRef node As fNode) As String
    Dim s As String : s = ""
    Dim n As fNode
    s = ReplaceKey(node) & " " 'if not found -> ErrRename:

    Dim b1 As String : b1 = ""
    Dim b2 As String : b2 = ""
    If node.name_ = "nroot" Then
        b1 = "["
        b2 = "] "
    End If

    Dim b As Boolean : b = True
    For Each n In node.children
        s = s & IIf(b, b1, "") & BeginNode(n)
        s = s & Print_(n)
        s = s & EndNode(n) & IIf(b, b2, "")
        If b Then b = False
    Next
    NodeTune2 = s & " "
End Function

Public Function Matrix(ByRef node As fNode) As String
    Dim n As fNode, s As String : s =  CHR$(10) & _
    "\begin{matrix} " & CHR$(10) & _
    " " & Print_(node.children.Item(1)) & CHR$(10) & _
    "\end{matrix} " & CHR$(10)
    Matrix = s
End Function

Public Function Print_(ByRef node As fNode) As String
    Dim s As String : s = ""      
    Dim n As fNode
    For Each n In node.children
        If SF_Array.IndexOf(mMath.NodeTune, n.name_) > -1 Then
            s = s & NodeTune(n)
        ElseIf SF_Array.IndexOf(mMath.NodeTune2, n.name_) > -1 Then
            s = s & NodeTune2(n)
        ElseIf n.name_ = "matrix" Then
            s = s & Matrix(n)
        Else
            s = s & BeginNode(n)
            s = s & Print_(n)
            s = s & EndNode(n)
        End If
    Next
    Print_ = s
End Function

Public Function PrintMe(ByRef node As fNode) As String
    PrintMe = _
    "\begin{align} " & CHR$(10) & _
    Print_(node) & CHR$(10) & _
    "\end{align}"
End Function

