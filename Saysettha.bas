Attribute VB_Name = "Saysettha"
Dim statement() As String
Dim line, maxFor, maxIf As Integer
Dim ForCount(1 To 1000000, 1 To 2), IfCount(1 To 1000000, 1 To 2) As Variant

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
Public Sub ErrorHandle()
If Slide1.errcount <> 0 Then Slide1.console = "Runtime Error. Check the code again."
End Sub
Public Sub ConsoleWrite(ByVal message As String)
Slide1.console = Slide1.console + message
End Sub
Public Sub CountFor()
    Dim v As Variant
    Dim ForCountReverse(1 To 1000000) As Integer
    Dim i, c1, c2 As Integer
    i = -1
    j = 0
    For Each v In statement
        i = i + 1
        If (Not v Like "//*") Then
            If v Like "for ?* :: ?* >> ?*" Then
                c1 = c1 + 1
                ForCount(c1, 1) = i
            End If
            If v Like "next*" Then
                c2 = c2 + 1
                ForCountReverse(c2) = i
            End If
        End If
    Next v
    If c1 <> c2 Then
        Slide1.errcount = Slide1.errcount + 1
    Else
        If c1 = 0 Then Exit Sub
        For i = c2 To 1 Step -1
            j = j + 1
            ForCount(j, 2) = ForCountReverse(i)
        Next
        For j = 1 To c1
            ''MsgBox ForCount(j, 1) & ">" & ForCount(j, 2)
        Next
        maxFor = c1
    End If
End Sub
Public Sub CountIf()
    Dim v As Variant
    Dim ForCountReverse(1 To 1000000) As Integer
    Dim i, c1, c2 As Integer
    i = -1
    j = 0
    For Each v In statement
        i = i + 1
        If (Not v Like "//*") Then
            If v Like "if ?*" Then
                c1 = c1 + 1
                IfCount(c1, 1) = i
            End If
            If v Like "endif" Then
                c2 = c2 + 1
                ForCountReverse(c2) = i
            End If
        End If
    Next v
    If c1 <> c2 Then
        Slide1.errcount = Slide1.errcount + 1
    Else
        If c1 = 0 Then Exit Sub
        For i = c2 To 1 Step -1
            j = j + 1
            IfCount(j, 2) = ForCountReverse(i)
        Next
        For j = 1 To c1
            ''MsgBox IfCount(j, 1) & ">" & IfCount(j, 2)
        Next
        maxIf = c1
    End If
End Sub

Public Function GetNode(ByVal nodeArray As String, ByVal lineStatement As Integer, ByVal typeReturn As String)
    If nodeArray = "for" Then
        For i = 1 To maxFor
            If ForCount(i, 1) = lineStatement Then
                If typeReturn = "end" Then
                    GetNode = ForCount(i, 2)
                Else
                    GetNode = i
                End If
                Exit Function
            End If
        Next
    Else
    End If
End Function
''
Public Sub resetVariables()
'Dim i As Integer
'Dim v As Variant
'Dim s() As String
's = Split(Slide1.Shapes("$$Saysettha~~VariablesStack").TextFrame2.TextRange.Text, ",")
'For i = 1 To ArrayLen(s) - 2
'    Slide1.Shapes(s(i)).Delete
'Next
'For i = 1 To ArrayLen(s) - 2
'    Slide1.Shapes(s(i)).Delete
'Next
    Dim s As Shape
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
End Sub


Public Sub BuildProject(ByVal codeInput As String)
    ' Xu ly cac so lieu  va reset cac thong so
    Slide1.errcount = 0
    Slide1.console = "" ''"Saysettha Pre-processed Language (C) 2017 - 2022" & vbNewLine
    Dim v As Variant
    statement = Split(codeInput, vbNewLine)
    line = ArrayLen(statement)
    i = -1
    For Each v In statement
        i = i + 1
        statement(i) = StringRefine.Trim(v)
    Next v
    Slide1.Shapes("$$Saysettha~~VariablesStack").TextFrame2.TextRange.Text = ","
    
    ''Call CountFor
    ''Call errorHandle
    ''Call CountIf
    ''Call errorHandle
    Call resetVariables
    'MsgBox "1"
    ' Xu ly code
    Dim kHandle As String
    
    For i = 0 To line - 1
        If Not statement(i) Like "//*" And Not statement(i) = "" Then
            If statement(i) Like "$?*" Then
                getReturnOneTime = Variables.Confirm(statement(i))
                Select Case getReturnOneTime
                    Case 0
                        getReturnOneTime = Variables.DeclareVariable(statement(i))
                        If getReturnOneTime = 2 Then
                            Report.ErrorHandle ("Warning :: Compile Error: declaring variable [" & _
                            StringRefine.Trim(statement(i)) & "] in line " & CStr(i + 1) & " is invalid. ")
                            Exit For
                        ElseIf getReturnOneTime = 3 Then
                            Report.ErrorHandle ("Warning :: Compile Error: Math error [" & _
                            StringRefine.Trim(statement(i)) & "] in line " & CStr(i + 1) & " is invalid. ")
                            Exit For
                        ElseIf getReturnOneTime Like "value_assign:?*" Then
                            kHandle = StringRefine.Trim(Split(statement(i), "=", 2)(1))
                            Report.ErrorHandle ("Warning :: Compile Error: assigning with the [" & _
                            kHandle & "] value in line " & CStr(i + 1) & " is invalid. ")
                            Exit For
                        End If
                    Case 2
                        Report.ErrorHandle ("Warning :: Compile Error: declaring variable [" & _
                        StringRefine.Trim(statement(i)) & "] in line " & CStr(i + 1) & " is invalid. ")
                        Exit For
                End Select
            ElseIf StringRefine.CustomRegexChecker(statement(i), "^die[ ]*[(]*[ ]*[)]*[ ]*$") = 0 Then
                Exit For
            ElseIf statement(i) Like "cout*" Then
                If StringRefine.CustomRegexChecker(statement(i), "^cout[ ]*<[ ]*.*") = 0 Then
                    getReturnOneTime = Cout.Process(statement(i))
                    If getReturnOneTime = "quote" Then
                        Report.ErrorHandle ("Warning :: Compile Error: missing expected quote(s) in [" & _
                        StringRefine.Trim(statement(i)) & "], line " & CStr(i + 1) & " is invalid. ")
                        Exit For
                    ElseIf getReturnOneTime Like "quote:?*" Then
                        kHandle = StringRefine.Trim(Split(getReturnOneTime, ":", 2)(1))
                        Report.ErrorHandle ("Warning :: Compile Error: invalid [" & _
                        kHandle & "] string in line " & CStr(i + 1) & ". ")
                        Exit For
                    ElseIf getReturnOneTime Like "variable:?*" Then
                        kHandle = StringRefine.Trim(Split(getReturnOneTime, ":", 2)(1))
                        Report.ErrorHandle ("Warning :: Compile Error: invalid [" & _
                        kHandle & "] variable in line " & CStr(i + 1) & ". ")
                        Exit For
                    End If
                Else
                    Report.ErrorHandle ("Warning :: Compile Error: calling the [" & _
                    StringRefine.Trim(statement(i)) & "] statement in line " & CStr(i + 1) & " is invalid. ")
                    Exit For
                End If
            End If
        End If
    Next
    ConsoleWrite (vbNewLine & vbNewLine & vbNewLine & "Program exits with the 0 code . . .")
End Sub

