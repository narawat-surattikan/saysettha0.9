Attribute VB_Name = "StringRefine"
Public Function Trim(ByVal str As String) As String
    If Len(str) <> 0 Then
        While Mid(str, 1, 1) = " " Or Mid(str, 1, 1) = vbTab
            str = Right(str, Len(str) - 1)
        Wend
        If Len(str) = 0 Then Exit Function
        While Mid(str, Len(str), 1) = " " Or Mid(str, 1, 1) = vbTab
            str = Left(str, Len(str) - 1)
        Wend
        Trim = str
    Else
        Trim = ""
    End If
End Function

Public Function VerifyUserdefinedName(ByVal strIn As String) As String
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Pattern = "^\w+$"
        If .Test(strIn) = True Then
            VerifyUserdefinedName = 0
        Else
            VerifyUserdefinedName = 1
        End If
    End With
End Function
Public Function VerifyNumber(ByVal strIn As String) As String
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Pattern = "^-?\d*[.]?[0-9]+$"
        If .Test(strIn) = True Then
            VerifyNumber = 0
        Else
            VerifyNumber = 1
        End If
    End With
End Function
Public Function CustomRegexChecker(ByVal strIn As String, ByVal strPattern As String) As String
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Pattern = strPattern
        If .Test(strIn) = True Then
            CustomRegexChecker = 0
        Else
            CustomRegexChecker = 1
        End If
    End With
End Function
Public Function VerifyStringValue(ByVal strIn As String)
    strIn = Replace(strIn, "\" & Chr(34), "<replace>")
    If InStr(1, strIn, Chr(34)) <> 0 Then
        VerifyStringValue = 1
    Else
        VerifyStringValue = 0
    End If
End Function


