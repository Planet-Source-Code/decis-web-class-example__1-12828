Attribute VB_Name = "mod_Strings"
Function ChangeCase(ValueIn As String) As String
Dim i As Integer
' Change case of string i.e. HELLO becomes Hello
    For i = 1 To Len(ValueIn)
        If i = 1 Then
            ChangeCase = UCase((Mid$(ValueIn, i, 1)))
        ElseIf Mid$(ValueIn, i, 1) = " " Then
            ChangeCase = ChangeCase & Mid$(ValueIn, i, 1)
            i = i + 1
            ChangeCase = ChangeCase & UCase(Mid$(ValueIn, i, 1))
        Else
            ChangeCase = ChangeCase & LCase(Mid$(ValueIn, i, 1))
        End If
    Next i
End Function
