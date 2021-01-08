Option Explicit
Dim name

name = InputBox("What is your age?" , "Information:")

If IsNumeric(name) And name="15" then
msgbox "correct"
ElseIf name = "" Then
MsgBox "please do not leave blank."
ElseIf Not IsNumeric(name) then
msgbox "please type in a number value."
ElseIf Not name = "15" then
msgbox "wrong age"
End If