Attribute VB_Name = "Module1"
Sub main()
'MsgBox "Hi"
'Dim N As Integer
Dim n%
n = Val(InputBox("Enter any Nuber", "N = ", "5"))
'n = Val(txtBox1.Text())
If n > 0 Then
    MsgBox "+ve Number"
ElseIf n < 0 Then
    MsgBox "-Ve Number"
Else
    MsgBox " Number is 0"
End If

End Sub
