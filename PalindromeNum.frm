VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Segoe Fluent Icons"
      Size            =   30
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   2760
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim R%, N%, Temp%, Sum%
    N = Val(InputBox("Enter number", "EnterValue", 10))
    Sum = 0
    Temp = N
    While N > 0
        R = N Mod 10
        Sum = (Sum * 10) + R
        N = N / 10
    Wend
    If Temp = Sum Then
        MsgBox "Nuber is palindeome ", vbOKCancel, "Result"
    Else
        MsgBox "Nuber is not palindeome ", vbOKCancel, "Result"
    End If
End Sub
