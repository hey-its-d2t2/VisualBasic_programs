VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6810
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
   ScaleHeight     =   5295
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   3240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim N%
    N = Val(InputBox("Enter number", "Enter value", 10))
    MsgBox "Factorial = " & Fact(N), vbOKCancel, "Result"
End Sub
Private Function Fact(N)
    Dim I%, F%
    F = 1
    For I = 1 To N
        F = F * I
    Fact = F
    Next I
End Function
