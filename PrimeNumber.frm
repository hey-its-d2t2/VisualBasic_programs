VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6960
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
   ScaleHeight     =   4500
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim Check As Integer
  Dim Num As Integer
  Check = 1
  Num = Val(InputBox("Enter Number ", "Enter Value", 10))
  For i = 2 To (Num - 1)
    If Num Mod i = 0 Then
        Check = 0
        Exit For
    End If
        Next
    If Check = 0 Then
        MsgBox "Not a prime number", vbOKCancel, "Result"
    Else
        MsgBox "A prime number", vbOKCancel, "Result"
    End If
  
End Sub
