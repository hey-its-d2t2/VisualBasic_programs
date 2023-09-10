VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7170
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
   ScaleHeight     =   5280
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   1095
      Left            =   2160
      TabIndex        =   0
      Top             =   3240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim N As Integer
Dim Cbin As Long
Dim i As Double
Dim c As Long
N = Val(InputBox("Enter Integer Number", "Input", 12))
If N = 0 Then
        Cbin = 0
    ElseIf N > 0 Then
        i = 2 ^ CLng(Log(N) / Log(2) + 0.1)
        Do While i >= 1
            c = Fix(N / i)
            Cbin = Cbin & c
            N = N - i * c
            i = i / 2
        Loop
    End If
    MsgBox "Binary = " & Cbin, vbOKCancel, "Result"
End Sub
