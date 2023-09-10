VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "N = "
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x, g, n, i, sum As Integer
    n = Val(Text1.Text)
    x = 0
    y = 1
    Print x;
    Print y;
    For i = 3 To n
    sum = x + y
    Print sum;
    x = y
    y = sum
    y = sum
    Next i
End Sub
