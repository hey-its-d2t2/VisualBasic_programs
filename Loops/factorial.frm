VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Factorial "
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   120
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "N = "
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim n As Long
    Dim i As Integer, f As Integer
    n = Val(Text1.Text())
    f = 1
    For i = 1 To n
        f = f * i
    Next
    Label2.Caption = "Factorial = " & f
End Sub

