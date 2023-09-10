VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pattern 10"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9975
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
   ScaleHeight     =   7335
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click ME"
      Height          =   735
      Left            =   7440
      TabIndex        =   0
      Top             =   6120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Me.FontSize = 32
    For i = 1 To 5
        For k = 1 To 5 - i
            Print "    ";
        Next
        For j = 1 To i - 1
            Print j;
        Next
        For j = i To 1 Step -1
        Print j;
        Next
    Print
    Next
End Sub
