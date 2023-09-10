VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtBox1 
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Height          =   1335
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter A Number :"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'MsgBox "Hi"
'MsgBox "Hi"
'Dim N As Integer
Dim n%
'n = Val(InputBox("Enter any Nuber", "N = ", "5"))
n = Val(txtBox1.Text())
If n > 0 Then
    MsgBox "+ve Number"
ElseIf n < 0 Then
    MsgBox "-Ve Number"
Else
    MsgBox " Number is 0"
End If

End Sub


