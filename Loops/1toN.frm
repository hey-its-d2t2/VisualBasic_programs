VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   Caption         =   "1 To N"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5040
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
   ScaleHeight     =   6060
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   3000
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ok"
      Height          =   735
      Left            =   1440
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Last Digit :"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "1 to N"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1950
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim N%, I%
    N = Val(Text1.Text)
    For I = 1 To N
        List1.AddItem I
    Next
End Sub
Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text())
    Text1.Text = ""
    List1.Clear
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim N%, I%
    If KeyAscii = 13 Then
        Command1.SetFocus
        N = Val(Text1.Text)
        For I = 1 To N
            List1.AddItem I
      Next
    End If
End Sub

