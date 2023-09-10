VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1 To N"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Segoe Fluent Icons"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "1toNVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Ok"
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      ItemData        =   "1toNVer.frx":0442
      Left            =   120
      List            =   "1toNVer.frx":0444
      TabIndex        =   0
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   4320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Last Digit :"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2745
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1 To N"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2572
      TabIndex        =   2
      Top             =   90
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6255
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
    N = Val(Text1.Text)
    If KeyAscii = 13 Then
    Command1.SetFocus
    For I = 1 To N
        List1.AddItem I
    Next
    End If
End Sub
