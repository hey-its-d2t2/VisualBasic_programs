VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Odd Numbers up to N"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4650
   FillColor       =   &H0000FF00&
   BeginProperty Font 
      Name            =   "Segoe Fluent Icons"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OddUpToN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   3075
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton cmdClick 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Click Me"
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtN 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Last Digit :"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Odd Numbers Up To N"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   660
      TabIndex        =   0
      Top             =   90
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClick_Click()
    Dim N%, I%
    N = Val(txtN.Text())
    For I = 1 To N
        If I Mod 2 > 0 Then
        List1.AddItem I
        End If
    Next
End Sub

Private Sub txtN_GotFocus()
    txtN.SelStart = 0
    txtN.SelLength = Len(txtN.Text())
    txtN.Text = ""
    List1.Clear
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdClick.SetFocus
        cmdClick_Click
    End If
End Sub
