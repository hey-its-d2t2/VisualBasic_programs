VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Your Name N Time"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Segoe Fluent Icons"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NameUpToN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Height          =   3510
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   3615
   End
   Begin VB.CommandButton cmdClick 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Click Me"
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtN 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   3240
      MaxLength       =   100
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   4560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Tast Number : "
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3090
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Print Your Name N Times"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   465
      TabIndex        =   0
      Top             =   150
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   735
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
        List1.AddItem "Deepak Singh"
    Next
End Sub
Private Sub txtN_KeyPress(KeyAscii As Integer)

End Sub
