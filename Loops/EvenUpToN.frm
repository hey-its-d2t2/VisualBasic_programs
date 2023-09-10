VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Even Num Up To N"
   ClientHeight    =   6690
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
   Icon            =   "EvenUpToN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   3075
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton cmdClick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Click Me"
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   4560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Last Digit :"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Even Num Up To N"
      Height          =   435
      Left            =   667
      TabIndex        =   0
      Top             =   150
      Width           =   3360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF00FF&
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
        If I Mod 2 = 0 Then
            List1.AddItem I
        End If
    Next
End Sub
Private Sub txtN_GotFocus()
    txtN.Text = ""
    cmdClick_Click
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        For I = 1 To N
        If I Mod 2 = 0 Then
            List1.AddItem I
        End If
    Next
    cmdClick.SetFocus
        End If
    
End Sub
