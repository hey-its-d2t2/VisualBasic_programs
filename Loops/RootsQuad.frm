VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Roots of quadratic equation "
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
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
   ScaleHeight     =   4950
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Width           =   105
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C = "
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   90
      TabIndex        =   3
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "B = "
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   135
      TabIndex        =   2
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "A = "
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim R1 As Double
    Dim R2 As Double
    Dim RP As Double
    Dim IP As Double
    A = Val(Text1.Text())
    B = Val(Text2.Text())
    C = Val(Text3.Text())
    Dim dis As Double
    dis = B * B - 4 * A * C
    If dis > 0 Then
        R1 = (-B + Math.Sqr(dis) / (2 * A))
        R2 = (-B - Math.Sqr(dis) / (2 * A))
        Label4.Caption = "Root 2 = " & R1 + "Root 2 = " & R2
    End If
        
    
    
    

End Sub
