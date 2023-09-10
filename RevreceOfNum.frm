VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sum of Digit"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4650
   Icon            =   "RevreceOfNum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      MaskColor       =   &H000000FF&
      Picture         =   "RevreceOfNum.frx":0442
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtN 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   960
      MaxLength       =   10
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   1680
      X2              =   3840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Result : "
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   960
      X2              =   4080
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N : "
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rev()

    Dim N%, R%, S%
    S = 0
    N = Val(txtN.Text)
     While N > 0
            R = N Mod 10
            Sum = Sum + N
            N = N / 10
    End While
End Sub
Private Sub Command1_Click()
    
    rev()
    
End Sub

