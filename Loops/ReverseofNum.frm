VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Reverse  of number"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Click ME"
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "   "
      Height          =   585
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "    "
      Height          =   585
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "N =  "
      Height          =   585
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim N As Integer, Reminder As Integer, Rev As Integer
    N = Val(txt1.Text())
   Do While N <> 0
    Reminder = N Mod 10
     Rev = Rev * 10 + Reminder
     N = N / 10
   Loop
   Label2.Caption = "Reverse "
   Label3.Caption = Rev
    
    
End Sub
