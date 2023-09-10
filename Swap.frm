VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Swap"
   ClientHeight    =   4365
   ClientLeft      =   9750
   ClientTop       =   3735
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Segoe Fluent Icons"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Swap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5850
   Begin VB.CommandButton cmdSwap 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Swap"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MaskColor       =   &H000000C0&
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtB 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
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
      Height          =   735
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtA 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
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
      Height          =   735
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblResultB 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblResultA 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   " B  :"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   " A  :"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   " B  :"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   " A  :"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSwap_Click()
    Dim A%, B%
    A = Val(txtA.Text)
    B = Val(txtB.Text)
    A = A + B
    B = A - B
    A = A - B
    lblResultA.Caption = A
    lblResultB.Caption = B
End Sub
