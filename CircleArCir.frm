VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circle"
   ClientHeight    =   6960
   ClientLeft      =   7725
   ClientTop       =   4155
   ClientWidth     =   4860
   Icon            =   "CircleArCir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H8000000D&
      Caption         =   "Calculate "
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
      Left            =   1200
      MaskColor       =   &H8000000D&
      TabIndex        =   3
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtRadious 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00000"
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
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   4800
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label lblResultCir 
      Alignment       =   2  'Center
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Circumference: "
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   3840
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label lblResultArea 
      Alignment       =   2  'Center
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Area : "
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   4320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Radious : "
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblMsgCircle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Circle"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1935
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim R#
    Dim Ar#, Cir#
    R = Val(txtRadious.Text)
    Ar = 3.141 * R * R
    Cir = 2 * 3.141 * R
    lblResultArea.Caption = Ar
    lblResultCir.Caption = Cir
End Sub
