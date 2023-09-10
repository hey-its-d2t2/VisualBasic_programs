VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SumAverage"
   ClientHeight    =   5700
   ClientLeft      =   8130
   ClientTop       =   3315
   ClientWidth     =   5775
   Icon            =   "SumAverage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5775
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FF0000&
      Caption         =   "Calculate"
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
      Left            =   1920
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.TextBox txtB 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   2
      ToolTipText     =   "Enter value of A "
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   1
      ToolTipText     =   "Enter value of A "
      Top             =   960
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      X1              =   1320
      X2              =   4800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   1320
      X2              =   4800
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   5520
      Width           =   5775
   End
   Begin VB.Label lblAverageResult 
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
      Height          =   735
      Left            =   2280
      TabIndex        =   9
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Average :"
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
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label lblSumResult 
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
      Height          =   735
      Left            =   1560
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label lblSum 
      BackStyle       =   0  'Transparent
      Caption         =   "Sum :"
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
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "B :"
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
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblMsgA 
      BackStyle       =   0  'Transparent
      Caption         =   "A :"
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
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Sum && Average"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim A As Double
    Dim B As Double
    Dim S#, Av#
    A = Val(txtA.Text)
    B = Val(txtB.Text)
    S = (A + B)
    Av = S / 2
    lblSumResult.Caption = S
    lblAverageResult.Caption = Av
End Sub


