VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rectangle"
   ClientHeight    =   6555
   ClientLeft      =   6915
   ClientTop       =   4155
   ClientWidth     =   5040
   Icon            =   "AreaPeriRectangle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCalculate 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   1440
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtBreadth 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtlength 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblResultPerimeter 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   2160
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblResultArea 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Perimeter : "
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
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Rectangle"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   2040
      X2              =   4320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   2040
      X2              =   4320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblBreadth 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Breadth : "
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
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblLength 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Length : "
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
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   3840
      X2              =   3840
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   3840
      X2              =   3840
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   2640
      X2              =   3480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   1440
      X2              =   2280
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   485
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      Height          =   975
      Left            =   1440
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim L#, B#
    Dim Ar#, Pr#
    L = Val(txtlength.Text)
    B = Val(txtBreadth.Text)
    Ar = L * B
    Pr = 2 * (L + B)
    lblResultArea.Caption = Ar
    lblResultPerimeter.Caption = Pr
End Sub
