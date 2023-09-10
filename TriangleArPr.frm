VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Triangle "
   ClientHeight    =   7275
   ClientLeft      =   8130
   ClientTop       =   2055
   ClientWidth     =   4890
   Icon            =   "TriangleArPr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   4890
   Begin VB.TextBox txtHeight 
      BackColor       =   &H8000000D&
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
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtBase 
      BackColor       =   &H8000000D&
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
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3120
      Width           =   2415
   End
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
      Left            =   1920
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      X1              =   2160
      X2              =   4560
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label lblResultAr 
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
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000E&
      X1              =   1800
      X2              =   4200
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   1800
      X2              =   4200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Height : "
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
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Base : "
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
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Triangle"
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
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   2400
      X2              =   3600
      Y1              =   360
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   2400
      X2              =   1200
      Y1              =   360
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   3600
      X2              =   1200
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim B#, H#
    Dim Ar#
    B = Val(txtBase.Text)
    H = Val(txtHeight.Text)
    Ar = (B * H) / 2
    lblResultAr.Caption = Ar
    
End Sub
