VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "Testing of Massage Box", 3 + 16 + 512, "Testing of Title"
End Sub
