VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Leap Year Checker"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6105
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Click Me"
      BeginProperty Font 
         Name            =   "Segoe Fluent Icons"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim X%, Y%, Z%, Year%
    Me.Font.Size = 30
    Year = Val(InputBox("Eter Year", "Input Year", 2022))
    X = Year Mod 4
    Y = Year Mod 100
    Z = Year Mod 400
    If ((X = 0 And Not (Y = 0)) Or Z = 0) Then
        MsgBox "This is a leap year", vbOKCancel, "Result"
    Else
        MsgBox "This is not a leap year", vbOKCancel, "Result"
    End If

End Sub
