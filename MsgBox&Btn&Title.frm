VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   9810
   ClientTop       =   5325
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'MsgBox as a Subroutine

   ' MsgBox "Hello!" & Chr(13) & "Good Morning Sir..."
    'A = 10
    'B = 20
    'c = A + B
    'MsgBox "The Sum Of " & A & " and " & B & " is  " & c
    ' MsgBox "Hello Tilte", , "My Title"
    ' MsgBox "Hello Button", vbRetryCancel, "My Title"
    ' MsgBox "Hello", vbYesNo + 32, "My Title"
   ' MsgBox "Hello", vbAbortRetryIgnore + vbCritical, "My Title"
    'MsgBox "Hello Default Button ", vbAbortRetryIgnore + vbCritical + vbDefaultButton3
   ' MsgBox "Hello Modility", 36 + vbSystemModal, "My Title"
 
 ' MsgBos as a Function
    
    Ans = MsgBox("My Msg Function", vbOKCancel + vbQuestion, "My Title")
    If Ans = vbOK Then
        MsgBox ("Hi You Pressed Ok ")
        
    End If
    
End Sub
