VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cdlg with Pictures"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Returned Path"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sfile As String
Text1.Text = ""
'just a single line to call
sfile = ShowOpen(Me.hwnd, "Image files " + Chr(0) + "*.bmp;*.gif;*.jpg", 5, , "Cdlg with Pictures", True)
If sfile <> "" Then Text1.Text = sfile

End Sub
