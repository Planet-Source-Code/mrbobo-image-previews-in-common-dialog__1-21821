VERSION 5.00
Begin VB.Form frmPic 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   1455
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   -60
      Width           =   4575
      Begin VB.Label lblh 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height : Not Available"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblw 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width : Not Available"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Size : Not Available"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1665
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2520
      Top             =   1080
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1335
      Left            =   4800
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   30
      Width           =   1335
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   1560
      Left            =   2640
      Picture         =   "frmPic.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1560
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the form we overlay onto
'the standard Commondialog after
'first increasing the height of the
'Commondialog
Public Sub LoadImage(mfile As String)
'Load the picture the user clicked on in the
'Commondialog to display a preview
Picture1.Picture = LoadPicture(mfile)
SizeImage
lblh = "Image Height :" + Str(Picture1.Height / Screen.TwipsPerPixelY) + " Pixels"
lblw = "Image Width :" + Str(Picture1.Width / Screen.TwipsPerPixelX) + " Pixels"
lblSize = "File Size :" + Str(Format(FileLen(mfile) / 1024, "0.00")) + " Kb"

End Sub

Private Sub Form_Load()
SizeImage
End Sub
Private Sub Form_Resize()
Picture2.Left = Me.Width - Picture2.Width
End Sub

Private Sub Timer1_Timer()
'this call keeps this form in the correct
'position in relation to the Commondialog
'and ontop of the Commondialog but not ontop of
'other windows
MovePics
End Sub

Private Sub SizeImage()
'Create the thumbnail picture of the selected file
If Picture1.Width >= Picture1.Height Then
    Image1.Height = 1095 * (Picture1.Height / Picture1.Width)
Else
    Image1.Width = 1095 * (Picture1.Width / Picture1.Height)
End If
Image1.Left = Picture2.Width / 2 - Image1.Width / 2
Image1.Top = Picture2.Height / 2 - Image1.Height / 2
Image1.Picture = Picture1.Image

End Sub
