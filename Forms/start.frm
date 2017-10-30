VERSION 5.00
Begin VB.Form welcome 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome To Heritage Quizmaster"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   855
      Left            =   1440
      TabIndex        =   6
      Top             =   6720
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Check Previous Scores"
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   5760
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   855
      Left            =   1440
      TabIndex        =   3
      Top             =   7680
      Width           =   4215
   End
   Begin VB.TextBox Nm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   4200
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proceed"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   4800
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "ENTER YOUR NAME "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   1320
      TabIndex        =   5
      Top             =   3720
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
score = 0
ques = 0
nme = Nm.Text
If Nm.Text = "" Then
MsgBox "Enter Your Name First !", , "ERROR !"
Else
topics.Show
Me.Hide
End If
End Sub

Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
scores.Show
End Sub
Private Sub Command4_Click()
MsgBox "Designed and coded by Garima Sancheti, Vatsal Agarwal, Shounak Dawn, Suvoparno Banerjee" & Chr(13) & "Created using Visual Basic 6", , "About Us"
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture("C:\Documents and Settings\Suvoparno\Desktop\Project\herlogo.bmp")
Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, Picture1.Picture.Width / 26.46, _
    Picture1.Picture.Height / 26.46
    Picture1.Picture = Picture1.Image
End Sub

