VERSION 5.00
Begin VB.Form GQ3 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question 2"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   21034.95
   ScaleMode       =   0  'User
   ScaleWidth      =   10172.31
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   8640
      Width           =   6615
   End
   Begin VB.CommandButton pss 
      Caption         =   "Pass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   7
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton flp 
      Caption         =   "Flip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      TabIndex        =   6
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton option4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   6615
   End
   Begin VB.CommandButton fifty 
      Caption         =   "50-50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton option1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   6615
   End
   Begin VB.CommandButton option2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   6615
   End
   Begin VB.CommandButton option3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   6615
   End
   Begin VB.Label question 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "GQ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
result.Show
End Sub

Private Sub fifty_Click()
option1.Visible = False
option3.Visible = False
score = score - 5
End Sub

Private Sub Form_Load()
question.Caption = "Gameworks is a suite of gaming technologies made by which company ?"
option1.Caption = "AMD"
option2.Caption = "Nvidia"
option3.Caption = "Qualcomm"
option4.Caption = "Microsoft"
End Sub

Private Sub option1_Click()
option1.BackColor = RGB(256, 0, 0)
option2.BackColor = RGB(0, 256, 0)
option1.Enabled = False
option2.Enabled = False
option3.Enabled = False
option4.Enabled = False
fifty.Enabled = False
pss.Enabled = False
flp.Enabled = False
score = score - 10
Sleep 2000
Me.Hide
GQ3.Show
End Sub

Private Sub option2_Click()
option2.BackColor = RGB(0, 256, 0)
option1.Enabled = False
option2.Enabled = False
option3.Enabled = False
option4.Enabled = False
fifty.Enabled = False
pss.Enabled = False
flp.Enabled = False
score = score + 10
Sleep 2000
Me.Hide
GQ3.Show
End Sub

Private Sub option3_Click()
option3.BackColor = RGB(256, 0, 0)
option2.BackColor = RGB(0, 256, 0)
option1.Enabled = False
option2.Enabled = False
option3.Enabled = False
option4.Enabled = False
fifty.Enabled = False
pss.Enabled = False
flp.Enabled = False
score = score - 10
Sleep 2000
Me.Hide
GQ3.Show
End Sub

Private Sub option4_Click()
option4.BackColor = RGB(256, 0, 0)
option2.BackColor = RGB(0, 256, 0)
option1.Enabled = False
option2.Enabled = False
option3.Enabled = False
option4.Enabled = False
fifty.Enabled = False
pss.Enabled = False
flp.Enabled = False
score = score - 10
Sleep 2000
Me.Hide
GQ3.Show
End Sub

Private Sub pss_Click()
option2.BackColor = RGB(0, 256, 0)
option1.Enabled = False
option2.Enabled = False
option3.Enabled = False
option4.Enabled = False
fifty.Enabled = False
flp.Enabled = False
pss.Enabled = False
Sleep 5000
Me.Hide
result.Show
End Sub

