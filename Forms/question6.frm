VERSION 5.00
Begin VB.Form SQ2 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question 2"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   17685.45
   ScaleMode       =   0  'User
   ScaleWidth      =   10172.31
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
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
      TabIndex        =   0
      Top             =   6960
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
      Left            =   4200
      TabIndex        =   8
      Top             =   5640
      Width           =   2535
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
      Left            =   2760
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   3255
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
      TabIndex        =   6
      Top             =   5640
      Width           =   2535
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
      TabIndex        =   2
      Top             =   3240
      Width           =   3255
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
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
      TabIndex        =   4
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Time Remaining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   6855
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
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
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "SQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Integer
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Sub Command1_Click()
Me.Hide
result.Show
End Sub

Private Sub fifty_Click()
option3.Visible = False
option4.Visible = False
score = score - 5
End Sub

Private Sub flp_Click()
flip = 1
Me.Hide
FSQ1.Show
End Sub

Private Sub Form_Load()
If flip = 1 Then
    flp.Visible = False
End If
last = last + 1
question.Caption = "Which of the following is used in pencils?"
option1.Caption = "Graphite"
option2.Caption = "Silicon"
option3.Caption = "Charcol"
option4.Caption = "Phosphorous"
time = 60
End Sub

Private Sub option1_Click()
option1.BackColor = RGB(0, 256, 0)
option1.Enabled = Falsef
option2.Enabled = False
option3.Enabled = False
option4.Enabled = False
fifty.Enabled = False
pss.Enabled = Faflse
flp.Enabled = False
score = score + 10
Sleep 2000
Me.Hide
SQ3.Show
End Sub

Private Sub option2_Click()
option2.BackColor = RGB(256, 0, 0)
option1.BackColor = RGB(0, 256, 0)
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
SQ3.Show
End Sub

Private Sub option3_Click()
option3.BackColor = RGB(256, 0, 0)
option1.BackColor = RGB(0, 256, 0)
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
SQ3.Show
End Sub

Private Sub option4_Click()
option1.BackColor = RGB(0, 256, 0)
option4.BackColor = RGB(256, 0, 0)
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
SQ3.Show
End Sub

Private Sub pss_Click()
option1.BackColor = RGB(0, 256, 0)
option1.Enabled = False
option2.Enabled = False
option3.Enabled = False
option4.Enabled = False
fifty.Enabled = False
flp.Enabled = False
Sleep 5000
Me.Hide
SQ3.Show
End Sub

Private Sub Timer1_Timer()
Label1.Caption = time
time = time - 1
If time = -1 Then
    MsgBox "Game Over !", , "Time's Up!"
    option3.BackColor = RGB(0, 256, 0)
    Timer1.Enabled = False
    option1.Enabled = False
    option2.Enabled = False
    option3.Enabled = False
    option4.Enabled = False
    fifty.Enabled = False
    pss.Enabled = False
    flp.Enabled = False
End If
End Sub
