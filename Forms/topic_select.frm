VERSION 5.00
Begin VB.Form topics 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Topics"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      BackColor       =   &H00000000&
      Caption         =   "Back"
      Height          =   975
      Left            =   2280
      TabIndex        =   5
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton rule 
      Caption         =   "Rules"
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Gaming"
      Height          =   975
      Left            =   3600
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Environment and Science"
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Current Affairs"
      Height          =   975
      Left            =   3600
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IT"
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "topics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
GQ1.Show
Me.Hide
End Sub

Private Sub Command6_Click()
Me.Hide
welcome.Show
End Sub

Private Sub Form_Load()
Label1.Caption = "Welcome " & nme & Chr(13) & Chr(13) & "Please select a topic"
End Sub

Private Sub rule_Click()
rules.Show
End Sub
