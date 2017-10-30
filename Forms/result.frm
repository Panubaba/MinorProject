VERSION 5.00
Begin VB.Form result 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5970
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Proceed"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Here's Your Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   5415
      End
      Begin VB.Label label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5415
      End
   End
End
Attribute VB_Name = "result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Unload result
topics.Show
End Sub

Private Sub Form_Load()
If ques = 4 Then
Command2.Visible = False
End If
Label1.Caption = "Hello " & nme
Label2.Caption = "Your score is " & score
Dim sFileText As String
Dim iFileNo As Integer
iFileNo = FreeFile
Open "C:\score.txt" For Append As #iFileNo
    Print #iFileNo, "Name " & nme
    Print #iFileNo, "Score " & score
Close #iFileNo
End Sub

