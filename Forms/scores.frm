VERSION 5.00
Begin VB.Form scores 
   Caption         =   "Previous Scores"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   5040
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   4545
      ItemData        =   "scores.frx":0000
      Left            =   120
      List            =   "scores.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "scores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim sFileText As String
Dim iFileNo As Integer
iFileNo = FreeFile
Open "C:\score.txt" For Input As #iFileNo
    Do While Not EOF(iFileNo)
        Input #iFileNo, sFileText
        List1.AddItem sFileText
    Loop
Close #iFileNo
End Sub

Private Sub List1_Click()
Command1.SetFocus
List1.ListIndex = -1
End Sub

Private Sub List1_DblClick()
Call List1_Click
End Sub

