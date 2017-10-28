VERSION 5.00
Begin VB.Form rules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rules"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "rules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Label1.Caption = "Points : " & Chr(13) & "Correct Answer : +10" & Chr(13) & "Wrong Answer : -10" & Chr(13) & "50-50 : -5" & Chr(13) & "You can use flip to get a new question worth 5 points" & Chr(13) & "You can skip a question without penalty"
End Sub

Private Sub OKButton_Click()
Me.Hide
End Sub
