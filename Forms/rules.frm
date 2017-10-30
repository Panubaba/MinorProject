VERSION 5.00
Begin VB.Form rules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rules"
   ClientHeight    =   3765
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "The rules of this game are"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.Label Label1 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "rules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Label1.Caption = "You have 60 seconds to answer" & Chr(13) & Chr(13) & "Points : " & Chr(13) & "Correct Answer : +10" & Chr(13) & "Wrong Answer : -10" & Chr(13) & "50-50 : -5" & Chr(13) & Chr(13) & "You can use flip to get a new question worth 5 points once" & Chr(13) & "You can skip a question without penalty"
End Sub

Private Sub OKButton_Click()
Me.Hide
End Sub
