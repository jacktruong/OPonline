VERSION 5.00
Begin VB.Form frmChooser 
   Caption         =   "Choose Image Type to Import"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "New Special"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Images"
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Events"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Missions"
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "No Pic Specials"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdBasicUniverse 
      Caption         =   "Basic U"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdTraining 
      Caption         =   "Training"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDoubleshot 
      Caption         =   "Doubleshot"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdTeamwork 
      Caption         =   "Teamwork"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAspect 
      Caption         =   "Aspect"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdArtifact 
      Caption         =   "Artifact"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSpecial 
      Caption         =   "Special"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdPowerCard 
      Caption         =   "Power Card"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHomeBase 
      Caption         =   "Homebase"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCharacter 
      Caption         =   "Character"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlly 
      Caption         =   "Ally"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlly_Click()
frmAllyImage.Show 1

End Sub

Private Sub cmdArtifact_Click()
frmArtifactImages.Show 1

End Sub

Private Sub cmdAspect_Click()
frmAspectImages.Show 1

End Sub

Private Sub cmdBasicUniverse_Click()
frmBUImages.Show 1

End Sub

Private Sub cmdCharacter_Click()
frmCharacterImage.Show 1

End Sub

Private Sub cmdDoubleshot_Click()
frmDoubleshotimages.Show 1

End Sub

Private Sub cmdHomeBase_Click()
frmHomeBase.Show 1

End Sub

Private Sub cmdPowerCard_Click()
frmPowerCard.Show 1

End Sub

Private Sub cmdSpecial_Click()
frmSpecialImages.Show 1

End Sub

Private Sub cmdTeamwork_Click()
frmTeamworkimages.Show 1

End Sub

Private Sub cmdTraining_Click()
frmTrainingImages.Show 1

End Sub

Private Sub Command1_Click()
frmSpecialImages2.Show 1

End Sub

Private Sub Command2_Click()
frmMissionImages.Show 1

End Sub

Private Sub Command3_Click()
frmEventImages.Show 1

End Sub

Private Sub Command4_Click()
frmCheckSpecialImages.Show 1

End Sub

Private Sub Command5_Click()
frmNewSpecial.Show 1

End Sub

