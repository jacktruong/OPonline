VERSION 5.00
Begin VB.Form frmBasicMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Move Cards"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNo 
      Caption         =   "No"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton optDropType 
      Caption         =   "Add to TOP of Draw Pile"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.OptionButton optDropType 
      Caption         =   "Add to BOTTOM of Draw Pile"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.OptionButton optDropType 
      Caption         =   "Add and then shuffle Draw Pile"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblMoveType 
      Caption         =   "Move cards from Discard Pile to Draw Pile:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmBasicMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
chkNo.Value = 1
Me.Hide

End Sub

Private Sub OKButton_Click()
Me.Hide

End Sub
