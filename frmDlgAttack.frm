VERSION 5.00
Begin VB.Form frmDlgAttack 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Additional Attacks/Actions?"
   ClientHeight    =   2790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEndTurn 
      Caption         =   "End Turn"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkShowMessage 
      Caption         =   "Continue to show this message"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "More Actions"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "* You have played a Teamwork card and have additional          attacks"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "* You have played a special that allows for additional attacks"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "* You have played a special that requires additional actions       (such as drawing additional cards, removing hits, etc)."
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDlgAttack.frx":0000
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmDlgAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub chkShowMessage_Click()

If chkShowMessage.Value = 0 Then
    mySettings.ShowAttackMessage = False
Else
    mySettings.ShowAttackMessage = True
End If

End Sub

Private Sub cmdEndTurn_Click()
Check1.Value = 1
Me.Hide

End Sub

Private Sub Form_Load()
Check1.Value = 0

If mySettings.ShowAttackMessage = True Then
    chkShowMessage.Value = 1
Else
    chkShowMessage.Value = 0
End If

End Sub

Private Sub OKButton_Click()
Check1.Value = 0
Me.Hide

End Sub
