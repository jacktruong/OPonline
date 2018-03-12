VERSION 5.00
Begin VB.Form frmHistory 
   Caption         =   "Game History"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstHistory 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label lblHistoryItem 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   9135
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstHistory_Click()
lstHistory.ToolTipText = lstHistory.List(lstHistory.ListIndex)
lblHistoryItem.Caption = lstHistory.List(lstHistory.ListIndex)

End Sub
