VERSION 5.00
Begin VB.Form frmCardDetail 
   Caption         =   "Card Image"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgCard 
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmCardDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.Width = imgCard.Width + 150
Me.Height = imgCard.Height + 435

End Sub

