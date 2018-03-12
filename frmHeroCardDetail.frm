VERSION 5.00
Begin VB.Form frmHeroCardDetail 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Detail"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmHeroCardDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Click on picture to close this window"
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Image imgCard 
         BorderStyle     =   1  'Fixed Single
         Height          =   4260
         Left            =   100
         Picture         =   "frmHeroCardDetail.frx":1272
         Stretch         =   -1  'True
         Top             =   190
         Width           =   5955
      End
   End
End
Attribute VB_Name = "frmHeroCardDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgCard_Click()
Unload Me

End Sub
