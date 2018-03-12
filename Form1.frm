VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape1 
      FillColor       =   &H00000040&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1020
      Index           =   4
      Left            =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000040&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1020
      Index           =   1
      Left            =   960
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000040&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1020
      Index           =   2
      Left            =   840
      Top             =   1080
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000040&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1020
      Index           =   0
      Left            =   2040
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgDrawPile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   1680
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Tag             =   "Deck"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgDiscard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   2400
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgDeadPile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   3120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Dead"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgDefeated 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   1680
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Dead"
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgDiscard_OLEStartDrag(Data As DataObject, AllowedEffects As Long)

If cDiscardPile.Count = 0 Then Exit Sub

Set cCurrentDragSource = imgDiscard

End Sub

Private Sub imgDrawPile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbBar1.Panels(1).Text = "Draw Pile (" & cDrawPile.Count & ")"

End Sub

Private Sub imgDrawPile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'See where data is coming from
Select Case cCurrentDragSource.Tag

Case "Discard"
'moving a card from the deck to the discard pile
With frmMoveCards
.Target = "Draw"
.Source = "Discard"
.Show 1
End With

Case "Dead"
With frmMoveCards
.Target = "Draw"
.Source = "Dead"
.Show 1
End With

Case Else

End Select

UpdateDeckDisplay
End Sub

Private Sub imgDrawPile_OLEStartDrag(Data As DataObject, AllowedEffects As Long)

If cDrawPile.Count = 0 Then Exit Sub

Set cCurrentDragSource = imgDrawPile

End Sub

