VERSION 5.00
Begin VB.Form frmWhoGoesFirst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who Goes First?"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDrawAgain 
      Caption         =   "Draw Again"
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpGoesFirst 
      Caption         =   "Name 2 Goes First"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdIGoFirst 
      Caption         =   "Name 1 Goes First"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblName2 
      Caption         =   "Name 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblName1 
      Caption         =   "Name 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image imgCardDetail2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4725
      Left            =   6360
      OLEDragMode     =   1  'Automatic
      Picture         =   "frmWhoGoesFirst.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3525
   End
   Begin VB.Image imgCardDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4725
      Left            =   120
      OLEDragMode     =   1  'Automatic
      Picture         =   "frmWhoGoesFirst.frx":B4DA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3525
   End
End
Attribute VB_Name = "frmWhoGoesFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDrawAgain_Click()
DrawCards

End Sub

Private Sub Form_Load()

DrawCards
lblName1.Caption = mySettings.PlayerName
cmdIGoFirst.Caption = mySettings.PlayerName & " goes first"

lblName2.Caption = sOpponentName
cmdOpGoesFirst.Caption = sOpponentName & " goes first"

If bHost = False Then
    Me.cmdDrawAgain.Enabled = False
    Me.cmdIGoFirst.Enabled = False
    Me.cmdOpGoesFirst.Enabled = False
End If

End Sub
Private Sub DrawCards()
Dim ccard
Dim ccard2

Randomize

p1 = Int((cDrawPile.Count * Rnd) + 1)
p2 = Int((cDrawPileO.Count * Rnd) + 1)
t1$ = cDrawPile.Item(p1).Title
t2$ = cDrawPileO.Item(p2).Title

imgCardDetail.ToolTipText = t1$
imgCardDetail2.ToolTipText = t2$

If cDrawPile.Item(p1).LoadImage(cDrawPile.Item(p1).ID) = True Then
    imgCardDetail.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgCardDetail.Picture = LoadPicture(sBlankImagePath)
End If

If cDrawPileO.Item(p2).LoadImage(cDrawPileO.Item(p2).ID) = True Then
    imgCardDetail2.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgCardDetail2.Picture = LoadPicture(sBlankImagePath)
End If

End Sub
