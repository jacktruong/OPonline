VERSION 5.00
Begin VB.Form frmTestDraw 
   Caption         =   "Test Draws"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDrawHand 
      Caption         =   "&Draw Hand"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblCard 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   9615
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   7
      Left            =   8520
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   6
      Left            =   7320
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   5
      Left            =   6120
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   4
      Left            =   4920
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   3
      Left            =   3720
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   2
      Left            =   2520
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   1
      Left            =   1320
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Index           =   0
      Left            =   120
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Tag             =   "Discard"
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "frmTestDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctest As Collection

Private Sub cmdCancel_Click()
Unload Me

End Sub
Private Sub cmdDrawHand_Click()
On Error Resume Next
Dim ccard

For i = 0 To 7
    Set imgCard(i).Picture = Nothing
    imgCard(i).ToolTipText = ""
Next i

Randomize

n = 8

Counter = 0

If n > ctest.Count Then n = ctest.Count

For i = 1 To n

c = Int(Rnd * ctest.Count) + 1
Set ccard = ctest.Item(c)

ctest.Remove c

Me.Caption = "Test Draws: Deck Count=" & Trim(Str(ctest.Count))

imgCard(Counter).ToolTipText = ccard.Title

If ccard.LoadImage(ccard.ID) = True Then
    imgCard(Counter).Picture = LoadPicture(App.Path & "\temppic.jpg")
End If

Counter = Counter + 1

Next i

lblCard.Caption = imgCard(0).ToolTipText


End Sub

Private Sub Form_Load()
Set ctest = New Collection

For i = 1 To cdeck.Count
    ctest.Add cdeck.Item(i)
Next i

End Sub

Private Sub imgCard_Click(Index As Integer)

lblCard.Caption = imgCard(Index).ToolTipText

End Sub

Private Sub imgCard_DblClick(Index As Integer)
    Load frmCardDetail
    frmCardDetail.imgCard.Picture = imgCard(Index).Picture
    frmCardDetail.Show 1
End Sub
