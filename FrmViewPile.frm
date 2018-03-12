VERSION 5.00
Begin VB.Form FrmViewPile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View:"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   11535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.ListBox lstPile 
      Height          =   5910
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Frame frmActions 
      Height          =   1455
      Index           =   3
      Left            =   9720
      TabIndex        =   17
      Top             =   3840
      Width           =   1695
      Begin VB.CommandButton cmdBattleSiteToHand 
         Caption         =   "To Hand"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBattleSiteToDiscard 
         Caption         =   "Discard"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame frmActions 
      Height          =   2055
      Index           =   4
      Left            =   9720
      TabIndex        =   20
      Top             =   3240
      Width           =   1695
      Begin VB.CommandButton cmdDefeatedToBattlesite 
         Caption         =   "To Battlesite"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdDefeatedToHand 
         Caption         =   "To Hand"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdResurrectChar 
         Caption         =   "Resurrect"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmActions 
      Height          =   2775
      Index           =   1
      Left            =   9720
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
      Begin VB.CommandButton cmdMyDiscardToDraw 
         Caption         =   "To Draw Pile"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdShuffleDiscard 
         Caption         =   "Shuffle"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMyDiscardToDead 
         Caption         =   "To Dead Pile"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdMyDiscardToHand 
         Caption         =   "To Hand"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame frmActions 
      Height          =   2775
      Index           =   0
      Left            =   9720
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
      Begin VB.CommandButton cmdMyDrawToMyHand 
         Caption         =   "To Hand"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdMyDrawToDead 
         Caption         =   "To Dead Pile"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdShuffleDraw 
         Caption         =   "Shuffle"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMyDrawToDiscard 
         Caption         =   "To Power Pack"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmActions 
      Height          =   2775
      Index           =   2
      Left            =   9720
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
      Begin VB.CommandButton cmdMyDeadToHand 
         Caption         =   "To Hand"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdMyDeadToPowerPack 
         Caption         =   "To Power Pack"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdShuffleDead 
         Caption         =   "Shuffle"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMyDeadToDraw 
         Caption         =   "To Draw Pile"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Image imgNormal 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5925
      Left            =   5400
      OLEDragMode     =   1  'Automatic
      Picture         =   "FrmViewPile.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4245
   End
   Begin VB.Image imgLandScape 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4245
      Left            =   5400
      OLEDragMode     =   1  'Automatic
      Picture         =   "FrmViewPile.frx":B4DA
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   5925
   End
End
Attribute VB_Name = "FrmViewPile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vShowPile As Collection
Dim ccard
Dim nType As Integer
Dim bAddedtomyhand As Boolean

Public Property Let PileType(ByVal vnewValue As Integer)
' 0 = my draw pile
' 1 = my discard pile
' 2 = my dead pile
' 3 = my Battlesite deck
' 4 = my defeated pile


' 3 = opponents draw
' 4 = opponents discard
' 5 = opponents dead
' 6 = opponents Battlesite deck

nType = vnewValue

End Property
Public Property Set ShowPile(ByVal vnewValue As Collection)

Set vShowPile = New Collection

Set vShowPile = vnewValue

End Property

Private Sub cmdBattleSiteToDraw_Click()

End Sub

Private Sub cmdBattleSiteToDiscard_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cDefeatedCharactersPile.Add myBattleSite.Deck_GetCard(a)
myBattleSite.RemoveDeckCard a
vShowPile.Remove a

ShowCards
End Sub

Private Sub cmdBattleSiteToHand_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cHand.Add myBattleSite.Deck_GetCard(a)

frmTable.tcpChannel.SendData "CSC:11:2:" & Trim(Str(a)) & ":|"

If cHandTags.Count = 0 Then
    cHandTags.Add "A"
Else
    cHandTags.Add Chr$(Asc(cHandTags.Item(cHandTags.Count)) + 1)
End If

myBattleSite.RemoveDeckCard a
vShowPile.Remove a

bAddedtomyhand = True
ShowCards
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdDefeatedToBattlesite_Click()
Dim ccard
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)

Set ccard = cDefeatedCharactersPile.Item(a)

If ccard.CardType = "Hero" Then
    X = MsgBox("Heroes may not be added to your Battlesite deck.", vbCritical, "Cannot Add to Hand.")
    Exit Sub
End If

End Sub

Private Sub cmdDefeatedToHand_Click()
Dim ccard
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)

Set ccard = cDefeatedCharactersPile.Item(a)

If ccard.CardType = "Hero" Then
    X = MsgBox("Heroes may not be added to your hand.", vbCritical, "Cannot Add to Hand.")
    Exit Sub
End If

End Sub

Private Sub cmdMyDeadToDraw_Click()
Dim ctemp As Collection

a = lstPile.ItemData(lstPile.ListIndex)

With frmBasicMove
.lblMoveType.Caption = "Move from Dead Pile to Draw Pile"
.Show 1

If .chkNo.Value <> 1 Then

If .optDropType(0).Value = True Then 'add to top
    Set ctemp = New Collection
    
    ctemp.Add cDeadPile.Item(a)
    cDeadPile.Remove a
    
    For i = 1 To cDrawPile.Count
    ctemp.Add cDrawPile.Item(i)
    Next i
    
    Set cDrawPile = New Collection
    Set cDrawPile = ctemp
    
    frmTable.tcpChannel.SendData "CSC:4:1:" & Trim(Str(a)) & ":|"
    frmTable.tcpChannel.SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"
    frmTable.tcpChannel.SendData "CEP:1:|"
    
End If

If .optDropType(1).Value = True Then 'add to bottom

    cDrawPile.Add cDeadPile.Item(a)
    cDeadPile.Remove a
    
    frmTable.tcpChannel.SendData "CSC:4:1:" & Trim(Str(a)) & ":|"

End If

If .optDropType(2).Value = True Then 'shuffle in

    cDrawPile.Add cDeadPile.Item(a)
    cDeadPile.Remove a
    ShufflePile 0

    frmTable.tcpChannel.SendData "CSC:4:1:" & Trim(Str(a)) & ":|"
    frmTable.tcpChannel.SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"
    frmTable.tcpChannel.SendData "CEP:1:|"
    
End If


End If

End With

Unload frmBasicMove

ShowCards

End Sub

Private Sub cmdMyDeadToHand_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cHand.Add cDeadPile.Item(a)

If cHandTags.Count = 0 Then
    cHandTags.Add "A"
Else
    cHandTags.Add Chr$(Asc(cHandTags.Item(cHandTags.Count)) + 1)
End If
cDeadPile.Remove a

frmTable.tcpChannel.SendData "CSC:4:2:" & Trim(Str(a)) & ":|"

bAddedtomyhand = True
ShowCards
End Sub

Private Sub cmdMyDeadToPowerPack_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cDiscardPile.Add cDeadPile.Item(a)
cDeadPile.Remove a

frmTable.tcpChannel.SendData "CSC:4:3:" & Trim(Str(a)) & ":|"

ShowCards
End Sub

Private Sub cmdMyDiscardToDead_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cDeadPile.Add cDiscardPile.Item(a)
cDiscardPile.Remove a

frmTable.tcpChannel.SendData "CSC:3:4:" & Trim(Str(a)) & ":|"

ShowCards
End Sub

Private Sub cmdMyDiscardToDraw_Click()

a = lstPile.ItemData(lstPile.ListIndex)

With frmBasicMove
.lblMoveType.Caption = "Move from Power Pack to Draw Pile"
.Show 1

If .chkNo.Value <> 1 Then

If .optDropType(0).Value = True Then 'add to top
    Set ctemp = New Collection
    
    ctemp.Add cDiscardPile.Item(a)
    cDiscardPile.Remove a
    
    For i = 1 To cDrawPile.Count
        ctemp.Add cDrawPile.Item(i)
    Next i
    
    Set cDrawPile = New Collection
    Set cDrawPile = ctemp

    frmTable.tcpChannel.SendData "CSC:3:1:" & Trim(Str(a)) & ":|"
    frmTable.tcpChannel.SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"
    frmTable.tcpChannel.SendData "CEP:1:|"
    
End If

If .optDropType(1).Value = True Then 'add to bottom

    cDrawPile.Add cDiscardPile.Item(a)
    cDiscardPile.Remove a

    frmTable.tcpChannel.SendData "CSC:3:1:" & Trim(Str(a)) & ":|"

End If

If .optDropType(2).Value = True Then 'shuffle in

    cDrawPile.Add cDiscardPile.Item(a)
    cDiscardPile.Remove a
    ShufflePile 0

    frmTable.tcpChannel.SendData "CSC:3:1:" & Trim(Str(a)) & ":|"
    frmTable.tcpChannel.SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"
    frmTable.tcpChannel.SendData "CEP:1:|"
    
End If


End If

End With

Unload frmBasicMove
ShowCards

End Sub

Private Sub cmdMyDiscardToHand_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cHand.Add cDiscardPile.Item(a)
cHandTags.Add Chr$(Asc(cHandTags.Item(cHandTags.Count)) + 1)
cDiscardPile.Remove a
bAddedtomyhand = True

frmTable.tcpChannel.SendData "CSC:3:2:" & Trim(Str(a)) & ":|"

ShowCards
End Sub

Private Sub cmdMyDrawToDead_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cDeadPile.Add cDrawPile.Item(a)
cDrawPile.Remove a

frmTable.tcpChannel.SendData "CSC:1:4:" & Trim(Str(a)) & ":|"

ShowCards
End Sub

Private Sub cmdMyDrawToDiscard_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cDiscardPile.Add cDrawPile.Item(a)
cDrawPile.Remove a

frmTable.tcpChannel.SendData "CSC:1:3:" & Trim(Str(a)) & ":|"

ShowCards

End Sub

Public Property Get AddedToMyHand() As Boolean
AddedToMyHand = bAddedtomyhand

End Property

Private Sub cmdMyDrawToMyHand_Click()
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)
cHand.Add cDrawPile.Item(a)

If cHandTags.Count = 0 Then
cHandTags.Add "A"
Else
cHandTags.Add Chr$(Asc(cHandTags.Item(cHandTags.Count)) + 1)
End If

frmTable.tcpChannel.SendData "CSC:1:2:" & Trim(Str(a)) & ":|"

cDrawPile.Remove a
bAddedtomyhand = True
ShowCards
End Sub

Private Sub cmdResurrectChar_Click()
Dim ccard
If lstPile.ListIndex = -1 Then Exit Sub

a = lstPile.ItemData(lstPile.ListIndex)

Set ccard = cDefeatedCharactersPile.Item(a)

If ccard.CardType <> "Hero" Then
    X = MsgBox("Only Heroes may be resurrected.", vbCritical, "Cannot Resurrect.")
    Exit Sub
End If

'Get ID in frontline

For i = 1 To 4

If cFrontLine.Character_Name(i) = ccard.Name And cFrontLine.isCharacterDead(i) = True Then
    cFrontLine.RessurrectCharacter i
    Me.cmdResurrectChar.Tag = i
'    SendData "CRC:" & Trim(Str(i)) & ":|"
    
End If

Next i

cDefeatedCharactersPile.Remove a
ShowCards


End Sub

Private Sub cmdShuffleDead_Click()
ShufflePile 2
Set vShowPile = cDeadPile

ShowCards

frmTable.tcpChannel.SendData "CDD:" & GetCode_CardString(cDeadPile) & "|"
frmTable.tcpChannel.SendData "CEP:1:|"

End Sub

Private Sub cmdShuffleDiscard_Click()
ShufflePile 1
Set vShowPile = cDiscardPile
ShowCards

frmTable.tcpChannel.SendData "CDI:" & GetCode_CardString(cDiscardPile) & "|"
frmTable.tcpChannel.SendData "CEP:1:|"

End Sub

Private Sub cmdShuffleDraw_Click()
ShufflePile 0
Set vShowPile = cDrawPile

frmTable.tcpChannel.SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"
frmTable.tcpChannel.SendData "CEP:1:|"

ShowCards

End Sub

Private Sub Form_Activate()
ShowCards
ShowCaption

End Sub

Private Sub ShowCards()
If vShowPile.Count = 0 Then
    Me.Hide
    Exit Sub
End If

ShowFrames
lstPile.Clear

For i = 1 To vShowPile.Count

Set ccard = vShowPile.Item(i)
lstPile.AddItem ccard.Title
lstPile.ItemData(lstPile.NewIndex) = i
Next i

If lstPile.ListCount > 0 Then lstPile.ListIndex = 0

End Sub
Private Sub ShowFrames()
On Error Resume Next

For i = 0 To frmActions.Count - 1
    frmActions(i).Visible = False
Next i

frmActions(nType).Visible = True

End Sub

Private Sub Form_Load()

bAddedtomyhand = False

End Sub

Private Sub lstPile_Click()
If lstPile.ListIndex = -1 Then Exit Sub

Set ccard = vShowPile.Item(lstPile.ItemData(lstPile.ListIndex))

If ccard.isLandscape = True Then

Me.Width = 13395

On Error Resume Next
For i = 0 To 8
frmActions(i).Left = 11400
Next i

cmdCancel.Left = 11520

Else

Me.Width = 11625

On Error Resume Next
For i = 0 To 8
frmActions(i).Left = 9720
Next i

cmdCancel.Left = 9840

End If

If ccard.LoadImage(ccard.ID) = True Then
    If ccard.isLandscape = True Then
        imgLandScape.Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgLandScape.Visible = True
        imgNormal.Visible = False
        
    Else
        imgNormal.Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgLandScape.Visible = False
        imgNormal.Visible = True
        
    End If

Else

    If ccard.isLandscape = True Then
        imgLandScape.Picture = LoadPicture(sBlankImagePath)
        imgLandScape.Visible = True
        imgNormal.Visible = False
    
    Else
        imgNormal.Picture = LoadPicture(sBlankImagePath)
        imgLandScape.Visible = False
        imgNormal.Visible = True
    
    End If
    
End If

ShowCaption

End Sub
Private Sub ShowCaption()
Select Case nType
Case 0
    Me.Caption = "My Draw Pile: " & vShowPile.Count
Case 1
    Me.Caption = "My Power Pack: " & vShowPile.Count
Case 2
    Me.Caption = "My Dead Pile: " & vShowPile.Count
Case 3
    Me.Caption = "Battlesite Deck: " & vShowPile.Count
Case 4
    Me.Caption = "My Defeated Characters Pile: " & vShowPile.Count
    
Case Else
End Select

End Sub
