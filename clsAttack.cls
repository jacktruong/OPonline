VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'cls holds all of the information about a particular attack
Dim cAttackCards As Collection
Dim cAttackSource As Collection
Dim cAttackID As Collection
Dim bFaceDown As Boolean

Dim HeroId As Integer
Dim OppId As Integer
Public Sub NewAttack()
Set cAttackCards = New Collection
Set cAttackSource = New Collection
Set cAttackID = New Collection

bFaceDown = False

HeroId = 0
OppId = 0

End Sub
Public Property Get Card_Count() As Variant

Card_Count = cAttackCards.Count

End Property
Public Property Get DefenderID() As Integer
DefenderID = OppId

End Property
Public Property Let DefenderID(vnewValue As Integer)
OppId = vnewValue

End Property
Public Property Get AttackerID() As Integer
AttackerID = HeroId

End Property
Public Property Let AttackerID(vnewValue As Integer)
HeroId = vnewValue

End Property
Public Property Get Card_isPlaced(Index) As Boolean

Set ccard = cAttackCards.Item(Index)

If ccard.Attack_isPlaced = True Then
    Card_isPlaced = True
Else
    Card_isPlaced = False
End If

End Property
Public Property Get GetCard(Index)

Set GetCard = cAttackCards.Item(Index)

End Property
Public Property Get Card_ID(Index)

If Index > cAttackID.Count Or Index < 1 Then Exit Sub

Card_ID = cAttackID.Item(Index)

End Property
Public Property Get Attack_isFaceDown() As Boolean
Attack_isFaceDown = bFaceDown

End Property
Public Property Let Attack_isFaceDown(vnewValue As Boolean)

bFaceDown = vnewValue

End Property
Public Property Get Card_Source(Index)

Card_Source = cAttackSource.Item(Index)

End Property
Public Property Get Card_SourceTitle(Index, bisOpponentAttack As Boolean)

Select Case cAttackSource.Item(Index)
Case "P1"

    If bisOpponentAttack = True Then
        Card_SourceTitle = "Placed on " & cOpponent.Character_Name(1)
    Else
        Card_SourceTitle = "Placed on " & cFrontLine.Character_Name(1)
    End If

Case "P2"

    If bisOpponentAttack = True Then
        Card_SourceTitle = "Placed on " & cOpponent.Character_Name(2)
    Else
        Card_SourceTitle = "Placed on " & cFrontLine.Character_Name(2)
    End If
    
Case "P3"

    If bisOpponentAttack = True Then
        Card_SourceTitle = "Placed on " & cOpponent.Character_Name(3)
    Else
        Card_SourceTitle = "Placed on " & cFrontLine.Character_Name(3)
    End If
    
Case "P4"

    If bisOpponentAttack = True Then
        Card_SourceTitle = "Placed on " & cOpponent.Character_Name(4)
    Else
        Card_SourceTitle = "Placed on " & cFrontLine.Character_Name(4)
    End If
    
Case "P5"

    If bisOpponentAttack = True Then
        Card_SourceTitle = "Placed on " & OpHomebase.Name
    Else
        Card_SourceTitle = "Placed on " & myHomebase.Name
    End If

Case "H"
    Card_SourceTitle = "Hand"
    
Case Else

End Select

End Property
Public Sub AddCard(ccard, sSource, nIndex)
'sSource = P1, P2, P3, P4 or "H"
'nindex = index of card to be removed

cAttackCards.Add ccard
cAttackSource.Add sSource
cAttackID.Add nIndex

Select Case sSource
Case "P1"
    cFrontLine.RemovePlacedCard 1, nIndex
Case "P2"
    cFrontLine.RemovePlacedCard 2, nIndex
Case "P3"
    cFrontLine.RemovePlacedCard 3, nIndex
Case "P4"
    cFrontLine.RemovePlacedCard 4, nIndex
Case "P5"
    myHomebase.RemovePlacedCard nIndex
Case "H"
    cHand.Remove nIndex
    cHandTags.Remove nIndex
End Select


End Sub
Public Sub AddCardO(ccard, sSource, nIndex)

cAttackCards.Add ccard
cAttackSource.Add sSource
cAttackID.Add nIndex

End Sub
Public Sub RemoveCardO(Index)

cAttackCards.Remove Index
cAttackSource.Remove Index
cAttackID.Remove Index

End Sub
Public Sub RemoveCard(Index)
Dim ccard

Set ccard = cAttackCards.Item(Index)

Select Case cAttackSource.Item(Index)
Case "P1"
    cFrontLine.PlaceCard 1, ccard
Case "P2"
    cFrontLine.PlaceCard 2, ccard
Case "P3"
    cFrontLine.PlaceCard 3, ccard
Case "P4"
    cFrontLine.PlaceCard 4, ccard
Case "P5"
    myHomebase.PlaceCard ccard, True
Case "H"
    cHand.Add ccard
    
    If cHandTags.Count = 0 Then
        cHandTags.Add "A"
    Else
        a = Asc(cHandTags.Item(cHandTags.Count)) + 1
        cHandTags.Add Chr$(a)
    End If

End Select

cAttackCards.Remove Index
cAttackSource.Remove Index
cAttackID.Remove Index

End Sub
