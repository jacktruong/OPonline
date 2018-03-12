VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrontLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cCharacters As Collection
Dim myh As clsHero
Dim bisDead(4) As Boolean
Dim bisReserve(4) As Boolean
Dim b3Grid(4) As Boolean

Dim cBattleEffect(4) As Collection 'Artifacts, battle and game effects
Dim cBattleEffectType(4) As Collection

Dim cBuffers(4) As Collection

Dim cplacedcards(4) As Collection 'Cards that are placed on the character
Dim cHitCards(4) As Collection 'Cards that have hit the character
Dim cBufferHits As Collection

Dim cPRCards(4) As Collection 'Cards in Characters permanent record

Private Sub Class_Initialize()

Set cCharacters = New Collection
Set cBufferHits = New Collection

Set cBuffers(1) = New Collection
Set cBuffers(2) = New Collection
Set cBuffers(3) = New Collection
Set cBuffers(4) = New Collection


Set cBattleEffect(1) = New Collection
Set cBattleEffect(2) = New Collection
Set cBattleEffect(3) = New Collection
Set cBattleEffect(4) = New Collection

Set cBattleEffectType(1) = New Collection
Set cBattleEffectType(2) = New Collection
Set cBattleEffectType(3) = New Collection
Set cBattleEffectType(4) = New Collection

Set cplacedcards(1) = New Collection
Set cplacedcards(2) = New Collection
Set cplacedcards(3) = New Collection
Set cplacedcards(4) = New Collection

Set cHitCards(1) = New Collection
Set cHitCards(2) = New Collection
Set cHitCards(3) = New Collection
Set cHitCards(4) = New Collection

Set cPRCards(1) = New Collection
Set cPRCards(2) = New Collection
Set cPRCards(3) = New Collection
Set cPRCards(4) = New Collection

bisDead(1) = False
bisDead(2) = False
bisDead(3) = False
bisDead(4) = False

bisReserve(1) = False
bisReserve(2) = False
bisReserve(3) = False
bisReserve(4) = False

End Sub
Public Sub AddCharacter(nID, bInReserve As Boolean, bUse3Grid As Boolean)

Set myh = New clsHero

nc = cCharacters.Count + 1

myh.Load nID

cCharacters.Add myh
bisReserve(nc) = bInReserve
bisDead(nc) = False
b3Grid(nc) = bUse3Grid

Set cplacedcards(nc) = New Collection
Set cHitCards(nc) = New Collection
Set cBattleEffect(nc) = New Collection
Set cBattleEffectType(nc) = New Collection
Set cBuffers(nc) = New Collection

End Sub
Public Property Get LiveCharacterCount() As Integer
lc = 0

For i = 1 To 4
If bisDead(i) = False Then lc = lc + 1
Next i

LiveCharacterCount = lc

End Property
Public Property Get isCharacterDead(Index) As Boolean
If Index < 1 Or Index > 4 Then Exit Property

isCharacterDead = bisDead(Index)

End Property
Public Property Get isCharacterReserve(Index) As Boolean
If Index < 1 Or Index > 4 Then Exit Property

isCharacterReserve = bisReserve(Index)

End Property
Public Property Let isCharacterReserve(Index, vnewValue As Boolean)
If Index < 1 Or Index > 4 Then Exit Property

For i = 1 To 4
bisReserve(i) = False
Next i

bisReserve(Index) = vnewValue

End Property
Public Sub KillCharacter(Index)
Dim ccard

If Index < 1 Or Index > 4 Then Exit Sub

bisDead(Index) = True
bisReserve(Index) = False

Set myh = New clsHero
Set myh = cCharacters.Item(Index)
cDefeatedCharactersPile.Add myh
Set myh = Nothing

For i = 1 To cplacedcards(Index).Count
    Set ccard = cplacedcards(Index).Item(i)
    
    Select Case ccard.cardtype
    Case "Power Card"
        cDiscardPile.Add ccard
    Case Else
        cDeadPile.Add ccard
    End Select
Next i

For i = 1 To cBattleEffect(Index).Count
    Set ccard = cBattleEffect(Index).Item(i)
    cDeadPile.Add ccard
Next i

For i = 1 To cPRCards(Index).Count
    Set ccard = cPRCards(Index).Item(i)
    
    Select Case ccard.cardtype
    Case "Power Card"
        cDiscardPileO.Add ccard
    Case Else
        cDeadPileO.Add ccard
    End Select
Next i

Set cBattleEffect(Index) = New Collection
Set cBattleEffectType(Index) = New Collection

Set cplacedcards(Index) = New Collection
Set cHitCards(Index) = New Collection
Set cPRCards(Index) = New Collection


End Sub
Public Sub RessurrectCharacter(Index)

If Index < 1 Or Index > 4 Then Exit Sub

bisDead(Index) = False

Set cplacedcards(Index) = New Collection
Set cHitCards(Index) = New Collection
Set cBattleEffect(Index) = New Collection
Set cBattleEffectType(Index) = New Collection

a = 0
For i = 1 To 4
If bisReserve(i) = True Then a = i
Next i

If a = 0 Then bisReserve(Index) = True

End Sub
Public Property Get Character_Name(Index) As String

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

Character_Name = myh.Name

End Property
Public Property Get Character_Inherent(Index) As String

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

Character_Inherent = myh.InherentAbility

End Property
Public Property Get Character_HasInherent(Index) As Boolean

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

Character_HasInherent = myh.HasInherent


End Property


Public Property Get Character_ID(Index) As Integer

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

Character_ID = myh.ID

End Property
Public Property Get Character_Energy(Index) As Integer

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

If b3Grid(Index) = True Then
    Character_Energy = myh.Energy3
Else
    Character_Energy = myh.Energy
End If

End Property
Public Property Get Character_Fighting(Index) As Integer

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

If b3Grid(Index) = True Then
    Character_Fighting = myh.Fighting3
Else
    Character_Fighting = myh.Fighting
End If

End Property
Public Property Get Character_Intellect(Index) As Integer

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

Character_Intellect = myh.Intellect

End Property
Public Property Get Character_Strength(Index) As Integer

If Index < 1 Or Index > 4 Then Exit Property

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

If b3Grid(Index) = True Then
    Character_Strength = myh.Strength3
Else
    Character_Strength = myh.Strength
End If

End Property
Public Function LoadImage(Index) As Boolean
If Index < 1 Or Index > 4 Then Exit Function

Set myh = New clsHero
Set myh = cCharacters.Item(Index)
LoadImage = LoadImageFromDatabase("SELECT * FROM Characters WHERE ((Characters.id)=" & myh.ID & ");")

End Function
Public Sub PlaceCard(Index, ccard)
If Index < 1 Or Index > 4 Then Exit Sub

cplacedcards(Index).Add ccard

End Sub
Public Sub RemovePlacedCard(charIndex, cardindex)
If charIndex < 1 Or charIndex > 4 Then Exit Sub
If cardindex > cplacedcards(charIndex).Count Then Exit Sub
If cardindex < 1 Then Exit Sub

cplacedcards(charIndex).Remove cardindex

End Sub
Public Property Get Placed_Count(Index)

Placed_Count = cplacedcards(Index).Count

End Property
Public Function PlacedCard(Index, ccardNumber)
If Index < 1 Or Index > 4 Then Exit Function

If ccardNumber > cplacedcards(Index).Count Then Exit Function
If ccardNumber < 1 Then Exit Function

Set PlacedCard = cplacedcards(Index).Item(ccardNumber)

End Function
Public Property Get Placed_Type(Index, ccardNumber)
Dim ccard
If Index < 1 Or Index > 4 Then Exit Property

Set ccard = cplacedcards(Index).Item(ccardNumber)

Placed_Type = ccard.cardtype

End Property
Public Function CanCharacterPlayCard(Index, ccard) As Boolean
If Index < 1 Or Index > 4 Then Exit Function

CanCharacterPlayCard = False

Select Case ccard.cardtype

Case "Basic Universe", "Ally Card", "Power Card", "Training", "Teamwork"
If ccard.CanIPlay(Character_Energy(Index), Character_Fighting(Index), Character_Strength(Index), Character_Intellect(Index)) = True Then
    CanCharacterPlayCard = True
End If

Case "Special Card"

If ccard.Character = "*ANY CHARACTER*" Then
    CanCharacterPlayCard = True
    Exit Function
End If

Set myh = New clsHero
Set myh = cCharacters.Item(Index)

For i = 1 To myh.Special_Count
    If myh.Special_ID(i) = ccard.ID Then
        CanCharacterPlayCard = True
        Exit Function
    End If
Next i

Case Else
End Select

End Function
Public Property Get HitsToCurrentBattle_Count(Index) As Integer
If Index < 1 Or Index > 4 Then Exit Property
HitsToCurrentBattle_Count = cHitCards(Index).Count

End Property
Public Function HitsToCurrentBattle_GetCard(heroindex, cardindex)

If heroindex < 1 Or heroindex > 4 Then Exit Function
If cardindex > cHitCards(heroindex).Count Then Exit Function
If cardindex < 1 Then Exit Function

Set HitsToCurrentBattle_GetCard = cHitCards(heroindex).Item(cardindex)

End Function
Public Sub HitsToCurrentBattle_RemoveCard(heroindex, cardindex, bDontReturntoOpponent As Boolean)
Dim ccard
Dim ctemp As Collection

If heroindex < 1 Or heroindex > 4 Then Exit Sub
If cardindex > cHitCards(heroindex).Count Then Exit Sub
If cardindex < 1 Then Exit Sub

If bDontReturntoOpponent = False Then

Set ccard = cHitCards(heroindex).Item(cardindex)

Select Case ccard.cardtype
Case "Power Card"
'Return to opponent's power pack
    cDiscardPileO.Add ccard
'    cDiscardPileO.Add ccard, , 1
    
Case Else
'Put in Opponents dead pile
    Set ctemp = New Collection
    ctemp.Add ccard
    
    For i = 1 To cDeadPileO.Count
        ctemp.Add cDeadPileO.Item(i)
    Next i
    
    Set cDeadPileO = ctemp
    
End Select

End If

cHitCards(heroindex).Remove cardindex

End Sub
Public Sub HitsToCurrentBattle_AddCard(heroindex, ccard)
If heroindex < 1 Or heroindex > 4 Then Exit Sub

cHitCards(heroindex).Add ccard

End Sub
Public Property Get Modifiers_Count(Index) As Integer
If Index < 1 Or Index > 4 Then Exit Property
Modifiers_Count = cBattleEffect(Index).Count

End Property
Public Function Modifiers_GetCard(heroindex, cardindex)
If heroindex < 1 Or heroindex > 4 Then Exit Function
If cardindex > cBattleEffect(heroindex).Count Then Exit Function
If cardindex < 1 Then Exit Function

Set Modifiers_GetCard = cBattleEffect(heroindex).Item(cardindex)

End Function
Public Sub Modifiers_RemoveCard(heroindex, cardindex)
Dim ccard
Dim ctemp As Collection

If heroindex < 1 Or heroindex > 4 Then Exit Sub
If cardindex > cBattleEffect(heroindex).Count Then Exit Sub
If cardindex < 1 Then Exit Sub

cBattleEffect(heroindex).Remove cardindex
cBattleEffectType(heroindex).Remove cardindex

End Sub
Public Sub Modifiers_AddCard(heroindex, ccard, cardtype As ModifierType)
If heroindex < 1 Or heroindex > 4 Then Exit Sub

cBattleEffect(heroindex).Add ccard
cBattleEffectType(heroindex).Add cardtype

End Sub
Public Property Get Modifiers_TypeText(heroindex, cardindex)

If heroindex < 1 Or heroindex > 4 Then Exit Property
If cardindex > cBattleEffect(heroindex).Count Then Exit Property

Select Case cBattleEffectType(heroindex).Item(cardindex)

Case 1
    Modifiers_TypeText = "BATTLE"
Case 2
    Modifiers_TypeText = "GAME"
Case 3
    Modifiers_TypeText = "ARTIFACT"
Case Else
    Modifiers_TypeText = "Unknown"
End Select

End Property
Public Property Get Modifiers_Type(heroindex, cardindex)

If heroindex < 1 Or heroindex > 4 Then Exit Property
If cardindex > cBattleEffect(heroindex).Count Then Exit Property

Modifiers_Type = cBattleEffectType(heroindex).Item(cardindex)

End Property

Public Property Get PermanentRecord_Count(Index) As Integer

If Index < 1 Or Index > 4 Then Exit Property
PermanentRecord_Count = cPRCards(Index).Count

End Property
Public Function PermanentRecord_GetCard(heroindex, cardindex)

If heroindex < 1 Or heroindex > 4 Then Exit Function
If cardindex > cPRCards(heroindex).Count Then Exit Function
If cardindex < 1 Then Exit Function

Set PermanentRecord_GetCard = cPRCards(heroindex).Item(cardindex)

End Function
Public Sub PermanentRecord_RemoveCard(heroindex, cardindex, bDontReturntoOpponent As Boolean)
Dim ccard
Dim ctemp As Collection

If heroindex < 1 Or heroindex > 4 Then Exit Sub
If cardindex > cPRCards(heroindex).Count Then Exit Sub
If cardindex < 1 Then Exit Sub

If bDontReturntoOpponent = False Then
Set ccard = cPRCards(heroindex).Item(cardindex)

Select Case ccard.cardtype
Case "Power Card"
'Return to opponent's power pack
    cDiscardPileO.Add ccard, , 1
    
Case Else
'Put in Opponents dead pile
    Set ctemp = New Collection
    ctemp.Add ccard
    
    For i = 1 To cDeadPileO.Count
        ctemp.Add cDeadPileO.Item(i)
    Next i
    
    Set cDeadPileO = ctemp
    
End Select

End If

cPRCards(heroindex).Remove cardindex

End Sub
Public Sub PermanentRecord_AddCard(heroindex, ccard)
If heroindex < 1 Or heroindex > 4 Then Exit Sub

cPRCards(heroindex).Add ccard

End Sub
Public Property Get Buffers_Count(Index) As Integer
If Index < 1 Or Index > 4 Then Exit Property
Buffers_Count = cBuffers(Index).Count

End Property
Public Function Buffers_GetCard(heroindex, cardindex)
If heroindex < 1 Or heroindex > 4 Then Exit Function
If cardindex > cBuffers(heroindex).Count Then Exit Function
If cardindex < 1 Then Exit Function

Set Buffers_GetCard = cBuffers(heroindex).Item(cardindex)

End Function
Public Sub Buffers_RemoveCard(heroindex, cardindex)
Dim ccard
Dim ctemp As Collection

If heroindex < 1 Or heroindex > 4 Then Exit Sub
If cardindex > cBuffers(heroindex).Count Then Exit Sub
If cardindex < 1 Then Exit Sub

cBuffers(heroindex).Remove cardindex

End Sub
Public Sub Buffers_AddCard(heroindex, ccard)
If heroindex < 1 Or heroindex > 4 Then Exit Sub

cBuffers(heroindex).Add ccard

End Sub
Public Sub BufferHits_AddCard(ccard)

cBufferHits.Add ccard

End Sub
Public Function BufferHits_Count()
BufferHits_Count = cBufferHits.Count

End Function
Public Sub BufferHits_Clear()
Set cBufferHits = New Collection

End Sub

Public Function BufferHits_GetCard(Index)

Set BufferHits_GetCard = cBufferHits.Item(Index)

End Function
