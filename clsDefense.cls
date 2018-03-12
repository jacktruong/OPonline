VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDefense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'cls holds all of the information about a particular Defense
Dim cDefenseCards As Collection
Dim cDefenseSource As Collection
Dim HeroID As Integer
Dim OppId As Integer
Public Sub NewDefense()
Set cDefenseCards = New Collection
Set cDefenseSource = New Collection
HeroID = 0
OppId = 0

End Sub
Public Property Get Card_Count() As Variant

Card_Count = cDefenseCards.Count

End Property
Public Property Get Card_Source(Index)

If Index < 1 Or Index > cDefenseSource.Count Then Exit Sub

Card_Source = cDefenseSource(Index)
End Property
Public Property Get GetCard(Index)
Set GetCard = cDefenseCards.Item(Index)

End Property
Public Sub AddCard(ccard, sSource, nIndex)
'sSource = P1, P2, P3, P4 or "H"
'nindex = index of card to be removed

cDefenseCards.Add ccard
cDefenseSource.Add sSource

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
Public Sub AddCardO(ccard)

cDefenseCards.Add ccard

End Sub
Public Sub RemoveCard(Index)
Dim ccard

Set ccard = cDefenseCards.Item(Index)

Select Case cDefenseSource.Item(Index)
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

cDefenseCards.Remove Index
cDefenseSource.Remove Index

End Sub

Public Sub DiscardCard(Index)
Dim ccard

Set ccard = cDefenseCards.Item(Index)

If ccard.cardtype = "Power Card" Then
    cDiscardPile.Add ccard
Else
    cDeadPile.Add ccard
End If

cDefenseCards.Remove Index
cDefenseSource.Remove Index

End Sub
