VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBattlesite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nId As Integer
Dim sname
Dim sCharacters
Dim sEffect

Dim cSpecials As Collection
Dim cHitCards As Collection 'Cards that have hit the character
Dim cPRCards As Collection 'Cards in Characters permanent record
Public Sub NewBattlesite()
nId = 0
sname = ""
sCharacters = ""
sEffect = ""
Set cPRCards = New Collection
Set cHitCards = New Collection

Set cSpecials = New Collection

End Sub
Public Property Get Title() As String

Title = "HOMEBASE: " & sname

End Property
Public Property Get isLandscape() As Boolean

isLandscape = True

End Property
Public Property Get Description() As String
stemp = ""

stemp = "HOMEBASE: " & sname & ".  CHARACTERS: " & sCharacters & ".  EFFECT: " & sEffect

Description = stemp
End Property
Public Property Get ID() As Integer
ID = nId

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Homebases WHERE ((Homebases.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset
Dim myAspect As clsAspect

NewBattlesite

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Homebases WHERE Homebases.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nId = dbRec.Fields("ID").Value
sname = dbRec.Fields("Name").Value
sEffect = dbRec.Fields("Effect").Value
sCharacters = dbRec.Fields("Characters").Value

dbRec.Close

'Get Aspects
Set cAspects = New Collection

dbRec.Open "SELECT * FROM Aspect WHERE Aspect.HomeBase=" & Chr(34) & sname & Chr(34) & ";", db

If dbRec.EOF = False Then

While dbRec.EOF = False
    Set myAspect = New clsAspect
    myAspect.Load dbRec.Fields("ID").Value
    cAspects.Add myAspect
    dbRec.MoveNext
Wend

End If

dbRec.Close

db.Close

End Sub
Public Property Get Effect() As String
Effect = sEffect

End Property

Public Property Get Name() As String
Name = sname
End Property
Public Property Get Characters() As String

Characters = sCharacters

End Property
Public Property Get CardType() As String
CardType = "Battle Site"

End Property

Public Property Get Deck_Count() As Integer
Deck_Count = cSpecials.Count

End Property
Public Property Get Deck_Type(ccardNumber)
Dim ccard
If ccardNumber = 0 Or ccardNumber > cSpecials.Count Then Exit Property

Set ccard = cSpecials.Item(ccardNumber)

Placed_Type = ccard.CardType

End Property
Public Sub Deck_AddCard(ccard)

cSpecials.Add ccard

End Sub
Public Sub RemoveDeckCard(cardindex)
If cardindex > cSpecials.Count Then Exit Sub
If cardindex < 1 Then Exit Sub

cSpecials.Remove cardindex

End Sub
Public Function Deck_GetCard(ccardNumber)

If ccardNumber > cSpecials.Count Then Exit Function
If ccardNumber < 1 Then Exit Function

Set Deck_GetCard = cSpecials.Item(ccardNumber)

End Function

Public Property Get HitsToCurrentBattle_Count() As Integer
HitsToCurrentBattle_Count = cHitCards.Count

End Property
Public Function HitsToCurrentBattle_GetCard(cardindex)
If cardindex > cHitCards.Count Then Exit Function
If cardindex < 1 Then Exit Function

Set HitsToCurrentBattle_GetCard = cHitCards.Item(cardindex)

End Function
Public Sub HitsToCurrentBattle_RemoveCard(cardindex, bDontReturntoOpponent As Boolean)
Dim ccard
Dim ctemp As Collection

If cardindex > cHitCards.Count Then Exit Sub
If cardindex < 1 Then Exit Sub

If bDontReturntoOpponent = False Then

Set ccard = cHitCards.Item(cardindex)

Select Case ccard.CardType
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

cHitCards.Remove cardindex

End Sub
Public Sub HitsToCurrentBattle_AddCard(ccard)

cHitCards.Add ccard

End Sub
Public Property Get PermanentRecord_Count() As Integer

PermanentRecord_Count = cPRCards.Count

End Property
Public Function PermanentRecord_GetCard(cardindex)

If cardindex > cPRCards.Count Then Exit Function
If cardindex < 1 Then Exit Function

Set PermanentRecord_GetCard = cPRCards.Item(cardindex)

End Function
Public Sub PermanentRecord_RemoveCard(cardindex, bDontReturntoOpponent As Boolean)
Dim ccard
Dim ctemp As Collection

If cardindex > cPRCards.Count Then Exit Sub
If cardindex < 1 Then Exit Sub

If bDontReturntoOpponent = False Then
Set ccard = cPRCards.Item(cardindex)

Select Case ccard.CardType
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

cPRCards.Remove cardindex

End Sub
Public Sub PermanentRecord_AddCard(ccard)

cPRCards.Add ccard

End Sub
Public Sub HitsToCurrentBattle_MoveAllToPermanentRecord()

For k = 1 To cHitCards.Count
    Set ccard = cHitCards.Item(k)
    cPRCards.Add ccard
Next k

For k = 1 To cHitCards.Count
    cHitCards.Remove 1
Next k

End Sub
