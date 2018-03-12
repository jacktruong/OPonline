VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHomebase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim sName
Dim sCharacters
Dim sEffect

Dim cPlaced As Collection
Dim cPlacedGameEffect As Collection
Private Sub NewHomebase()
nID = 0
sName = ""
sCharacters = ""
sEffect = ""

Set cPlaced = New Collection

End Sub
Public Property Get Title() As String

Title = "HOMEBASE: " & sName

End Property
Public Property Get isLandscape() As Boolean

isLandscape = True

End Property
Public Property Get Description() As String
stemp = ""

stemp = "HOMEBASE: " & sName & ".  CHARACTERS: " & sCharacters & ".  EFFECT: " & sEffect

Description = stemp
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Homebases WHERE ((Homebases.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset
Dim myAspect As clsAspect

NewHomebase

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

nID = dbRec.Fields("ID").Value
sName = dbRec.Fields("Name").Value
sEffect = dbRec.Fields("Effect").Value
sCharacters = dbRec.Fields("Characters").Value

dbRec.Close

'Get Aspects
Set cAspects = New Collection

dbRec.Open "SELECT * FROM Aspect WHERE Aspect.HomeBase=" & Chr(34) & sName & Chr(34) & ";", db

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
Name = sName
End Property
Public Property Get Characters() As String

Characters = sCharacters

End Property
Public Property Get Placed_Count() As Integer
Placed_Count = cPlaced.Count

End Property

Public Property Get cardtype() As String
cardtype = "Homebase"

End Property
Public Property Get Placed_Type(ccardNumber)
Dim ccard
If ccardNumber = 0 Or ccardNumber > cPlaced.Count Then Exit Sub

Set ccard = cPlaced.Item(ccardNumber)

Placed_Type = ccard.cardtype

End Property
Public Sub PlaceCard(ccard, isGameEffect As Boolean)

cPlaced.Add ccard

If isGameEffect = True Then
    cPlacedGameEffect.Add "1"
Else
    cPlacedGameEffect.Add "0"
End If

End Sub
Public Sub RemovePlacedCard(cardindex)
If cardindex > cPlaced.Count Then Exit Sub
If cardindex < 1 Then Exit Sub

cPlaced.Remove cardindex
cPlacedGameEffect.Remove cardindex

End Sub
Public Function PlacedCard(ccardNumber)

If ccardNumber > cPlaced.Count Then Exit Function
If ccardNumber < 1 Then Exit Function

Set PlacedCard = cPlaced.Item(ccardNumber)

End Function

Private Sub Class_Initialize()
Set cPlaced = New Collection
Set cPlacedGameEffect = New Collection

End Sub
Private Function PlacedCard_isGameEffect(Index) As Boolean

If cPlacedGameEffect.Item(Index) = "1" Then
    PlacedCard_isGameEffect = True
Else
    PlacedCard_isGameEffect = False
End If

End Function
