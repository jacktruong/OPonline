VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim sType1 As String
Dim sType2 As String
Dim nBonus As Integer
Dim sPictured As String
Private Sub NewTrainingCard()
nID = 0
nBonus = 0
sType1 = ""
sType2 = ""
sPictured = ""

End Sub
Public Function GetPlayableTrainingCards(nLowEnergy, nLowFighting, nLowStrength, nLowIntellect) As Collection
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset
Dim ctempc As Collection
Dim bFound As Boolean

Set ctempc = New Collection
Set db = New ADODB.Connection
db.ConnectionString = dbName

db.Open

Set dbRec = New ADODB.Recordset

'Energy

If nLowEnergy < 6 Then

    strSQL = "SELECT * From Training WHERE (((Training.PWR1)='E')) OR (((Training.PWR2)='E'));"
    
    dbRec.Open strSQL, db
    
    While dbRec.EOF = False
    nID = dbRec.Fields("ID").Value
    
    bFound = False
    
    For i = 1 To ctempc.Count
        If ctempc.Item(i) = nID Then
            bFound = True
        End If
    Next i
    
    If bFound = False Then ctempc.Add nID
    
    dbRec.MoveNext
    Wend
    
    dbRec.Close

End If


'Fighting

If nLowFighting < 6 Then

    strSQL = "SELECT * From Training WHERE (((Training.PWR1)='F')) OR (((Training.PWR2)='F'));"
    
    dbRec.Open strSQL, db
    
    While dbRec.EOF = False
    nID = dbRec.Fields("ID").Value
    
    bFound = False
    
    For i = 1 To ctempc.Count
        If ctempc.Item(i) = nID Then
            bFound = True
        End If
    Next i
    
    If bFound = False Then ctempc.Add nID
    
    dbRec.MoveNext
    Wend
    
    dbRec.Close

End If

'Strength

If nLowStrength < 6 Then

    strSQL = "SELECT * From Training WHERE (((Training.PWR1)='S')) OR (((Training.PWR2)='S'));"
    
    dbRec.Open strSQL, db
    
    While dbRec.EOF = False
    nID = dbRec.Fields("ID").Value
    
    bFound = False
    
    For i = 1 To ctempc.Count
        If ctempc.Item(i) = nID Then
            bFound = True
        End If
    Next i
    
    If bFound = False Then ctempc.Add nID
    
    dbRec.MoveNext
    Wend
    
    dbRec.Close

End If

'Intellect

If nLowIntellect < 6 Then

    strSQL = "SELECT * From Training WHERE (((Training.PWR1)='I')) OR (((Training.PWR2)='I'));"
    
    dbRec.Open strSQL, db
    
    While dbRec.EOF = False
    nID = dbRec.Fields("ID").Value
    
    bFound = False
    
    For i = 1 To ctempc.Count
        If ctempc.Item(i) = nID Then
            bFound = True
        End If
    Next i
    
    If bFound = False Then ctempc.Add nID
    
    dbRec.MoveNext
    Wend
    
    dbRec.Close

End If

db.Close

Set GetPlayableTrainingCards = ctempc
Set ctempc = Nothing

End Function
Public Property Get Title() As String

Title = "TRAINING: " & sType1 & sType2 & " + " & Trim(Str(nBonus))

End Property
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Public Property Get Description() As String
stemp = ""

stemp = "Training Universe card.  Character must have a 5 or less in "
stemp = stemp & ConvertPowerCode(sType1) & " OR " & ConvertPowerCode(sType2) & " to use. "
stemp = stemp & "Provides a +" & Trim(Str(nBonus)) & " bonus to attack."

Description = stemp
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Training WHERE ((Training.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewTrainingCard

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Training WHERE Training.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value

nBonus = dbRec.Fields("Bonus").Value
sType1 = dbRec.Fields("PWR1").Value
sType2 = dbRec.Fields("PWR2").Value
sPictured = dbRec.Fields("Pictured").Value

dbRec.Close
db.Close

End Sub
Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect) As Boolean
CanIPlay = False

If sType1 = "E" And nEnergy <= 5 Then CanIPlay = True: Exit Function
If sType2 = "E" And nEnergy <= 5 Then CanIPlay = True: Exit Function

If sType1 = "F" And nFighting <= 5 Then CanIPlay = True: Exit Function
If sType2 = "F" And nFighting <= 5 Then CanIPlay = True: Exit Function

If sType1 = "S" And nStrength <= 5 Then CanIPlay = True: Exit Function
If sType2 = "S" And nStrength <= 5 Then CanIPlay = True: Exit Function

If sType1 = "I" And nIntellect <= 5 Then CanIPlay = True: Exit Function
If sType2 = "I" And nIntellect <= 5 Then CanIPlay = True: Exit Function

End Function
Public Property Get Skill_1() As String
Skill_1 = ConvertPowerCode(sType1)

End Property
Public Property Get Skill_2() As String
Skill_2 = ConvertPowerCode(sType2)

End Property
Public Property Get Bonus() As Integer
Bonus = nBonus

End Property
Public Property Get Pictured() As String
Pictured = sPictured

End Property
Public Property Get CardType() As String
CardType = "Training"

End Property

Public Property Get Attack_isPlaced() As Boolean

Attack_isPlaced = False

End Property
Public Property Get Attack_VentureValue() As Integer

Attack_VentureValue = 0

End Property
Public Property Get Attack_NegativeVentureValue() As Integer

Attack_NegativeVentureValue = 0

End Property

Public Property Get Attack_isStringAttack() As Boolean

Attack_isStringAttack = False

End Property

