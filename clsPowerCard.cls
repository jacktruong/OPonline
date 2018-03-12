VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPowerCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim nValue As Integer
Dim sType As String
Dim cPlayablePowerCards As Collection
Public Function GetPlayablePowerCards(nHighEnergy, nHighFighting, nHighStrength, nHighIntellect) As Collection
Set cPlayablePowerCards = New Collection

Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName

db.Open

Set dbRec = New ADODB.Recordset

'Energy
strSQL = "SELECT *, Val([Power]![VALUE]) AS Power From Power WHERE (((Val([Power]![VALUE]))<= " & Trim(Str(nHighEnergy)) & ") AND ((Power.E)=True));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

cPlayablePowerCards.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Fighting
strSQL = "SELECT *, Val([Power]![VALUE]) AS Power From Power WHERE (((Val([Power]![VALUE]))<= " & Trim(Str(nHighFighting)) & ") AND ((Power.F)=True));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

cPlayablePowerCards.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Strength
strSQL = "SELECT *, Val([Power]![VALUE]) AS Power From Power WHERE (((Val([Power]![VALUE]))<= " & Trim(Str(nHighStrength)) & ") AND ((Power.S)=True));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

cPlayablePowerCards.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Intellect
strSQL = "SELECT *, Val([Power]![VALUE]) AS Power From Power WHERE (((Val([Power]![VALUE]))<= " & Trim(Str(nHighIntellect)) & ") AND ((Power.I)=True));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

cPlayablePowerCards.Add nID

dbRec.MoveNext
Wend

dbRec.Close

nmaxpower = nHighEnergy
If nHighFighting > nmaxpower Then nmaxpower = nHighFighting
If nHighStrength > nmaxpower Then nmaxpower = nHighStrength
If nHighIntellect > nmaxpower Then nmaxpower = nHighIntellect

'Anypower
strSQL = "SELECT *, Val([Power]![VALUE]) AS Power From Power WHERE (((Val([Power]![VALUE]))<= " & Trim(Str(nmaxpower)) & ") AND ((Power.A)=True));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

cPlayablePowerCards.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Multipower
strSQL = "SELECT *, Val([Power]![VALUE]) AS Power From Power WHERE (((Val([Power]![VALUE]))<= " & Trim(Str(nmaxpower)) & ") AND ((Power.M)=True));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

cPlayablePowerCards.Add nID

dbRec.MoveNext
Wend

dbRec.Close

db.Close

Set GetPlayablePowerCards = cPlayablePowerCards
Set cPlayablePowerCards = Nothing

End Function
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Private Sub NewPowerCard()
nID = 0
nValue = 0
sType = ""

End Sub
Public Property Get Title() As String

If isEnergy = True Then Title = "POWER: ENERGY " & Trim(Str(nValue))
If isStrength = True Then Title = "POWER: STRENGTH " & Trim(Str(nValue))
If isIntellect = True Then Title = "POWER: INTELLECT " & Trim(Str(nValue))
If isFighting = True Then Title = "POWER: FIGHTING " & Trim(Str(nValue))
If isAnyPower = True Then Title = "POWER: ANY-POWER " & Trim(Str(nValue))
If isMultiPower = True Then Title = "POWER: MULTI-POWER " & Trim(Str(nValue))

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Power WHERE ((Power.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Power WHERE Power.ID=" & GetID & ";", db

NewPowerCard

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value
nValue = dbRec.Fields("Value").Value

If dbRec.Fields("E").Value = True Then sType = "E"
If dbRec.Fields("F").Value = True Then sType = "F"
If dbRec.Fields("S").Value = True Then sType = "S"
If dbRec.Fields("I").Value = True Then sType = "I"
If dbRec.Fields("A").Value = True Then sType = "A"
If dbRec.Fields("M").Value = True Then sType = "M"

dbRec.Close
db.Close

End Sub
Public Property Get ID() As Variant
ID = nID

End Property
Public Property Get Power() As Variant

Power = nValue
End Property
Public Property Get isEnergy() As Boolean

If sType = "E" Then isEnergy = True Else isEnergy = False

End Property
Public Property Get isFighting() As Boolean

If sType = "F" Then isFighting = True Else isFighting = False

End Property
Public Property Get isStrength() As Boolean

If sType = "S" Then isStrength = True Else isStrength = False

End Property
Public Property Get isIntellect() As Boolean

If sType = "I" Then isIntellect = True Else isIntellect = False

End Property
Public Property Get isMultiPower() As Boolean

If sType = "M" Then isMultiPower = True Else isMultiPower = False

End Property
Public Property Get isAnyPower() As Boolean

If sType = "A" Then isAnyPower = True Else isAnyPower = False

End Property
Public Property Get PowerType() As String

Select Case sType

Case "E"
PowerType = "Energy"
Case "F"
PowerType = "Fighting"
Case "S"
PowerType = "Strength"
Case "I"
PowerType = "Intellect"
Case "A"
PowerType = "Any-Power"
Case "M"
PowerType = "Multi-Power"
Case Else
End Select

End Property
Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect)
CanIPlay = False

If sType = "E" And nEnergy >= nValue Then CanIPlay = True: Exit Function
If sType = "F" And nFighting >= nValue Then CanIPlay = True: Exit Function
If sType = "S" And nStrength >= nValue Then CanIPlay = True: Exit Function
If sType = "I" And nIntellect >= nValue Then CanIPlay = True: Exit Function

If sType = "A" And (nEnergy >= nValue Or nFighting >= nValue Or nStrength >= nValue Or nIntellect >= nValue) Then
    CanIPlay = True
    Exit Function
End If

If sType = "M" And (nEnergy >= nValue Or nFighting >= nValue Or nStrength >= nValue Or nIntellect >= nValue) Then
    CanIPlay = True
    Exit Function
End If


End Function
Public Property Get CardType() As String
CardType = "Power Card"

End Property
Public Property Get Attack_isPlaced() As Boolean

Attack_isPlaced = True

End Property
Public Property Get Attack_VentureValue() As Integer

Attack_VentureValue = nValue

End Property
Public Property Get Attack_NegativeVentureValue() As Integer

Attack_NegativeVentureValue = 0

End Property

Public Property Get Attack_isStringAttack() As Boolean

Attack_isStringAttack = False

End Property

