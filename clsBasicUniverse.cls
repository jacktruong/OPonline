VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBasicUniverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim nValue As Integer
Dim sType1 As String
Dim nBonus As Integer
Dim sPictured As String
Dim sDescription As String
Private Sub NewBasic()
nID = 0
nValue = 0
sType1 = ""
nBonus = 0
sPictured = ""
sDescription = ""

End Sub
Public Function GetPlayableBasicUniverseCards(nHighE, nHighF, nHighS, nHighI) As Collection
Dim ctempc As Collection
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set ctempc = New Collection

Set db = New ADODB.Connection
db.ConnectionString = dbName

db.Open

Set dbRec = New ADODB.Recordset

'Energy
strSQL = "SELECT * From [Basic Universe] WHERE (((Val([Basic Universe]![Requires]))<=" & Trim(Str(nHighE)) & ") AND (([Basic Universe].Skill)='E'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Fighting
strSQL = "SELECT * From [Basic Universe] WHERE (((Val([Basic Universe]![Requires]))<=" & Trim(Str(nHighF)) & ") AND (([Basic Universe].Skill)='F'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close


'Strength
strSQL = "SELECT * From [Basic Universe] WHERE (((Val([Basic Universe]![Requires]))<=" & Trim(Str(nHighS)) & ") AND (([Basic Universe].Skill)='S'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Intellect
strSQL = "SELECT * From [Basic Universe] WHERE (((Val([Basic Universe]![Requires]))<=" & Trim(Str(nHighI)) & ") AND (([Basic Universe].Skill)='I'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close
db.Close

Set GetPlayableBasicUniverseCards = ctempc
Set ctempc = Nothing

End Function
Public Property Get Title() As String

Title = "BASIC UNIVERSE: " & Trim(Str(nValue)) & sType1 & "+" & Trim(Str(nBonus))

End Property
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Public Property Get Description() As String

stemp = "Basic Universe card.  Character must have at least " & Trim(Str(nValue))
stemp = stemp & " " & ConvertPowerCode(sType1) & ".  Use card to receive a  +" & Trim(Str(nBonus)) & " bonus to attack."

Description = stemp
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM [Basic Universe] WHERE (([Basic Universe].id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewBasic

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM [Basic Universe] WHERE [Basic Universe].ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value

nValue = dbRec.Fields("Requires").Value
sType1 = dbRec.Fields("Skill").Value
nBonus = dbRec.Fields("Bonus").Value
sPictured = dbRec.Fields("Character").Value
sDescription = dbRec.Fields("Description").Value

dbRec.Close
db.Close

End Sub
Public Property Get Skill() As String
Skill = ConvertPowerCode(sType1)

End Property
Public Property Get Requires() As Integer
Requires = nValue

End Property
Public Property Get Bonus() As Integer
Bonus = nBonus

End Property
Public Property Get Pictured() As String
Pictured = sPictured
End Property
Public Property Get Name() As String
Name = sDescription
End Property
Public Property Get CardType() As String
CardType = "Basic Universe"

End Property
Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect)
CanIPlay = False

If sType1 = "E" And nEnergy >= nValue Then CanIPlay = True: Exit Function
If sType1 = "F" And nFighting >= nValue Then CanIPlay = True: Exit Function
If sType1 = "S" And nStrength >= nValue Then CanIPlay = True: Exit Function
If sType1 = "I" And nIntellect >= nValue Then CanIPlay = True: Exit Function

If sType = "A" Or sType = "M" Then

    If nEnergy >= nValue Or nFighting >= nValue Or nStrength >= nValue Or nIntellect >= nValue Then
        CanIPlay = True
    End If
    

End If


End Function
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

