VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTeamwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim sPictured As String
Dim nValue1 As Integer
Dim sType1 As String
Dim sType2 As String
Dim Stype3 As String
Dim nBonus1 As Integer
Dim nBonus2 As Integer
Public Function GetPlayableTeamworkCards(nHighE, nHighF, nHighS, nHighI) As Collection
Dim ctempc As Collection
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set ctempc = New Collection

Set db = New ADODB.Connection
db.ConnectionString = dbName

db.Open

Set dbRec = New ADODB.Recordset

'Energy
strSQL = "SELECT * From Teamwork WHERE (((Val(Teamwork![T1_PW]))<=" & Trim(Str(nHighE)) & ") AND ((Teamwork![T1_SK])='E'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Fighting
strSQL = "SELECT * From Teamwork WHERE (((Val(Teamwork![T1_PW]))<=" & Trim(Str(nHighF)) & ") AND ((Teamwork![T1_SK])='F'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close


'Strength
strSQL = "SELECT * From Teamwork WHERE (((Val(Teamwork![T1_PW]))<=" & Trim(Str(nHighS)) & ") AND ((Teamwork![T1_SK])='S'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close

'Intellect
strSQL = "SELECT * From Teamwork WHERE (((Val(Teamwork![T1_PW]))<=" & Trim(Str(nHighI)) & ") AND ((Teamwork![T1_SK])='I'));"

dbRec.Open strSQL, db

While dbRec.EOF = False
nID = dbRec.Fields("ID").Value

ctempc.Add nID

dbRec.MoveNext
Wend

dbRec.Close
db.Close

Set GetPlayableTeamworkCards = ctempc
Set ctempc = Nothing

End Function
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Private Sub NewTeamwork()

nID = 0
sPictured = ""
nValue1 = 0
sType1 = ""
sType2 = ""
Stype3 = ""
nBonus1 = 0
nBonus2 = 0

End Sub

Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewTeamwork

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Teamwork WHERE Teamwork.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value

If IsNull(dbRec.Fields("Pictured").Value) = False Then
    sPictured = dbRec.Fields("Pictured").Value
Else
    sPictured = "Unknown"
End If

nValue1 = dbRec.Fields("T1_PW").Value
sType1 = dbRec.Fields("T1_SK").Value
sType2 = dbRec.Fields("T2_SK").Value
Stype3 = dbRec.Fields("T3_SK").Value
nBonus1 = dbRec.Fields("Bonus1").Value
nBonus2 = dbRec.Fields("Bonus2").Value

dbRec.Close
db.Close

End Sub
Public Property Get Title() As String

Title = "TEAMWORK: " & Trim(Str(nValue1)) & sType1 & "/+" & Trim(Str(nBonus1)) & ", +" & Trim(Str(nBonus2)) & " " & sType2 & Stype3

End Property
Public Property Get Description() As String

stemp = ""
stemp = "Teamwork Universe Card: First character must have at least " & Trim(Str(nValue1)) & " " & ConvertPowerCode(sType1) & ".  Acts as a 6 " & ConvertPowerCode(sType1) & " attack. Second teammate may make a " & ConvertPowerCode(sType2) & " or " & ConvertPowerCode(Stype3) & " attack at +" & Trim(Str(nBonus1)) & "."
stemp = stemp & " Third teammate may make a " & ConvertPowerCode(sType2) & " or " & ConvertPowerCode(Stype3) & " attack at +" & Trim(Str(nBonus2)) & "."

Description = stemp

End Property
Public Property Get Pictured() As String

Pictured = sPictured

End Property
Public Property Get Bonus_1() As Integer
Bonus_1 = nBonus1

End Property
Public Property Get Bonus_2() As Integer
Bonus_2 = nBonus2

End Property
Public Property Get Skill_1() As String
Skill_1 = ConvertPowerCode(sType1)

End Property
Public Property Get Skill_2() As String
Skill_2 = ConvertPowerCode(sType2)

End Property
Public Property Get Skill_3() As String
Skill_3 = ConvertPowerCode(Stype3)

End Property

Public Property Get Require_1() As Integer
Require_1 = nValue1

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Teamwork WHERE ((Teamwork.id)=" & GetID & ");")

End Function
Public Property Get ID() As Integer
ID = nID

End Property
Public Property Get CardType() As String
CardType = "Teamwork"

End Property
Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect)
CanIPlay = False

If cFrontLine.LiveCharacterCount = 1 Then
    CanIPlay = False
    Exit Function
End If

Select Case sType1
Case "E"
    If nEnergy >= nValue1 Then CanIPlay = True
Case "F"
    If nFighting >= nValue1 Then CanIPlay = True
Case "S"
    If nStrength >= nValue1 Then CanIPlay = True
Case "I"
    If nIntellect >= nValue1 Then CanIPlay = True
Case "A", "M"
    If nEnergy >= nValue1 Or nFighting >= nValue1 Or nStrength >= nValue1 Or nIntellect >= nValue1 Then
        CanIPlay = True
    End If
    
Case Else
End Select
End Function
Public Property Get Attack_isPlaced() As Boolean

Attack_isPlaced = True

End Property
Public Property Get Attack_VentureValue() As Integer

Attack_VentureValue = 6

End Property
Public Property Get Attack_NegativeVentureValue() As Integer

Attack_NegativeVentureValue = 0

End Property

Public Property Get Attack_isStringAttack() As Boolean

Attack_isStringAttack = True

End Property

