VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArtifact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim nValue As Integer
Dim sType1 As String
Dim nValue2 As Integer
Dim sType2 As String
Dim sName
Dim sEffect As String


Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect) As Boolean
CanIPlay = False
If cFrontLine.LiveCharacterCount = 1 Then
    CanIPlay = False
    Exit Function
End If

If sType1 = "E" And nEnergy >= nValue Then CanIPlay = True
If sType1 = "F" And nFighting >= nValue Then CanIPlay = True
If sType1 = "S" And nStrength >= nValue Then CanIPlay = True
If sType1 = "I" And nIntellect >= nValue Then CanIPlay = True

End Function
Private Sub NewArtifact()
nID = 0
nValue = 0
sType1 = ""
sType2 = ""
nValue2 = 0
sName = ""
sEffect = ""

End Sub
Public Property Get Effect() As String

Effect = sEffect

End Property

Public Property Get Title() As String

Title = "ARTIFACT: " & sName

End Property
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Public Property Get Description() As String
stemp = ""

stemp = "Artifact card.  Character must have greater than or equal to "

stemp = stemp & Trim(Str(nValue1)) & " " & ConvertPowerCode(sType1) & " to use.  May play on to teammate with "
stemp = stemp & Trim(Str(nValue2)) & " " & ConvertPowerCode(sType2) & " or less."

Description = stemp
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Artifact WHERE ((Artifact.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewArtifact

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Artifact WHERE Artifact.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value

nValue = dbRec.Fields("T1_PW").Value
sType1 = dbRec.Fields("T1_SK").Value
sType2 = dbRec.Fields("T2_SK").Value
nValue2 = dbRec.Fields("T2_PW").Value
sName = dbRec.Fields("Character").Value
sEffect = dbRec.Fields("Effect").Value

dbRec.Close
db.Close

End Sub
Public Property Get Skill_1() As String
Skill_1 = ConvertPowerCode(sType1)

End Property
Public Property Get Skill_2() As String
Skill_2 = ConvertPowerCode(sType2)

End Property
Public Property Get Require_1() As Integer
Require_1 = nValue

End Property
Public Property Get Require_2() As Integer
Require_2 = nValue2

End Property
Public Property Get cardtype() As String
cardtype = "Artifact"

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
Public Property Get Attack_Frontline_BattleBonus() As Boolean
Attack_Frontline_BattleBonus = False

End Property
Public Property Get Attack_Frontline_GameBonus() As Boolean
Attack_Frontline_GameBonus = True

End Property
Public Property Get Attack_Frontline_Allies() As Boolean
Attack_Frontline_Allies = False

End Property
