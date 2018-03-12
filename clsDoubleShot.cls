VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoubleShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim nValue As Integer
Dim sType1 As String
Dim nValue2 As Integer
Dim sType2 As String
Dim sPictured As String
Dim Stype3 As String
Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect) As Boolean
CanIPlay = False
If cFrontLine.LiveCharacterCount = 1 Then
    CanIPlay = False
    Exit Function
End If

End Function
Private Sub NewDoubleShot()

nID = 0
nValue = 0
sType1 = ""
sType2 = ""
nValue2 = 0
sPictured = ""
Stype3 = ""

End Sub
Public Property Get Title() As String

Title = "DOUBLESHOT: " & Trim(Str(nValue)) & sType1 & "/" & Trim(Str(nValue2)) & sType2

End Property
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Public Property Get Description() As String
stemp = Trim(Str(nValue)) & " " & ConvertPowerCode(sType1) & " to use.  Acts as a level 4 " & ConvertPowerCode(Stype3) & " Power Card.  "
stemp = stemp & "Teammate with " & Trim(Str(nValue2)) & " " & ConvertPowerCode(sType2) & " must combine with " & ConvertPowerCode(sType2) & " Power Card 1 through 4 for a single attack or defense."

Description = stemp
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Doubleshot WHERE ((Doubleshot.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewDoubleShot

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Doubleshot WHERE doubleshot.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value
nValue = dbRec.Fields("T1_PW").Value
nValue2 = Val(dbRec.Fields("T2_PW").Value)
sType1 = dbRec.Fields("T1_SK").Value
sType2 = dbRec.Fields("T2_SK").Value
Stype3 = dbRec.Fields("Effect").Value

sPictured = dbRec.Fields("Character").Value

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
Public Property Get Skill_3() As String
Skill_3 = ConvertPowerCode(Stype3)

End Property
Public Property Get Effect() As String
stemp = Trim(Str(nValue)) & " " & ConvertPowerCode(sType1) & " to use.  Acts as a level 4 " & ConvertPowerCode(Stype3) & " Power Card.  "
stemp = stemp & "Teammate with " & Trim(Str(nValue2)) & " " & ConvertPowerCode(sType2) & " must combine with " & ConvertPowerCode(sType2) & " Power Card 1 through 4 for a single attack or defense."

Effect = stemp

End Property

Public Property Get cardtype() As String
cardtype = "Double Shot"

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

