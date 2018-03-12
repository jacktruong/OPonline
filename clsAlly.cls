VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim bLessThan As Boolean
Dim bMoreThan As Boolean
Dim nValue As Integer
Dim sType1 As String
Dim sType2 As String
Dim nValue2 As Integer
Dim sPictured As String
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
Private Sub NewAlly()
nID = 0
bLessThan = False
bMoreThan = False
nValue = 0
sType1 = ""
sType2 = ""
nValue2 = 0
sPictured = ""

End Sub
Public Property Get Title() As String

If bLessThan = True Then
    stemp = "<="
End If

If bMoreThan = True Then
    stemp = ">="
End If

stemp = stemp & Trim(Str(nValue)) & sType1 & ", " & Trim(Str(nValue2)) & sType2


Title = "ALLY: " & stemp & " [" & sPictured & "]"


End Property
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Public Property Get Description() As String
stemp = ""

stemp = "Ally Universe card.  Character must have "

If bLessThan = True Then
    stemp = stemp & " less than or equal to "
End If

If bMoreThan = True Then
    stemp = stemp & " greater than or equal to "
End If

stemp = stemp & Trim(Str(nValue1)) & " " & ConvertPowerCode(sType1) & " to use.  Acts as a "
stemp = stemp & Trim(Str(nValue2)) & " " & ConvertPowerCode(sType2) & " attack."

Description = stemp
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Ally WHERE ((Ally.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewAlly

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Ally WHERE Ally.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value

If dbRec.Fields("CR_LT").Value = ">=" Then
    bMoreThan = True
End If

If dbRec.Fields("CR_LT").Value = "<=" Then
    bLessThan = True
End If

nValue = dbRec.Fields("CR_PW").Value
sType1 = dbRec.Fields("CR_SK").Value
sType2 = dbRec.Fields("EF_SK").Value
nValue2 = dbRec.Fields("EF_PW").Value
sPictured = dbRec.Fields("Character").Value

dbRec.Close
db.Close

End Sub

Public Property Get GreaterThan() As Boolean
GreaterThan = bMoreThan
End Property
Public Property Get LessThan() As Boolean
LessThan = bLessThan

End Property
Public Property Get Skill_1() As String
Skill_1 = ConvertPowerCode(sType1)

End Property
Public Property Get Skill_2() As String
Skill_2 = ConvertPowerCode(sType2)

End Property
Public Property Get Require_1() As Integer
Require_1 = nValue

End Property
Public Property Get Effect() As Integer
Effect = nValue2

End Property
Public Property Get cardtype() As String
cardtype = "Ally Card"

End Property
Public Property Get Attack_isPlaced() As Boolean

Attack_isPlaced = True

End Property
Public Property Get Attack_VentureValue() As Integer

Attack_VentureValue = nValue2

End Property
Public Property Get Attack_NegativeVentureValue() As Integer

Attack_NegativeVentureValue = 0

End Property
Public Property Get Attack_isStringAttack() As Boolean

Attack_isStringAttack = True

End Property

