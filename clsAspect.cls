VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim sName As String
Dim sHomeBase As String
Dim sEffect As String
Dim bStringAttack As Boolean
Dim bCharEffect As Boolean
Dim bGameEffect As Boolean
Dim bBattleEffect As Boolean
Dim bPlaceAttack As Boolean
Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect) As Boolean
CanIPlay = True

End Function
Private Sub NewAspect()
nID = 0
sName = ""
sHomeBase = ""
sEffect = ""

bStringAttack = False
bCharEffect = False
bGameEffect = False
bBattleEffect = False
End Sub
Public Property Get Title() As String

Title = "ASPECT: " & sName & " [" & sHomeBase & "]"

End Property
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Public Property Get Description() As String
Description = sEffect
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Aspect WHERE ((Aspect.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewAspect

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Aspect WHERE Aspect.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value

sName = dbRec.Fields("Name").Value
sHomeBase = dbRec.Fields("Homebase").Value
sEffect = dbRec.Fields("Effect").Value
bStringAttack = dbRec.Fields("isStringAttack").Value
bCharEffect = dbRec.Fields("isCharacterEffect").Value
bGameEffect = dbRec.Fields("isGameEffect").Value
bBattleEffect = dbRec.Fields("isBattleEffect").Value
bPlaceAttack = dbRec.Fields("isPlacedAttack").Value

dbRec.Close
db.Close

End Sub
Public Property Get Effect() As String
Effect = sEffect

End Property
Public Property Get Name() As String
Name = sName
End Property
Public Property Get HomeBase() As String
HomeBase = sHomeBase
End Property
Public Property Get cardtype() As String
cardtype = "Aspect Card"

End Property
Public Property Get Attack_isPlaced() As Boolean

Attack_isPlaced = bPlaceAttack

End Property
Public Property Get Attack_VentureValue() As Integer

Attack_VentureValue = 0

End Property
Public Property Get Attack_NegativeVentureValue() As Integer

Attack_NegativeVentureValue = 0

End Property
Public Property Get Attack_isStringAttack() As Boolean

Attack_isStringAttack = bStringAttack

End Property
Public Property Get Attack_isGameEffect() As Boolean

Attack_isGameEffect = bGameEffect

End Property
Public Property Get Attack_isBattleEffect() As Boolean

Attack_isBattleEffect = bBattleEffect

End Property
Public Property Get Attack_isCharacterModifier() As Boolean

Attack_isCharacterModifier = bCharEffect

End Property
