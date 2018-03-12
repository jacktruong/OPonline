VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim sName As String
Dim sEffect As String
Dim smission As String
Dim nMissionID As Integer
Public Function CanIPlay(nEnergy, nFighting, nStrength, nIntellect) As Boolean
CanIPlay = True

End Function
Public Property Get isLandscape() As Boolean

isLandscape = True

End Property
Public Property Get Mission() As Boolean

Mission = smission

End Property
Public Property Get Name() As String
Name = sName
End Property
Public Property Get MissionID() As Integer
MissionID = nMissionID

End Property
Private Sub NewEvent()
nID = 0
sName = ""
sEffect = ""

End Sub
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewEvent

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Events WHERE Events.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value
sName = dbRec.Fields("Name").Value
sEffect = dbRec.Fields("Effect").Value
smission = dbRec.Fields("Mission").Value
nMissionID = dbRec.Fields("MissionID").Value

dbRec.Close
db.Close

End Sub
Public Property Get Title() As String

Title = "EVENT: " & sName

End Property
Public Property Get Description() As String
Description = sEffect

End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Events WHERE ((Events.id)=" & GetID & ");")

End Function
Public Property Get cardtype() As String
cardtype = "Event"

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

