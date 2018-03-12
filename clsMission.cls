VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nID As Integer
Dim sName As String
Dim nNumber As Integer
Dim cEvents As Collection
Private Sub NewMission()
nID = 0
sName = ""
nNumber = 0
Set cEvents = New Collection

End Sub
Public Property Get Events_Count()
Events_Count = cEvents.Count

End Property
Public Property Get Events_Name(Index)
Dim ccard

If Index < 1 Or Index > cEvents.Count Then Exit Property

Set ccard = cEvents.Item(Index)
Events_Name = ccard.Name

End Property
Public Property Get Events_Effect(Index)
Dim ccard

If Index < 1 Or Index > cEvents.Count Then Exit Property

Set ccard = cEvents.Item(Index)
Events_Effect = ccard.Effect

End Property
Public Property Get Events_ID(Index)
Dim ccard

If Index < 1 Or Index > cEvents.Count Then Exit Property

Set ccard = cEvents.Item(Index)
Events_ID = ccard.ID

End Property

Public Sub LoadEvents()
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

Set cEvents = New Collection

strSQL = "SELECT Events.ID, Events.MissionID, Events.Mission, Events.Name, Events.Effect From Events WHERE (((Events.MissionID)=0)) OR (((Events.MissionID)=" & Trim(Str(nID)) & ")) Order By Events.Name;"
dbRec.Open strSQL, db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

While Not dbRec.EOF

eid = dbRec.Fields("ID").Value

Set myEvent = New clsEvent
myEvent.Load eid
cEvents.Add myEvent

dbRec.MoveNext
Wend


dbRec.Close
db.Close
End Sub
Public Property Get Title() As String

Title = "MISSION: " & sName & " (" & Trim(Str(nNumber)) & " OF 7)"


End Property
Public Property Get isLandscape() As Boolean

isLandscape = True

End Property
Public Property Get Description() As String
Description = "MISSION: " & sName & " (" & Trim(Str(nNumber)) & " OF 7)"
End Property
Public Property Get ID() As Integer
ID = nID

End Property
Public Property Get Number() As Integer
Number = nNumber

End Property
Public Property Get Name() As String
Name = sName

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Missions WHERE ((Missions.id)=" & GetID & ");")

End Function
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewMission

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Missions WHERE Missions.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nID = dbRec.Fields("ID").Value
nNumber = dbRec.Fields("Number").Value
sName = dbRec.Fields("Name").Value

dbRec.Close
db.Close

End Sub

Public Property Get CardType() As String
CardType = "Mission"

End Property

