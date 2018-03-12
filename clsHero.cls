VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim sname
Dim nId
Dim nEnergy
Dim nIntellect
Dim nFighting
Dim nStrength
Dim cSpecials As Collection
Dim sInherent
Dim bInherent As Boolean
Dim nCloneId As Integer

Public Sub Load(HeroId)

Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Characters WHERE Characters.ID=" & HeroId & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nId = dbRec.Fields("ID").Value
sname = dbRec.Fields("Character").Value
nStrength = dbRec.Fields("S").Value
nIntellect = dbRec.Fields("I").Value
nFighting = dbRec.Fields("F").Value
nEnergy = dbRec.Fields("E").Value
sInherent = dbRec.Fields("Inherent").Value
bInherent = dbRec.Fields("HasInherent").Value


If dbRec.Fields("Clone").Value = True Then
    nCloneId = GetIDFromName(dbRec.Fields("CloneCharacter").Value)
Else
    nCloneId = 0
End If

dbRec.Close
db.Close


End Sub
Public Property Get Name() As String

Name = sname

End Property
Public Property Get ID()

ID = nId

End Property

Public Function GetIDFromName(sFindName) As String
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

strSQL = "SELECT * FROM CHARACTERS WHERE (((Characters.Character)=" & Chr(34) & sFindName & Chr(34) & "));"

dbRec.Open strSQL, db

If dbRec.EOF = True Then
    GetIDFromName = 0
    dbRec.Close
    db.Close
    Exit Function
End If

GetIDFromName = dbRec.Fields("ID").Value

dbRec.Close
db.Close


End Function
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Characters WHERE ((Characters.id)=" & GetID & ");")

End Function
Public Property Get Strength() As Integer

Strength = nStrength

End Property
Public Property Get Energy() As Integer

Energy = nEnergy

End Property
Public Property Get Fighting() As Integer

Fighting = nFighting

End Property
Public Property Get Intellect() As Integer

Intellect = nIntellect
End Property
Public Property Get Special_Count()
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

If nCloneId > 0 Then
    dbRec.Open "SELECT * FROM Specials WHERE Specials.CharID=" & nCloneId & ";", db
Else
    dbRec.Open "SELECT * FROM Specials WHERE Specials.CharID=" & nId & ";", db
End If

Set cSpecials = New Collection

If dbRec.EOF = True Then
    Special_Count = 0
    dbRec.Close
    db.Close
    Exit Property
End If

Do Until dbRec.EOF

Set myspecial = New clsSpecial
myspecial.Load dbRec.Fields("ID").Value
cSpecials.Add myspecial

dbRec.MoveNext
Loop

dbRec.Close
db.Close

Special_Count = cSpecials.Count

End Property
Public Property Get Special_Name(nSpecID)

Special_Name = cSpecials.Item(nSpecID).Name


End Property
Public Property Get Special_ID(nSpecID)

Special_ID = cSpecials.Item(nSpecID).ID


End Property

Public Property Get Special_OPD(nSpecID) As Boolean

If cSpecials.Item(nSpecID).isOPD = True Then
    Special_OPD = True
Else
    Special_OPD = False
End If

End Property
Public Property Get InherentAbility() As String
    InherentAbility = sInherent
End Property
Public Property Get HasInherent() As Boolean
    HasInherent = bInherent
    
End Property

Public Property Get Strength3() As Integer

Strength3 = nStrength

End Property
Public Property Get Energy3() As Integer

Energy3 = nEnergy

End Property
Public Property Get Fighting3() As Integer

Fighting3 = nFighting

End Property
Public Property Get Title()

Title = "CHARACTER: " & sname

End Property
Public Property Get isLandscape() As Boolean
isLandscape = True

End Property
Public Property Get CardType()

CardType = "Hero"

End Property
