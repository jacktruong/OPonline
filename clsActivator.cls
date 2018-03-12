VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActivator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nId
Dim sName As String

Dim cSpecials As Collection
Public Property Get isLandscape() As Boolean

isLandscape = True

End Property
Public Property Get Title() As String

Title = "ACTIVATOR: " & sName

End Property
Public Property Get ID() As Variant
ID = nId
End Property
Public Property Get Name() As Variant
Name = sName

End Property
Private Sub NewActivator()
nId = 0
sName = ""

Set cSpecials = New Collection

End Sub
Private Sub Class_Initialize()

NewActivator

End Sub
Public Sub Load(HeroId)
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset
Dim myspecial As clsSpecial

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

objRS.Open "SELECT * FROM Characters WHERE Characters.ID=" & HeroId & ";", db


NewActivator

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

nId = objRS.Fields("ID").Value

If objRS.Fields("Clone").Value = True Then
    sName = objRS.Fields("CloneCharacter").Value
Else
    sName = objRS.Fields("Character").Value
End If


objRS.Close

'LoadSpecials
If sClone <> "" Then

    objRS.Open "SELECT * FROM Specials WHERE Specials.Character=" & Chr(34) & sClone & Chr(34) & ";", db

Else
    objRS.Open "SELECT * FROM Specials WHERE Specials.Character=" & Chr(34) & sName & Chr(34) & ";", db

End If

If objRS.EOF = False Then


Do Until objRS.EOF = True

Set myspecial = New clsSpecial
myspecial.Load objRS.Fields("ID").Value

cSpecials.Add myspecial

objRS.MoveNext
Loop


End If

objRS.Close
db.Close

End Sub
Public Property Get Special_Count() As Integer
Special_Count = cSpecials.Count

End Property
Public Property Get Special_Name(Index) As String
Dim myspecial As clsSpecial

If Index < 0 Or Index > cSpecials.Count Then
    Index = 1
End If

Set myspecial = New clsSpecial
Set myspecial = cSpecials.Item(Index)

Special_Name = myspecial.Name


End Property
Public Property Get Special_ID(Index) As Integer
Dim myspecial As clsSpecial

If Index < 0 Or Index > cSpecials.Count Then
    Index = 1
End If

Set myspecial = New clsSpecial
Set myspecial = cSpecials.Item(Index)

Special_ID = myspecial.ID


End Property
Public Property Get Special_Effect(Index) As Variant
Dim myspecial As clsSpecial

If Index < 0 Or Index > cSpecials.Count Then
    Index = 1
End If

Set myspecial = New clsSpecial
Set myspecial = cSpecials.Item(Index)

Special_Effect = myspecial.Effect


End Property
Public Property Get Special_Code(Index) As Variant
Dim myspecial As clsSpecial

If Index < 0 Or Index > cSpecials.Count Then
    Index = 1
End If

Set myspecial = New clsSpecial
Set myspecial = cSpecials.Item(Index)

Special_Code = myspecial.Code

End Property
Public Property Get Special_OPD(Index) As Boolean
Dim myspecial As clsSpecial

If Index < 0 Or Index > cSpecials.Count Then
    Index = 1
End If

Set myspecial = New clsSpecial
Set myspecial = cSpecials.Item(Index)

Special_OPD = myspecial.isOPD

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Characters WHERE ((Characters.id)=" & GetID & ");")

End Function
Public Property Get CardType() As String
CardType = "Activator"

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
