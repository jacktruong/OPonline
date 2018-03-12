VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpecialEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim nId
Dim bAttackisPlaced As Boolean
Dim nVentureValue As Integer
Dim nNegativeVentureValue As Integer
Dim bStringAttack As Boolean

Dim bEffectMe As Boolean
Dim bAllies As Boolean
Dim bBattleBonus As Boolean
Dim bGameBonus As Boolean
Dim bConcedeStop As Boolean
Dim bConcedeAttack As Boolean

Dim sCharacter As String
Dim sName As String
Dim sEffect As String
Dim Scode As String * 2
Dim bOPD As Boolean
Dim sImagePath As String
Public Property Get Title() As String
Title = "SPECIAL: " & sName & " (" & sCharacter & ")"

End Property
Public Function LoadImage(GetID) As Boolean

LoadImage = LoadImageFromDatabase("SELECT * FROM Specials WHERE ((Specials.id)=" & GetID & ");")

End Function
Public Property Get isLandscape() As Boolean

isLandscape = False

End Property
Public Sub Load(GetID)
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

NewSpecial

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open "SELECT * FROM Specials WHERE Specials.ID=" & GetID & ";", db

If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

nId = dbRec.Fields("ID").Value
sCharacter = dbRec.Fields("Character").Value
sName = dbRec.Fields("Description").Value
sEffect = dbRec.Fields("Effect").Value
bOPD = dbRec.Fields("OPD").Value
bAttackisPlaced = dbRec.Fields("Attack").Value
nVentureValue = dbRec.Fields("Venturevalue").Value
nNegativeVentureValue = dbRec.Fields("VenNegValue").Value
bStringAttack = dbRec.Fields("StringAttack").Value
bEffectMe = dbRec.Fields("EffectMe").Value
bAllies = dbRec.Fields("Allies").Value
bBattleBonus = dbRec.Fields("BattleBonus").Value
bGameBonus = dbRec.Fields("GameBonus").Value
bConcedeStop = dbRec.Fields("ConcedeStop").Value
bConcedeAttack = dbRec.Fields("ConcedeAttack").Value

If IsNull(dbRec.Fields("Code").Value) = True Then
Scode = ""
Else
Scode = dbRec.Fields("Code").Value
End If

dbRec.Close
db.Close

End Sub
Private Sub NewSpecial()

nId = 0
sCharacter = ""
sName = ""
sEffect = ""
Scode = ""
bOPD = False
sImagePath = ""
bAttackisPlaced = False
nVentureValue = 0
nNegativeVentureValue = 0

End Sub
Public Property Get ID() As Variant
ID = nId
End Property
Public Property Get Character() As Variant

Character = sCharacter

End Property
Public Property Get Name() As Variant

Name = sName
End Property
Public Property Get Description() As Variant
Description = sEffect
End Property
Public Property Get Effect() As Variant
Effect = sEffect
End Property
Public Property Get Code() As Variant
Code = Scode

End Property
Public Property Get ImagePath() As Variant
ImagePath = sImagePath

End Property
Public Property Get isOPD() As Boolean
isOPD = bOPD

End Property
Public Function CanIPlay() As Boolean
'CanIPlay = False
'
'For i = 1 To 4
'    If cFrontLine.isCharacterDead(i) = False And (cFrontLine.Character_Name(i) = sCharacter) Then
        CanIPlay = True
'    End If
'Next i
    
End Function
Public Property Get cardtype() As String
cardtype = "Special Card"

End Property
Public Property Get Attack_isPlaced() As Boolean

Attack_isPlaced = bAttackisPlaced

End Property
Public Property Get Attack_VentureValue() As Integer

Attack_VentureValue = nVentureValue

End Property
Public Property Get Attack_NegativeVentureValue() As Integer

Attack_NegativeVentureValue = nNegativeVentureValue

End Property
Public Property Get Attack_isStringAttack() As Boolean

Attack_isStringAttack = bStringAttack

End Property
Public Property Get Attack_EffectsFrontline() As Boolean

Attack_EffectsFrontline = bEffectMe

End Property
Public Property Get Attack_Frontline_Allies() As Boolean

Attack_Frontline_Allies = bAllies

End Property
Public Property Get Attack_Frontline_BattleBonus() As Boolean

Attack_Frontline_BattleBonus = bBattleBonus

End Property
Public Property Get Attack_Frontline_GameBonus() As Boolean

Attack_Frontline_GameBonus = bGameBonus

End Property
Public Property Get Attack_StopsConcede() As Boolean
Attack_StopsConcede = bConcedeStop

End Property
Public Property Get Attack_PostConcessionAttack() As Boolean

Attack_PostConcessionAttack = bConcedeAttack

End Property

