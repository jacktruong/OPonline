VERSION 5.00
Begin VB.Form frmSpecialEditor 
   Caption         =   "Edit Special Characteristics"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAlpha 
      Caption         =   "Alpha"
      Height          =   255
      Left            =   8760
      TabIndex        =   21
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Low ID"
      Height          =   255
      Left            =   9960
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtBeginID 
      Height          =   285
      Left            =   8760
      TabIndex        =   19
      Text            =   "2139"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkConcedeAttack 
      Caption         =   "Make an attack after concession"
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   840
      Width           =   2775
   End
   Begin VB.CheckBox chkConcedeStop 
      Caption         =   "Stop opponent from conceding"
      Height          =   255
      Left            =   8280
      TabIndex        =   17
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CheckBox chkGameBonus 
      Caption         =   "Game Bonus"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CheckBox chkBattleBonus 
      Caption         =   "Battle Bonus"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CheckBox chkAllies 
      Caption         =   """Allies from the Deep"" type special"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.CheckBox chkStringAttack 
      Caption         =   "Treat as String Attack"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CheckBox chkPR 
      Caption         =   "Place to Permanent Record upon hit"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox txtNegVenValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   5
      Text            =   "0"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtVentureValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   3
      Text            =   "0"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ListBox lstSpecials 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Effects Concession:"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Effect to Opponent:"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Effect to Frontline:"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label label1 
      Caption         =   "Negative Venture Value:"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   4
      Top             =   2085
      Width           =   1935
   End
   Begin VB.Label label1 
      Caption         =   "Venture Value:"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label lblEffect 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmSpecialEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myspecial As clsSpecialEdit


Private Sub Check1_Click()

End Sub

Private Sub cmdFind_Click()
If txtFind.Text = "" Then Exit Sub

Dim db As Database
Dim dbRec As Recordset

Set db = OpenDatabase(App.Path & "\overpower.mdb")

lstSpecials.Clear

strqry = "SELECT * From Specials WHERE (((Specials.Effect) Like '*" & txtFind.Text & "*')) ORDER BY SPECIALS.CHARACTER, SPECIALS.DESCRIPTION;"

Set dbRec = db.OpenRecordset(strqry, dbOpenDynaset)

dbRec.MoveFirst

Do Until dbRec.EOF

a$ = dbRec.Fields("Character").Value & "-->" & dbRec.Fields("Description").Value
lstSpecials.AddItem a$
lstSpecials.ItemData(lstSpecials.NewIndex) = dbRec.Fields("ID").Value

dbRec.MoveNext
Loop

dbRec.Close
db.Close

End Sub

Private Sub cmdSave_Click()
Dim db As Database
Dim dbRec As Recordset

Set db = OpenDatabase(App.Path & "\overpower.mdb")

Set dbRec = db.OpenRecordset("SELECT * FROM Specials WHERE Specials.ID=" & Trim(Str(lstSpecials.ItemData(lstSpecials.ListIndex))) & ";", dbOpenDynaset)


If dbRec.EOF = True Then
    dbRec.Close
    db.Close
    Exit Sub
End If

dbRec.Edit
dbRec.Fields("VentureValue").Value = Val(txtVentureValue.Text)
dbRec.Fields("VenNegValue").Value = Val(txtNegVenValue.Text)

If chkPR.Value = 1 Then
    dbRec.Fields("Attack").Value = True
Else
    dbRec.Fields("Attack").Value = False
End If

If chkStringAttack.Value = 1 Then
    dbRec.Fields("StringAttack").Value = True
Else
    dbRec.Fields("StringAttack").Value = False
End If

If chkAllies.Value = 1 Then
    dbRec.Fields("Allies").Value = True
Else
    dbRec.Fields("allies").Value = False
End If

If chkBattleBonus.Value = 1 Then
    dbRec.Fields("battlebonus").Value = True
Else
    dbRec.Fields("battlebonus").Value = False
End If

If chkGameBonus.Value = 1 Then
    dbRec.Fields("gamebonus").Value = True
Else
    dbRec.Fields("gamebonus").Value = False
End If

If chkGameBonus.Value = 1 Or chkAllies.Value = 1 Or chkBattleBonus.Value = 1 Then
    dbRec.Fields("Effectme").Value = True
Else
    dbRec.Fields("Effectme").Value = False
End If

If chkConcedeStop.Value = 1 Then
    dbRec.Fields("ConcedeStop").Value = True
Else
    dbRec.Fields("ConcedeStop").Value = False
End If

If chkConcedeAttack.Value = 1 Then
    dbRec.Fields("ConcedeAttack").Value = True
Else
    dbRec.Fields("ConcedeAttack").Value = False
End If

dbRec.Update
dbRec.Close
db.Close

End Sub

Private Sub Command1_Click()

Dim db As Database
Dim dbRec As Recordset

Set db = OpenDatabase(App.Path & "\overpower.mdb")

lstSpecials.Clear

If chkAlpha.Value = 0 Then
    strqry = "SELECT * FROM SPECIALS WHERE (SPECIALS.ID >=" & txtBeginID.Text & ") ORDER BY SPECIALS.ID;"
Else
    strqry = "SELECT * FROM SPECIALS WHERE (SPECIALS.ID >=" & txtBeginID.Text & ") ORDER BY SPECIALS.CHARACTER, SPECIALS.DESCRIPTION;"
End If

Set dbRec = db.OpenRecordset(strqry, dbOpenDynaset)

dbRec.MoveFirst

Do Until dbRec.EOF

a$ = dbRec.Fields("Character").Value & "-->" & dbRec.Fields("Description").Value
lstSpecials.AddItem a$
lstSpecials.ItemData(lstSpecials.NewIndex) = dbRec.Fields("ID").Value

dbRec.MoveNext
Loop

dbRec.Close
db.Close

End Sub

Private Sub Form_Load()
dbName = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Overpower.mdb"

Dim db As Database
Dim dbRec As Recordset

Set db = OpenDatabase(App.Path & "\overpower.mdb")

lstSpecials.Clear

strqry = "SELECT * FROM SPECIALS ORDER BY SPECIALS.CHARACTER, SPECIALS.DESCRIPTION;"

Set dbRec = db.OpenRecordset(strqry, dbOpenDynaset)

dbRec.MoveFirst

Do Until dbRec.EOF

a$ = dbRec.Fields("Character").Value & "-->" & dbRec.Fields("Description").Value
lstSpecials.AddItem a$
lstSpecials.ItemData(lstSpecials.NewIndex) = dbRec.Fields("ID").Value

dbRec.MoveNext
Loop

dbRec.Close
db.Close

End Sub

Private Sub lstSpecials_Click()
Set myspecial = New clsSpecialEdit

myspecial.Load lstSpecials.ItemData(lstSpecials.ListIndex)

lblEffect.Caption = myspecial.Effect

txtVentureValue.Text = myspecial.Attack_VentureValue
txtNegVenValue.Text = myspecial.Attack_NegativeVentureValue

If myspecial.Attack_isPlaced = True Then
    chkPR.Value = 1
Else
    chkPR.Value = 0
End If

If myspecial.Attack_isStringAttack = True Then
    chkStringAttack.Value = 1
Else
    chkStringAttack.Value = 0
End If

If myspecial.Attack_Frontline_Allies = True Then
    chkAllies.Value = 1
Else
    chkAllies.Value = 0
End If

If myspecial.Attack_Frontline_BattleBonus = True Then
    chkBattleBonus.Value = 1
Else
    chkBattleBonus.Value = 0
End If

If myspecial.Attack_Frontline_GameBonus = True Then
    chkGameBonus.Value = 1
Else
    chkGameBonus.Value = 0
End If

If myspecial.Attack_StopsConcede = True Then
    chkConcedeStop.Value = 1
Else
    chkConcedeStop.Value = 0
End If

If myspecial.Attack_PostConcessionAttack = True Then
    chkConcedeAttack.Value = 1
Else
    chkConcedeAttack.Value = 0
End If

txtVentureValue.SetFocus
txtVentureValue.SelStart = 0
txtVentureValue.SelLength = Len(txtVentureValue.Text)

End Sub
