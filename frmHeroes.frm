VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHeroes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Heroes"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7500
   Icon            =   "frmHeroes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbLoad 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   6000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   5760
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Character"
      TabPicture(0)   =   "frmHeroes.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgHero"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgStat(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgStat(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgStat(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgStat(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblStat(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblStat(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblStat(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblStat(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtInherent"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk3Grid"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbCharacters"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Specials"
      TabPicture(1)   =   "frmHeroes.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgSpecial"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblSpecials"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblCode"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblOPD"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstSpecials"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtEffect"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtEffect 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2280
         Width           =   3495
      End
      Begin VB.ListBox lstSpecials 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   -74760
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.ComboBox cmbCharacters 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   6615
      End
      Begin VB.CheckBox chk3Grid 
         Caption         =   "3 Grid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtInherent 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   4080
         Width           =   6735
      End
      Begin VB.Label lblOPD 
         Caption         =   "* One Per Deck *"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -71040
         TabIndex        =   14
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblCode 
         Caption         =   "(AB)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68640
         TabIndex        =   13
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label lblSpecials 
         Caption         =   "Specials:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   3255
      End
      Begin VB.Image imgSpecial 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4095
         Left            =   -71040
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblStat 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   6360
         TabIndex        =   4
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblStat 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   6360
         TabIndex        =   3
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblStat 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   6360
         TabIndex        =   2
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblStat 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   6360
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
      Begin VB.Image imgStat 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Index           =   3
         Left            =   5760
         Picture         =   "frmHeroes.frx":12AA
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image imgStat 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Index           =   2
         Left            =   5760
         Picture         =   "frmHeroes.frx":30B3
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image imgStat 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Index           =   1
         Left            =   5760
         Picture         =   "frmHeroes.frx":4CA8
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgStat 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Index           =   0
         Left            =   5760
         Picture         =   "frmHeroes.frx":6A81
         Stretch         =   -1  'True
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgHero 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Left            =   240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   5415
      End
   End
   Begin VB.Label lblLoading 
      Caption         =   "Loading Characters:"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Setup"
      Begin VB.Menu mnuEditCardImages 
         Caption         =   "Edit Character/Special Card Images"
      End
   End
End
Attribute VB_Name = "frmHeroes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCurChar As Integer
Dim cCharacters As Collection

Private Sub chk3Grid_Click()
If chk3Grid.Value = 0 Then
    ShowCharacter
Else
    ShowCharacter True
End If

End Sub

Private Sub cmbCharacters_Click()
ShowCharacter

End Sub

Private Sub cmdBack_Click()
a = cmbCharacters.ListIndex

If a = 0 Then Exit Sub

cmbCharacters.ListIndex = cmbCharacters.ListIndex - 1

ShowCharacter
End Sub

Private Sub cmdForward_Click()
a = cmbCharacters.ListIndex

If a = cmbCharacters.ListCount - 1 Then Exit Sub

cmbCharacters.ListIndex = cmbCharacters.ListIndex + 1

ShowCharacter

End Sub


Private Sub Form_Activate()
DoEvents

LoadCharacters
LoadHeroList
cmbCharacters.ListIndex = 0


ShowCharacter
End Sub

Private Sub LoadHeroList()

Dim myhero As clsHero

For i = 1 To cCharacters.Count
Set myhero = New clsHero
Set myhero = cCharacters(i)
cmbCharacters.AddItem myhero.Name
cmbCharacters.ItemData(cmbCharacters.NewIndex) = myhero.ID
Next i


End Sub
Private Sub ShowCharacter(Optional Show3Grid As Boolean)
Dim myhero As clsHero

If IsMissing(Show3Grid) = True Then Show3Grid = False

Set myhero = New clsHero

nCurChar = cmbCharacters.ListIndex + 1

If nCurChar = 0 Then nCurChar = 1

If nCurChar > cCharacters.Count Then
    nCurChar = cCharacters.Count
    
End If

Set myhero = cCharacters(nCurChar)

Me.Caption = "Heroes: " & myhero.Name

If Show3Grid = False Then

If myhero.ImagePath <> "" Then
    X = Dir(myhero.ImagePath)
    If X = "" Then
        imgHero.Picture = LoadPicture(App.Path & "\NotFound.jpg")
    Else
        imgHero.Picture = LoadPicture(myhero.ImagePath)
    End If
End If

lblStat(0).Caption = myhero.Energy
lblStat(1).Caption = myhero.Fighting
lblStat(2).Caption = myhero.Strength
lblStat(3).Caption = myhero.Intellect

If myhero.Has3Grid = True Then
    chk3Grid.Visible = True
    chk3Grid.Value = 0
Else
    chk3Grid.Visible = False
End If

Else

If myhero.Image3Path <> "" Then
    X = Dir(myhero.Image3Path)
    If X = "" Then imgHero.Picture = LoadPicture(App.Path & "\NotFound.jpg")
    imgHero.Picture = LoadPicture(myhero.Image3Path)
End If

lblStat(0).Caption = myhero.Energy3
lblStat(1).Caption = myhero.Fighting3
lblStat(2).Caption = myhero.Strength3
lblStat(3).Caption = "-"

End If

If myhero.HasInherent = True Then
    txtInherent.Text = myhero.InherentAbility
Else
    txtInherent.Text = "No Inherent Ability"
End If

lstSpecials.Clear
txtEffect.Text = ""
imgSpecial.Picture = LoadPicture(sBlankImagePath)
lblCode.Caption = ""

lblSpecials.Caption = "Specials (" & myhero.Special_Count & "):"

For i = 1 To myhero.Special_Count
    If myhero.Special_OPD(i) = True Then
        lstSpecials.AddItem myhero.Special_Name(i) & " [OPD]"
    Else
        lstSpecials.AddItem myhero.Special_Name(i)
    End If
    
Next i


If lstSpecials.ListCount > 0 Then lstSpecials.ListIndex = 0
End Sub
Private Sub lstSpecials_Click()
LoadSpecial
End Sub
Private Sub LoadSpecial()
Dim myhero As clsHero
Dim myspecial As clsSpecial

Set myhero = New clsHero

nCurChar = cmbCharacters.ListIndex + 1

If nCurChar = 0 Then nCurChar = 1

If nCurChar > cCharacters.Count Then
    nCurChar = cCharacters.Count
    
End If

Set myhero = cCharacters(nCurChar)
Set myspecial = New clsSpecial

a = lstSpecials.ListIndex + 1

txtEffect.Text = myhero.Special_Effect(a)
imgSpecial.Picture = LoadPicture(myhero.Special_Image(a))
lblCode.Caption = "(" & myhero.Special_Code(a) & ")"

lblOPD.Visible = myhero.Special_OPD(a)

    

End Sub
Public Sub LoadCharacters()
Dim myhero As clsHero
Dim db As Database
Dim dbRec As Recordset

Set cCharacters = New Collection

Set db = OpenDatabase(dbName)
Set dbRec = db.OpenRecordSet("SELECT * FROM Characters ORDER BY Characters.Character;", dbOpenDynaset)

dbRec.MoveLast
dbRec.MoveFirst

pbLoad.Max = dbRec.RecordCount
pbLoad.Refresh

For i = 1 To dbRec.RecordCount

Set myhero = New clsHero
myhero.Load dbRec.Fields("ID").Value

cCharacters.Add myhero
pbLoad.Value = i
pbLoad.Refresh

dbRec.MoveNext
Next i

dbRec.Close
db.Close

lblLoading.Visible = False
pbLoad.Visible = False

End Sub

