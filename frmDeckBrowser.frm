VERSION 5.00
Begin VB.Form frmDeckBrowser 
   Caption         =   "Deck Browser"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox lstBSDeck 
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3720
      Width           =   5895
   End
   Begin VB.ComboBox lstDeck 
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3360
      Width           =   5895
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Preview Decks"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ListBox lstDecks 
      Height          =   4155
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   4440
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "BSD (0):"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   3735
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Deck (0):"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   3380
      Width           =   975
   End
   Begin VB.Image imgFrontLine 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   0
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image imgFrontLine 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   1
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image imgFrontLine 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   2
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image ImgReserve 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Image imgBattlesite 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image imgHomebase 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   7200
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Available Decks: (0)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   9
      Left            =   7200
      Top             =   480
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   8
      Left            =   5400
      Top             =   480
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   7
      Left            =   3600
      Top             =   480
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   10
      Left            =   5400
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1215
      Index           =   1
      Left            =   7200
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1215
      Index           =   0
      Left            =   9000
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmDeckBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkPreview_Click()

If chkPreview.Value = 0 Then
For i = 0 To 2
    Set imgFrontLine(i).Picture = Nothing
Next i

Set imgHomebase.Picture = Nothing
Set imgBattlesite.Picture = Nothing
Set ImgReserve.Picture = Nothing

lstDeck.Clear
Label2.Caption = "Deck (0):"

End If

End Sub

Private Sub cmdCancel_Click()
cmdOpen.Tag = ""
Me.Hide

End Sub

Private Sub cmdOpen_Click()
If lstDecks.ListIndex = -1 Then Exit Sub

cmdOpen.Tag = lstDecks.List(lstDecks.ListIndex)
Me.Hide

End Sub

Private Sub Form_Load()
LoadDecks
End Sub
Private Sub LoadDecks()

X = Dir(App.Path & "\decks\*.dat")

looper:

If X <> "" Then

a$ = Left(X, Len(X) - 4)
lstDecks.AddItem a$

X = Dir()
GoTo looper

End If


Label1(0).Caption = "Available Decks (" & Trim(Str(lstDecks.ListCount)) & "):"

End Sub
Private Sub PreviewDeck(sFileName)
Dim myh As clsHero
Dim heroes(4)

For i = 0 To 2
    Set imgFrontLine(i).Picture = Nothing
Next i

Set ImgReserve.Picture = Nothing
Set imgBattlesite.Picture = Nothing
Set imgHomebase.Picture = Nothing

lblStatus.Caption = "Loading.."
lstDeck.Clear
lstBSDeck.Clear
Label2.Caption = "Deck (0):"
Label3.Caption = "BSD (0):"

X = FreeFile

Open sFileName For Input As #X

'read 4 heroes
Set myh = New clsHero

For i = 1 To 4
Line Input #X, a$
heroes(i) = Val(GetVal(a$))
Next i

Line Input #X, a$
nreserve = Val(GetVal(a$))

fc = 0

For i = 1 To 4
    If i = nreserve Then
        If myh.LoadImage(heroes(i)) = True Then
            ImgReserve.Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            ImgReserve.Picture = LoadPicture(sBlankImagePath)
        End If

    
    Else
        If myh.LoadImage(heroes(i)) = True Then
            imgFrontLine(fc).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            imgFrontLine(fc).Picture = LoadPicture(sBlankImagePath)
        End If
        
        fc = fc + 1
    End If
    
Me.Refresh
LStatus
DoEvents
Next i

Dim sHomeBase As clsHomebase
Set sHomeBase = New clsHomebase

Line Input #X, a$

n = Val(GetVal(a$))
If n = 0 Then
    Set imgHomebase.Picture = Nothing
Else
    If sHomeBase.LoadImage(n) = True Then
            imgHomebase.Picture = LoadPicture(App.Path & "\temppic.jpg")
   Else
            imgHomebase.Picture = LoadPicture(sBlankImagePath)
    End If
End If

Set sHomeBase = Nothing
LStatus
Me.Refresh
DoEvents

Dim sbattlesite As clsBattlesite
Set sbattlesite = New clsBattlesite

Line Input #X, a$

n = Val(GetVal(a$))
If n = 0 Then
    Set imgBattlesite.Picture = Nothing
Else
    If sbattlesite.LoadImage(n) = True Then
            imgBattlesite.Picture = LoadPicture(App.Path & "\temppic.jpg")
   Else
            imgBattlesite.Picture = LoadPicture(sBlankImagePath)
    End If
End If

Set sbattlesite = Nothing
LStatus
Me.Refresh
DoEvents

''Get Mission
Line Input #X, a$

'Get number of cards in deck
Line Input #X, a$
ncards = Val(GetVal(a$))

Label2.Caption = "Deck (" & Trim(Str(ncards)) & "):"

For i = 1 To ncards
LStatus
DoEvents

Line Input #X, a$
scardtype = GetVal(a$)

Line Input #X, a$
ncardid = Val(GetVal(a$))

Select Case scardtype

Case "Activator"
    Set myActivator = New clsActivator
    myActivator.Load ncardid
    lstDeck.AddItem myActivator.Title

Case "Artifact"
    Set myArtifact = New clsArtifact
    myArtifact.Load ncardid
    lstDeck.AddItem myArtifact.Title

Case "Ally Card"
    Set myAlly = New clsAlly
    myAlly.Load ncardid
    lstDeck.AddItem myAlly.Title

Case "Aspect Card"
    Set myAspect = New clsAspect
    myAspect.Load ncardid
    lstDeck.AddItem myAspect.Title

Case "Basic Universe"
    Set myBasic = New clsBasicUniverse
    myBasic.Load ncardid
    lstDeck.AddItem myBasic.Title

Case "Double Shot"
    Set myDoubleShot = New clsDoubleShot
    myDoubleShot.Load ncardid
    lstDeck.AddItem myDoubleShot.Title

Case "Event"
    Set myEvent = New clsEvent
    myEvent.Load ncardid
    lstDeck.AddItem myEvent.Title

Case "Power Card"
    Set myPower = New clsPowerCard
    myPower.Load ncardid
    lstDeck.AddItem myPower.Title

Case "Special Card"
    Set myspecial = New clsSpecial
    myspecial.Load ncardid
    lstDeck.AddItem myspecial.Title

Case "Teamwork"
    Set myTeamwork = New clsTeamwork
    myTeamwork.Load ncardid
    lstDeck.AddItem myTeamwork.Title

Case "Training"
    Set myTraining = New clsTraining
    myTraining.Load ncardid
    lstDeck.AddItem myTraining.Title

Case Else
End Select

Next i

If lstDeck.ListCount > 0 Then lstDeck.ListIndex = 0


'Load Battlesite deck

Line Input #X, a$
nbd = Val(GetVal(a$))

Label3.Caption = "BSD (" & Trim(Str(nbd)) & "):"

For i = 1 To nbd

Line Input #X, a$
scardtype = GetVal(a$)

Line Input #X, a$
ncardid = Val(GetVal(a$))

Select Case scardtype

Case "Special Card"
    Set myspecial = New clsSpecial
    myspecial.Load ncardid
    lstBSDeck.AddItem myspecial.Title

Case Else
End Select

LStatus
DoEvents
Next i

If lstBSDeck.ListCount > 0 Then lstBSDeck.ListIndex = 0

Close #X

lblStatus.Caption = ""
'
End Sub
Private Sub LStatus()

a$ = lblStatus.Caption
If Len(a$) > 40 Then
    lblStatus.Caption = "Loading.."
Else
    lblStatus.Caption = a$ & "."
End If


End Sub
Private Sub lstDecks_Click()
Me.Caption = "Deck Browser: " & lstDecks.List(lstDecks.ListIndex)

If lstDecks.ListIndex = -1 Then
    cmdOpen.Enabled = False
    cmdOpen.Tag = ""
Else
    cmdOpen.Enabled = True
End If

If chkPreview.Value = 1 Then PreviewDeck App.Path & "\Decks\" & lstDecks.List(lstDecks.ListIndex) & ".dat"

End Sub
