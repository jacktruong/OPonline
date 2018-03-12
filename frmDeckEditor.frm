VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDeckEditor 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overpower Deck Editor"
   ClientHeight    =   6135
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11130
   ControlBox      =   0   'False
   Icon            =   "frmDeckEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Overpower\Overpower.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Characters"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      TabsPerRow      =   11
      TabHeight       =   520
      TabCaption(0)   =   "Heroes"
      TabPicture(0)   =   "frmDeckEditor.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "shpReserve"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgHero(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgHero(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgHero(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgHero(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblHeroes(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblHeroes(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lstGridLimit"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkEnforce"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdReserve"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdRemove"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAdd"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lstShow"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmD1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DBGrid1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Locations"
      TabPicture(1)   =   "frmDeckEditor.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblBSDeckCount"
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(2)=   "Label1(4)"
      Tab(1).Control(3)=   "lblBattlesite"
      Tab(1).Control(4)=   "lblHBEffect"
      Tab(1).Control(5)=   "lblHBCharacters"
      Tab(1).Control(6)=   "lblHomeBase"
      Tab(1).Control(7)=   "Label1(3)"
      Tab(1).Control(8)=   "Label1(2)"
      Tab(1).Control(9)=   "cmdRemoveBSCard"
      Tab(1).Control(10)=   "cmdClearBattlesite"
      Tab(1).Control(11)=   "lstBSDeck"
      Tab(1).Control(12)=   "cmdClearHomebase"
      Tab(1).Control(13)=   "cmdBattlesite"
      Tab(1).Control(14)=   "cmdHomebase"
      Tab(1).Control(15)=   "lstHomebaseCharacters"
      Tab(1).Control(16)=   "txtHomeBaseEffect"
      Tab(1).Control(17)=   "lstLocations"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Specials"
      TabPicture(2)   =   "frmDeckEditor.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(0)"
      Tab(2).Control(1)=   "Label1(1)"
      Tab(2).Control(2)=   "imgCardDetail"
      Tab(2).Control(3)=   "cmdAddToBattlesite"
      Tab(2).Control(4)=   "lstCharacters"
      Tab(2).Control(5)=   "lstSpecials"
      Tab(2).Control(6)=   "txtSpecialEffect"
      Tab(2).Control(7)=   "cmdAddSpecial"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Mission"
      TabPicture(3)   =   "frmDeckEditor.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(5)"
      Tab(3).Control(1)=   "Label1(6)"
      Tab(3).Control(2)=   "imgMissionCard"
      Tab(3).Control(3)=   "lblCurrentMission(7)"
      Tab(3).Control(4)=   "lblMission"
      Tab(3).Control(5)=   "lstMissions"
      Tab(3).Control(6)=   "lstEvents"
      Tab(3).Control(7)=   "cmdUseMission"
      Tab(3).Control(8)=   "txtEventEffect"
      Tab(3).Control(9)=   "cmdAddEvent"
      Tab(3).Control(10)=   "cmdFindEvent"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Power Cards"
      TabPicture(4)   =   "frmDeckEditor.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(7)"
      Tab(4).Control(1)=   "imgPowerCard"
      Tab(4).Control(2)=   "lstPowerCards"
      Tab(4).Control(3)=   "lstShowPower"
      Tab(4).Control(4)=   "chkPlayable"
      Tab(4).Control(5)=   "cmdAddPower"
      Tab(4).Control(6)=   "Command2"
      Tab(4).Control(7)=   "lstHeroStats"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Universe"
      TabPicture(5)   =   "frmDeckEditor.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(8)"
      Tab(5).Control(1)=   "Label1(10)"
      Tab(5).Control(2)=   "Label1(9)"
      Tab(5).Control(3)=   "imgUniverse"
      Tab(5).Control(4)=   "Label1(14)"
      Tab(5).Control(5)=   "lstBasicUniverse"
      Tab(5).Control(6)=   "cmdAddBasic"
      Tab(5).Control(7)=   "chkPlayBasic"
      Tab(5).Control(8)=   "lstTeamwork"
      Tab(5).Control(9)=   "cmdAddTeamwork"
      Tab(5).Control(10)=   "chkPlayTeamwork"
      Tab(5).Control(11)=   "lstTraining"
      Tab(5).Control(12)=   "cmdAddTraining"
      Tab(5).Control(13)=   "chkPlayTraining"
      Tab(5).Control(14)=   "cmdAddAlly"
      Tab(5).Control(15)=   "lstAllys"
      Tab(5).Control(16)=   "lstherostats2"
      Tab(5).ControlCount=   17
      TabCaption(6)   =   "Tactic"
      TabPicture(6)   =   "frmDeckEditor.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1(11)"
      Tab(6).Control(1)=   "imgArtifact"
      Tab(6).Control(2)=   "Label1(12)"
      Tab(6).Control(3)=   "Label1(13)"
      Tab(6).Control(4)=   "lstArtifacts"
      Tab(6).Control(5)=   "txtArtifactEffect"
      Tab(6).Control(6)=   "cmdAddArtifact"
      Tab(6).Control(7)=   "lstAspects"
      Tab(6).Control(8)=   "cmdAddAspect"
      Tab(6).Control(9)=   "lstDoubleshot"
      Tab(6).Control(10)=   "cmdAddDoubleShot"
      Tab(6).ControlCount=   11
      TabCaption(7)   =   "Find Specials"
      TabPicture(7)   =   "frmDeckEditor.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label4"
      Tab(7).Control(1)=   "lblMatches"
      Tab(7).Control(2)=   "imgFindSpecial"
      Tab(7).Control(3)=   "lstFindSpecials"
      Tab(7).Control(4)=   "txtFind"
      Tab(7).Control(5)=   "cmdFind"
      Tab(7).Control(6)=   "cmdFindSpecialAdd"
      Tab(7).ControlCount=   7
      TabCaption(8)   =   "Find Heroes"
      TabPicture(8)   =   "frmDeckEditor.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Command1"
      Tab(8).Control(1)=   "cmdFind2"
      Tab(8).Control(2)=   "txtTotal"
      Tab(8).Control(3)=   "Frame1"
      Tab(8).Control(4)=   "txtMinI"
      Tab(8).Control(5)=   "txtMinS"
      Tab(8).Control(6)=   "txtMinF"
      Tab(8).Control(7)=   "txtMinE"
      Tab(8).Control(8)=   "lstHeroMatches"
      Tab(8).Control(9)=   "Label5(3)"
      Tab(8).Control(10)=   "Label7"
      Tab(8).Control(11)=   "Label5(2)"
      Tab(8).Control(12)=   "Label5(1)"
      Tab(8).Control(13)=   "Label6"
      Tab(8).Control(14)=   "Label5(0)"
      Tab(8).ControlCount=   15
      TabCaption(9)   =   "Find Locations"
      TabPicture(9)   =   "frmDeckEditor.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "cmdAddHero"
      Tab(9).Control(1)=   "cmdRandomLocation"
      Tab(9).Control(2)=   "cmdFindLocs"
      Tab(9).Control(3)=   "txtFindLoc"
      Tab(9).Control(4)=   "cmdShowAllLocs"
      Tab(9).Control(5)=   "cmdUseBattlesite2"
      Tab(9).Control(6)=   "cmdUseHomebase2"
      Tab(9).Control(7)=   "Frame2"
      Tab(9).Control(8)=   "lstLocCharacters"
      Tab(9).Control(9)=   "txtLocEffect"
      Tab(9).Control(10)=   "lstFindLocations"
      Tab(9).Control(11)=   "Label9"
      Tab(9).Control(12)=   "Label8"
      Tab(9).Control(13)=   "Label1(15)"
      Tab(9).ControlCount=   14
      TabCaption(10)  =   "Deck Overview"
      TabPicture(10)  =   "frmDeckEditor.frx":055A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lblCard(9)"
      Tab(10).Control(1)=   "lblCard(8)"
      Tab(10).Control(2)=   "lblCard(7)"
      Tab(10).Control(3)=   "lblCard(6)"
      Tab(10).Control(4)=   "lblCard(5)"
      Tab(10).Control(5)=   "lblCard(4)"
      Tab(10).Control(6)=   "lblCard(3)"
      Tab(10).Control(7)=   "lblCard(2)"
      Tab(10).Control(8)=   "lblCard(1)"
      Tab(10).Control(9)=   "lblCard(0)"
      Tab(10).Control(10)=   "Label2(11)"
      Tab(10).Control(11)=   "Label2(10)"
      Tab(10).Control(12)=   "Label2(9)"
      Tab(10).Control(13)=   "Label2(8)"
      Tab(10).Control(14)=   "Label2(7)"
      Tab(10).Control(15)=   "Label2(6)"
      Tab(10).Control(16)=   "Label2(5)"
      Tab(10).Control(17)=   "Label2(4)"
      Tab(10).Control(18)=   "Label2(3)"
      Tab(10).Control(19)=   "Label2(2)"
      Tab(10).Control(20)=   "Label2(1)"
      Tab(10).Control(21)=   "Label2(0)"
      Tab(10).Control(22)=   "Label2(12)"
      Tab(10).Control(23)=   "lblCard(10)"
      Tab(10).Control(24)=   "cmdRemoveFromDeck"
      Tab(10).Control(25)=   "lstDeck"
      Tab(10).ControlCount=   26
      Begin VB.CommandButton cmdAddHero 
         Caption         =   "Add Hero"
         Height          =   375
         Left            =   -66000
         TabIndex        =   144
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmdRandomLocation 
         Caption         =   "Random Location"
         Height          =   375
         Left            =   -66000
         TabIndex        =   143
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdFindLocs 
         Caption         =   "&Find Locations"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66000
         TabIndex        =   141
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtFindLoc 
         Height          =   855
         Left            =   -70800
         TabIndex        =   140
         Top             =   960
         Width           =   4575
      End
      Begin VB.CommandButton cmdShowAllLocs 
         Caption         =   "Show All Locations"
         Height          =   375
         Left            =   -66000
         TabIndex        =   138
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdUseBattlesite2 
         Caption         =   "Use as Battlesite"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66000
         TabIndex        =   137
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton cmdUseHomebase2 
         Caption         =   "Use as Homebase"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66000
         TabIndex        =   136
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Specials:"
         Height          =   2895
         Left            =   -70920
         TabIndex        =   133
         Top             =   2520
         Width           =   4695
         Begin VB.ListBox lstLocSpecials 
            Height          =   1230
            Left            =   120
            TabIndex        =   135
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtLocSpecEffect 
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   134
            Top             =   1680
            Width           =   4455
         End
      End
      Begin VB.ListBox lstLocCharacters 
         Height          =   1230
         Left            =   -74760
         TabIndex        =   131
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtLocEffect 
         Height          =   1455
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   130
         Top             =   3960
         Width           =   3735
      End
      Begin VB.ListBox lstFindLocations 
         Height          =   1815
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   129
         Top             =   720
         Width           =   3735
      End
      Begin VB.ListBox lstDeck 
         Height          =   4155
         Left            =   -74760
         TabIndex        =   104
         Top             =   840
         Width           =   5775
      End
      Begin VB.CommandButton cmdRemoveFromDeck 
         Caption         =   "Remove Card"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70200
         TabIndex        =   103
         Top             =   5160
         Width           =   1215
      End
      Begin VB.ListBox lstherostats2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   -66600
         TabIndex        =   102
         Top             =   720
         Width           =   2295
      End
      Begin VB.ListBox lstHeroStats 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   -67200
         TabIndex        =   101
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdFindEvent 
         Caption         =   "Find Event"
         Height          =   375
         Left            =   -71640
         TabIndex        =   100
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Quick Power Card Tool"
         Height          =   495
         Left            =   -67200
         TabIndex        =   99
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   -71040
         TabIndex        =   98
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdFind2 
         Caption         =   "Find"
         Height          =   375
         Left            =   -67800
         TabIndex        =   95
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   -68640
         TabIndex        =   94
         Text            =   "99"
         Top             =   600
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Specials:"
         Height          =   3735
         Left            =   -69840
         TabIndex        =   92
         Top             =   1680
         Width           =   5295
         Begin VB.TextBox Text1 
            Height          =   1575
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   97
            Top             =   1800
            Width           =   4935
         End
         Begin VB.ListBox lstFindSpec2 
            Height          =   1425
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.TextBox txtMinI 
         Height          =   285
         Left            =   -71160
         TabIndex        =   90
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtMinS 
         Height          =   285
         Left            =   -71160
         TabIndex        =   88
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtMinF 
         Height          =   285
         Left            =   -73320
         TabIndex        =   86
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtMinE 
         Height          =   285
         Left            =   -73320
         TabIndex        =   84
         Text            =   "0"
         Top             =   580
         Width           =   495
      End
      Begin VB.ListBox lstHeroMatches 
         Height          =   3180
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   82
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton cmdFindSpecialAdd 
         Caption         =   "&Add To Deck"
         Height          =   375
         Left            =   -70440
         TabIndex        =   81
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Go!"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69360
         TabIndex        =   80
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   -72600
         TabIndex        =   79
         Top             =   580
         Width           =   3135
      End
      Begin VB.ListBox lstFindSpecials 
         Height          =   3570
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   77
         Top             =   1440
         Width           =   5655
      End
      Begin VB.ListBox lstAllys 
         Height          =   1230
         Left            =   -70560
         Sorted          =   -1  'True
         TabIndex        =   74
         Top             =   3600
         Width           =   3855
      End
      Begin VB.CommandButton cmdAddAlly 
         Caption         =   "Add to Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67920
         TabIndex        =   73
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddDoubleShot 
         Caption         =   "Add to Deck"
         Height          =   375
         Left            =   -69360
         TabIndex        =   72
         Top             =   5040
         Width           =   1215
      End
      Begin VB.ListBox lstDoubleshot 
         Height          =   3960
         Left            =   -70920
         Sorted          =   -1  'True
         TabIndex        =   70
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton cmdAddAspect 
         Caption         =   "Add to Deck"
         Height          =   375
         Left            =   -72360
         TabIndex        =   69
         Top             =   5040
         Width           =   1215
      End
      Begin VB.ListBox lstAspects 
         Height          =   1425
         Left            =   -74760
         TabIndex        =   68
         Top             =   3480
         Width           =   3615
      End
      Begin VB.CommandButton cmdAddArtifact 
         Caption         =   "Add to Deck"
         Height          =   375
         Left            =   -72360
         TabIndex        =   66
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtArtifactEffect 
         Height          =   1575
         Left            =   -67920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   65
         Top             =   3960
         Width           =   3615
      End
      Begin VB.ListBox lstArtifacts 
         Height          =   1620
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   63
         Top             =   915
         Width           =   3615
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmDeckEditor.frx":0576
         Height          =   1935
         Left            =   240
         OleObjectBlob   =   "frmDeckEditor.frx":058A
         TabIndex        =   62
         Top             =   3360
         Width           =   10335
      End
      Begin MSComDlg.CommonDialog cmD1 
         Left            =   840
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open Overpower Deck"
         FileName        =   "*.dat"
         Filter          =   "Overpower Decks|*.dat"
      End
      Begin VB.CheckBox chkPlayTraining 
         Caption         =   "Show only playable cards"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   4980
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddTraining 
         Caption         =   "Add to Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72120
         TabIndex        =   61
         Top             =   4920
         Width           =   1215
      End
      Begin VB.ListBox lstTraining 
         Height          =   1620
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   59
         Top             =   3240
         Width           =   3855
      End
      Begin VB.CheckBox chkPlayTeamwork 
         Caption         =   "Show only playable cards"
         Height          =   255
         Left            =   -70560
         TabIndex        =   58
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddTeamwork 
         Caption         =   "Add to Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68040
         TabIndex        =   57
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ListBox lstTeamwork 
         Height          =   2010
         Left            =   -70560
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox chkPlayBasic 
         Caption         =   "Show only playable cards"
         Height          =   255
         Left            =   -74760
         TabIndex        =   54
         Top             =   2460
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddBasic 
         Caption         =   "Add to Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72120
         TabIndex        =   53
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ListBox lstBasicUniverse 
         Height          =   1620
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmdAddPower 
         Caption         =   "Add to Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72240
         TabIndex        =   50
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CheckBox chkPlayable 
         Caption         =   "Show only playable cards"
         Height          =   255
         Left            =   -74640
         TabIndex        =   49
         Top             =   5040
         Width           =   2295
      End
      Begin VB.ComboBox lstShowPower 
         Height          =   315
         ItemData        =   "frmDeckEditor.frx":15ED
         Left            =   -73080
         List            =   "frmDeckEditor.frx":1606
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox lstPowerCards 
         Height          =   3960
         Left            =   -74640
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   960
         Width           =   3615
      End
      Begin VB.CommandButton cmdAddEvent 
         Caption         =   "Add to Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70320
         TabIndex        =   43
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtEventEffect 
         Height          =   855
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   4080
         Width           =   5655
      End
      Begin VB.CommandButton cmdUseMission 
         Caption         =   "Use Mission"
         Height          =   375
         Left            =   -70320
         TabIndex        =   41
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ListBox lstEvents 
         Height          =   840
         Left            =   -74760
         TabIndex        =   40
         Top             =   3120
         Width           =   5655
      End
      Begin VB.ListBox lstMissions 
         Height          =   1620
         Left            =   -74760
         TabIndex        =   38
         Top             =   840
         Width           =   5655
      End
      Begin VB.CommandButton cmdAddSpecial 
         Caption         =   "&Add To Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69000
         TabIndex        =   34
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox txtSpecialEffect 
         Height          =   2295
         Left            =   -71520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2640
         Width           =   3855
      End
      Begin VB.ListBox lstSpecials 
         Height          =   1815
         Left            =   -71520
         TabIndex        =   32
         Top             =   720
         Width           =   3855
      End
      Begin VB.ListBox lstCharacters 
         Height          =   4545
         Left            =   -74760
         TabIndex        =   31
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton cmdAddToBattlesite 
         Caption         =   "Add to Battlesite Deck"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71160
         TabIndex        =   30
         Top             =   5040
         Width           =   1935
      End
      Begin VB.ListBox lstLocations 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   21
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtHomeBaseEffect 
         Height          =   1455
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   3960
         Width           =   3735
      End
      Begin VB.ListBox lstHomebaseCharacters 
         Height          =   1230
         Left            =   -74760
         TabIndex        =   19
         Top             =   2640
         Width           =   3735
      End
      Begin VB.CommandButton cmdHomebase 
         Caption         =   "Use as Homebase"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67680
         TabIndex        =   18
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdBattlesite 
         Caption         =   "Use as Battlesite"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70680
         TabIndex        =   17
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton cmdClearHomebase 
         Caption         =   "Clear Homebase"
         Height          =   375
         Left            =   -66000
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ListBox lstBSDeck 
         Height          =   1230
         Left            =   -70680
         TabIndex        =   15
         Top             =   3720
         Width           =   6255
      End
      Begin VB.CommandButton cmdClearBattlesite 
         Caption         =   "Clear Battlesite"
         Height          =   375
         Left            =   -69000
         TabIndex        =   14
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemoveBSCard 
         Caption         =   "Remove Card"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66000
         TabIndex        =   13
         Top             =   5040
         Width           =   1575
      End
      Begin VB.ComboBox lstShow 
         Height          =   315
         ItemData        =   "frmDeckEditor.frx":1678
         Left            =   4200
         List            =   "frmDeckEditor.frx":169A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2600
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7440
         TabIndex        =   6
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8520
         TabIndex        =   5
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdReserve 
         Caption         =   "Reserve"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.CheckBox chkEnforce 
         Caption         =   "Enforce Grid Limit of:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox lstGridLimit 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDeckEditor.frx":1719
         Left            =   2160
         List            =   "frmDeckEditor.frx":172C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2600
         Width           =   975
      End
      Begin VB.Label Label9 
         Height          =   375
         Left            =   -70800
         TabIndex        =   142
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label8 
         Caption         =   "Find a location with special(s) containing the following text:"
         Height          =   255
         Left            =   -70800
         TabIndex        =   139
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Locations:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   132
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   10
         Left            =   -65760
         TabIndex        =   128
         Tag             =   "Ally Card"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Universe: Ally"
         Height          =   255
         Index           =   12
         Left            =   -68640
         TabIndex        =   127
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Cards:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   126
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "By Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -68640
         TabIndex        =   125
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Specials:"
         Height          =   255
         Index           =   2
         Left            =   -68640
         TabIndex        =   124
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Power Cards:"
         Height          =   255
         Index           =   3
         Left            =   -68640
         TabIndex        =   123
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Universe: Basic"
         Height          =   255
         Index           =   4
         Left            =   -68640
         TabIndex        =   122
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Universe: Teamwork"
         Height          =   255
         Index           =   5
         Left            =   -68640
         TabIndex        =   121
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Universe: Training"
         Height          =   255
         Index           =   6
         Left            =   -68640
         TabIndex        =   120
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tactic: DoubleShot"
         Height          =   255
         Index           =   7
         Left            =   -68640
         TabIndex        =   119
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tactic: Artifact"
         Height          =   255
         Index           =   8
         Left            =   -68640
         TabIndex        =   118
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tactic: Aspect"
         Height          =   255
         Index           =   9
         Left            =   -68640
         TabIndex        =   117
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Activators:"
         Height          =   255
         Index           =   10
         Left            =   -68640
         TabIndex        =   116
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Events:"
         Height          =   255
         Index           =   11
         Left            =   -68640
         TabIndex        =   115
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   -65760
         TabIndex        =   114
         Tag             =   "Special Card"
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   -65760
         TabIndex        =   113
         Tag             =   "Power Card"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   -65760
         TabIndex        =   112
         Tag             =   "Basic Universe"
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   -65760
         TabIndex        =   111
         Tag             =   "Teamwork"
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   -65760
         TabIndex        =   110
         Tag             =   "Training"
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   5
         Left            =   -65760
         TabIndex        =   109
         Tag             =   "Artifact"
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   6
         Left            =   -65760
         TabIndex        =   108
         Tag             =   "Aspect Card"
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   7
         Left            =   -65760
         TabIndex        =   107
         Tag             =   "Double Shot"
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   8
         Left            =   -65760
         TabIndex        =   106
         Tag             =   "Activator"
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblCard 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   9
         Left            =   -65760
         TabIndex        =   105
         Tag             =   "Event"
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Maximum Grid:"
         Height          =   255
         Index           =   3
         Left            =   -70080
         TabIndex        =   93
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Matches:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   91
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Minimum Intellect:"
         Height          =   255
         Index           =   2
         Left            =   -72600
         TabIndex        =   89
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Minimum Strength:"
         Height          =   255
         Index           =   1
         Left            =   -72600
         TabIndex        =   87
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Minimum Fighting:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   85
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Minimum Energy:"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   83
         Top             =   600
         Width           =   1335
      End
      Begin VB.Image imgFindSpecial 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4440
         Left            =   -67800
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmDeckEditor.frx":1744
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3180
      End
      Begin VB.Label lblMatches 
         Caption         =   "Matches (0):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   78
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Find Specials containing text:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   76
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Ally:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -70560
         TabIndex        =   75
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Doubleshot:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -70920
         TabIndex        =   71
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Aspects:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   67
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Image imgArtifact 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3210
         Left            =   -67920
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Tag             =   "Discard"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Artifacts:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   64
         Top             =   600
         Width           =   1215
      End
      Begin VB.Image imgUniverse 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3210
         Left            =   -66600
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Tag             =   "Discard"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Training Cards:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   60
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Teamwork:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -70560
         TabIndex        =   56
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Basic Universe:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   52
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image imgPowerCard 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4440
         Left            =   -70680
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmDeckEditor.frx":CC1E
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "Power Cards:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74640
         TabIndex        =   47
         Top             =   640
         Width           =   1215
      End
      Begin VB.Label lblMission 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -67080
         TabIndex        =   45
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblCurrentMission 
         Caption         =   "Current Mission:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -68640
         TabIndex        =   44
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgMissionCard 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3045
         Left            =   -68640
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmDeckEditor.frx":180F8
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   4245
      End
      Begin VB.Label Label1 
         Caption         =   "Events:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   39
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Missions:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   37
         Top             =   600
         Width           =   2535
      End
      Begin VB.Image imgCardDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4440
         Left            =   -67440
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmDeckEditor.frx":608CF
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "Specials:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -71520
         TabIndex        =   36
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Characters:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   35
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Locations:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Homebase:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -70680
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblHomeBase 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69480
         TabIndex        =   27
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblHBCharacters 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -70680
         TabIndex        =   26
         Top             =   840
         Width           =   6255
      End
      Begin VB.Label lblHBEffect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -70680
         TabIndex        =   25
         Top             =   1560
         Width           =   6255
      End
      Begin VB.Label lblBattlesite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69480
         TabIndex        =   24
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Battlesite:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -70680
         TabIndex        =   23
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   -70680
         X2              =   -64320
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label lblBSDeckCount 
         Caption         =   "Battlesite Deck (0):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70680
         TabIndex        =   22
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Show:"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblHeroes 
         Caption         =   "Current Heroes:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   520
         Width           =   6375
      End
      Begin VB.Label lblHeroes 
         Caption         =   "Available Heroes:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   6135
      End
      Begin VB.Image imgHero 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1650
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2370
      End
      Begin VB.Image imgHero 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1650
         Index           =   1
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2370
      End
      Begin VB.Image imgHero 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1650
         Index           =   2
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2370
      End
      Begin VB.Image imgHero 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1650
         Index           =   3
         Left            =   8160
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2370
      End
      Begin VB.Shape shpReserve 
         BorderColor     =   &H00008000&
         BorderWidth     =   4
         Height          =   1695
         Left            =   240
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.ListBox lstTemp 
      Height          =   2595
      Left            =   3000
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgStore 
      Height          =   1335
      Index           =   0
      Left            =   360
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Deck"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Deck"
      End
      Begin VB.Menu mnuFileCap 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save Deck"
      End
      Begin VB.Menu mnuFileSaveDeckAs 
         Caption         =   "Save Deck As..."
      End
      Begin VB.Menu mnuFileCap2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveBattlesiteDeck 
         Caption         =   "Save Battlesite Deck"
      End
      Begin VB.Menu mnuLoadBattlesiteDeck 
         Caption         =   "Load Battlesite Deck"
      End
      Begin VB.Menu mnucap2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuAddBeyond 
         Caption         =   "Add Beyonder Activator"
      End
      Begin VB.Menu mnuToolsTestDraws 
         Caption         =   "Test Draws"
      End
      Begin VB.Menu mnucap4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRandomCharacterChallenge 
         Caption         =   "Random Character Challenge"
      End
   End
End
Attribute VB_Name = "frmDeckEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cdeckheroes As Collection
Dim nBattlesite As Integer
Dim nHomebase As Integer

Dim nReserve As Integer
Dim nMission As Integer

Dim sdeckname As String


Dim ncurimage As Integer

Dim strGridSQL As String
Dim strGridValSQL As String
Dim strCharSQL As String
Private Sub chkEnforce_Click()
If chkEnforce.Value = 1 Then
    lstGridLimit.Enabled = True
    EnforceGridLimit
    
Else
    lstGridLimit.Enabled = False
    Data1.RecordSource = strCharSQL
    Data1.Refresh
    lblHeroes(1).Caption = "Available Heroes:"
End If

End Sub
Private Sub EnforceGridLimit()
Dim myh As clsHero

ntot = 0

For i = 1 To cdeckheroes.Count
Set myh = New clsHero
myh.Load cdeckheroes.Item(i)
ntot = ntot + myh.Energy + myh.Fighting + myh.Strength + myh.Intellect
Next i

eval = lstGridLimit.List(lstGridLimit.ListIndex)

lblHeroes(1).Caption = "Available Heroes (" & Trim(Str(eval - ntot)) & " and under)"

Data1.RecordSource = ReplaceAllInString(strGridSQL, "[NUM]", Trim(Str((eval - ntot) + 1)))
Data1.Refresh

End Sub

Private Sub chkPlayable_Click()
Dim ctemp As Collection

If chkPlayable.Value = 1 Then
    lstShowPower.ListIndex = 0
    lstShowPower.Enabled = False
        
    ShowPlayablePowerCards
Else
    lstShowPower.Enabled = True
End If

End Sub
Private Sub ShowPlayablePowerCards()
Dim myh As clsHero

ne = 0
nf = 0
ns = 0
ni = 0

    Set myPower = New clsPowerCard
    
    For i = 1 To cdeckheroes.Count
        Set myh = New clsHero
        myh.Load cdeckheroes.Item(i)
    
        If myh.Energy > ne Then ne = myh.Energy
        If myh.Fighting > nf Then nf = myh.Fighting
        If myh.Strength > ns Then ns = myh.Strength
        If myh.Intellect > ni Then ni = myh.Intellect
        
    Next i
    
    Set myh = Nothing
    
    Set ctemp = myPower.GetPlayablePowerCards(ne, nf, ns, ni)
    
    lstPowerCards.Clear
    
    For i = 1 To ctemp.Count
        Set myPower = New clsPowerCard
        myPower.Load ctemp.Item(i)
        lstPowerCards.AddItem myPower.Title
        lstPowerCards.ItemData(lstPowerCards.NewIndex) = myPower.ID
    
    Next i

End Sub

Private Sub chkPlayBasic_Click()
If chkPlayBasic.Value = 1 Then
    ShowPlayableBasicUniverseCards
Else
    ShowAllBasicUniverse
End If

End Sub

Private Sub chkPlayTeamwork_Click()
If chkPlayTeamwork.Value = 1 Then
    ShowPlayableTeamworkCards
Else
    ShowAllTeamwork
End If

End Sub

Private Sub chkPlayTraining_Click()

If chkPlayTraining.Value = 1 Then
    ShowPlayableTrainingCards
Else
    ShowAllTrainingCards
End If

End Sub

Private Sub cmdAdd_Click()
Dim myh As clsHero

If cdeckheroes.Count = 4 Then
    MsgBox "You already have four heroes. Please select and delete a hero before adding a new one.", vbCritical, "Roster is Full."
    Exit Sub
End If

DBGrid1.Col = 0
a$ = DBGrid1.Text

If a$ = "" Then Exit Sub
If DBGrid1.Row < 0 Then Exit Sub

Set myh = New clsHero
nId = myh.GetIDFromName(a$)

If nId > 0 Then
cdeckheroes.Add nId
showdeckheroes
updateSpecialHeroes

DoEvents
Me.Refresh

If chkPlayTraining.Value = 1 Then
    ShowPlayableTrainingCards
Else
    ShowAllTrainingCards
End If

If chkPlayBasic.Value = 1 Then
    ShowPlayableBasicUniverseCards
Else
    ShowAllBasicUniverse
End If

If chkPlayTeamwork.Value = 1 Then
    ShowPlayableTeamworkCards
Else
    ShowAllTeamwork
End If

If chkPlayable.Value = 1 Then ShowPlayablePowerCards

If chkEnforce.Value = 1 Then EnforceGridLimit

End If


End Sub

Private Sub cmdAddAlly_Click()
If lstAllys.ListIndex = -1 Then Exit Sub

Set myAlly = New clsAlly
myAlly.Load lstAllys.ItemData(lstAllys.ListIndex)
cdeck.Add myAlly
Set myAlly = Nothing
ShowDeckCount
End Sub

Private Sub cmdAddArtifact_Click()
If lstArtifacts.ListIndex = -1 Then Exit Sub

Set myArtifact = New clsArtifact
myArtifact.Load lstArtifacts.ItemData(lstArtifacts.ListIndex)
cdeck.Add myArtifact

Set myArtifact = Nothing
ShowDeckCount
End Sub

Private Sub cmdAddAspect_Click()
If lstAspects.ListIndex = -1 Then Exit Sub

Set myAspect = New clsAspect
myAspect.Load lstAspects.ItemData(lstAspects.ListIndex)
cdeck.Add myAspect

Set myAspect = Nothing
ShowDeckCount
End Sub

Private Sub cmdAddBasic_Click()
If lstBasicUniverse.ListIndex = -1 Then Exit Sub

Set myBasic = New clsBasicUniverse
myBasic.Load lstBasicUniverse.ItemData(lstBasicUniverse.ListIndex)
cdeck.Add myBasic
Set myBasic = Nothing
ShowDeckCount
End Sub

Private Sub cmdAddDoubleShot_Click()
If lstDoubleshot.ListIndex = -1 Then Exit Sub

Set myDoubleShot = New clsDoubleShot
myDoubleShot.Load lstDoubleshot.ItemData(lstDoubleshot.ListIndex)
cdeck.Add myDoubleShot

Set myDoubleShot = Nothing
ShowDeckCount
End Sub

Private Sub cmdAddEvent_Click()
If lstEvents.ListIndex = -1 Then Exit Sub

Set myEvent = New clsEvent
myEvent.Load lstEvents.ItemData(lstEvents.ListIndex)
cdeck.Add myEvent
Set myEvent = Nothing
ShowDeckCount

End Sub

Private Sub cmdAddHero_Click()
If lstLocCharacters.ListIndex = -1 Then Exit Sub

a = lstLocCharacters.ItemData(lstLocCharacters.ListIndex)

If a < 1 Then Exit Sub

cdeckheroes.Add a
showdeckheroes
updateSpecialHeroes
End Sub

Private Sub cmdAddPower_Click()
If lstPowerCards.ListIndex = -1 Then Exit Sub

Set myPower = New clsPowerCard
myPower.Load lstPowerCards.ItemData(lstPowerCards.ListIndex)
cdeck.Add myPower
Set myPower = Nothing
ShowDeckCount
End Sub

Private Sub cmdAddSpecial_Click()
If lstSpecials.ListIndex = -1 Then Exit Sub

a = lstSpecials.ItemData(lstSpecials.ListIndex)
Set myspecial = New clsSpecial
myspecial.Load a

cdeck.Add myspecial

ShowDeckCount

End Sub

Private Sub cmdAddTeamwork_Click()
If lstTeamwork.ListIndex = -1 Then Exit Sub

Set myTeamwork = New clsTeamwork
myTeamwork.Load lstTeamwork.ItemData(lstTeamwork.ListIndex)
cdeck.Add myTeamwork
Set myTeamwork = Nothing
ShowDeckCount
End Sub

Private Sub cmdAddToBattlesite_Click()

sid = lstSpecials.ItemData(lstSpecials.ListIndex)
cid = lstCharacters.ItemData(lstCharacters.ListIndex)

Set myspecial = New clsSpecial
myspecial.Load sid

cbattlesitedeck.Add myspecial

'add an activator

If cid > 0 Then
    Set myActivator = New clsActivator
    myActivator.Load cid
    cdeck.Add myActivator
    Set myActivator = Nothing
End If

Set myspecial = Nothing

ShowDeckCount
ShowCurrentBattleSite

End Sub

Private Sub cmdAddTraining_Click()
If lstTraining.ListIndex = -1 Then Exit Sub

Set myTraining = New clsTraining
myTraining.Load lstTraining.ItemData(lstTraining.ListIndex)
cdeck.Add myTraining
Set myTraining = Nothing
ShowDeckCount
End Sub

Private Sub cmdBattlesite_Click()
nBattlesite = 0
ShowCurrentBattleSite

nBattlesite = lstLocations.ItemData(lstLocations.ListIndex)
ShowCurrentBattleSite
updateSpecialHeroes
End Sub

Private Sub cmdClearBattlesite_Click()
nBattlesite = 0
Set cbattlesitedeck = New Collection

ShowCurrentBattleSite
updateSpecialHeroes
End Sub

Private Sub cmdClearHomebase_Click()
nHomebase = 0
ShowCurrentHomeBase

End Sub
Private Sub cmdFind_Click()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
db.Open

Set objRS = New ADODB.Recordset

lstFindSpecials.Clear
cmdFind.Enabled = False

objRS.Open "SELECT * From Specials WHERE (((Specials.Effect) Like '%" & txtFind.Text & "%'));", db

Do Until objRS.EOF

lstFindSpecials.AddItem objRS.Fields("Character").Value & "-->" & objRS.Fields("Description").Value
lstFindSpecials.ItemData(lstFindSpecials.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close

lblMatches.Caption = "Matches (" & Trim(Str(lstFindSpecials.ListCount)) & ")"

cmdFind.Enabled = True

End Sub

Private Sub cmdFind2_Click()
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

strSQL = "SELECT * From Characters WHERE (((Val([Characters]![E]))>=[ENERGY]) AND ((Val([Characters]![F]))>=[FIGHTING]) AND ((Val([Characters]![S]))>=[STRENGTH]) AND ((Val([Characters]![I]))>=[INTELLECT]) AND ((Val([Characters]![E])+Val([Characters]![F])+Val([Characters]![S])+Val([Characters]![I]))<=[TOTAL]));"

strSQL = ReplaceAllInString(strSQL, "[ENERGY]", txtMinE.Text)
strSQL = ReplaceAllInString(strSQL, "[FIGHTING]", txtMinF.Text)
strSQL = ReplaceAllInString(strSQL, "[STRENGTH]", txtMinS.Text)
strSQL = ReplaceAllInString(strSQL, "[INTELLECT]", txtMinI.Text)
strSQL = ReplaceAllInString(strSQL, "[TOTAL]", txtTotal.Text)

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open strSQL, db

If dbRec.EOF = True Then
Label7.Caption = "Matches (" & Trim(Str(lstHeroMatches.ListCount)) & ")"
    dbRec.Close
    db.Close
    Exit Sub
End If

lstHeroMatches.Clear

Do Until dbRec.EOF
lstHeroMatches.AddItem dbRec.Fields("Character").Value & " [" & dbRec.Fields("E").Value & "/" & dbRec.Fields("F").Value & "/" & dbRec.Fields("S").Value & "/" & dbRec.Fields("I").Value & "]"
lstHeroMatches.ItemData(lstHeroMatches.NewIndex) = dbRec.Fields("id").Value
dbRec.MoveNext

Loop

dbRec.Close
db.Close


Label7.Caption = "Matches (" & Trim(Str(lstHeroMatches.ListCount)) & ")"

End Sub

Private Sub cmdFindEvent_Click()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

a$ = InputBox$("Find Events containing the text:", "Find Events", "")
If a$ = "" Then Exit Sub

Set db = New ADODB.Connection
db.ConnectionString = dbName
db.Open

Set objRS = New ADODB.Recordset

lstFindSpecials.Clear
cmdFind.Enabled = False

objRS.Open "SELECT * From Events WHERE (((Events.Effect) Like '%" & a$ & "%'));", db

If objRS.EOF = True Then
    x = MsgBox("No events found with text " & Chr(34) & a$ & Chr(34), vbCritical, "No Events Found")
    objRS.Close
    db.Close
    Exit Sub
End If

lstEvents.Clear

Do Until objRS.EOF

lstEvents.AddItem objRS.Fields("Name").Value & " [" & objRS.Fields("Mission").Value & "]"
lstEvents.ItemData(lstEvents.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close


End Sub

Private Sub cmdFindLocs_Click()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset
Dim objRS2 As ADODB.Recordset
Dim ctemp As Collection
Dim myh As clsHero

On Error Resume Next

Set db = New ADODB.Connection
db.ConnectionString = dbName
db.Open

Label9.Caption = "Searching..."

lstFindLocations.Clear

Set objRS = New ADODB.Recordset
objRS.Open "SELECT * From Homebases;", db

Do Until objRS.EOF

'Get Characters from homebase
Set myHomebase = New clsHomebase
myHomebase.Load objRS.Fields("ID").Value

Set ctemp = New Collection

a$ = myHomebase.Characters & ","

If a$ = "ANY TOURNAMENT LEGAL TEAM, USING NORMAL DECK-BUILDING RULES.," Or a$ = "ANY CHARACTERS," Or a$ = "" Or a$ = "ANY TOURNAMENT LEGAL TEAM USING NORMAL DECKBUILDING RULES," Then GoTo skipit

x = InStr(a$, ",")
While x > 0

b$ = Trim(Left(a$, x - 1))

Set myh = New clsHero
nId = myh.GetIDFromName(b$)
Set myh = Nothing

If nId > 0 Then ctemp.Add nId

a$ = Right$(a$, Len(a$) - x)
x = InStr(a$, ",")

Wend


'Have character ids.  Put together query
strSQL = "SELECT * FROM SPECIALS WHERE ("

For i = 1 To ctemp.Count

strSQL = strSQL & "(Specials.CharID=" & Trim(Str(ctemp.Item(i))) & ") OR "

Next i

strSQL = Left(strSQL, Len(strSQL) - 4) & ") AND ((Specials.Effect) Like ""%"
strSQL = strSQL & Trim(txtFindLoc.Text) & "%" & Chr(34) & ");"

Set objRS2 = New ADODB.Recordset

objRS2.Open strSQL, db

If objRS2.EOF = True Then
    'not found
Else
    Counter = 0
    
    Do Until objRS2.EOF
    Counter = Counter + 1
    objRS2.MoveNext
    Loop
    
    lstFindLocations.AddItem objRS.Fields("NAME").Value & " (" & Trim(Str(Counter)) & ")"
    
    lstFindLocations.ItemData(lstFindLocations.NewIndex) = objRS.Fields("ID").Value
End If

objRS2.Close

skipit:
objRS.MoveNext

lstFindLocations.Refresh
Label1(15).Caption = "Locations (" & Trim(Str(lstFindLocations.ListCount)) & ")"
Label1(15).Refresh

Loop

objRS.Close
Set objRS = Nothing

db.Close
Set db = Nothing

Set objRS2 = Nothing

If lstFindLocations.ListCount > 0 Then lstFindLocations.ListIndex = 0
Label9.Caption = "Search Complete"

End Sub

Private Sub cmdFindSpecialAdd_Click()
If lstFindSpecials.ListIndex = -1 Then Exit Sub


a = lstFindSpecials.ItemData(lstFindSpecials.ListIndex)
Set myspecial = New clsSpecial
myspecial.Load a

cdeck.Add myspecial

ShowDeckCount

End Sub

Private Sub cmdHomebase_Click()
nHomebase = 0
ShowCurrentHomeBase

nHomebase = lstLocations.ItemData(lstLocations.ListIndex)
ShowCurrentHomeBase
End Sub
Private Sub ShowCurrentBattleSite()
Dim ccard

If nBattlesite = 0 Then
    lblBattlesite.Caption = ""
    lstBSDeck.Clear
    lblBSDeckCount.Caption = "Battlesite Deck (0):"
    StatusBar1.Panels(4).Text = "Battlesite: None"
    Exit Sub
End If

lstBSDeck.Clear

Set myBattleSite = New clsBattlesite
myBattleSite.Load nBattlesite

lblBattlesite.Caption = myBattleSite.Name

'load battlesite deck
For i = 1 To cbattlesitedeck.Count
    Set ccard = cbattlesitedeck.Item(i)
    lstBSDeck.AddItem ccard.Title
Next i

StatusBar1.Panels(4).Text = "Battlesite: " & myBattleSite.Name & " (" & cbattlesitedeck.Count & ")"
lblBSDeckCount.Caption = "Battlesite Deck (" & cbattlesitedeck.Count & "):"

Set myBattleSite = Nothing



End Sub
Private Sub ShowCurrentHomeBase()

If nHomebase = 0 Then
    lblHomeBase.Caption = "None"
    lblHBEffect.Caption = ""
    lblHBCharacters.Caption = ""
    StatusBar1.Panels(3).Text = "Homebase: None"

    Exit Sub
End If

Set myHomebase = New clsHomebase
myHomebase.Load nHomebase

lblHomeBase.Caption = myHomebase.Name
lblHBCharacters.Caption = myHomebase.Characters
lblHBEffect.Caption = myHomebase.Effect
StatusBar1.Panels(3).Text = "Homebase: " & myHomebase.Name
Set myHomebase = Nothing

End Sub

Private Sub cmdRandomLocation_Click()
lstFindLocations.Clear
lstLocations.Clear
LoadLocations

Randomize

a = Int(Rnd * lstFindLocations.ListCount) + 1
lstFindLocations.ListIndex = a - 1

End Sub

Private Sub cmdRemove_Click()

If ncurimage < 0 Then Exit Sub

cdeckheroes.Remove ncurimage + 1
showdeckheroes

ShowPlayableTrainingCards
ShowPlayableBasicUniverseCards
ShowPlayableTeamworkCards

If chkPlayable.Value = 1 Then ShowPlayablePowerCards

If chkEnforce.Value = 1 Then EnforceGridLimit

updateSpecialHeroes

End Sub

Private Sub cmdRemoveBSCard_Click()
If lstBSDeck.ListIndex = -1 Then Exit Sub

cbattlesitedeck.Remove lstBSDeck.ListIndex + 1
ShowCurrentBattleSite

End Sub

Private Sub cmdRemoveFromDeck_Click()
a = lstDeck.ListIndex
If a = -1 Then Exit Sub

cdeck.Remove a + 1


ShowDeckCount
If lstDeck.ListCount = 0 Then Exit Sub

If a <= (lstDeck.ListCount - 1) Then
    lstDeck.ListIndex = a
Else
    If lstDeck.ListCount > 0 Then lstDeck.ListIndex = 0
End If

End Sub

Private Sub cmdReserve_Click()
If ncurimage < 0 Then Exit Sub

nReserve = ncurimage + 1
showdeckheroes

End Sub

Private Sub cmdShowAllLocs_Click()

lstFindLocations.Clear
lstLocations.Clear
LoadLocations

End Sub

Private Sub cmdUseBattlesite2_Click()
nBattlesite = 0
ShowCurrentBattleSite

nBattlesite = lstFindLocations.ItemData(lstFindLocations.ListIndex)
ShowCurrentBattleSite
updateSpecialHeroes
End Sub

Private Sub cmdUseHomebase2_Click()
nHomebase = 0
ShowCurrentHomeBase

nHomebase = lstFindLocations.ItemData(lstFindLocations.ListIndex)
ShowCurrentHomeBase
End Sub

Private Sub cmdUseMission_Click()
If lstMissions.ListIndex = -1 Then Exit Sub
nMission = lstMissions.ItemData(lstMissions.ListIndex)
ShowCurrentMission

End Sub

Private Sub Command1_Click()
Dim myh As clsHero

If lstHeroMatches.ListIndex = -1 Then Exit Sub

nId = lstHeroMatches.ItemData(lstHeroMatches.ListIndex)

If cdeckheroes.Count = 4 Then
    MsgBox "You already have four heroes. Please select and delete a hero before adding a new one.", vbCritical, "Roster is Full."
    Exit Sub
End If

If nId > 0 Then
cdeckheroes.Add nId
showdeckheroes
updateSpecialHeroes

DoEvents
Me.Refresh

If chkPlayTraining.Value = 1 Then
    ShowPlayableTrainingCards
Else
    ShowAllTrainingCards
End If

If chkPlayBasic.Value = 1 Then
    ShowPlayableBasicUniverseCards
Else
    ShowAllBasicUniverse
End If

If chkPlayTeamwork.Value = 1 Then
    ShowPlayableTeamworkCards
Else
    ShowAllTeamwork
End If

If chkPlayable.Value = 1 Then ShowPlayablePowerCards

If chkEnforce.Value = 1 Then EnforceGridLimit

End If


End Sub

Private Sub Command2_Click()

Load frmQuickPC

With frmQuickPC

.imgHero(0).Picture = imgHero(0).Picture
.imgHero(1).Picture = imgHero(1).Picture
.imgHero(2).Picture = imgHero(2).Picture
.imgHero(3).Picture = imgHero(3).Picture

.Show 1

'Energy
For i = 0 To 7

z = Val(.txtEnergy(i).Text)
If z > 0 Then
    tg = Val(.txtEnergy(i).Tag)
    Set myPower = New clsPowerCard
    myPower.Load tg
    For n = 1 To z
        cdeck.Add myPower
    Next n
End If

Next i

'Fighting
For i = 0 To 7

z = Val(.txtFighting(i).Text)
If z > 0 Then
    tg = Val(.txtFighting(i).Tag)
    Set myPower = New clsPowerCard
    myPower.Load tg
    For n = 1 To z
        cdeck.Add myPower
    Next n
End If

Next i

'Strength
For i = 0 To 7

z = Val(.txtStrength(i).Text)
If z > 0 Then
    tg = Val(.txtStrength(i).Tag)
    Set myPower = New clsPowerCard
    myPower.Load tg
    For n = 1 To z
        cdeck.Add myPower
    Next n
End If

Next i

'Intellect
For i = 0 To 7

z = Val(.txtIntellect(i).Text)
If z > 0 Then
    tg = Val(.txtIntellect(i).Tag)
    Set myPower = New clsPowerCard
    myPower.Load tg
    For n = 1 To z
        cdeck.Add myPower
    Next n
End If

Next i


For i = 0 To 7

z = Val(.txtMulti(i).Text)
If z > 0 Then
    tg = Val(.txtMulti(i).Tag)
    Set myPower = New clsPowerCard
    myPower.Load tg
    For n = 1 To z
        cdeck.Add myPower
    Next n
End If

Next i


End With

Unload frmQuickPC
ShowDeckCount
Set myPower = Nothing

End Sub

Private Sub DBGrid1_Click()

ShowHeroInfo



End Sub
Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
Dim strSQL As String

Select Case ColIndex
Case 0
strSQL = "SELECT * FROM Characters ORDER BY Characters.Character;"
Case 1
strSQL = "SELECT * FROM Characters ORDER BY Characters.E DESC, Characters.Character;"
Case 2
strSQL = "SELECT * FROM Characters ORDER BY Characters.F DESC, Characters.Character;"
Case 3
strSQL = "SELECT * FROM Characters ORDER BY Characters.S DESC, Characters.Character;"
Case 4
strSQL = "SELECT * FROM Characters ORDER BY Characters.I DESC, Characters.Character;"

Case Else
End Select

Data1.RecordSource = strSQL
Data1.Refresh

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
ShowHeroInfo

End Sub

Private Sub DBGrid1_SelChange(Cancel As Integer)

ShowHeroInfo

End Sub
Private Sub ShowHeroInfo()
tot = 0

With DBGrid1
.Col = 0
sname = .Text

.Col = 1
tot = tot + Val(.Text)
.Col = 2
tot = tot + Val(.Text)
.Col = 3
tot = tot + Val(.Text)
.Col = 4
tot = tot + Val(.Text)

lblHeroes(1).Caption = sname & ": (" & Trim(Str(tot)) & ")"


End With

If cdeckheroes.Count < 4 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
End If

cmdRemove.Enabled = False
cmdReserve.Enabled = False

End Sub
Private Sub Form_Load()
dbName = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Overpower.mdb"
sBlankImagePath = App.Path & "\NotFound.jpg"

strGridSQL = "SELECT Characters.ID, Characters.Character, Characters.Inherent, Characters.E, Characters.F, Characters.S, Characters.I, Val([Characters]![E])+Val([Characters]![F])+Val([Characters]![S])+Val([Characters]![I]) AS Grid From Characters WHERE (((Val([Characters]![E])+Val([Characters]![F])+Val([Characters]![S])+Val([Characters]![I]))<[NUM]));"
strCharSQL = "SELECT * From Characters ORDER By Characters.Character;"
strGridValSQL = "SELECT Characters.ID, Characters.Character, Characters.Inherent, Characters.E, Characters.F, Characters.S, Characters.I, Val([Characters]![E])+Val([Characters]![F])+Val([Characters]![S])+Val([Characters]![I]) AS Grid From Characters WHERE (((Val([Characters]![E])+Val([Characters]![F])+Val([Characters]![S])+Val([Characters]![I]))=[NUM]));"

Data1.DatabaseName = App.Path & "\Overpower.mdb"
Data1.RecordSource = strCharSQL
Data1.Refresh

Me.Refresh

tot = 0

With DBGrid1
.Col = 0
sname = .Text

.Col = 1
tot = tot + Val(.Text)
.Col = 2
tot = tot + Val(.Text)
.Col = 3
tot = tot + Val(.Text)
.Col = 4
tot = tot + Val(.Text)

lblHeroes(1).Caption = sname & ": (" & Trim(Str(tot)) & ")"

End With

DoEvents

SplashCaption "Loading Hero list..."

LoadHeroList
DoEvents

SplashCaption "Loading Locations..."

LoadLocations
DoEvents

SplashCaption "Loading Missions..."
LoadMissions
DoEvents

SplashCaption "Loading cards--Artifacts..."
LoadArtifacts
DoEvents

SplashCaption "Loading cards--Aspects..."
LoadAspects
DoEvents

SplashCaption "Loading cards--Doubleshots..."
LoadDoubleShots
DoEvents

LoadAllys
DoEvents
SplashCaption "Loading cards--Ally..."

newdeck
lstGridLimit.ListIndex = 4
lstShow.ListIndex = 0
lstShowPower.ListIndex = 0

SplashCaption "Loading cards--Specials..."
updateSpecialHeroes

Unload frmSplash
Me.Show

End Sub
Private Sub SplashCaption(sCap)
frmSplash.Label1(1).Caption = sCap

frmSplash.Label1(1).Refresh
End Sub
Private Sub ShowPlayableTrainingCards()
Dim ctemp As Collection
Dim myh As clsHero

Set myTraining = New clsTraining
Set ctemp = New Collection

If chkPlayTraining.Value = 0 Or cdeckheroes.Count = 0 Then

    Set ctemp = myTraining.GetPlayableTrainingCards(0, 0, 0, 0)

Else

    nlowe = 10
    nlowf = 10
    nlows = 10
    nlowi = 10
    
    For i = 1 To cdeckheroes.Count
        Set myh = New clsHero
        myh.Load cdeckheroes.Item(i)
        
        If myh.Energy < nlowe Then nlowe = myh.Energy
        If myh.Fighting < nlowf Then nlowf = myh.Fighting
        If myh.Strength < nlows Then nlows = myh.Strength
        If myh.Intellect < nlowi Then nlowi = myh.Intellect
    Next i
    
    Set myh = Nothing
    
    Set ctemp = myTraining.GetPlayableTrainingCards(nlowe, nlowf, nlows, nlowi)
    

End If

lstTraining.Clear

For i = 1 To ctemp.Count
    Set myTraining = New clsTraining
    myTraining.Load ctemp.Item(i)
    lstTraining.AddItem myTraining.Title
    lstTraining.ItemData(lstTraining.NewIndex) = myTraining.ID
Next i

Set myTraining = Nothing
Set ctemp = Nothing
Set myh = Nothing

End Sub
Private Sub ShowPlayableBasicUniverseCards()
Dim ctemp As Collection
Dim myh As clsHero

Set myBasic = New clsBasicUniverse
Set ctemp = New Collection

If chkPlayBasic.Value = 0 Or cdeckheroes.Count = 0 Then

    Set ctemp = myBasic.GetPlayableBasicUniverseCards(10, 10, 10, 10)
Else

    nhe = 0
    nhf = 0
    nhs = 0
    nhi = 0
    
    For i = 1 To cdeckheroes.Count
        Set myh = New clsHero
        myh.Load cdeckheroes.Item(i)
        
        If myh.Energy > nhe Then nhe = myh.Energy
        If myh.Fighting > nhf Then nhf = myh.Fighting
        If myh.Strength > nhs Then nhs = myh.Strength
        If myh.Intellect > nhi Then nhi = myh.Intellect
    Next i
    
    Set myh = Nothing

    
    Set ctemp = myBasic.GetPlayableBasicUniverseCards(nhe, nhf, nhs, nhi)
    

End If

lstBasicUniverse.Clear

For i = 1 To ctemp.Count
    Set myBasic = New clsBasicUniverse
    myBasic.Load ctemp.Item(i)
    lstBasicUniverse.AddItem myBasic.Title
    lstBasicUniverse.ItemData(lstBasicUniverse.NewIndex) = myBasic.ID
Next i

Set myBasic = Nothing
Set ctemp = Nothing
Set myh = Nothing

End Sub
Private Sub newdeck()
Set cdeckheroes = New Collection
Set cdeck = New Collection

For i = 0 To 3
    Set imgHero(i).Picture = Nothing
Next i

sdeckname = "New Deck"
Me.Caption = "Overpower Deck Editor: New Deck"

ShowDeckCount

nReserve = 0
nBattlesite = 0
nHomebase = 0

Set cbattlesitedeck = New Collection

chkPlayTeamwork.Value = 0
chkPlayBasic.Value = 0
chkPlayTraining.Value = 0
chkPlayable.Value = 0

ShowCurrentHomeBase
ShowCurrentBattleSite

DoEvents

SplashCaption "Loading cards--Training..."
ShowAllTrainingCards

SplashCaption "Loading cards--Basic Universe..."
ShowAllBasicUniverse
SplashCaption "Loading cards--Teamwork..."
ShowAllTeamwork

StatusBar1.Panels(1).Text = "Grid Count (0)"

nMission = 1
ShowCurrentMission
End Sub
Private Sub ShowCurrentMission()

Set myMission = New clsMission
myMission.Load nMission
lblMission.Caption = myMission.Name

End Sub
Private Sub ShowDeckCount()
Dim ccard

StatusBar1.Panels(2).Text = "Deck = " & cdeck.Count

For i = 0 To 10
    lblCard(i).Caption = "0"
Next i

lstDeck.Clear
For i = 1 To cdeck.Count

Set ccard = cdeck.Item(i)
lstDeck.AddItem ccard.Title

For k = 0 To 10
    If lblCard(k).Tag = ccard.CardType Then
        
        c = Val(lblCard(k).Caption) + 1
        lblCard(k).Caption = Trim(Str(c))
    End If
Next k

Next i

If lstDeck.ListCount = 0 Then
    cmdRemoveFromDeck.Enabled = False
Else
    lstDeck.ListIndex = 0
End If

End Sub
Private Sub showdeckheroes()
Dim myh As clsHero

nval = 0

StatusBar1.Panels(1).Text = "Loading heroes..."

ncurimage = -1

For i = 0 To 3
Set imgHero(i).Picture = Nothing
imgHero(i).Tag = -1
Next i

shpReserve.Visible = False

lstHeroStats.Clear
lstHeroStats.AddItem "HERO STATS"
lstHeroStats.AddItem "================="

lstherostats2.Clear
lstherostats2.AddItem "HERO STATS"
lstherostats2.AddItem "================="

For i = 1 To cdeckheroes.Count
Set myh = New clsHero
myh.Load cdeckheroes.Item(i)

If FindImage(myh.ID) = -1 Then

    Load imgStore(imgStore.Count)
    
    If myh.LoadImage(cdeckheroes.Item(i)) = True Then
        imgStore(imgStore.Count - 1).Picture = LoadPicture(App.Path & "\temppic.jpg")
    Else
        imgStore(imgStore.Count - 1).Picture = LoadPicture(sBlankImagePath)
    End If
    
    imgHero(i - 1).Picture = imgStore(imgStore.Count - 1).Picture
    
Else

    imgHero(i - 1).Picture = imgStore(FindImage(myh.ID))

End If

imgHero(i - 1).Tag = myh.ID

nval = nval + myh.Energy + myh.Fighting + myh.Strength + myh.Intellect

lstHeroStats.AddItem "E" & Trim(Str(myh.Energy)) & "/F" & Trim(Str(myh.Fighting)) & "/S" & Trim(Str(myh.Strength)) & "/I" & Trim(Str(myh.Intellect))
lstherostats2.AddItem "E" & Trim(Str(myh.Energy)) & "/F" & Trim(Str(myh.Fighting)) & "/S" & Trim(Str(myh.Strength)) & "/I" & Trim(Str(myh.Intellect))

If i = nReserve Then
    shpReserve.Left = imgHero(i - 1).Left
    shpReserve.Visible = True
End If

Me.Refresh

Next i

StatusBar1.Panels(1).Text = ""

If cdeckheroes.Count = 4 Then
    z$ = ""
Else

    
    nval2 = 76 - nval
    nval3 = Format(nval2 / (4 - cdeckheroes.Count), "#.##")
    If cdeckheroes.Count = 3 Then
        z$ = " | Room: " & Trim(Str(nval2))
    Else
        z$ = " | Room: " & Trim(Str(nval2)) & " | Room Per: " & Trim(Str(nval3))
    End If
    
End If

lblHeroes(0).Caption = "Current Heroes (" & Trim(Str(nval)) & ")" & z$
StatusBar1.Panels(1).Text = "Grid Count (" & Trim(Str(nval)) & ")"

End Sub
Private Sub updateSpecialHeroes()
Dim myh As clsHero

lstCharacters.Clear

lstCharacters.AddItem "* ANY CHARACTER *"
lstCharacters.ItemData(lstCharacters.NewIndex) = -2
lstCharacters.AddItem "====================================="


For i = 1 To cdeckheroes.Count
Set myh = New clsHero
myh.Load cdeckheroes.Item(i)
lstCharacters.AddItem myh.Name
lstCharacters.ItemData(lstCharacters.NewIndex) = myh.ID
Next i

If cdeckheroes.Count > 0 Then lstCharacters.AddItem "====================================="

If nBattlesite > 0 Then

Set myBattleSite = New clsBattlesite
myBattleSite.Load nBattlesite

a$ = myBattleSite.Characters & ","

x = InStr(a$, ",")

While x > 0

c$ = Trim(Left(a$, x - 1))

Set myh = New clsHero
cid = myh.GetIDFromName(c$)
If cid > 0 And cid <> "#ERROR#" Then
    myh.Load cid
    lstCharacters.AddItem myh.Name
    lstCharacters.ItemData(lstCharacters.NewIndex) = myh.ID
End If

a$ = Right(a$, Len(a$) - x)
x = InStr(a$, ",")

Wend

lstCharacters.AddItem "====================================="
End If


lstCharacters.ItemData(lstCharacters.NewIndex) = 0

For i = 0 To lstTemp.ListCount - 1
lstCharacters.AddItem lstTemp.List(i)
lstCharacters.ItemData(lstCharacters.NewIndex) = lstTemp.ItemData(i)
Next i

End Sub


Private Sub imgHero_Click(Index As Integer)
If Val(imgHero(Index).Tag) < 1 Then
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdReserve.Enabled = False
Else
    cmdAdd.Enabled = False
    cmdRemove.Enabled = True
    cmdReserve.Enabled = True
    ncurimage = Index
End If

End Sub

Private Sub imgHero_DblClick(Index As Integer)
If Val(imgHero(Index).Tag) > 0 Then
    Load frmCardDetail
    frmCardDetail.imgCard.Picture = imgHero(Index).Picture
    frmCardDetail.Show 1

'    Load frmHeroCardDetail
'    frmHeroCardDetail.imgCard.Picture = imgHero(Index).Picture
'    frmHeroCardDetail.Show 1
End If
End Sub

Private Sub lstAllys_Click()
If lstAllys.ListIndex = -1 Then
    cmdAddAlly.Enabled = False
    Exit Sub
End If

Set myAlly = New clsAlly
myAlly.Load lstAllys.ItemData(lstAllys.ListIndex)

If myAlly.LoadImage(lstAllys.ItemData(lstAllys.ListIndex)) = True Then
    imgUniverse.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgUniverse.Picture = LoadPicture(sBlankImagePath)
End If

Set myAlly = Nothing
cmdAddAlly.Enabled = True
End Sub

Private Sub lstArtifacts_Click()
If lstArtifacts.ListIndex = -1 Then Exit Sub

Set myArtifact = New clsArtifact
nId = lstArtifacts.ItemData(lstArtifacts.ListIndex)
myArtifact.Load nId

If myArtifact.LoadImage(nId) = True Then
    imgArtifact.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgArtifact.Picture = LoadPicture(sBlankImagePath)
End If

imgArtifact.ToolTipText = myArtifact.Effect
txtArtifactEffect.Text = myArtifact.Effect

End Sub


Private Sub lstAspects_Click()
If lstAspects.ListIndex = -1 Then Exit Sub

Set myAspect = New clsAspect

nId = lstAspects.ItemData(lstAspects.ListIndex)
myAspect.Load nId

If myAspect.LoadImage(nId) = True Then
    imgArtifact.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgArtifact.Picture = LoadPicture(sBlankImagePath)
End If

imgArtifact.ToolTipText = myAspect.Effect
txtArtifactEffect.Text = myAspect.Effect
End Sub

Private Sub lstBasicUniverse_Click()
If lstBasicUniverse.ListIndex = -1 Then
    cmdAddBasic.Enabled = False
    Exit Sub
End If

Set myBasic = New clsBasicUniverse
myBasic.Load lstBasicUniverse.ItemData(lstBasicUniverse.ListIndex)

If myBasic.LoadImage(lstBasicUniverse.ItemData(lstBasicUniverse.ListIndex)) = True Then
    imgUniverse.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgUniverse.Picture = LoadPicture(sBlankImagePath)
End If

Set myBasic = Nothing
cmdAddBasic.Enabled = True
End Sub

Private Sub lstBSDeck_Click()
If lstBSDeck.ListIndex <> -1 Then
    cmdRemoveBSCard.Enabled = True
Else
    cmdRemoveBSCard.Enabled = False
End If

End Sub

Private Sub lstCharacters_Click()
Dim myh As clsHero

If lstCharacters.List(lstCharacters.ListIndex) = "BEYONDER" Then
    x = MsgBox("Would you like to add a Beyonder Activator to your deck?", vbYesNoCancel, "BEYONDER")
    If x <> 6 Then Exit Sub
        
    a$ = InputBox$("How many Beyonder Activators do you want to add?", "Add Beyonder Activators", "1")
    
    na = Val(a$)
    If na = 0 Then Exit Sub
    
    For i = 1 To na
    Set myActivator = New clsActivator
    myActivator.Load lstCharacters.ItemData(lstCharacters.ListIndex)
    cdeck.Add myActivator
    Next i
    
    ShowDeckCount
    
    Exit Sub
End If

a = lstCharacters.ItemData(lstCharacters.ListIndex)

If a = -1 Or a = 0 Then
    txtSpecialEffect.Text = ""
    lstSpecials.Clear
    Exit Sub
End If

If a = -2 Then
lstSpecials.Clear

LoadAnyCharacterSpecials

Else
Set myh = New clsHero
myh.Load a

lstSpecials.Clear

For i = 1 To myh.Special_Count
c$ = myh.Special_Name(i)
If myh.Special_OPD(i) = True Then
    c$ = c$ & " [OPD]"
End If

lstSpecials.AddItem c$
lstSpecials.ItemData(lstSpecials.NewIndex) = myh.Special_ID(i)
Next i
End If

If lstSpecials.ListCount > 0 Then
    lstSpecials.ListIndex = 0
    ShowSpecial
End If

End Sub
Private Sub LoadAnyCharacterSpecials()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

lstSpecials.Clear

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

objRS.Open "SELECT * FROM Specials WHERE Specials.CharID=0 ORDER BY Specials.Description;", db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If


Do Until objRS.EOF
nId = objRS.Fields("ID").Value

lstSpecials.AddItem objRS.Fields("Description").Value
lstSpecials.ItemData(lstSpecials.NewIndex) = nId

objRS.MoveNext
Loop


objRS.Close
db.Close


End Sub
Private Sub LoadHeroList()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

lstTemp.Clear

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

objRS.Open "SELECT * FROM Characters ORDER BY Characters.Character;", db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

Do Until objRS.EOF

lstTemp.AddItem objRS.Fields("Character").Value
lstTemp.ItemData(lstTemp.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close

End Sub
Private Sub LoadLocations()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

lstLocations.Clear

objRS.Open "SELECT * FROM Homebases ORDER BY Homebases.Name;", db

'If objRS.EOF = True Then
'    objRS.Close
'    db.Close
'    Exit Sub
'End If

objRS.MoveFirst

Do Until objRS.EOF

lstLocations.AddItem objRS.Fields("Name").Value
lstLocations.ItemData(lstLocations.NewIndex) = objRS.Fields("ID").Value

lstFindLocations.AddItem objRS.Fields("Name").Value
lstFindLocations.ItemData(lstFindLocations.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close

If lstLocations.ListCount > 0 Then
    lstLocations.ListIndex = 0
End If

If lstFindLocations.ListCount > 0 Then
    lstFindLocations.ListIndex = 0
End If

Label1(2).Caption = "Locations (" & Trim(Str(lstLocations.ListCount)) & ")"
Label1(15).Caption = "Locations (" & Trim(Str(lstFindLocations.ListCount)) & ")"

End Sub
Private Sub lstDeck_Click()
If lstDeck.ListIndex = -1 Then
    cmdRemoveFromDeck.Enabled = False
Else
    cmdRemoveFromDeck.Enabled = True
End If

End Sub

Private Sub lstDeck_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 And lstDeck.ListIndex > -1 Then cmdRemoveFromDeck = True

End Sub

Private Sub lstDoubleshot_Click()
If lstDoubleshot.ListIndex = -1 Then Exit Sub

Set myDoubleShot = New clsDoubleShot
nId = lstDoubleshot.ItemData(lstDoubleshot.ListIndex)
myDoubleShot.Load nId

If myDoubleShot.LoadImage(nId) = True Then
    imgArtifact.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgArtifact.Picture = LoadPicture(sBlankImagePath)
End If

imgArtifact.ToolTipText = myDoubleShot.Effect
txtArtifactEffect.Text = myDoubleShot.Effect

End Sub

Private Sub lstEvents_Click()
If lstEvents.ListIndex = -1 Then
    cmdAddEvent.Enabled = False
    Exit Sub
End If

Set myEvent = New clsEvent
a = lstEvents.ItemData(lstEvents.ListIndex)
myEvent.Load a

If myEvent.LoadImage(a) = True Then
    imgMissionCard.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgMissionCard.Picture = LoadPicture(sBlankImagePath)
End If

txtEventEffect.Text = myEvent.Description

cmdAddEvent.Enabled = True

Set myEvent = Nothing

End Sub

Private Sub lstFindLocations_Click()
Dim myh As clsHero

If lstFindLocations.ListIndex = -1 Then
    txtLocEffect.Text = ""
    lstLocCharacters.Clear
    cmdUseHomebase2.Enabled = False
    cmdUseBattlesite2.Enabled = False
    Exit Sub
Else
    cmdUseHomebase2.Enabled = True
    cmdUseBattlesite2.Enabled = True

End If

Set myHomebase = New clsHomebase
myHomebase.Load lstFindLocations.ItemData(lstFindLocations.ListIndex)
txtLocEffect.Text = myHomebase.Effect

a$ = myHomebase.Characters & ","

lstLocCharacters.Clear

x = InStr(a$, ",")
While x > 0

b$ = Trim(Left(a$, x - 1))
Set myh = New clsHero
nId = myh.GetIDFromName(b$)

myh.Load nId

If myh.Name = "" Then
    lstLocCharacters.AddItem b$
Else
    lstLocCharacters.AddItem myh.Name & " [E" & myh.Energy & "/F" & myh.Fighting & "/S" & myh.Strength & "/I" & myh.Intellect & "]"
End If

a$ = Right$(a$, Len(a$) - x)
x = InStr(a$, ",")


lstLocCharacters.ItemData(lstLocCharacters.NewIndex) = nId


Wend

If lstLocCharacters.ListCount > 0 Then lstLocCharacters.ListIndex = 0

Set myHomebase = Nothing
End Sub

Private Sub lstFindSpec2_Click()
With lstFindSpec2
 
If .ListIndex = -1 Then Exit Sub

a = .ItemData(.ListIndex)

Set myspecial = New clsSpecial
myspecial.Load a
Text1.Text = myspecial.Effect

End With

End Sub

Private Sub lstFindSpecials_Click()
If lstFindSpecials.ListIndex = -1 Then
    imgFindSpecial.Picture = LoadPicture(sBlankImagePath)
    imgFindSpecial.ToolTipText = ""
Else

    a = lstFindSpecials.ItemData(lstFindSpecials.ListIndex)

    Set myspecial = New clsSpecial
    myspecial.Load a
    
    If myspecial.LoadImage(a) = True Then
        imgFindSpecial.Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgFindSpecial.ToolTipText = myspecial.Effect
    Else
        imgFindSpecial.Picture = LoadPicture(sBlankImagePath)
        imgFindSpecial.ToolTipText = ""
    End If
    
End If
End Sub

Private Sub lstGridLimit_Click()
If chkEnforce.Value = 1 Then EnforceGridLimit

End Sub

Private Sub lstHeroMatches_Click()
Dim myh2 As clsHero

If lstHeroMatches.ListIndex = -1 Then Exit Sub

lstFindSpec2.Clear
Set myh2 = New clsHero

myh2.Load lstHeroMatches.ItemData(lstHeroMatches.ListIndex)

For i = 1 To myh2.Special_Count

If myh2.Special_OPD(i) = True Then
    lstFindSpec2.AddItem myh2.Special_Name(i) & " [OPD]"
Else
    lstFindSpec2.AddItem myh2.Special_Name(i)
End If

lstFindSpec2.ItemData(lstFindSpec2.NewIndex) = myh2.Special_ID(i)

Next i

Set myh2 = Nothing

If lstFindSpec2.ListCount > 0 Then lstFindSpec2.ListIndex = 0

End Sub

Private Sub lstLocations_Click()

ShowHomeBase

End Sub
Private Sub ShowHomeBase()
If lstLocations.ListIndex = -1 Then
    txtHomeBaseEffect.Text = ""
    lstHomebaseCharacters.Clear
    cmdHomebase.Enabled = False
    cmdBattlesite.Enabled = False
    Exit Sub
End If

Set myHomebase = New clsHomebase
myHomebase.Load lstLocations.ItemData(lstLocations.ListIndex)
txtHomeBaseEffect.Text = myHomebase.Effect

a$ = myHomebase.Characters & ","

lstHomebaseCharacters.Clear

x = InStr(a$, ",")
While x > 0

lstHomebaseCharacters.AddItem Trim(Left(a$, x - 1))
a$ = Right$(a$, Len(a$) - x)
x = InStr(a$, ",")

Wend

cmdHomebase.Enabled = True
cmdBattlesite.Enabled = True

Set myHomebase = Nothing

End Sub

Private Sub lstLocCharacters_Click()
Dim myh2 As clsHero

If lstLocCharacters.ListIndex = -1 Then Exit Sub

If lstLocCharacters.ItemData(lstLocCharacters.ListIndex) = 0 Then
    lstLocSpecials.Clear
    txtLocSpecEffect.Text = ""
    Exit Sub
End If

lstLocSpecials.Clear
Set myh2 = New clsHero

myh2.Load lstLocCharacters.ItemData(lstLocCharacters.ListIndex)

For i = 1 To myh2.Special_Count

If myh2.Special_OPD(i) = True Then
    lstLocSpecials.AddItem myh2.Special_Name(i) & " [OPD]"
Else
    lstLocSpecials.AddItem myh2.Special_Name(i)
End If

lstLocSpecials.ItemData(lstLocSpecials.NewIndex) = myh2.Special_ID(i)

Next i

Set myh2 = Nothing

If lstLocSpecials.ListCount > 0 Then lstLocSpecials.ListIndex = 0

End Sub

Private Sub lstLocSpecials_Click()
With lstLocSpecials
 
If .ListIndex = -1 Then Exit Sub

a = .ItemData(.ListIndex)

Set myspecial = New clsSpecial
myspecial.Load a
txtLocSpecEffect.Text = myspecial.Effect

End With
End Sub

Private Sub lstMissions_Click()
If lstMissions.ListIndex = -1 Then
    lstEvents.Clear
    
    Exit Sub
End If

a = lstMissions.ItemData(lstMissions.ListIndex)
Set myMission = New clsMission
myMission.Load a

If myMission.LoadImage(a) = True Then
    imgMissionCard.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgMissionCard.Picture = LoadPicture(sBlankImagePath)
End If

myMission.LoadEvents

lstEvents.Clear

For i = 1 To myMission.Events_Count
    lstEvents.AddItem myMission.Events_Name(i)
    lstEvents.ItemData(lstEvents.NewIndex) = myMission.Events_ID(i)
Next i

End Sub

Private Sub lstPowerCards_Click()
If lstPowerCards.ListIndex = -1 Then
    cmdAddPower.Enabled = False
    Exit Sub
End If

Set myPower = New clsPowerCard
a = lstPowerCards.ItemData(lstPowerCards.ListIndex)
myPower.Load a

If myPower.LoadImage(a) = True Then
    imgPowerCard.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgPowerCard.Picture = LoadPicture(sBlankImagePath)
End If

cmdAddPower.Enabled = True

Set myPower = Nothing
End Sub

Private Sub lstShow_Click()
Select Case lstShow.ListIndex
Case 0
    Data1.RecordSource = strCharSQL
Case 1
    Data1.RecordSource = ReplaceAllInString(strGridValSQL, "[NUM]", "16")
Case 2
    Data1.RecordSource = ReplaceAllInString(strGridValSQL, "[NUM]", "17")
Case 3
    Data1.RecordSource = ReplaceAllInString(strGridValSQL, "[NUM]", "18")
Case 4
    Data1.RecordSource = ReplaceAllInString(strGridValSQL, "[NUM]", "19")
Case 5
    Data1.RecordSource = "SELECT Characters.ID, Characters.Character, Characters.Inherent, Characters.E, Characters.F, Characters.S, Characters.I, Val([Characters]![E])+Val([Characters]![F])+Val([Characters]![S])+Val([Characters]![I]) AS Grid From Characters WHERE (((Val([Characters]![E])+Val([Characters]![F])+Val([Characters]![S])+Val([Characters]![I]))>=20));"
Case 6
    Data1.RecordSource = "SELECT Characters.ID, Characters.Character, Characters.Inherent, Characters.E, Characters.F, Characters.S, Characters.I From Characters WHERE (Val(Characters.E) >6);"
Case 7
    Data1.RecordSource = "SELECT Characters.ID, Characters.Character, Characters.Inherent, Characters.E, Characters.F, Characters.S, Characters.I From Characters WHERE (Val(Characters.F) >6);"
Case 8
    Data1.RecordSource = "SELECT Characters.ID, Characters.Character, Characters.Inherent, Characters.E, Characters.F, Characters.S, Characters.I From Characters WHERE (Val(Characters.S) >6);"
Case 9
    Data1.RecordSource = "SELECT Characters.ID, Characters.Character, Characters.Inherent, Characters.E, Characters.F, Characters.S, Characters.I From Characters WHERE (Val(Characters.I) >6);"

Case Else
End Select

Data1.Refresh
End Sub

Private Sub lstShowPower_Click()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

lstPowerCards.Clear

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

Select Case lstShowPower.ListIndex
Case 0
    strSQL = "SELECT * FROM Power;"
Case 1
    strSQL = "SELECT * FROM Power WHERE Power.E = True;"
Case 2
    strSQL = "SELECT * FROM Power WHERE Power.F = True;"
Case 3
    strSQL = "SELECT * FROM Power WHERE Power.S = True;"

Case 4
    strSQL = "SELECT * FROM Power WHERE Power.I = True;"

Case 5
    strSQL = "SELECT * FROM Power WHERE Power.A = True;"

Case 6
    strSQL = "SELECT * FROM Power WHERE Power.M = True;"


Case Else
End Select

objRS.Open strSQL, db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

Do Until objRS.EOF

Set myPower = New clsPowerCard
myPower.Load objRS.Fields("ID").Value
lstPowerCards.AddItem myPower.Title
lstPowerCards.ItemData(lstPowerCards.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close

Set myPower = Nothing

End Sub

Private Sub lstSpecials_Click()
If lstSpecials.ListIndex = -1 Then
    cmdAddSpecial.Enabled = False
    cmdAddToBattlesite.Enabled = False
    txtSpecialEffect.Text = ""
    imgCardDetail.Picture = LoadPicture(sBlankImagePath)
Else
    ShowSpecial
    cmdAddSpecial.Enabled = True
    
    If nBattlesite = 0 Then
        cmdAddToBattlesite.Enabled = False
    Else
        cmdAddToBattlesite.Enabled = True
    End If
    
End If

End Sub
Private Sub ShowSpecial()
a = lstSpecials.ItemData(lstSpecials.ListIndex)

Set myspecial = New clsSpecial
myspecial.Load a

txtSpecialEffect.Text = myspecial.Effect & " (" & myspecial.Code & ")"

If myspecial.LoadImage(a) = True Then
    imgCardDetail.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgCardDetail.Picture = LoadPicture(sBlankImagePath)
End If
End Sub

Private Sub lstTeamwork_Click()
If lstTeamwork.ListIndex = -1 Then
    cmdAddTeamwork.Enabled = False
    Exit Sub
End If

Set myTeamwork = New clsTeamwork
myTeamwork.Load lstTeamwork.ItemData(lstTeamwork.ListIndex)

If myTeamwork.LoadImage(lstTeamwork.ItemData(lstTeamwork.ListIndex)) = True Then
    imgUniverse.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgUniverse.Picture = LoadPicture(sBlankImagePath)
End If

Set myTeamwork = Nothing
cmdAddTeamwork.Enabled = True
End Sub

Private Sub lstTraining_Click()
If lstTraining.ListIndex = -1 Then
    cmdAddTraining.Enabled = False
    Exit Sub
End If

Set myTraining = New clsTraining
myTraining.Load lstTraining.ItemData(lstTraining.ListIndex)

If myTraining.LoadImage(lstTraining.ItemData(lstTraining.ListIndex)) = True Then
    imgUniverse.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgUniverse.Picture = LoadPicture(sBlankImagePath)
End If

Set myTraining = Nothing
cmdAddTraining.Enabled = True

End Sub

Private Sub mnuAddBeyond_Click()
    Set myActivator = New clsActivator
    myActivator.Load 18
    cdeck.Add myActivator
    ShowDeckCount
End Sub

Private Sub mnuFile_Click()
If cbattlesitedeck.Count = 0 Then
    mnuSaveBattlesiteDeck.Enabled = False
Else
    mnuSaveBattlesiteDeck.Enabled = True
End If

End Sub

Private Sub mnuFileExit_Click()
End

End Sub
Private Function FindImage(HeroTag)
a = -1

For i = 1 To imgStore.Count - 1

If Val(imgStore(i).Tag) = HeroTag Then
    a = i
End If

Next i

FindImage = a
End Function
Private Sub LoadMissions()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

lstMissions.Clear

strSQL = "SELECT First(Missions.ID) AS FirstOfID, Missions.Name From Missions GROUP BY Missions.Name;"

objRS.Open strSQL, db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

Do Until objRS.EOF

lstMissions.AddItem objRS.Fields("Name").Value
lstMissions.ItemData(lstMissions.NewIndex) = objRS.Fields("FirstOfID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close

If lstMissions.ListCount > 0 Then lstMissions.ListIndex = 0


End Sub
Private Sub ShowPlayableTeamworkCards()
Dim ctemp As Collection
Dim myh As clsHero

Set myTeamwork = New clsTeamwork
Set ctemp = New Collection

If chkPlayTeamwork.Value = 0 Or cdeckheroes.Count = 0 Then

    Set ctemp = myTeamwork.GetPlayableTeamworkCards(10, 10, 10, 10)
Else

    nhe = 0
    nhf = 0
    nhs = 0
    nhi = 0
    
    For i = 1 To cdeckheroes.Count
        Set myh = New clsHero
        myh.Load cdeckheroes.Item(i)
        
        If myh.Energy > nhe Then nhe = myh.Energy
        If myh.Fighting > nhf Then nhf = myh.Fighting
        If myh.Strength > nhs Then nhs = myh.Strength
        If myh.Intellect > nhi Then nhi = myh.Intellect
    Next i
    
    Set myh = Nothing

    
    Set ctemp = myTeamwork.GetPlayableTeamworkCards(nhe, nhf, nhs, nhi)
    

End If

lstTeamwork.Clear

For i = 1 To ctemp.Count
    Set myTeamwork = New clsTeamwork
    myTeamwork.Load ctemp.Item(i)
    lstTeamwork.AddItem myTeamwork.Title
    lstTeamwork.ItemData(lstTeamwork.NewIndex) = myTeamwork.ID
Next i

Set myTeamwork = Nothing
Set ctemp = Nothing
Set myh = Nothing

End Sub

Private Sub mnuFileNew_Click()

If cdeck.Count > 0 Then
    x = MsgBox("Save changes to current deck?", vbYesNoCancel, "Save Deck?")

    If x = 2 Then Exit Sub
       
    If x = 6 Then SaveDeck

End If

newdeck

End Sub
Private Sub SaveDeck()
Dim myh As clsHero
Dim ccard

If sdeckname = "New Deck" Then
    c$ = ""
    For i = 1 To cdeckheroes.Count
        Set myh = New clsHero
        myh.Load cdeckheroes.Item(i)
        c$ = c$ & " " & myh.Name
    Next i
    
    c$ = ReplaceAllInString(c$, "/", "_")
    c$ = ReplaceAllInString(c$, "\", "_")
    c$ = ReplaceAllInString(c$, ":", "_")
    c$ = ReplaceAllInString(c$, "*", "_")
    c$ = ReplaceAllInString(c$, "?", "_")
    c$ = ReplaceAllInString(c$, Chr(34), "_")
    c$ = ReplaceAllInString(c$, "<", "_")
    c$ = ReplaceAllInString(c$, ">", "_")
    c$ = ReplaceAllInString(c$, "|", "_")
    c$ = Trim$(c$)
    
    sdeckname = InputBox$("Please enter a name for this deck:", "Save Deck", c$)

    c$ = sdeckname
    
    c$ = ReplaceAllInString(c$, "/", "_")
    c$ = ReplaceAllInString(c$, "\", "_")
    c$ = ReplaceAllInString(c$, ":", "_")
    c$ = ReplaceAllInString(c$, "*", "_")
    c$ = ReplaceAllInString(c$, "?", "_")
    c$ = ReplaceAllInString(c$, Chr(34), "_")
    c$ = ReplaceAllInString(c$, "<", "_")
    c$ = ReplaceAllInString(c$, ">", "_")
    c$ = ReplaceAllInString(c$, "|", "_")
    
    sdeckname = c$
    
    If sdeckname = "" Then
        sdeckname = "New Deck"
        Me.Caption = "Overpower Deck Editor: New Deck"
        Exit Sub
    End If
    
End If


SaveDeck2


End Sub
Private Sub SaveDeck2()
x = FreeFile

Open App.Path & "\Decks\" & sdeckname & ".dat" For Output As #x

For i = 1 To 4
    If i > cdeckheroes.Count Then
        Print #x, "HERO=-1"
    Else
        Print #x, "HERO=" & Trim(Str(cdeckheroes.Item(i)))
    End If
Next i


If nReserve = 0 Then nReserve = 4

Print #x, "RESERVE=" & Trim(Str(nReserve))


'Locations
Print #x, "HOMEBASE=" & Trim(Str(nHomebase))
Print #x, "BATTLESITE=" & Trim(Str(nBattlesite))

'MISSION
Print #x, "MISSION=" & Trim(Str(nMission))

'Deck count
Print #x, "DECK=" & Trim(Str(cdeck.Count))

For i = 1 To cdeck.Count

Set ccard = cdeck.Item(i)

Print #x, "CARD=" & ccard.CardType
Print #x, "CARDID=" & Trim(Str(ccard.ID))

Next i

Print #x, "BATTLESITE DECK=" & Trim(Str(cbattlesitedeck.Count))

For i = 1 To cbattlesitedeck.Count
Set ccard = cbattlesitedeck.Item(i)

Print #x, "CARD=" & ccard.CardType
Print #x, "CARDID=" & Trim(Str(ccard.ID))

Next i


Close #x

MsgBox "Deck saved."
Me.Caption = "Overpower Deck Editor: " & sdeckname

End Sub
Private Sub ShowAllTrainingCards()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

lstTraining.Clear

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

strSQL = "SELECT * FROM Training;"

objRS.Open strSQL, db
While Not objRS.EOF

lstTraining.AddItem "TRAINING: " & objRS.Fields("PWR1").Value & objRS.Fields("PWR2") & " + " & objRS.Fields("Bonus").Value
lstTraining.ItemData(lstTraining.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext

Wend

objRS.Close
db.Close

End Sub
Private Sub ShowAllBasicUniverse()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

lstBasicUniverse.Clear

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

strSQL = "SELECT * FROM [Basic Universe];"

objRS.Open strSQL, db
While Not objRS.EOF

lstBasicUniverse.AddItem "BASIC UNIVERSE: " & Trim(Str(objRS.Fields("requires").Value)) & objRS.Fields("Skill").Value & "+" & Trim(Str(objRS.Fields("Bonus").Value))
lstBasicUniverse.ItemData(lstBasicUniverse.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext

Wend

objRS.Close
db.Close

End Sub
Private Sub ShowAllTeamwork()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

lstTeamwork.Clear

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

strSQL = "SELECT * FROM Teamwork;"

objRS.Open strSQL, db
While Not objRS.EOF

lstTeamwork.AddItem "TEAMWORK: " & objRS.Fields("T1_PW").Value & objRS.Fields("T1_SK").Value & "/+" & Trim(Str(objRS.Fields("Bonus1").Value)) & ", +" & Trim(Str(objRS.Fields("Bonus2").Value)) & " " & objRS.Fields("T2_SK").Value & objRS.Fields("T3_SK").Value
lstTeamwork.ItemData(lstTeamwork.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext

Wend

objRS.Close
db.Close

End Sub

Private Sub mnuFileOpen_Click()
With cmD1

.InitDir = App.Path & "\Decks"
.FileName = "*.dat"
.Action = 1
If .FileName = "*.dat" Then Exit Sub

OpenDeck .FileName, .FileTitle


End With


End Sub
Private Sub OpenDeck(sfilename, sfiletitle)

newdeck

x = FreeFile

Open sfilename For Input As #x

'read 4 heroes
For i = 1 To 4

Line Input #x, a$
    cdeckheroes.Add Val(GetVal(a$))
Next i

Line Input #x, a$
nReserve = Val(GetVal(a$))

showdeckheroes

Line Input #x, a$
nHomebase = Val(GetVal(a$))

Line Input #x, a$
nBattlesite = Val(GetVal(a$))

'Get Mission
Line Input #x, a$
nMission = Val(GetVal(a$))

'Need to set up mission

'Get number of cards in deck
Line Input #x, a$
ncards = Val(GetVal(a$))

For i = 1 To ncards

Line Input #x, a$
scardtype = GetVal(a$)

Line Input #x, a$
ncardid = Val(GetVal(a$))

Select Case scardtype

Case "Activator"
    Set myActivator = New clsActivator
    myActivator.Load ncardid
    cdeck.Add myActivator
    
Case "Ally Card"
    Set myAlly = New clsAlly
    myAlly.Load ncardid
    cdeck.Add myAlly
    
Case "Artifact"
    Set myArtifact = New clsArtifact
    myArtifact.Load ncardid
    cdeck.Add myArtifact
    
Case "Aspect Card"
    Set myAspect = New clsAspect
    myAspect.Load ncardid
    cdeck.Add myAspect
    
Case "Basic Universe"
    Set myBasic = New clsBasicUniverse
    myBasic.Load ncardid
    cdeck.Add myBasic
    
Case "Double Shot"
    Set myDoubleShot = New clsDoubleShot
    myDoubleShot.Load ncardid
    cdeck.Add myDoubleShot
    
Case "Event"
    Set myEvent = New clsEvent
    myEvent.Load ncardid
    cdeck.Add myEvent
    
Case "Power Card"
    Set myPower = New clsPowerCard
    myPower.Load ncardid
    cdeck.Add myPower
    
Case "Special Card"
    Set myspecial = New clsSpecial
    myspecial.Load ncardid
    cdeck.Add myspecial
    
Case "Teamwork"
    Set myTeamwork = New clsTeamwork
    myTeamwork.Load ncardid
    cdeck.Add myTeamwork
    
    
Case "Training"
    Set myTraining = New clsTraining
    myTraining.Load ncardid
    cdeck.Add myTraining

Case Else
End Select

Next i

'Get battlesite deck cards
Line Input #x, a$
ncards = Val(GetVal(a$))

For i = 1 To ncards

Line Input #x, a$
scardtype = GetVal(a$)

Line Input #x, a$
ncardid = Val(GetVal(a$))

Select Case scardtype

Case "Activator"
    Set myActivator = New clsActivator
    myActivator.Load ncardid
    cbattlesitedeck.Add myActivator
    
Case "Ally Card"
    Set myAlly = New clsAlly
    myAlly.Load ncardid
    cbattlesitedeck.Add myAlly
    
Case "Aspect Card"
    Set myAspect = New clsAspect
    myAspect.Load ncardid
    cbattlesitedeck.Add myAspect
    
Case "Basic Universe"
    Set myBasic = New clsBasicUniverse
    myBasic.Load ncardid
    cbattlesitedeck.Add myBasic
    
Case "Double Shot"
    Set myDoubleShot = New clsDoubleShot
    myDoubleShot.Load ncardid
    cbattlesitedeck.Add myDoubleShot
    
Case "Event"
    Set myEvent = New clsEvent
    myEvent.Load ncardid
    cbattlesitedeck.Add myEvent
    
Case "Power Card"
    Set myPower = New clsPowerCard
    myPower.Load ncardid
    cbattlesitedeck.Add myPower
    
Case "Special Card"
    Set myspecial = New clsSpecial
    myspecial.Load ncardid
    cbattlesitedeck.Add myspecial
    
Case "Teamwork"
    Set myTeamwork = New clsTeamwork
    myTeamwork.Load ncardid
    cbattlesitedeck.Add myTeamwork
    
Case "Training"
    Set myTraining = New clsTraining
    myTraining.Load ncardid
    cbattlesitedeck.Add myTraining

Case Else
End Select

Next i

Close #x
sdeckname = Left(sfiletitle, Len(sfiletitle) - 4)

ShowDeckCount
ShowCurrentHomeBase
ShowCurrentBattleSite

updateSpecialHeroes

For i = 0 To lstMissions.ListCount - 1
    If lstMissions.ItemData(i) = nMission Then
        lstMissions.ListIndex = i
    End If
Next i

If lstMissions.ListIndex = -1 Then lstMissions.ListIndex = 0


End Sub

Private Sub mnuFileSave_Click()

SaveDeck

End Sub

Private Sub mnuFileSaveDeckAs_Click()
    c$ = InputBox$("Save deck as:", "Save Deck As", sdeckname)

    c$ = ReplaceAllInString(c$, "/", "_")
    c$ = ReplaceAllInString(c$, "\", "_")
    c$ = ReplaceAllInString(c$, ":", "_")
    c$ = ReplaceAllInString(c$, "*", "_")
    c$ = ReplaceAllInString(c$, "?", "_")
    c$ = ReplaceAllInString(c$, Chr(34), "_")
    c$ = ReplaceAllInString(c$, "<", "_")
    c$ = ReplaceAllInString(c$, ">", "_")
    c$ = ReplaceAllInString(c$, "|", "_")
    
    If c$ = "" Then Exit Sub
    
    sdeckname = c$
    Me.Caption = "Overpower Deck Editor: " & sdeckname
    
    SaveDeck2
    

End Sub

Private Sub LoadArtifacts()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

lstArtifacts.Clear


objRS.Open "SELECT * FROM Artifact ORDER BY Artifact.Character;", db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

Do Until objRS.EOF

lstArtifacts.AddItem objRS.Fields("Character").Value
lstArtifacts.ItemData(lstArtifacts.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close

If lstArtifacts.ListCount > 0 Then
    lstArtifacts.ListIndex = 0
End If

End Sub
Private Sub LoadAspects()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

lstAspects.Clear


objRS.Open "SELECT * FROM Aspect ORDER BY Aspect.Name;", db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

Do Until objRS.EOF

lstAspects.AddItem objRS.Fields("Name").Value & " [" & objRS.Fields("Homebase").Value & "]"
lstAspects.ItemData(lstAspects.NewIndex) = objRS.Fields("ID").Value

objRS.MoveNext
Loop


objRS.Close
db.Close


End Sub
Private Sub LoadDoubleShots()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

lstDoubleshot.Clear

objRS.Open "SELECT * FROM Doubleshot ORDER BY Doubleshot.ID;", db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

Do Until objRS.EOF

Set myDoubleShot = New clsDoubleShot
myDoubleShot.Load objRS.Fields("ID").Value
lstDoubleshot.AddItem myDoubleShot.Title
lstDoubleshot.ItemData(lstDoubleshot.NewIndex) = myDoubleShot.ID

objRS.MoveNext
Loop


objRS.Close
db.Close

End Sub
Sub LoadAllys()

Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

lstAllys.Clear

objRS.Open "SELECT * FROM Ally;", db

If objRS.EOF = True Then
    objRS.Close
    db.Close
    Exit Sub
End If

Do Until objRS.EOF

Set myAlly = New clsAlly
myAlly.Load objRS.Fields("ID").Value
lstAllys.AddItem myAlly.Title
lstAllys.ItemData(lstAllys.NewIndex) = myAlly.ID


objRS.MoveNext
Loop


objRS.Close
db.Close


End Sub

Private Sub mnuLoadBattlesiteDeck_Click()
With cmD1

.InitDir = App.Path & "\Decks"
.FileName = "*.bsd"

.Action = 1

If .FileName = "" Then Exit Sub

OpenBattlesiteDeck .FileName, .FileTitle


End With
End Sub
Private Sub mnuRandomCharacterChallenge_Click()
Dim db As ADODB.Connection
Dim objRS As ADODB.Recordset
Dim ctemp As Collection
Dim ctemp2 As Collection
Dim cchars As Collection

Set ctemp = New Collection
Set ctemp2 = New Collection
Set cchars = New Collection

Do Until cdeckheroes.Count = 0
    cdeckheroes.Remove 1
Loop

Randomize

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set objRS = New ADODB.Recordset

objRS.Open "SELECT * FROM Characters ORDER BY Characters.Character;", db

Do Until objRS.EOF

ctemp.Add objRS.Fields("ID").Value
ctemp2.Add objRS.Fields("Character").Value

objRS.MoveNext
Loop

objRS.Close
db.Close

Randomize

For i = 1 To 4

cdeckheroes.Add ctemp.Item(Int(Rnd * ctemp.Count) + 1)

Next i

showdeckheroes
updateSpecialHeroes

DoEvents
Me.Refresh


Exit Sub

End Sub

Private Sub mnuSaveBattlesiteDeck_Click()
SaveBattleSiteDeck

End Sub

Private Sub mnuTools_Click()
If cdeck.Count < 8 Then
    mnuToolsTestDraws.Enabled = False
Else
    mnuToolsTestDraws.Enabled = True
End If

End Sub

Private Sub mnuToolsTestDraws_Click()
frmTestDraw.Show 1

End Sub

Private Sub txtFind_Change()
If Trim(txtFind.Text) = "" Then
    cmdFind.Enabled = False
Else
    cmdFind.Enabled = True
End If

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(txtFind.Text) <> "" Then cmdFind = True

End Sub

Private Sub txtFindLoc_Change()
If Trim(txtFindLoc.Text) <> "" Then
    cmdFindLocs.Enabled = True
Else
    cmdFindLocs.Enabled = False
End If

End Sub

Private Sub txtFindLoc_GotFocus()
With txtFindLoc
.SelStart = 0
.SelLength = Len(.Text)
End With


End Sub

Private Sub txtMinE_GotFocus()
txtMinE.SelStart = 0
txtMinE.SelLength = Len(txtMinE.Text)

End Sub

Private Sub txtMinE_LostFocus()
txtMinE.Text = Trim(txtMinE.Text)
If txtMinE.Text = "" Then txtMinE.Text = "0"

End Sub


Private Sub txtMinF_GotFocus()
txtMinF.SelStart = 0
txtMinF.SelLength = Len(txtMinF.Text)

End Sub
Private Sub txtMinF_LostFocus()
txtMinF.Text = Val(Trim(txtMinF.Text))
If txtMinF.Text = "" Then txtMinF.Text = "0"

End Sub
Private Sub txtMinI_GotFocus()
txtMinI.SelStart = 0
txtMinI.SelLength = Len(txtMinI.Text)

End Sub

Private Sub txtMinI_LostFocus()
txtMinI.Text = Val(Trim(txtMinI.Text))
If txtMinI.Text = "" Then txtMinI.Text = "0"

End Sub

Private Sub txtMinS_GotFocus()
txtMinS.SelStart = 0
txtMinS.SelLength = Len(txtMinS.Text)

End Sub

Private Sub txtMinS_LostFocus()
txtMinS.Text = Val(Trim(txtMinS.Text))
If txtMinS.Text = "" Then txtMinS.Text = "0"

End Sub

Private Sub txtTotal_GotFocus()
txtTotal.SelStart = 0
txtTotal.SelLength = Len(txtTotal.Text)

End Sub

Private Sub txtTotal_LostFocus()
txtTotal.Text = Val(Trim(txtTotal.Text))
If txtTotal.Text = "" Then txtTotal.Text = "99"

End Sub
Private Sub SaveBattleSiteDeck()
Set myBattleSite = New clsBattlesite
myBattleSite.Load nBattlesite

c$ = myBattleSite.Name

Set myBattleSite = Nothing

   
        
    c$ = ReplaceAllInString(c$, "/", "_")
    c$ = ReplaceAllInString(c$, "\", "_")
    c$ = ReplaceAllInString(c$, ":", "_")
    c$ = ReplaceAllInString(c$, "*", "_")
    c$ = ReplaceAllInString(c$, "?", "_")
    c$ = ReplaceAllInString(c$, Chr(34), "_")
    c$ = ReplaceAllInString(c$, "<", "_")
    c$ = ReplaceAllInString(c$, ">", "_")
    c$ = ReplaceAllInString(c$, "|", "_")
    c$ = Trim$(c$)
    
    sbdeckname = InputBox$("Please enter a name for this Battlesite Deck:", "Save Battlesite Deck", c$)

    c$ = sbdeckname
    
    c$ = ReplaceAllInString(c$, "/", "_")
    c$ = ReplaceAllInString(c$, "\", "_")
    c$ = ReplaceAllInString(c$, ":", "_")
    c$ = ReplaceAllInString(c$, "*", "_")
    c$ = ReplaceAllInString(c$, "?", "_")
    c$ = ReplaceAllInString(c$, Chr(34), "_")
    c$ = ReplaceAllInString(c$, "<", "_")
    c$ = ReplaceAllInString(c$, ">", "_")
    c$ = ReplaceAllInString(c$, "|", "_")
    
    sbdeckname = c$
    

SaveBattleSiteDeck2 sbdeckname


End Sub
Private Sub SaveBattleSiteDeck2(sname)

x = FreeFile

Open App.Path & "\Decks\" & sname & ".bsd" For Output As #x

'Battlesite
Print #x, "BATTLESITE=" & Trim(Str(nBattlesite))
Print #x, "BATTLESITE DECK=" & Trim(Str(cbattlesitedeck.Count))

For i = 1 To cbattlesitedeck.Count
Set ccard = cbattlesitedeck.Item(i)

Print #x, "CARD=" & ccard.CardType
Print #x, "CARDID=" & Trim(Str(ccard.ID))

Next i


Close #x

MsgBox "Battlesite Deck saved."

End Sub
Private Sub OpenBattlesiteDeck(sfilename, sfiletitle)
cap$ = Me.Caption

z$ = Left(sfiletitle, Len(sfiletitle) - 4)

Me.Caption = "Overpower Deck Editor: Loading " & z$ & " Battlesite Deck"
Me.Refresh

x = FreeFile

Open sfilename For Input As #x

Line Input #x, a$
nBattlesite = Val(GetVal(a$))


'Get battlesite deck cards
Line Input #x, a$
ncards = Val(GetVal(a$))

For i = 1 To ncards

Line Input #x, a$
scardtype = GetVal(a$)

Line Input #x, a$
ncardid = Val(GetVal(a$))

Select Case scardtype
    
Case "Special Card"
    Set myspecial = New clsSpecial
    myspecial.Load ncardid
    cbattlesitedeck.Add myspecial

    Dim myh As clsHero
    
    Set myh = New clsHero
    nId = myh.GetIDFromName(myspecial.Character)
    
    Set myh = Nothing
    
    Set myActivator = New clsActivator
    myActivator.Load nId
    cdeck.Add myActivator
    
Case Else
End Select

Next i

Close #x

ShowDeckCount
ShowCurrentBattleSite

updateSpecialHeroes

Me.Caption = cap$

End Sub
