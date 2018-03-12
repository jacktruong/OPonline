VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTable 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overpower Online"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   810
   ClientWidth     =   15015
   ForeColor       =   &H00400040&
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   15015
   Begin VB.ListBox lstHandTips 
      Height          =   2010
      Left            =   13560
      TabIndex        =   183
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame frmActInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Available Specials in Battlesite Deck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   6120
      TabIndex        =   179
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.ListBox lstAvailableActivators 
         Height          =   1620
         Left            =   120
         TabIndex        =   180
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.ListBox lstCodes 
      Height          =   2400
      Left            =   13560
      TabIndex        =   115
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5520
      Top             =   600
   End
   Begin VB.Frame frmString 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "String Attack"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   4320
      TabIndex        =   169
      Top             =   1080
      Visible         =   0   'False
      Width           =   1620
      Begin VB.Label lblPile1 
         BackStyle       =   0  'Transparent
         Caption         =   "String Attack:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   75
         TabIndex        =   170
         Top             =   105
         Width           =   1455
      End
      Begin VB.Image imgStringAttack 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2025
         Left            =   60
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSWinsockLib.Winsock tcpChannel 
      Left            =   4440
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame frmAttack 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   6120
      TabIndex        =   98
      ToolTipText     =   "Attack Box"
      Top             =   3960
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CheckBox chkPlayFaceDown 
         BackColor       =   &H008080FF&
         Caption         =   "Play face down"
         Height          =   315
         Left            =   5280
         TabIndex        =   173
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelAction 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5400
         TabIndex        =   100
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdOKAction 
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   99
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape shpAction 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Height          =   1575
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Image imgAction 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   0
         Left            =   240
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Tag             =   "Discard"
         Top             =   55
         Width           =   1095
      End
      Begin VB.Image imgAction 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   1
         Left            =   1440
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Tag             =   "Discard"
         Top             =   55
         Width           =   1095
      End
      Begin VB.Image imgAction 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   2
         Left            =   2640
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Tag             =   "Discard"
         Top             =   55
         Width           =   1095
      End
      Begin VB.Image imgAction 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   3
         Left            =   3840
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Tag             =   "Discard"
         Top             =   55
         Width           =   1095
      End
      Begin VB.Shape shpActionBorder 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   1455
         Left            =   240
         Top             =   55
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   4335
      Begin VB.Frame frmHero 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   0
         TabIndex        =   43
         Top             =   3080
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtInherent 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   480
            Width           =   4095
         End
         Begin VB.CommandButton cmdKOCharacter 
            Caption         =   "K.O. Character"
            Height          =   495
            Left            =   1560
            TabIndex        =   45
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton cmdTakeAction 
            Caption         =   "Action"
            Height          =   495
            Left            =   120
            TabIndex        =   44
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton cmdReserveToFrontline 
            Caption         =   "To Frontline"
            Height          =   495
            Left            =   3000
            TabIndex        =   47
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdSwitchWithReserve 
            Caption         =   "To Reserve"
            Height          =   495
            Left            =   3000
            TabIndex        =   48
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Inherent Ability:"
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
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame frmEvent 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   94
         Top             =   3120
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdPlayEvent 
            Caption         =   "Play && Draw 1"
            Height          =   495
            Left            =   1320
            TabIndex        =   97
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdDiscardEvent1 
            Caption         =   "Discard && Draw 1"
            Height          =   495
            Left            =   1320
            TabIndex        =   96
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdDiscardEvent 
            Caption         =   "Discard"
            Height          =   495
            Left            =   1320
            TabIndex        =   95
            Top             =   1560
            Width           =   1575
         End
      End
      Begin VB.Frame frmHomebase 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   0
         TabIndex        =   38
         Top             =   3080
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtHomebaseBonus 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   1440
            Width           =   4095
         End
         Begin VB.ListBox lstHomeBaseChars 
            Height          =   645
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Characters:"
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
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bonus:"
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
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   3975
         End
      End
      Begin VB.Frame frmActivator 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   91
         Top             =   3120
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdExchangeActivator 
            Caption         =   "Exchange for Card"
            Height          =   495
            Left            =   1320
            TabIndex        =   93
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdDiscardActivator 
            Caption         =   "Discard Activator"
            Height          =   495
            Left            =   1320
            TabIndex        =   92
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame frmVentured 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   66
         Top             =   3120
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CheckBox chkMoveAll 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Move All"
            Height          =   255
            Left            =   2160
            TabIndex        =   70
            Top             =   1920
            Width           =   1815
         End
         Begin VB.CommandButton cmdReturnToMissions 
            Caption         =   "Return to Reserve"
            Height          =   615
            Left            =   360
            TabIndex        =   69
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton cmdDefeated 
            Caption         =   "To Defeated"
            Height          =   615
            Left            =   360
            TabIndex        =   68
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdCompleted 
            Caption         =   "To Completed"
            Height          =   615
            Left            =   360
            TabIndex        =   67
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frmBattlesite 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   58
         Top             =   3120
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdKOBattlesite 
            Caption         =   "K.O. Battlesite"
            Height          =   615
            Left            =   1320
            TabIndex        =   60
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CommandButton cmdViewBattlesiteDeck 
            Caption         =   "View Battlesite Deck"
            Height          =   615
            Left            =   1320
            TabIndex        =   59
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame frmVentureC 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   75
         Top             =   3120
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdVenCToReserve 
            Caption         =   "To Reserve"
            Height          =   615
            Left            =   360
            TabIndex        =   79
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdVenCToDefeated 
            Caption         =   "To Defeated"
            Height          =   615
            Left            =   360
            TabIndex        =   78
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdVenCReturn 
            Caption         =   "Return to Completed"
            Height          =   615
            Left            =   360
            TabIndex        =   77
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox chkMoveAllVenC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Move All"
            Height          =   255
            Left            =   2400
            TabIndex        =   76
            Top             =   1920
            Width           =   1215
         End
      End
      Begin VB.Frame frmMission 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   54
         Top             =   3120
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdVenture 
            Caption         =   "Venture 3 Cards"
            Height          =   495
            Index           =   2
            Left            =   600
            TabIndex        =   57
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdVenture 
            Caption         =   "Venture 2 Cards"
            Height          =   495
            Index           =   1
            Left            =   2160
            TabIndex        =   56
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdVenture 
            Caption         =   "Venture 1 Card"
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   55
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame frmDefeated 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         TabIndex        =   71
         Top             =   3000
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CheckBox chkMoveAllD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Move All"
            Height          =   255
            Left            =   1320
            TabIndex        =   74
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CommandButton chkReturnMissionD 
            Caption         =   "Return to Reserve"
            Height          =   615
            Left            =   1320
            TabIndex        =   73
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton chkMoveCompletedD 
            Caption         =   "To Completed"
            Height          =   615
            Left            =   1320
            TabIndex        =   72
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.Frame frmCompletedMission 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         TabIndex        =   61
         Top             =   3000
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdVentureC 
            Caption         =   "Venture"
            Height          =   615
            Left            =   1320
            TabIndex        =   65
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdDefeatedC 
            Caption         =   "To Defeated"
            Height          =   615
            Left            =   1320
            TabIndex        =   64
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton cmdReturnToMissionC 
            Caption         =   "To Reserve Missions"
            Height          =   615
            Left            =   1320
            TabIndex        =   63
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkMoveAllC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Move All"
            Height          =   255
            Left            =   1320
            TabIndex        =   62
            Top             =   2640
            Width           =   1815
         End
      End
      Begin VB.Frame frmPlaced 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   80
         Top             =   5160
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdPlayPlaced 
            Caption         =   "Play"
            Height          =   375
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReturnToHand 
            Caption         =   "Return to Hand"
            Height          =   375
            Left            =   2760
            TabIndex        =   82
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdDiscardPlaced 
            Caption         =   "Discard"
            Height          =   375
            Left            =   1440
            TabIndex        =   81
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame frmHandCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   50
         Top             =   5280
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdDiscard 
            Caption         =   "Discard"
            Height          =   375
            Left            =   2640
            TabIndex        =   53
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlace 
            Caption         =   "Place"
            Height          =   375
            Left            =   1320
            TabIndex        =   52
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdAttack 
            Caption         =   "Play"
            Height          =   375
            Left            =   0
            TabIndex        =   51
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame frmDefenseCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   112
         Top             =   5280
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdRemoveDefenseCard 
            Caption         =   "Remove Card"
            Height          =   495
            Left            =   1200
            TabIndex        =   113
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame frmModifier 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   174
         Top             =   5280
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdPlaceModifier 
            Caption         =   "Place"
            Height          =   375
            Left            =   1440
            TabIndex        =   177
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdDiscardModifier 
            Caption         =   "Discard"
            Height          =   375
            Left            =   2760
            TabIndex        =   175
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayModifier 
            Caption         =   "Play"
            Height          =   375
            Left            =   120
            TabIndex        =   176
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame frmPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   88
         Top             =   5160
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdPRRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPRMove 
            Caption         =   "Move"
            Height          =   375
            Left            =   1440
            TabIndex        =   89
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame frmAttackCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   171
         Top             =   5280
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdRemoveAttackCard 
            Caption         =   "Remove Card"
            Height          =   495
            Left            =   1200
            TabIndex        =   172
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame frmHTCB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   84
         Top             =   5160
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdHTCBMove 
            Caption         =   "Move"
            Height          =   375
            Left            =   1320
            TabIndex        =   87
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdHTCBToPR 
            Caption         =   "To Perm. Record"
            Height          =   375
            Left            =   2640
            TabIndex        =   86
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdRemoveHTCB 
            Caption         =   "Remove"
            Height          =   375
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Image imgCardDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   5205
         Left            =   240
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmTable.frx":1272
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3720
      End
      Begin VB.Image imgMissionCard 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3045
         Left            =   0
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmTable.frx":C74C
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.Image imgHeroCard 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3045
         Left            =   0
         OLEDragMode     =   1  'Automatic
         Picture         =   "frmTable.frx":54F23
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   4245
      End
   End
   Begin MSComDlg.CommonDialog cmD1 
      Left            =   4920
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   10005
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Me: 0"
            TextSave        =   "Me: 0"
            Key             =   "MyVenture"
            Object.ToolTipText     =   "My Current Venture Total"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Opp: 0"
            TextSave        =   "Opp: 0"
            Key             =   "OppVenture"
            Object.ToolTipText     =   "Current Opponent Venture Total"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      TabIndex        =   101
      Top             =   5760
      Width           =   4335
      Begin VB.CommandButton cmdSendMessage 
         Caption         =   "Send &Message"
         Height          =   375
         Left            =   2880
         TabIndex        =   105
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ListBox lstMessages 
         Height          =   1230
         Left            =   120
         TabIndex        =   104
         Top             =   600
         Width           =   3975
      End
      Begin VB.ListBox lstGameHistory 
         Height          =   1425
         Left            =   120
         TabIndex        =   102
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Messages:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   168
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblPhase 
         BackStyle       =   0  'Transparent
         Caption         =   "PHASE:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Game History:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   103
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Frame frmDefense 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   6480
      TabIndex        =   108
      ToolTipText     =   "Defense Box"
      Top             =   5400
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdShiftAttack 
         Caption         =   "Shift Attack"
         Height          =   375
         Left            =   3840
         TabIndex        =   178
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox imgHideDefense 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   6350
         Picture         =   "frmTable.frx":9D6FA
         ScaleHeight     =   165
         ScaleWidth      =   210
         TabIndex        =   109
         Top             =   1000
         Width           =   210
      End
      Begin VB.CommandButton cmdOKDefense 
         Caption         =   "OK"
         Height          =   375
         Left            =   5040
         TabIndex        =   111
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAcceptDefense 
         Caption         =   "Accept"
         Height          =   375
         Left            =   5040
         TabIndex        =   116
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNoDefense 
         Caption         =   "No Defense"
         Height          =   375
         Left            =   5040
         TabIndex        =   110
         Top             =   680
         Width           =   1215
      End
      Begin VB.CommandButton cmdChallengeDefense 
         Caption         =   "Challenge"
         Height          =   375
         Left            =   5040
         TabIndex        =   117
         Top             =   680
         Width           =   1215
      End
      Begin VB.Image imgDefense 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   4
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgDefense 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   3
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgDefense 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   2
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgDefense 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   1
         Left            =   960
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgDefense 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frmDiscardPhase 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Phase: Discard"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   6120
      TabIndex        =   125
      Top             =   3840
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdFinishedDiscarding 
         Caption         =   "Finished"
         Height          =   375
         Left            =   5520
         TabIndex        =   127
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Event Cards should also be played at this time."
         Height          =   255
         Left            =   240
         TabIndex        =   130
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblDiscard1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   129
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblDiscard1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   128
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTable.frx":9D7E4
         Height          =   615
         Left            =   240
         TabIndex        =   126
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Frame frmVenturePhase 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Phase: Venture"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   6120
      TabIndex        =   137
      Top             =   3840
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdFinishedVenture 
         Caption         =   "Finished"
         Height          =   375
         Left            =   5520
         TabIndex        =   138
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTable.frx":9D8F0
         Height          =   615
         Left            =   240
         TabIndex        =   142
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label lblDiscard1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   141
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblDiscard1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   140
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the ""Finished"" button when you are done venturing cards."
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   1080
         Width           =   5415
      End
   End
   Begin VB.Frame frmPlacingPhase 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Phase: Placing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   6120
      TabIndex        =   131
      Top             =   3840
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdFinishedPlacing 
         Caption         =   "Finished"
         Height          =   375
         Left            =   5520
         TabIndex        =   132
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: [Player 1] Should place first"
         Height          =   255
         Left            =   240
         TabIndex        =   136
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblDiscard1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   135
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblDiscard1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   134
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTable.frx":9D9E2
         Height          =   615
         Left            =   240
         TabIndex        =   133
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Frame frmWhoGoesFirst 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Who Goes First?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   4560
      TabIndex        =   118
      Top             =   3840
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton cmdWGFDraw2 
         Caption         =   "Draw Again"
         Height          =   495
         Left            =   5880
         TabIndex        =   124
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdWGFDraw1 
         Caption         =   "Draw Again"
         Height          =   495
         Left            =   5880
         TabIndex        =   123
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdWGF2 
         Caption         =   "He Goes First"
         Height          =   495
         Left            =   3600
         TabIndex        =   122
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdWGF1 
         Caption         =   "I Go First"
         Height          =   495
         Left            =   3600
         TabIndex        =   121
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblWGF2 
         BackStyle       =   0  'Transparent
         Caption         =   "my Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1920
         TabIndex        =   120
         Top             =   315
         Width           =   1455
      End
      Begin VB.Image imgWGF2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2025
         Left            =   1920
         OLEDropMode     =   1  'Manual
         Picture         =   "frmTable.frx":9DAB0
         Stretch         =   -1  'True
         Tag             =   "Hand"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblWGF1 
         BackStyle       =   0  'Transparent
         Caption         =   "my Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   119
         Top             =   315
         Width           =   1455
      End
      Begin VB.Image imgWGF1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2025
         Left            =   240
         OLEDropMode     =   1  'Manual
         Picture         =   "frmTable.frx":A8F8A
         Stretch         =   -1  'True
         Tag             =   "Hand"
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame frmResolveVenturePhase 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Phase: Resolve Venture"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   6120
      TabIndex        =   143
      Top             =   3840
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtMyVentureTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   148
         Text            =   "--"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdSendMyVenTotal 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2760
         TabIndex        =   147
         Top             =   920
         Width           =   735
      End
      Begin VB.TextBox txtOppVentureTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   146
         Text            =   "--"
         Top             =   1365
         Width           =   615
      End
      Begin VB.CommandButton cmdAcceptOppVenTotal 
         Caption         =   "&Accept"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2760
         TabIndex        =   145
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chkVTTotalAccepted 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total accepted by opponent"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3840
         TabIndex        =   144
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTable.frx":B4464
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   151
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "My Venture Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   150
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Opponent's Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   149
         Top             =   1410
         Width           =   1695
      End
   End
   Begin VB.Frame frmMoveVentureCards 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Phase: Move Venture Cards"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   6120
      TabIndex        =   152
      Top             =   3840
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdFinishedMovingVenture 
         Caption         =   "Finished"
         Height          =   375
         Left            =   5520
         TabIndex        =   153
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "[PLAYER 1] has won the venture."
         Height          =   255
         Left            =   240
         TabIndex        =   157
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label lblMoveVenture 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   156
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Move your ventured Mission cards appropriately and click on ""Finished"" when done."
         Height          =   615
         Left            =   240
         TabIndex        =   155
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label lblMoveVenture 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Me: Discarding..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   154
         Top             =   1560
         Width           =   2415
      End
   End
   Begin VB.Label lblEvent 
      Caption         =   "Event Label1"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   182
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblEvent 
      Caption         =   "Event Label1"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   181
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgHitBuffer 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHitBufferOP 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpBuffer4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpBuffer3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpBuffer2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOPBuffer1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBuffer4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBuffer3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   8760
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBuffer2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBGameEffect2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBuffer1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHitOpBS 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpEffect4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpEffect3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpEffect2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpEffect1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBGameEffect4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBGameEffect3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   8760
      Stretch         =   -1  'True
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBGameEffect1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   4
      Left            =   10080
      TabIndex        =   162
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   4680
      TabIndex        =   158
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   6480
      TabIndex        =   159
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   8280
      TabIndex        =   160
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   6480
      TabIndex        =   161
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   9
      Left            =   6240
      TabIndex        =   167
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   8
      Left            =   9840
      TabIndex        =   166
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   7
      Left            =   11640
      TabIndex        =   165
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   6
      Left            =   9840
      TabIndex        =   164
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblKO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "K.O."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   5
      Left            =   8040
      TabIndex        =   163
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgHomebase 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   8040
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Image imgBattlesite 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   9840
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   9
      Visible         =   0   'False
      X1              =   10755
      X2              =   8880
      Y1              =   6840
      Y2              =   5520
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   16
      Left            =   5400
      TabIndex        =   107
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblPile1 
      Caption         =   "OP Hand:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   4560
      TabIndex        =   106
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgOpPlace4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOppReserve 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image imgOPHit4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   0
      Left            =   11400
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOPHit3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   6
      Visible         =   0   'False
      X1              =   9360
      X2              =   12240
      Y1              =   3960
      Y2              =   2880
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   5
      Visible         =   0   'False
      X1              =   9360
      X2              =   10440
      Y1              =   3960
      Y2              =   2880
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   4
      Visible         =   0   'False
      X1              =   9360
      X2              =   8640
      Y1              =   3960
      Y2              =   2880
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   8
      Visible         =   0   'False
      X1              =   9360
      X2              =   6720
      Y1              =   3960
      Y2              =   2880
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   7
      Visible         =   0   'False
      X1              =   9360
      X2              =   10440
      Y1              =   3960
      Y2              =   1320
   End
   Begin VB.Image imgOpponent 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   2
      Left            =   11400
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image imgOpponent 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   1
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image imgOpponent 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   0
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   3
      Visible         =   0   'False
      X1              =   7080
      X2              =   9345
      Y1              =   8520
      Y2              =   5520
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   2
      Visible         =   0   'False
      X1              =   8880
      X2              =   9360
      Y1              =   6840
      Y2              =   5520
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   7080
      X2              =   9360
      Y1              =   6840
      Y2              =   5520
   End
   Begin VB.Line lnFrontLine 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   5280
      X2              =   9480
      Y1              =   6840
      Y2              =   5520
   End
   Begin VB.Image imgEffect 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   1
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgEffect 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   0
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgHand 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2025
      Index           =   0
      Left            =   13440
      OLEDropMode     =   1  'Manual
      Picture         =   "frmTable.frx":B4500
      Stretch         =   -1  'True
      Tag             =   "Hand"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000003&
      BorderWidth     =   3
      Index           =   1
      X1              =   13320
      X2              =   13320
      Y1              =   0
      Y2              =   9240
   End
   Begin VB.Image imgHit4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHitBS 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHit3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHit2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgHit1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   15
      Left            =   12720
      TabIndex        =   36
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   14
      Left            =   12600
      TabIndex        =   35
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblPile1 
      Caption         =   "Defeated:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   11520
      TabIndex        =   34
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblPile1 
      Caption         =   "Completed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   11520
      TabIndex        =   33
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblPile1 
      Caption         =   "Reserve:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   11520
      TabIndex        =   32
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   13
      Left            =   12600
      TabIndex        =   31
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   12
      Left            =   6960
      TabIndex        =   30
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblPile1 
      Caption         =   "Draw:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   6120
      TabIndex        =   29
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblPile1 
      Caption         =   "Discard: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   6120
      TabIndex        =   28
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblPile1 
      Caption         =   "Dead:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   27
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   11
      Left            =   6960
      TabIndex        =   26
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   10
      Left            =   6960
      TabIndex        =   25
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   9
      Left            =   6960
      TabIndex        =   24
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblPile1 
      Caption         =   "Battlesite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   6120
      TabIndex        =   23
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   8
      Left            =   6960
      TabIndex        =   22
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblPile1 
      Caption         =   "Defeated:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   21
      Top             =   840
      Width           =   855
   End
   Begin VB.Image imgOpPlacedHomebase 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpHomeBase 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   7800
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Homebase"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8025
      TabIndex        =   20
      Top             =   525
      Width           =   1215
   End
   Begin VB.Image imgOpBattlesite 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Battlesite"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   19
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblPile1 
      Caption         =   "Defeated:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   18
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   17
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label lblPile1 
      Caption         =   "Battlesite:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   16
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   15
      Top             =   9600
      Width           =   375
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   10800
      TabIndex        =   14
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label lblPile1 
      Caption         =   "Reserve:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9960
      TabIndex        =   13
      Top             =   8760
      Width           =   975
   End
   Begin VB.Image imgPR4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPlacedHomeBase 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Battlesite"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   10155
      TabIndex        =   12
      Top             =   7275
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1215
      Index           =   0
      Left            =   9840
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Shape shpOpponent 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Height          =   1215
      Index           =   3
      Left            =   9600
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape shpOpponent 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Height          =   1215
      Index           =   2
      Left            =   11400
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape shpOpponent 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Height          =   1215
      Index           =   1
      Left            =   9600
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image imgOpPlace3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpPlace2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOpPlace1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpOpponent 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Height          =   1215
      Index           =   0
      Left            =   7800
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblPile1 
      Caption         =   "Completed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9960
      TabIndex        =   10
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label lblPile1 
      Caption         =   "Defeated:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9960
      TabIndex        =   9
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   10920
      TabIndex        =   8
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   11040
      TabIndex        =   7
      Top             =   9000
      Width           =   255
   End
   Begin VB.Image imgPlaced4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPlaced3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPlaced2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPlaced1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgTemp 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Index           =   0
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   11880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   5
      Top             =   9000
      Width           =   375
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label lblPile1 
      Caption         =   "Dead:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   2
      Top             =   9000
      Width           =   495
   End
   Begin VB.Label lblPile1 
      Caption         =   "Discard: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblPile1 
      Caption         =   "Draw:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   8520
      Width           =   495
   End
   Begin VB.Image imgCompletedMissions 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1365
   End
   Begin VB.Image imgMissions 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   11640
      Picture         =   "frmTable.frx":BF9DA
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1365
   End
   Begin VB.Image imgDeadMissions 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   1365
   End
   Begin VB.Image ImgReserve 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Image imgFrontLine 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   2
      Left            =   8040
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Image imgFrontLine 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   1
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Image imgFrontLine 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Index           =   0
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000040&
      FillStyle       =   4  'Upward Diagonal
      Height          =   975
      Index           =   6
      Left            =   11640
      Top             =   7920
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   7
      Left            =   4440
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   8
      Left            =   6240
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   9
      Left            =   8040
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00004000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   10
      Left            =   6240
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label lblPile 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   4
      Top             =   8760
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   11
      Left            =   7800
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   12
      Left            =   9600
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   13
      Left            =   11400
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1215
      Index           =   14
      Left            =   9600
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image imgVenture 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   11880
      Stretch         =   -1  'True
      Top             =   7920
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Homebase"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8265
      TabIndex        =   11
      Top             =   8925
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1215
      Index           =   1
      Left            =   8040
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   3
      Left            =   11640
      Top             =   6840
      Width           =   1365
   End
   Begin VB.Image imgVentureC 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   11880
      Stretch         =   -1  'True
      Top             =   6840
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   5
      Left            =   11640
      Top             =   9000
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1215
      Index           =   3
      Left            =   7800
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1215
      Index           =   2
      Left            =   6000
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image imgPR1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPR2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPR3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPRBS 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOPHit2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgOPHit1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPROPBS 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPROP4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPROP1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPROP2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPROP3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenDeck 
         Caption         =   "Open Deck"
      End
      Begin VB.Menu mnuFileCap 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDeck 
      Caption         =   "&Me"
      Begin VB.Menu mnuShowDuplicates 
         Caption         =   "Check for Duplicates"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnusepHand1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEndTurn 
         Caption         =   "End Turn"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu mnucap4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBattleSiteDEck 
         Caption         =   "View Battlesite Deck"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuCapo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDrawPile 
         Caption         =   "View Draw Pile"
      End
      Begin VB.Menu mnuViewDiscardPile 
         Caption         =   "View Power Pack"
      End
      Begin VB.Menu mnuViewDeadPile 
         Caption         =   "View Dead Pile"
      End
      Begin VB.Menu mnuViewDefeatedCharacters 
         Caption         =   "View Defeated Characters Pile"
      End
      Begin VB.Menu mnucap34 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPassToOpponent 
         Caption         =   "Pass to opponent"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConcedeToOpponent 
         Caption         =   "Concede to opponent"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCapView 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawCard 
         Caption         =   "Draw Card"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucap2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayOpenHanded 
         Caption         =   "Play Open-Handed"
      End
      Begin VB.Menu mnucap7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShuffleDiscardsIntoDraw 
         Caption         =   "Shuffle Power Pack into Draw Pile"
      End
      Begin VB.Menu mnuShuffleDeadIntoDraw 
         Caption         =   "Shuffle Dead Pile into Draw Pile"
      End
   End
   Begin VB.Menu mnuOpponent 
      Caption         =   "&Opponent"
      Begin VB.Menu mnuToolsConnect 
         Caption         =   "Connect to Opponent"
      End
      Begin VB.Menu mnucappass 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOppViewHand 
         Caption         =   "*View Hand"
      End
      Begin VB.Menu mnuOppViewBattleSiteDeck 
         Caption         =   "*View Battlesite Deck"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOppViewDrawPile 
         Caption         =   "*View Draw Pile"
      End
      Begin VB.Menu mnuViewOpPowerPack 
         Caption         =   "*View Power Pack"
      End
      Begin VB.Menu mnuViewOPDeadPile 
         Caption         =   "*View Dead Pile"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsWhoGoesFirst 
         Caption         =   "Who Goes First?"
      End
      Begin VB.Menu mnucap 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuDrawTestHand 
         Caption         =   "Draw Test Hand"
      End
      Begin VB.Menu mnuCapDE 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeckEditor 
         Caption         =   "Deck Editor"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Overpower Online Help"
      End
      Begin VB.Menu mnuHelpRules 
         Caption         =   "Overpower Rules"
      End
      Begin VB.Menu mnucapHelp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cCurMenuIndex As Integer
Dim bSaveVisible As Boolean


Private Sub chkMoveCompletedD_Click()

MoveVentureCards cDeadMissions, cCompletedMissions, CBool(chkMoveAllD.Value), imgDeadMissions, "MISSIONS: DEFEATED"

End Sub

Private Sub chkPlayFaceDown_Click()

If chkPlayFaceDown.Value = 1 Then
    myattack.Attack_isFaceDown = True
Else
    myattack.Attack_isFaceDown = False
End If

End Sub

Private Sub chkReturnMissionD_Click()

MoveVentureCards cDeadMissions, cMissions, CBool(chkMoveAllD.Value), imgDeadMissions, "MISSIONS: DEFEATED"


End Sub

Private Sub cmdAcceptDefense_Click()
Dim ccard

frmAttack.Visible = False
frmDefense.Visible = False

For i = 0 To lnFrontLine.Count - 1
    lnFrontLine(i).Visible = False
Next i

If bOpponentConceded = True Then
    Set ccard = myattack.GetCard(1)
    
    If ccard.CardType = "Special Card" Then
        If ccard.Attack_StopsConcede = True Then
            X = MsgBox("Opponent has blocked your attempt to stop him from conceding the battle.", vbInformation, "Opponent Concedes Battle")
        End If
    End If
    
End If

SendData "CDA:1:|"

For i = 1 To myattack.Card_Count
    Set ccard = myattack.GetCard(i)
    
    Select Case ccard.CardType
    
    Case "Power Card"
        cDiscardPile.Add ccard
    
    Case Else
    
        cDeadPile.Add ccard
        
    End Select
Next i

For i = 1 To cIncomingDefense.Count
    Set ccard = cIncomingDefense.Item(i)
    
    Select Case ccard.CardType
    Case "Power Card"
        cDiscardPileO.Add ccard
    Case Else
        cDeadPileO.Add ccard
    End Select
Next i

myattack.NewAttack
    
Set cIncomingDefense = New Collection

History_Add mySettings.PlayerName & ": DEFENSE ACCEPTED"
UpdateOpponentDeckDisplay
UpdateDeckDisplay

If bOpponentConceded = True Then
    ResolveConcession False
End If

CheckForAdditionalAttack

End Sub

Private Sub cmdAcceptOppVenTotal_Click()

cmdAcceptOppVenTotal.Tag = "1"
SendData "CVA:1:|"

'set tag on button to indicate that the total has been accepted

If chkVTTotalAccepted.Value = 1 Then
    'both totals have been accepted
    ShowMoveVentureFrame

End If

End Sub

Private Sub cmdAttack_Click()

cmdAttack.Enabled = False

PlayCard

cmdAttack.Enabled = True

End Sub
Private Sub cmdCancelAction_Click()

For i = 1 To myattack.Card_Count

Select Case myattack.Card_Source(i)

Case "P1"
    SendData "CSC:16:12:1:|"
Case "P2"
    SendData "CSC:16:13:1:|"
Case "P3"
    SendData "CSC:16:14:1:|"
Case "P4"
    SendData "CSC:16:15:1:|"
Case "H"
    SendData "CSC:16:2:1:|"
Case "P5"
    SendData "CSC:16:18:1:|"
Case Else
End Select

myattack.RemoveCard 1
Next i

myattack.NewAttack

For i = 0 To 8
    lnFrontLine(i).Visible = False
Next i

shpAction.Visible = False
frmAttack.Visible = False

ShowPlacedCards

FetchHandImages
ShowHand

HeroClick 0

End Sub

Private Sub cmdChallengeDefense_Click()

X = MsgBox("Do you wish to dispute your opponent's proposed defense?", vbYesNo, "Challenge Defense?")

If X <> 6 Then Exit Sub

a$ = InputBox$("Enter your challenge:", "Challenge Opponent's Defense", "You Suck!")

If a$ = "" Then Exit Sub

SendData "CCH:" & a$ & ":|"

cmdAcceptDefense.Enabled = False
cmdChallengeDefense.Enabled = False

If sSounds(17) <> "" Then PlaySound sSounds(17)

History_Add mySettings.PlayerName & " CHALLENGES-->" & a$
frmDefense.Visible = False


End Sub

Private Sub cmdCompleted_Click()

'If chkMoveAll.Value = 1 Then
'    senddata "CSC:9:8:" & Trim(Str(cVenturedMissions.Count)) & ":|"
'Else
'    senddata "CSC:9:8:1:|"
'End If

MoveVentureCards cVenturedMissions, cCompletedMissions, CBool(chkMoveAll.Value), imgVenture, "MISSION: VENTURED CARDS"

End Sub

Private Sub cmdDefeated_Click()
'If chkMoveAll.Value = 1 Then
'    senddata "CSC:9:7:" & Trim(Str(cVenturedMissions.Count)) & ":|"
'Else
'    senddata "CSC:9:7:1:|"
'End If

MoveVentureCards cVenturedMissions, cDeadMissions, CBool(chkMoveAll.Value), imgVenture, "MISSION: VENTURED CARDS"

End Sub

Private Sub cmdDefeatedC_Click()

'senddata "CSC:8:7:" & Trim(Str(cCompletedMissions.Count)) & ":|"

MoveVentureCards cCompletedMissions, cDeadMissions, False, imgCompletedMissions, "MISSIONS: COMPLETED"

End Sub

Private Sub cmdDiscard_Click()
On Error Resume Next

cmdDiscard.Enabled = False

Index = Val(imgCardDetail.Tag)

Select Case cHand.Item(Index).CardType
Case "Power Card"

If sSounds(15) <> "" Then PlaySound sSounds(15)

History_Add "DISCARD: " & cHand.Item(Index).Title

cDiscardPile.Add cHand.Item(Index)
cHand.Remove Index
cHandTags.Remove Index

SendData "CSC:2:3:" & Trim(Str(Index)) & ":|"

Case Else

If sSounds(15) <> "" Then PlaySound sSounds(15)
History_Add "DISCARD: " & cHand.Item(Index).Title

cDeadPile.Add cHand.Item(Index)
cHand.Remove Index
cHandTags.Remove Index

SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"

End Select

imgCardDetail.Picture = LoadPicture(sBlankImagePath)
imgCardDetail.Tag = ""


ShowHand
UpdateDeckDisplay
cmdDiscard.Enabled = True

End Sub

Private Sub cmdDiscardActivator_Click()
Index = Val(imgHeroCard.Tag)

SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"

cDeadPile.Add cHand.Item(Index)
History_Add "DISCARD: " & cHand.Item(Index).Title
If sSounds(15) <> "" Then PlaySound sSounds(15)
cHand.Remove Index
cHandTags.Remove Index

ShowHand
UpdateDeckDisplay

End Sub


Private Sub DiscardAspect()
If frmHandCard.Tag = "H" Then

    Index = Val(imgCardDetail.Tag)
    
    cDeadPile.Add cHand.Item(Index)
    History_Add "DISCARD: " & cHand.Item(Index).Title
    
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    
    SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"
    
    cHand.Remove Index
    cHandTags.Remove Index
    ShowHand

Else
    Dim ccard
    
    Index = Val(imgCardDetail.Tag)
    Set ccard = myHomebase.PlacedCard(Index)
    History_Add "DISCARD: " & ccard.Title
    
    myHomebase.RemovePlacedCard Index
    cDeadPile.Add ccard
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    SendData "CHR:" & Trim(Str(Index)) & ":|"
    
    ShowPlacedCards
    
End If
End Sub
Private Sub cmdDiscardEvent_Click()
Index = Val(imgHeroCard.Tag)

cDeadPile.Add cHand.Item(Index)

SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"
History_Add "DISCARD: " & cHand.Item(Index).Title

cHand.Remove Index
If sSounds(15) <> "" Then PlaySound sSounds(15)
cHandTags.Remove Index
ShowHand

End Sub

Private Sub cmdDiscardEvent1_Click()

Index = Val(imgHeroCard.Tag)

cDeadPile.Add cHand.Item(Index)

SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"

History_Add "DISCARD: " & cHand.Item(Index).Title
cHand.Remove Index
If sSounds(15) <> "" Then PlaySound sSounds(15)
cHandTags.Remove Index

DrawCard

End Sub

Private Sub cmdDiscardModifier_Click()
Index = Val(imgCardDetail.Tag)
cindex = Val(frmModifier.Tag)
Dim ccard

If cindex = -1 Then
    'being removed from hand
    Set ccard = cHand.Item(Index)
    
    History_Add "DISCARD: " & ccard.Title
    cDeadPile.Add ccard
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    cHand.Remove Index
    SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"
    UpdateDeckDisplay
    FetchHandImages
    ShowHand
    Exit Sub
End If
    
If cmdDiscardModifier.Tag = "BUFFER" Then
    
    nId = cindex + 21
    
    Set ccard = cFrontLine.Buffers_GetCard(cindex, Index)
    History_Add "DISCARD: " & ccard.Title
    
    SendData "CSC:" & Trim(Str(nId)) & ":4:" & Trim(Str(Index)) & ":|"
    
    cDeadPile.Add ccard
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    cFrontLine.Buffers_RemoveCard cindex, Index
    
    imgCardDetail.Picture = LoadPicture(sBlankImagePath)
    imgCardDetail.Tag = ""
    frmModifier.Tag = ""
    Me.Caption = "OVERPOWER ONLINE-->" & ""
    HideFrames False, False, True
    
    ShowBuffers
Else


    nId = cindex + 17
    
    Set ccard = cFrontLine.Modifiers_GetCard(cindex, Index)
    History_Add "DISCARD: " & ccard.Title
    
    SendData "CSC:" & Trim(Str(nId)) & ":4:" & Trim(Str(Index)) & ":|"
    
    cDeadPile.Add ccard
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    cFrontLine.Modifiers_RemoveCard cindex, Index
    
    imgCardDetail.Picture = LoadPicture(sBlankImagePath)
    imgCardDetail.Tag = ""
    frmModifier.Tag = ""
    Me.Caption = "OVERPOWER ONLINE-->" & ""
    HideFrames False, False, True
    
    ShowModifiers

End If

cmdDiscardModifier.Tag = ""

UpdateDeckDisplay
End Sub

Private Sub cmdDiscardPlaced_Click()
On Error Resume Next

Index = Val(imgCardDetail.Tag)
cindex = Val(frmPlaced.Tag)
Dim ccard

If cindex < 5 Then
    nId = cindex + 11
Else
    nId = 18
End If

If cindex < 5 Then

Set ccard = cFrontLine.PlacedCard(cindex, Index)

Select Case ccard.CardType

Case "Power Card"

    SendData "CSC:" & Trim(Str(nId)) & ":3:" & Trim(Str(Index)) & ":|"
    
    cDiscardPile.Add ccard
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    History_Add "DISCARD: " & ccard.Title
    cFrontLine.RemovePlacedCard cindex, Index
    
Case Else

    SendData "CSC:" & Trim(Str(nId)) & ":4:" & Trim(Str(Index)) & ":|"
    
    cDeadPile.Add ccard
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    History_Add "DISCARD: " & ccard.Title
    cFrontLine.RemovePlacedCard cindex, Index

End Select

Else

    Set ccard = myHomebase.PlacedCard(Index)
    
    Select Case ccard.CardType
    
    Case "Power Card"
    cDiscardPile.Add ccard
    myHomebase.RemovePlacedCard Index
    
    SendData "CSC:26:3:" & Trim(Str(Index)) & ":|"

    Case Else
    cDeadPile.Add ccard
    myHomebase.RemovePlacedCard Index
    
    SendData "CSC:26:4:" & Trim(Str(Index)) & ":|"

    End Select
    
    If sSounds(15) <> "" Then PlaySound sSounds(15)
    History_Add "DISCARD: " & ccard.Title

End If


imgCardDetail.Picture = LoadPicture(sBlankImagePath)
imgCardDetail.Tag = ""
frmPlaced.Tag = ""
Me.Caption = "OVERPOWER ONLINE-->" & ""
HideFrames False, False, True


ShowPlacedCards
UpdateDeckDisplay
End Sub



Private Sub cmdExchangeActivator_Click()
Dim ctemp As Collection
Dim ccard

If myBattleSite.Deck_Count = 0 Then
    MsgBox "You do not currently have any cards in your Battlesite deck.", vbCritical, "No Cards"
    Exit Sub
End If

With FrmViewPile

Set ctemp = New Collection
For i = 1 To myBattleSite.Deck_Count
Set ccard = myBattleSite.Deck_GetCard(i)
ctemp.Add ccard
Next i

Set .ShowPile = ctemp
.PileType = 3
.Show 1

If .AddedToMyHand = True Then
    Index = Val(imgHeroCard.Tag)
    cDeadPile.Add cHand.Item(Index)
    cHand.Remove Index
    cHandTags.Remove Index
    FetchHandImages
    
    History_Add "ACTIVATOR EXCHANGED FOR SPECIAL"
    SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"
    
End If

End With
Unload FrmViewPile

ShowHand
UpdateDeckDisplay
End Sub

Private Sub cmdFinishedDiscarding_Click()

cmdFinishedDiscarding.Enabled = False
Me.lblDiscard1(0).Caption = mySettings.PlayerName & ": Done"

SendData "CFD:1:|"

If lblDiscard1(1).Caption = sOpponentName & ": Done" Then
    frmDiscardPhase.Visible = False
    myPhase = nPhase_Place
    UpdatePhase
    ShowPlacingFrame
End If

End Sub

Private Sub cmdFinishedMovingVenture_Click()

lblMoveVenture(0).Caption = mySettings.PlayerName & ": Finished"
cmdFinishedMovingVenture.Enabled = False

SendData "CM1:1:|"

If lblMoveVenture(1).Caption = sOpponentName & ": Finished" Then
    'Other player has finished moving venture cards
        
    CheckMissionMessages
    
    EndTurn
  
End If

End Sub
Private Sub CheckMissionMessages()
    If cCompletedMissions.Count = 7 Then
        X = MsgBox(mySettings.PlayerName & " has completed his mission!  You may continue playing if desired.", vbInformation, mySettings.PlayerName & " completes mission")
    End If
    
    If cCompletedMissionsO.Count = 7 Then
        X = MsgBox(sOpponentName & " has completed his mission!  You may continue playing if desired.", vbInformation, sOpponentName & " completes mission")
    End If
    
    If cDeadMissions.Count = 7 Then
        X = MsgBox(mySettings.PlayerName & " has lost his mission!  You may continue playing if desired.", vbInformation, mySettings.PlayerName & " loses mission")
    End If
    
    If cDeadMissionsO.Count = 7 Then
        X = MsgBox(sOpponentName & " has lost his mission!  You may continue playing if desired.", vbInformation, sOpponentName & " loses mission")
    End If
    
End Sub
Private Sub cmdFinishedPlacing_Click()
cmdFinishedPlacing.Enabled = False
lblDiscard1(2).Caption = mySettings.PlayerName & ": Done"

SendData "CFP:1:|"

If lblDiscard1(3).Caption = sOpponentName & ": Done" Then
    frmPlacingPhase.Visible = False
    myPhase = nPhase_Venture
    UpdatePhase
    ShowVentureFrame
    
End If
End Sub

Private Sub cmdFinishedVenture_Click()
cmdFinishedVenture.Enabled = False
lblDiscard1(4).Caption = mySettings.PlayerName & ": Done"

SendData "CFV:1:|"

If lblDiscard1(5).Caption = sOpponentName & ": Done" Then
    frmVenturePhase.Visible = False
    
    If bIGoFirst = True Then
        myPhase = nPhase_Attack
        History_Add mySettings.PlayerName & " PREPARE YOUR ATTACK"
        
        If sSounds(18) <> "" Then PlaySound sSounds(18)
    Else
        myPhase = nPhase_Defend
        History_Add sOpponentName & " IS PREPARING ATTACK"
        
    End If
    
    bHavePassed = False
    bOppPassed = False
    
    UpdatePhase
    
End If
End Sub

Private Sub cmdHTCBMove_Click()
hindex = Val(frmHTCB.Tag)
cindex = Val(imgCardDetail.Tag)

If hindex = 0 Or cindex = 0 Then Exit Sub

Load frmChooseCharacter

If myBattleSite.ID > 0 Then
With frmChooseCharacter
    .optCharacter(4).Caption = "BATTLESITE: " & myBattleSite.Name
    .optCharacter(4).Visible = True
    .optCharacter(4).Tag = 5
End With

Else
frmChooseCharacter.optCharacter(4).Visible = False

End If

frmChooseCharacter.Show 1

If frmChooseCharacter.SelectedCharacter = -1 Then
    Unload frmChooseCharacter
    Exit Sub
End If

b = frmChooseCharacter.SelectedCharacter
Unload frmChooseCharacter

If hindex < 5 Then
    Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(hindex, cindex)
Else
    Set ccard = myBattleSite.HitsToCurrentBattle_GetCard(cindex)
End If

If b < 5 Then
    cFrontLine.HitsToCurrentBattle_AddCard b, ccard
Else
    myBattleSite.HitsToCurrentBattle_AddCard ccard
End If

If hindex < 5 Then
    cFrontLine.HitsToCurrentBattle_RemoveCard hindex, cindex, True
Else
    myBattleSite.HitsToCurrentBattle_RemoveCard cindex, True
End If

SendData "CMH:" & Trim(Str(hindex)) & ":" & Trim(Str(b)) & ":" & Trim(Str(cindex)) & ":|"

ShowHitsToCurrentBattle
UpdateOpponentDeckDisplay
HideAllBorders
HeroClick 0

End Sub
Private Sub cmdHTCBToPR_Click()
Dim ccard

hindex = Val(frmHTCB.Tag)
cindex = Val(imgCardDetail.Tag)

If hindex = 0 Or cindex = 0 Then Exit Sub

If hindex < 5 Then
    Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(hindex, cindex)
    cFrontLine.PermanentRecord_AddCard hindex, ccard
    cFrontLine.HitsToCurrentBattle_RemoveCard hindex, cindex, True
Else
    Set ccard = myBattleSite.HitsToCurrentBattle_GetCard(cindex)
    myBattleSite.PermanentRecord_AddCard ccard
    myBattleSite.HitsToCurrentBattle_RemoveCard cindex, True
    
End If

SendData "CPH:" & Trim(Str(hindex)) & ":" & Trim(Str(cindex)) & ":|"

ShowHitsToCurrentBattle
ShowPermanentRecord
ShowPlacedCards
UpdateOpponentDeckDisplay
HeroClick 0

End Sub
Private Sub cmdKOBattlesite_Click()
'discard battlesite deck cards

X = MsgBox("Are you sure you want to K.O. your Battlesite?", vbYesNoCancel, "KO Battlesite?")

If X <> 6 Then Exit Sub

lblKO(4).Visible = True
cmdKOBattlesite.Enabled = False
cmdViewBattlesiteDeck.Enabled = False

SendData "CBK:1:|"

End Sub
Private Sub cmdKOCharacter_Click()
a = Val(imgHeroCard.Tag)
If a = 0 Then Exit Sub

X = MsgBox("Are you sure you want to K.O. " & cFrontLine.Character_Name(a) & "?", vbYesNoCancel, "K.O. Character")
If X <> 6 Then Exit Sub

If cFrontLine.LiveCharacterCount = 1 Then
    'all characters have been KO'd
    X = MsgBox("All of your characters have been KO'd.  " & sOpponentName & " wins!", vbCritical, "You lose!")
End If

If sSounds(9) <> "" Then PlaySound sSounds(9)

For i = 0 To 2
    If imgFrontLine(i).Tag = imgHeroCard.Tag Then
        lblKO(i).Visible = True
        lblKO(i).ZOrder 0
    End If
Next i

If ImgReserve.Tag = imgHeroCard.Tag Then
    lblKO(3).Visible = True
    lblKO(3).ZOrder 0
End If

SendData "CKO:" & Trim(Str(a)) & ":|"

cmdTakeAction.Enabled = False
cmdKOCharacter.Enabled = False
cmdSwitchWithReserve.Enabled = False
cmdReserveToFrontline.Enabled = False

End Sub

Private Sub cmdNoDefense_Click()

If sSounds(8) <> "" Then PlaySound sSounds(8)

frmDefense.Visible = False
frmAttack.Visible = False

For i = 0 To lnFrontLine.Count - 1
    lnFrontLine(i).Visible = False
Next i

SendData "CND:1:|"

nId = OpAttack.DefenderID

If cFrontLine.Buffers_Count(nId) > 0 Then
    
    Set ccard = cFrontLine.Buffers_GetCard(nId, 1)
    
    cFrontLine.Buffers_RemoveCard nId, 1
    ShowBuffers
    
    For i = 1 To cIncomingAttack.Count
    
        cFrontLine.BufferHits_AddCard cIncomingAttack.Item(i)
        
    Next i
    
    ShowBufferHits
    
    History_Add ccard.Title & " KO'D"
    
    ShowVentureTotals
    
Else

    For i = 1 To cIncomingAttack.Count
    
    If OpAttack.DefenderID = 5 Then
        myBattleSite.HitsToCurrentBattle_AddCard cIncomingAttack.Item(i)
    Else
        cFrontLine.HitsToCurrentBattle_AddCard OpAttack.DefenderID, cIncomingAttack.Item(i)
    End If
    
    Next i

    ShowVentureTotals
End If
    

If bIHaveConceded = True Then

    For i = 1 To cIncomingAttack.Count
        Set ccard = cIncomingAttack.Item(i)
        
        If ccard.CardType = "Special Card" Then
            If ccard.Attack_PostConcessionAttack = True Then
                OpAttack.NewAttack
                ShowHitsToCurrentBattle
                UpdateOpponentDeckDisplay
                ShowVentureTotals
                ResolveConcession True
                Exit Sub
            End If
            
            If ccard.Attack_StopsConcede = True Then
            
                X = MsgBox("Opponent has stopped you from conceding.  Play will continue as normal.", vbInformation, "Concession Stopped")
                bIHaveConceded = False
                
            End If
            
        End If

    Next i
End If

OpAttack.NewAttack
Set cIncomingAttack = New Collection

ShowHitsToCurrentBattle
UpdateOpponentDeckDisplay

End Sub

Private Sub cmdOKAction_Click()

If myattack.Card_Count = 0 Then
    X = MsgBox("You need to have at least one card in your attack/action.", vbCritical, "No Cards Selected")
    Exit Sub
End If


'Check to see if opponent has conceded, and if so, if there is a concession effect card in attack
If bOpponentConceded = True Then
    For i = 1 To myattack.Card_Count
    Set ccard = myattack.GetCard(i)

    If ccard.CardType = "Special Card" Then
        If ccard.Attack_PostConcessionAttack = True Or ccard.Attack_StopsConcede = True Then
            GoTo lpContinue
        End If
    End If
    Next i

    X = MsgBox("Your opponent has conceded.  At this point you may only play a Special that stops the concession, or a Special that allows an attack after your opponent has conceded.", vbCritical, "Invalid Action")
    Exit Sub

End If

lpContinue:

If myattack.DefenderID = 0 Then
    X = MsgBox("Please select a target for this attack.", vbCritical, "No Target Selected")
    Exit Sub
End If

If sSounds(4) <> "" Then PlaySound sSounds(4)

Set ccard = myattack.GetCard(1)

If ccard.Attack_isStringAttack = True Then
    imgStringAttack.Picture = imgAction(0).Picture
    imgStringAttack.ToolTipText = ccard.Description
    ShowStringAttackFrame
End If

If chkPlayFaceDown.Value = 1 Then
    SendData "CAT:" & Trim(Str(myattack.AttackerID)) & ":" & Trim(Str(myattack.DefenderID)) & ":1:|"
Else
    SendData "CAT:" & Trim(Str(myattack.AttackerID)) & ":" & Trim(Str(myattack.DefenderID)) & ":0:|"
End If

cmdOKAction.Enabled = False
cmdCancelAction.Enabled = False

History_Add "===================================================================="
History_Add cFrontLine.Character_Name(myattack.AttackerID) & " ATTACKS " & cOpponent.Character_Name(myattack.DefenderID)
History_Add "-------------------------------------------------------------------------------------------------"

For i = 1 To myattack.Card_Count
    Set ccard = myattack.GetCard(i)
    History_Add Trim(Str(i)) & ". " & ccard.Title
Next i

chkPlayFaceDown.Enabled = False

History_Add "===================================================================="


History_Add "AWAITING RESPONSE TO ATTACK..."

HeroClick 0
End Sub

Private Sub cmdOKDefense_Click()

If myDefense.Card_Count = 0 Then
    X = MsgBox("Are you sure you want to submit this defense?  It currently has no cards.", vbYesNoCancel, "Submit Defense?")
    If X <> 6 Then Exit Sub
End If

History_Add "===================================================================="
History_Add "DEFENSE"
History_Add "-------------------------------------------------------------------------------------------------"

For i = 1 To myDefense.Card_Count
    Set ccard = myDefense.GetCard(i)
    History_Add Trim(Str(i)) & ". " & ccard.Title
Next i

History_Add "===================================================================="

SendData "CDF:1:|"

cmdOKDefense.Enabled = False
cmdNoDefense.Enabled = False

End Sub

Private Sub cmdPlaceModifier_Click()

PlaceCard
End Sub

Private Sub cmdPlayAspect_Click()


End Sub

Private Sub cmdPlayEvent_Click()

PlayEvent

End Sub
Private Sub ShowOpponentEvent(nId)
Dim ccard

If imgEffect(0).Visible = True Then ec = 1
If imgEffect(1).Visible = True Then ec = 2
If imgEffect(1).Visible = False And imgEffect(0).Visible = False Then ec = 0

If ec = 2 Then
    Exit Sub
End If

Set ccard = cHandO.Item(nId)

If ccard.LoadImage(cHandO.Item(nId).ID) = True Then
    imgEffect(ec).Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgEffect(ec).Picture = LoadPicture(sBlankImagePath)
End If

imgEffect(ec).Visible = True
imgEffect(ec).Tag = sOpponentName
lblEvent(ec).Caption = cHandO.Item(nId).Description

X = MsgBox(sOpponentName & " has played an Event:" & vbCrLf & ccard.Description, vbOKOnly, ccard.Title)

cDeadPileO.Add cHandO.Item(nId)
cHandO.Remove nId

UpdateOpponentDeckDisplay

End Sub
Private Sub PlayEvent()
If imgEffect(0).Visible = True Then ec = 1
If imgEffect(1).Visible = True Then ec = 2
If imgEffect(1).Visible = False And imgEffect(0).Visible = False Then ec = 0

If ec = 2 Then
    X = MsgBox("There are already 2 events in play.  Event cannot be played.", vbCritical, "Event Cannot Be Played")
    Exit Sub
End If

If ec = 1 Then
    imgEffect(1).Picture = imgHeroCard.Picture
    imgEffect(1).Visible = True
    imgEffect(1).Tag = mySettings.PlayerName
    lblEvent(1).Caption = cHand.Item(Val(imgHeroCard.Tag)).Description
End If

If ec = 0 Then
    imgEffect(0).Picture = imgHeroCard.Picture
    imgEffect(0).Visible = True
    imgEffect(0).Tag = mySettings.PlayerName
    lblEvent(0).Caption = cHand.Item(Val(imgHeroCard.Tag)).Description
End If

Index = Val(imgHeroCard.Tag)

cDeadPile.Add cHand.Item(Index)

SendData "CPE:" & Trim(Str(Index)) & ":|"

cHand.Remove Index
cHandTags.Remove Index
UpdateDeckDisplay

DrawCard

End Sub
Private Sub CheckModifiers()
'Check to see if battleeffect modifiers need to be removed at end of battle
Dim ccard

Checkagain:

For i = 1 To 4

For k = 1 To cFrontLine.Modifiers_Count(i)

If cFrontLine.Modifiers_Type(i, k) = 1 Then
    
    Set ccard = cFrontLine.Modifiers_GetCard(i, k)
    cDeadPile.Add ccard
    cFrontLine.Modifiers_RemoveCard i, k
    ShowModifiers
    GoTo Checkagain
End If

Next k

Next i

UpdateDeckDisplay

End Sub
Private Sub cmdPlayModifier_Click()

a = Val(imgCardDetail.Tag)
If a = 0 Then Exit Sub

Set ccard = cHand.Item(a)

If ccard.CardType = "Special Card" Then frmChooseCharacter.SpecialCharacter = ccard.Character

frmChooseCharacter.Show 1

If frmChooseCharacter.SelectedCharacter = -1 Then
    Unload frmChooseCharacter
    Exit Sub
End If

b = frmChooseCharacter.SelectedCharacter
Unload frmChooseCharacter

Set ccard = cHand.Item(a)

If ccard.Attack_Frontline_BattleBonus = True Then cFrontLine.Modifiers_AddCard b, ccard, modifies_battle
If ccard.Attack_Frontline_GameBonus = True Then cFrontLine.Modifiers_AddCard b, ccard, Modifies_Game

If ccard.Attack_Frontline_Allies = True Then
    cFrontLine.Buffers_AddCard b, ccard
    History_Add "==============================================================================================================="
    History_Add mySettings.PlayerName & ": PLAYS BUFFER TO " & cFrontLine.Character_Name(b)
    History_Add "---------------------------------------------------------------------------------------------------------------"
    History_Add ccard.Title
    History_Add "==============================================================================================================="
ElseIf ccard.CardType = "Artifact" Then
    History_Add "==============================================================================================================="
    History_Add mySettings.PlayerName & ": PLAYS ARTIFACT TO " & cFrontLine.Character_Name(b)
    History_Add "---------------------------------------------------------------------------------------------------------------"
    History_Add ccard.Title
    History_Add "==============================================================================================================="
Else

    History_Add "==============================================================================================================="
    History_Add mySettings.PlayerName & ": PLAYS MODIFIER [" & cFrontLine.Modifiers_TypeText(b, cFrontLine.Modifiers_Count(b)) & "] TO " & cFrontLine.Character_Name(b)
    History_Add "---------------------------------------------------------------------------------------------------------------"
    History_Add ccard.Title
    History_Add "==============================================================================================================="
End If

cHand.Remove a
cHandTags.Remove a

If ccard.Attack_Frontline_Allies = True Then
    SendData "CSC:2:" & Trim(Str(b + 21)) & ":" & Trim(Str(a)) & ":|"
    ShowBuffers
Else
    SendData "CSC:2:" & Trim(Str(b + 17)) & ":" & Trim(Str(a)) & ":|"
    ShowModifiers
End If

ShowHand
HideAllBorders

If myPhase <> nPhase_Defend Then CheckForAdditionalAttack

End Sub

Private Sub cmdPlayPlaced_Click()

cmdPlayPlaced.Enabled = False

playplaced

cmdPlayPlaced.Enabled = True


End Sub
Private Sub playplaced()
Dim ccard

If myPhase = nPhase_Defend Then
    
    If cmdOKDefense.Enabled = False Then
        X = MsgBox("Cards cannot be added to a defense once it has been submitted to your opponent.  Message your opponent and ask him to challenge your defense.", vbCritical, "Cannot add Defense Card")
        Exit Sub
    End If
    
    Index = Val(imgCardDetail.Tag)
    cindex = Val(frmPlaced.Tag)
    
    If cindex < 5 Then
        
        Set ccard = cFrontLine.PlacedCard(cindex, Index)
        
        myDefense.AddCard ccard, "P" & Trim(Str(cindex)), Index
        SendData "CSC:" & Trim(Str((11 + cindex))) & ":17:" & Trim(Str(Index)) & ":|"
        ShowPlacedCards
        
    Else
    
        Set ccard = myHomebase.PlacedCard(Index)
        myDefense.AddCard ccard, "P5", Index
        
        SendData "CSC:26:17:" & Trim(Str(Index)) & ":|"
        
    End If
    
    ShowDefenseCards
    
Else

    Index = Val(imgCardDetail.Tag)
    cindex = Val(frmPlaced.Tag)
    
    'Check to see if this is a battle effect
    If cindex < 5 Then
    
        Set ccard = cFrontLine.PlacedCard(cindex, Index)
        
        If ccard.CardType = "Special Card" Then
        
            If ccard.Attack_EffectsFrontline = True Then
                
                
                'hide attack frame if visible
                For i = 0 To 8
                    lnFrontLine(i).Visible = False
                Next i
                
                shpAction.Visible = False
                frmAttack.Visible = False
                
                If ccard.Attack_Frontline_BattleBonus = True Then
                    cFrontLine.Modifiers_AddCard cindex, ccard, modifies_battle
                    SendData "CSC:" & Trim(Str(cindex + 11)) & ":" & Trim(Str(cindex + 17)) & ":" & Trim(Str(Index)) & ":|"
                End If
                
                If ccard.Attack_Frontline_GameBonus = True Then
                    cFrontLine.Modifiers_AddCard cindex, ccard, Modifies_Game
                    SendData "CSC:" & Trim(Str(cindex + 11)) & ":" & Trim(Str(cindex + 17)) & ":" & Trim(Str(Index)) & ":|"
                End If
                
                If ccard.Attack_Frontline_Allies = True Then
                    cFrontLine.Buffers_AddCard cindex, ccard
                    SendData "CSC:" & Trim(Str(cindex + 11)) & ":" & Trim(Str(cindex + 21)) & ":" & Trim(Str(Index)) & ":|"
                End If
                
                cFrontLine.RemovePlacedCard cindex, Index
                
                       
                ShowPlacedCards
                ShowBuffers
                ShowModifiers
                History_Add "MODIFIER PLAYED TO: " & cFrontLine.Character_Name(cindex)
                
                CheckForAdditionalAttack
                
                Exit Sub
    
            End If
            
        End If
        
    
    End If
    
    If myattack.AttackerID < 1 Then
        X = MsgBox("Please start an Action with a character before playing cards.", vbCritical, "No Current Action")
        Exit Sub
    End If
    
    
    If cindex < 5 Then
    
        Set ccard = cFrontLine.PlacedCard(cindex, Index)
        
        SendData "CSC:" & Trim(Str((11 + cindex))) & ":16:" & Trim(Str(Index)) & ":|"
        
        myattack.AddCard ccard, "P" & Trim(Str(cindex)), Index
        
    Else
        
        Set ccard = myHomebase.PlacedCard(Index)
        myattack.AddCard ccard, "P5", Index
        SendData "CSC:26:16:" & Trim(Str(Index)) & ":|"
        
    End If

    ShowAttackCards
    
End If


ShowPlacedCards

End Sub
Private Sub cmdPRMove_Click()
hindex = Val(frmPR.Tag)
cindex = Val(imgCardDetail.Tag)

If hindex = 0 Or cindex = 0 Then Exit Sub

Load frmChooseCharacter

If myBattleSite.ID > 0 Then
With frmChooseCharacter
    .optCharacter(4).Caption = "BATTLESITE: " & myBattleSite.Name
    .optCharacter(4).Visible = True
    .optCharacter(4).Tag = 5
End With
Else
frmChooseCharacter.optCharacter(4).Visible = False

End If

frmChooseCharacter.Show 1

If frmChooseCharacter.SelectedCharacter = -1 Then
    Unload frmChooseCharacter
    Exit Sub
End If

b = frmChooseCharacter.SelectedCharacter
Unload frmChooseCharacter

If hindex < 5 Then
    Set ccard = cFrontLine.PermanentRecord_GetCard(hindex, cindex)
Else
    Set ccard = myBattleSite.PermanentRecord_GetCard(cindex)
End If

If b < 5 Then
    cFrontLine.PermanentRecord_AddCard b, ccard
Else
    myBattleSite.PermanentRecord_AddCard ccard
End If

If hindex < 5 Then
    cFrontLine.PermanentRecord_RemoveCard hindex, cindex, True
Else
    myBattleSite.PermanentRecord_RemoveCard cindex, True
End If

SendData "CPM:" & Trim(Str(hindex)) & ":" & Trim(Str(b)) & ":" & Trim(Str(cindex)) & ":|"

ShowPermanentRecord
UpdateOpponentDeckDisplay
HideAllBorders
HeroClick 0
End Sub

Private Sub cmdPRRemove_Click()
hindex = Val(frmPR.Tag)
cindex = Val(imgCardDetail.Tag)

If hindex = 0 Or cindex = 0 Then Exit Sub

If hindex < 5 Then
    cFrontLine.PermanentRecord_RemoveCard hindex, cindex, False
Else
    myBattleSite.PermanentRecord_RemoveCard cindex, False
End If

SendData "CPX:" & Trim(Str(hindex)) & ":" & Trim(Str(cindex)) & ":|"

History_Add cFrontLine.Character_Name(hindex) & ": PR card removed."

ShowPermanentRecord
ShowPlacedCards

UpdateOpponentDeckDisplay
HeroClick 0
End Sub

Private Sub cmdRemoveAttackCard_Click()
If cmdRemoveAttackCard.Tag = -1 Then Exit Sub

cmdRemoveAttackCard.Enabled = False

a$ = Trim(Str((Val(cmdRemoveAttackCard.Tag))))

Select Case myattack.Card_Source(Val(cmdRemoveAttackCard.Tag))

Case "P1"
    SendData "CSC:16:12:" & a$ & ":|"
Case "P2"
    SendData "CSC:16:13:" & a$ & ":|"
Case "P3"
    SendData "CSC:16:14:" & a$ & ":|"
Case "P4"
    SendData "CSC:16:15:" & a$ & ":|"
Case "H"
    SendData "CSC:16:2:1:" & a$ & ":|"
Case "P5"
    SendData "CSC:16:26:1:" & a$ & ":|"
Case Else
End Select

myattack.RemoveCard cmdRemoveAttackCard.Tag + 0

cmdOKAction.Enabled = True
cmdCancelAction.Enabled = True

ShowAttackCards
ShowPlacedCards
FetchHandImages
ShowHand

End Sub

Private Sub cmdRemoveDefenseCard_Click()
If cmdOKDefense.Enabled = False Then
    X = MsgBox("Cards cannot be removed from a defense once it has been submitted to your opponent.  Message your opponent and ask him to challenge your defense.", vbCritical, "Cannot Remove Defense Card")
    Exit Sub
End If

cmdRemoveDefenseCard.Enabled = False

If myDefense.Card_Count = 0 Then Exit Sub

Index = imgCardDetail.Tag + 0

Select Case myDefense.Card_Source(Index)

Case "P1"
    SendData "CSC:17:12:" & Trim(Str(Index)) & ":|"
Case "P2"
    SendData "CSC:17:13:" & Trim(Str(Index)) & ":|"
Case "P3"
    SendData "CSC:17:14:" & Trim(Str(Index)) & ":|"
Case "P4"
    SendData "CSC:17:15:" & Trim(Str(Index)) & ":|"
Case "H"
    SendData "CSC:17:2:" & Trim(Str(Index)) & ":|"
Case "P5"
    SendData "CSC:17:18:" & Trim(Str(Index)) & ":|"
Case Else

End Select

myDefense.RemoveCard Val(imgCardDetail.Tag)
ShowDefenseCards

FetchHandImages
ShowHand
ShowPlacedCards

If myDefense.Card_Count <> 0 Then
    DefenseCardDetail 0
End If


End Sub


Private Sub cmdRemoveHTCB_Click()
hindex = Val(frmHTCB.Tag)
cindex = Val(imgCardDetail.Tag)

If hindex = 0 Or cindex = 0 Then Exit Sub

If hindex < 5 Then
    cFrontLine.HitsToCurrentBattle_RemoveCard hindex, cindex, False
Else
    myBattleSite.HitsToCurrentBattle_RemoveCard cindex, False
End If

SendData "CRH:" & Trim(Str(hindex)) & ":" & Trim(Str(cindex)) & ":|"

ShowHitsToCurrentBattle
UpdateOpponentDeckDisplay
HeroClick 0

End Sub

Private Sub cmdReserveToFrontline_Click()

If cFrontLine.LiveCharacterCount = 4 Then
    X = MsgBox("There is no open slot in the Frontline.  If you are switching with a Frontline character, select Frontline character and then click on 'Switch with Reserve.", vbOKOnly, "No Space for Reserve Character.")
    Exit Sub
End If

a = Val(imgHeroCard.Tag)
If a = 0 Then Exit Sub

SendData "CRF:" & Trim(Str(a)) & ":|"

cFrontLine.isCharacterReserve(a) = False

LoadCharacters
ShowHitsToCurrentBattle
ShowPermanentRecord
ShowPlacedCards
ShowModifiers
ShowBuffers

HeroClick 0

End Sub

Private Sub cmdReturnToHand_Click()
Dim ccard

cmdReturnToHand.Enabled = False

If Val(frmPlaced.Tag) = 0 Or Val(imgCardDetail.Tag = 0) Then Exit Sub

a = Val(frmPlaced.Tag)
b = Val(imgCardDetail.Tag)

If a = 5 Then

Set ccard = myHomebase.PlacedCard(b)
cHand.Add ccard
AddHandImage
myHomebase.RemovePlacedCard b

SendData "CSC:26:2:" & Trim(Str(b)) & ":|"

Else

SendData "CSC:" & Trim(Str(a + 11)) & ":2:" & Trim(Str(b)) & ":|"

Set ccard = cFrontLine.PlacedCard(a, b)
cHand.Add ccard
AddHandImage

cFrontLine.RemovePlacedCard a, b

End If

ShowPlacedCards
ShowHand
HideAllBorders

cmdReturnToHand.Enabled = True

End Sub

Private Sub cmdReturnToMissionC_Click()

'senddata "CSC:8:6:1:|"

MoveVentureCards cCompletedMissions, cMissions, False, imgCompletedMissions, "MISSIONS: COMPLETED"

End Sub

Private Sub cmdReturnToMissions_Click()
'If chkMoveAll.Value = 1 Then
'    senddata "CSC:9:6:" & Trim(Str(cVenturedMissions.Count)) & ":|"
'Else
'    senddata "CSC:9:6:1:|"
'End If

MoveVentureCards cVenturedMissions, cMissions, CBool(chkMoveAll.Value), imgVenture, "MISSION: VENTURED CARDS"

End Sub
Private Sub cmdSendMessage_Click()
    a$ = InputBox$("Message")
    If a$ = "" Then Exit Sub
    
   SendData "M" & a$ & "|"
   lstMessages.AddItem mySettings.PlayerName & ": " & a$
   lstMessages.ListIndex = lstMessages.ListCount - 1
   
End Sub

Private Sub cmdSendMyVenTotal_Click()

SendData "CVT:" & txtMyVentureTotal.Text & "|"

End Sub

Private Sub cmdShiftAttack_Click()

X = MsgBox("Note: You may only shift an attack to another character if a Special card or Homebase allows you to do so.  Would you like to continue?", vbYesNoCancel, "Shift Attack?")

If X <> 6 Then Exit Sub

Load frmChooseCharacter
With frmChooseCharacter

.HideHomebase = True
.ShowBattlesite = False
.Show 1

If .SelectedCharacter = -1 Then
    Unload frmChooseCharacter
    Exit Sub
End If

lc = -1

a = .SelectedCharacter
Unload frmChooseCharacter

End With

For i = 0 To 3
    lnFrontLine(i).Visible = False
Next i

For i = 0 To 2
    If imgFrontLine(i).Tag = a Then
        lc = i
    End If
Next i

If lc = -1 Then
    If ImgReserve.Tag = a Then
        lc = 3
    End If
End If

lnFrontLine(lc).Visible = True

OpAttack.DefenderID = a

History_Add "ATTACK SHIFTED TO: " & cFrontLine.Character_Name(a)

SendData "CSW:" & Trim(Str(a)) & ":|"

End Sub

Private Sub cmdSwitchWithReserve_Click()
a = Val(imgHeroCard.Tag)
If a = 0 Then Exit Sub

cFrontLine.isCharacterReserve(a) = True

SendData "CFR:" & Trim(Str(a)) & ":|"

HeroClick 0

LoadCharacters
ShowPlacedCards
ShowHitsToCurrentBattle
ShowPermanentRecord

End Sub
Private Sub cmdPlace_Click()
If myPhase <> nPhase_Place Then
    X = MsgBox("Note: you may not place cards to a hero at this time unless a Special or Event or other card allows you to do so.", vbInformation, "Wrong Phase")
End If

PlaceCard

End Sub
Private Sub cmdTakeAction_Click()
    
If myPhase <> nPhase_Attack Then
    X = MsgBox("It is not your turn to attack.", vbCritical, "Wrong Phase")
    Exit Sub
End If

NewAction

End Sub
Private Sub cmdVenCReturn_Click()

'If chkMoveAllVenC.Value = 1 Then
'    senddata "CSC:10:8:" & Trim(Str(cVenturedC.Count)) & ":|"
'Else
'    senddata "CSC:10:8:1:|"
'End If

MoveVentureCards cVenturedC, cCompletedMissions, CBool(chkMoveAllVenC.Value), imgVentureC, "MISSIONS: VENTURED"


End Sub

Private Sub cmdVenCToDefeated_Click()

'If chkMoveAllVenC.Value = 1 Then
'    senddata "CSC:10:7:" & Trim(Str(cVenturedC.Count)) & ":|"
'Else
'    senddata "CSC:10:7:1:|"
'End If

MoveVentureCards cVenturedC, cDeadMissions, CBool(chkMoveAllVenC.Value), imgVentureC, "MISSIONS: VENTURED"


End Sub

Private Sub cmdVenCToReserve_Click()
'If chkMoveAllVenC.Value = 1 Then
'    senddata "CSC:10:6:" & Trim(Str(cVenturedC.Count)) & ":|"
'Else
'    senddata "CSC:10:6:1:|"
'End If

MoveVentureCards cVenturedC, cMissions, CBool(chkMoveAllVenC.Value), imgVentureC, "MISSIONS: VENTURED"

End Sub
Private Sub cmdVenture_Click(Index As Integer)
If myPhase <> nPhase_Venture Then
    X = MsgBox("Note: Unless a Special, Homebase, etc. allows you to do so, you are not allowed to change the Venture at this time.  Are you sure you want to continue?", vbYesNoCancel, "Wrong Phase")
    If X <> 6 Then Exit Sub
End If

If cMissions.Count = 0 Then Exit Sub

mc = Index + 1
If mc > cMissions.Count Then mc = cMissions.Count

For i = 1 To mc
cVenturedMissions.Add cMissions.Item(1)
cMissions.Remove 1
Next i

imgVenture.Visible = True

SendData "CV0:" & Trim(Str(cMissions.Count)) & ":" & Trim(Str(cCompletedMissions.Count)) & ":" & Trim(Str(cDeadMissions.Count)) & ":" & Trim(Str(cVenturedMissions.Count)) & ":" & Trim(Str(cVenturedC.Count)) & ":|"

UpdateDeckDisplay

If cMissions.Count > 0 Then
    imgMissionCard.Picture = imgMissions.Picture
    Me.Caption = "OVERPOWER ONLINE-->" & "MISSIONS: (" & cMissions.Count & ")"
Else
    imgMissionCard.Picture = Nothing
    HeroClick 0
End If

End Sub

Private Sub cmdVentureC_Click()

If myPhase <> nPhase_Venture Then
    X = MsgBox("You may not venture cards at this time.", vbCritical, "Wrong Phase")
    Exit Sub
End If
'
'senddata "CSC:8:10:" & Trim(Str(cCompletedMissions.Count)) & ":|"

MoveVentureCards cCompletedMissions, cVenturedC, False, imgCompletedMissions, "MISSIONS: COMPLETED"

End Sub
Private Sub cmdViewBattlesiteDeck_Click()
ViewBattleSiteDeck

End Sub

Private Sub cmdWGF1_Click()

bIGoFirst = True

SendData "CGF:0:|"
frmWhoGoesFirst.Visible = False

History_Add mySettings.PlayerName & " GOES FIRST"

myPhase = nPhase_Draw
UpdatePhase
DrawNewHand

ShowDiscardFrame

End Sub

Private Sub cmdWGF2_Click()
bIGoFirst = False

SendData "CGF:1:|"
frmWhoGoesFirst.Visible = False

History_Add sOpponentName & " GOES FIRST"

myPhase = nPhase_Draw
UpdatePhase
DrawNewHand

ShowDiscardFrame

End Sub

Private Sub cmdWGFDraw1_Click()
Randomize

p1 = Int((cDrawPile.Count * Rnd) + 1)
t1$ = cDrawPile.Item(p1).Title

imgWGF1.ToolTipText = t1$

If cDrawPile.Item(p1).LoadImage(cDrawPile.Item(p1).ID) = True Then
    imgWGF1.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgWGF1.Picture = LoadPicture(sBlankImagePath)
End If

imgWGF1.Tag = p1

SendData "CWG:" & Trim(Str(p1)) & ":" & imgWGF2.Tag & ":|"

End Sub

Private Sub cmdWGFDraw2_Click()
Randomize

p2 = Int((cDrawPileO.Count * Rnd) + 1)
t2$ = cDrawPileO.Item(p2).Title

imgWGF2.ToolTipText = t2$

If cDrawPileO.Item(p2).LoadImage(cDrawPileO.Item(p2).ID) = True Then
    imgWGF2.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgWGF2.Picture = LoadPicture(sBlankImagePath)
End If

imgWGF2.Tag = p2

SendData "CWG:" & imgWGF1.Tag & ":" & Trim(Str(p2)) & ":|"

End Sub













Private Sub Form_Load()
NewGame

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bMoveDetail = True Then
frmDetail.Left = X

frmDetail.Refresh
End If

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmActInfo.Visible = False

End Sub

Private Sub imgAction_Click(Index As Integer)

ActionClick Index

End Sub
Private Sub imgBattlesite_Click()
If myBattleSite.ID < 1 Then Exit Sub

If lblKO(4).Visible = True Then
    cmdViewBattlesiteDeck.Enabled = False
    cmdKOBattlesite.Enabled = False
End If

HideFrames False, False, False
imgHeroCard.Picture = imgBattlesite.Picture

imgHeroCard.ToolTipText = myBattleSite.Effect

frmBattlesite.Visible = True
imgHeroCard.Visible = True
Me.Caption = "OVERPOWER ONLINE-->" & "BATTLESITE: " & myBattleSite.Name
End Sub

Private Sub imgBattlesite_DblClick()
ViewBattleSiteDeck

End Sub

Private Sub imgBGameEffect1_Click(Index As Integer)

ModifierDetail Index, imgBGameEffect1(Index), imgFrontLine(0).Tag, True


End Sub
Private Sub ModifierDetail(Index, oPicture As Image, sTag, bCardisPlaced As Boolean)
Dim ccard

    imgCardDetail.Picture = oPicture.Picture
    imgCardDetail.Tag = Index
    frmModifier.Tag = sTag
    HideFrames False, False, True
    frmModifier.Visible = True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    cmdDiscardModifier.Tag = ""
    
    If bCardisPlaced = True Then
        cmdPlayModifier.Enabled = False
        cmdPlaceModifier.Enabled = False
        cmdDiscardModifier.Enabled = True
    Else
        cmdPlayModifier.Enabled = True
        cmdPlaceModifier.Enabled = True
        cmdDiscardModifier.Enabled = True
    End If
    
    sTag = Val(sTag)
    Set ccard = cFrontLine.Modifiers_GetCard(sTag, Index)
        
    Me.Caption = "OVERPOWER ONLINE-->" & "MODIFIER [" & cFrontLine.Modifiers_TypeText(sTag, Index) & "]: " & ccard.Title
    
    
    If ccard.CardType = "Special Card" Or ccard.CardType = "Artifact" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If
    
    HideAllBorders
    
End Sub

Private Sub imgBGameEffect2_Click(Index As Integer)
ModifierDetail Index, imgBGameEffect2(Index), imgFrontLine(1).Tag, True

End Sub

Private Sub imgBGameEffect3_Click(Index As Integer)
ModifierDetail Index, imgBGameEffect3(Index), imgFrontLine(2).Tag, True

End Sub

Private Sub imgBGameEffect4_Click(Index As Integer)
ModifierDetail Index, imgBGameEffect4(Index), ImgReserve.Tag, True

End Sub

Private Sub imgBuffer1_Click(Index As Integer)

BufferDetail Index, imgBuffer1(Index), imgFrontLine(0).Tag, True

End Sub

Private Sub imgBuffer2_Click(Index As Integer)
BufferDetail Index, imgBuffer2(Index), imgFrontLine(1).Tag, True

End Sub

Private Sub imgBuffer3_Click(Index As Integer)
BufferDetail Index, imgBuffer3(Index), imgFrontLine(2).Tag, True

End Sub

Private Sub imgBuffer4_Click(Index As Integer)
BufferDetail Index, imgBuffer4(Index), ImgReserve.Tag, True

End Sub
Private Sub imgCardDetail_DblClick()
Load frmCardDetail
frmCardDetail.imgCard.Picture = imgCardDetail.Picture
frmCardDetail.Show 1
End Sub

Private Sub imgCompletedMissions_Click()
If cCompletedMissions.Count = 0 Then Exit Sub

Me.Caption = "OVERPOWER ONLINE-->" & "MISSIONS: COMPLETED (" & cCompletedMissions.Count & ")"
imgMissionCard.Picture = imgCompletedMissions.Picture
HideFrames False, False, False
imgMissionCard.Visible = True
frmCompletedMission.Visible = True

End Sub

Private Sub imgCompletedMissions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbBar1.Panels(1).Text = "Completed Missions (" & cCompletedMissions.Count & ")"

End Sub

Private Sub imgDeadMissions_Click()
If cDeadMissions.Count = 0 Then Exit Sub

Me.Caption = "OVERPOWER ONLINE-->" & "MISSIONS: DEFEATED (" & cDeadMissions.Count & ")"
imgMissionCard.Picture = imgDeadMissions.Picture
HideFrames False, False, False
imgMissionCard.Visible = True
frmDefeated.Visible = True
End Sub

Private Sub imgDefense_Click(Index As Integer)

DefenseCardDetail Index

End Sub
Private Sub DefenseCardDetail(Index)
On Error Resume Next

If imgDefense(Index).Tag = -1 Then Exit Sub
HideFrames False, False, False
imgCardDetail.Picture = imgDefense(Index).Picture
Me.Caption = "OVERPOWER ONLINE-->" & "DEFENSE CARD"
imgCardDetail.Visible = True
frmDefenseCard.Visible = True
imgCardDetail.Tag = Index + 1

'On Error Resume Next
If myPhase = nPhase_Attack Then
    cmdRemoveDefenseCard.Enabled = False
    Set ccard = cIncomingDefense.Item(Index + 1)
Else
    cmdRemoveDefenseCard.Enabled = True
    Set ccard = myDefense.GetCard(Index + 1)
End If

If ccard.CardType = "Special Card" Or ccard.CardType = "Aspect Card" Then
    imgCardDetail.ToolTipText = ccard.Effect
Else
    imgCardDetail.ToolTipText = ccard.Title
End If

End Sub
Private Sub imgEffect_Click(Index As Integer)

Me.Caption = "OVERPOWER ONLINE-->" & "EVENT IN PLAY"
imgHeroCard.Picture = imgEffect(Index).Picture
HideFrames False, False, False
imgHeroCard.Visible = True
imgHeroCard.ToolTipText = lblEvent(Index).Caption

End Sub
Private Sub imgFrontLine_Click(Index As Integer)
    
HeroClick Index

End Sub
Private Sub imgFrontLine_DblClick(Index As Integer)

If myPhase <> nPhase_Attack Then Exit Sub

myattack.AttackerID = imgFrontLine(Index).Tag
For i = 0 To 3
    lnFrontLine(i).Visible = False
Next i

If frmAttack.Visible = True Then

    lnFrontLine(Index).Visible = True

Else
    NewAction
End If


End Sub

Private Sub imgHand_Click(Index As Integer)

HandClick Index

End Sub
Private Sub HandClick(Index)

cmdAttack.Enabled = True
cmdDiscard.Enabled = True
cmdPlace.Enabled = True


imgCardDetail.Picture = imgHand(Index).Picture
imgCardDetail.Refresh
imgCardDetail.Tag = Index
HideFrames False, False, True
frmHandCard.Visible = True

Select Case cHand.Item(Index).CardType

Case "Aspect Card"
'    Set ccard = cHand.Item(Index)
'    imgCardDetail.Picture = imgHand(Index).Picture
'    imgCardDetail.Tag = Index
'    HideFrames False, False, False
    Me.Caption = "OVERPOWER ONLINE-->" & cHand.Item(Index).Title
    imgCardDetail.ToolTipText = cHand.Item(Index).Effect
    imgCardDetail.Visible = True
'    frmAspect.Visible = True
    frmHandCard.Tag = "H"
    
Case "Artifact"
    Set ccard = cHand.Item(Index)
    imgCardDetail.Picture = imgHand(Index).Picture
    imgCardDetail.Tag = Index
    HideFrames False, False, False
    frmModifier.Visible = True
    cmdPlayModifier.Enabled = True
    cmdPlaceModifier.Enabled = True
    cmdDiscardModifier.Enabled = True
    imgCardDetail.Visible = True
    Me.Caption = "OVERPOWER ONLINE-->" & "ARTIFACT: " & cHand.Item(Index).Title
    imgCardDetail.ToolTipText = cHand.Item(Index).Effect
    frmModifier.Tag = -1
    
Case "Special Card"
    Me.Caption = "OVERPOWER ONLINE-->" & "SPECIAL: " & cHand.Item(Index).Name
    imgCardDetail.ToolTipText = cHand.Item(Index).Effect

    If cHand.Item(Index).Attack_Frontline_Allies = True Then
        cmdDiscardModifier.Tag = "BUFFER"
    Else
        cmdDiscardModifier.Tag = ""
    End If
    
    
    If cHand.Item(Index).Attack_EffectsFrontline = True Then
        imgCardDetail.Picture = imgHand(Index).Picture
        HideFrames False, False, False
        Me.Refresh
        frmModifier.Tag = -1
        imgCardDetail.Tag = Index
        frmModifier.Visible = True
        cmdPlayModifier.Enabled = True
        cmdPlaceModifier.Enabled = True
        cmdDiscardModifier.Enabled = True
        imgCardDetail.Visible = True
        Me.Caption = "OVERPOWER ONLINE-->" & "SPECIAL: " & cHand.Item(Index).Name
        imgCardDetail.ToolTipText = cHand.Item(Index).Effect
        Exit Sub
    End If
    
Case "Activator"
    Me.Caption = "OVERPOWER ONLINE-->" & "ACTIVATOR: " & cHand.Item(Index).Name
    imgHeroCard.Picture = imgHand(Index).Picture
    imgHeroCard.Tag = Index
    imgHeroCard.ToolTipText = cHand.Item(Index).Title
    HideFrames False, False, False
    imgHeroCard.Visible = True
    frmActivator.Visible = True
    
    If lblKO(4).Visible = True Then cmdExchangeActivator.Enabled = False

Case "Event"
    Me.Caption = "OVERPOWER ONLINE-->EVENT: " & cHand.Item(Index).Name
    imgHeroCard.Picture = imgHand(Index).Picture
    imgHeroCard.Tag = Index
    imgHeroCard.ToolTipText = cHand.Item(Index).Title
    HideFrames False, False, False
    imgHeroCard.Visible = True
    frmEvent.Visible = True
    
Case Else
    Me.Caption = "OVERPOWER ONLINE-->" & cHand.Item(Index).Title
    imgCardDetail.ToolTipText = cHand.Item(Index).Title
End Select

'[FIX ME]
If myPhase < nPhase_Place Then
    cmdAttack.Enabled = False
    cmdPlace.Enabled = False
End If

If myPhase <> nPhase_Attack And myPhase <> nPhase_Defend Then
    cmdAttack.Enabled = False
End If

If myPhase = nPhase_Venture Or myPhase = nPhase_Resolve Then
    cmdAttack.Enabled = False
    cmdPlace.Enabled = False
    cmdDiscard.Enabled = False
End If



End Sub
Private Sub HideAllBorders()


End Sub
Private Sub HideFrames(isHero As Boolean, isMission As Boolean, ishand As Boolean)
frmHandCard.Visible = False
frmPlaced.Visible = False

If isHero = False And ishand = False And isMission = False Then
    imgMissionCard.Visible = False
    imgHeroCard.Visible = False
    imgCardDetail.Visible = False
    frmHero.Visible = False
    frmMission.Visible = False
End If

frmVentureC.Visible = False
frmVentured.Visible = False
frmCompletedMission.Visible = False
frmDefeated.Visible = False
frmHomebase.Visible = False
frmBattlesite.Visible = False
frmHTCB.Visible = False
frmPR.Visible = False
frmActivator.Visible = False
frmEvent.Visible = False
frmDefenseCard.Visible = False
frmAttackCard.Visible = False
frmModifier.Visible = False
'frmAspect.Visible = False

If isHero = True Then
    imgMissionCard.Visible = False
    imgHeroCard.Visible = True
    imgCardDetail.Visible = False
    frmHero.Visible = True
    frmMission.Visible = False
End If

If ishand = True Then
    imgMissionCard.Visible = False
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmHero.Visible = False
    frmMission.Visible = False
End If

If isMission = True Then
    imgMissionCard.Visible = True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = False
    frmHero.Visible = False
    frmMission.Visible = True
End If

End Sub

Private Sub imgHand_DblClick(Index As Integer)

On Error Resume Next

Select Case cHand.Item(Index).CardType

Case "Aspect Card", "Special Card", "Ally Card", "Basic Universe", "Double Shot", "Power Card", "Teamwork", "Training"
   
Case Else

    Exit Sub
    
End Select

Select Case myPhase

Case nPhase_Attack

    If frmAttack.Visible = False Then Exit Sub
    
    SendData "CSC:2:16:" & Trim(Str(Index)) & ":|"
    
    Set ccard = cHand.Item(Index)
    myattack.AddCard ccard, "H", Index
    ShowAttackCards
    ShowHand
    
Case nPhase_Defend

    If frmDefense.Visible = False Then Exit Sub
    
    If cmdOKDefense.Enabled = False Then
        X = MsgBox("Cards cannot be added to a defense once it has been submitted to your opponent.  Message your opponent and ask him to challenge your defense.", vbCritical, "Cannot add Defense Card")
        Exit Sub
    End If
    
    SendData "CSC:2:17:" & Trim(Str(Index)) & ":|"
    
    Set ccard = cHand.Item(Index)
    myDefense.AddCard ccard, "H", Index
    
    ShowDefenseCards
    ShowHand
    
Case Else
    Exit Sub
End Select


End Sub

Private Sub imgHand_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Status ""

a$ = lstHandTips.List(Index - 1)

If Left(a$, 3) = "XXA" Then

    a$ = Right(a$, Len(a$) - 3)
    lstAvailableActivators.Clear

looper:
    X = InStr(a$, "^")

    If X > 0 Then
        b$ = Left(a$, X - 1)
        a$ = Right(a$, Len(a$) - X)
        X = InStr(b$, "~")
        cnum = Val(Right(b$, Len(b$) - X))
        b$ = Left$(b$, X - 1)

        lstAvailableActivators.AddItem b$
        lstAvailableActivators.ItemData(lstAvailableActivators.NewIndex) = cnum


        GoTo looper
    End If

    lstAvailableActivators.SetFocus
    If lstAvailableActivators.ListCount > 0 Then lstAvailableActivators.ListIndex = 0

'    a$ = ReplaceAllInString(a$, "/", vbCrLf)
'    a$ = Right$(a$, Len(a$) - 3)
'    Me.lblActivators.Caption = a$
    frmActInfo.Tag = Index
    frmActInfo.Top = imgHand(Index).Top
    frmActInfo.Visible = True

Else
    frmActInfo.Visible = False
    Status a$
End If

End Sub
Private Sub CheckPlayStatus(ccard)
Exit Sub

'For i = 0 To 2
'    a = Val(imgFrontLine(i).Tag)
'    If a > 0 Then
'
'        If cFrontLine.CanCharacterPlayCard(a, ccard) = True Then
'            shpBorder(i).Visible = True
'        Else
'            shpBorder(i).Visible = False
'        End If
'
'
'    Else
'
'    shpBorder(i).Visible = False
'
'    End If
'Next i
'
'    a = Val(ImgReserve.Tag)
'    If a > 0 Then
'
'
'    If cFrontLine.CanCharacterPlayCard(a, ccard) = True Then
'        shpBorder(3).Visible = True
'    Else
'        shpBorder(3).Visible = False
'    End If
'
'
'    End If
    
End Sub
Private Sub imgHand_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'See where data is coming from
Select Case cCurrentDragSource.Tag

Case "Deck"
'moving a card from the deck to the discard pile
cHand.Add cDrawPile.Item(1)

AddHandImage

cDrawPile.Remove 1

Case Else

End Select

UpdateDeckDisplay
ShowHand

End Sub
Private Sub AddHandImage()

If cHandTags.Count = 0 Then
    cHandTags.Add "A"
Else
    a = Asc(cHandTags.Item(cHandTags.Count)) + 1
    cHandTags.Add Chr$(a)
End If

Load imgTemp(imgTemp.Count)
imgTemp(imgTemp.Count - 1).Visible = True
imgTemp(imgTemp.Count - 1).Left = imgTemp(imgTemp.Count - 2).Left + 400
imgTemp(imgTemp.Count - 1).Tag = Chr$(a)

If cHand.Item(cHand.Count).LoadImage(cHand.Item(cHand.Count).ID) = True Then
        imgTemp(imgTemp.Count - 1).Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
        imgTemp(imgTemp.Count - 1).Picture = LoadPicture(sBlankImagePath)
End If

End Sub

Private Sub imgHand_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
Set cCurrentDragSource = imgHand(Index)
End Sub

Private Sub imgHeroCard_DblClick()
Load frmCardDetail
frmCardDetail.imgCard.Picture = imgHeroCard.Picture

frmCardDetail.Show 1
End Sub

Private Sub imgHideDefense_Click()

If frmDefense.Top = 5400 Then
    frmDefense.Top = 4480
Else
    frmDefense.Top = 5400
End If

End Sub

Private Sub imgHit1_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHit1(Index).Picture
    frmHTCB.Tag = imgFrontLine(0).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmHTCB.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HIT: " & cFrontLine.Character_Name(Val(imgFrontLine(0).Tag))

    Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(Val(imgFrontLine(0).Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgHit2_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHit2(Index).Picture
    frmHTCB.Tag = imgFrontLine(1).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmHTCB.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HIT: " & cFrontLine.Character_Name(Val(imgFrontLine(1).Tag))

    Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(Val(imgFrontLine(1).Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgHit3_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHit3(Index).Picture
    frmHTCB.Tag = imgFrontLine(2).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmHTCB.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HIT: " & cFrontLine.Character_Name(Val(imgFrontLine(2).Tag))

    Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(Val(imgFrontLine(2).Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgHit4_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHit4(Index).Picture
    frmHTCB.Tag = ImgReserve.Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmHTCB.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HIT: " & cFrontLine.Character_Name(Val(ImgReserve.Tag))

    Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(Val(ImgReserve.Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgHitBS_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHitBS(Index).Picture
    frmHTCB.Tag = 5
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmHTCB.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HIT: " & myBattleSite.Name

    Set ccard = myBattleSite.HitsToCurrentBattle_GetCard(Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgHitBuffer_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHitBuffer(Index).Picture
    imgCardDetail.Tag = Index
    HideFrames False, False, False
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HIT ON A BUFFER CARD"
    
    Set ccard = cFrontLine.BufferHits_GetCard(Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgHitBufferOP_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHitBufferOP(Index).Picture
    imgCardDetail.Tag = Index
    HideFrames False, False, False
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HIT ON A BUFFER CARD"
    
    Set ccard = cOpponent.BufferHits_GetCard(Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgHitOpBS_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgHitOpBS(Index).Picture
    frmPR.Tag = imgOpBattlesite.Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HTCB: " & OpBattlesite.Name
    
    Set ccard = OpBattlesite.HitsToCurrentBattle_GetCard(Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub
Private Sub imgHomebase_Click()
If myHomebase.ID < 1 Then Exit Sub

HideFrames False, False, False
imgHeroCard.Picture = imgHomebase.Picture
imgHeroCard.Visible = True
frmHomebase.Visible = True
Me.Caption = "OVERPOWER ONLINE-->" & "HOMEBASE: " & myHomebase.Name
txtHomebaseBonus.Text = myHomebase.Effect

lstHomeBaseChars.Clear
a$ = myHomebase.Characters & ","

While a$ <> ""

X = InStr(a$, ",")
lstHomeBaseChars.AddItem Trim(Left(a$, X - 1))
a$ = Right$(a$, Len(a$) - X)

Wend

End Sub

Private Sub imgMissionCard_DblClick()
Load frmCardDetail
frmCardDetail.imgCard.Picture = imgMissionCard.Picture

frmCardDetail.Show 1
End Sub

Private Sub imgMissions_Click()

If cMissions.Count = 0 Then Exit Sub

Me.Caption = "OVERPOWER ONLINE-->" & "MISSIONS: (" & cMissions.Count & ")"
imgMissionCard.Picture = imgMissions.Picture
HideFrames False, True, False

End Sub

Private Sub imgMissions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbBar1.Panels(1).Text = "Missions Pile (" & cMissions.Count & ")"

End Sub

Private Sub imgOpBattlesite_Click()
On Error Resume Next
If OpBattlesite.ID < 1 Then Exit Sub

HideFrames False, False, False
imgHeroCard.Picture = imgOpBattlesite.Picture

imgHeroCard.ToolTipText = OpBattlesite.Effect

'frmBattlesite.Visible = True
imgHeroCard.Visible = True
Me.Caption = "OVERPOWER ONLINE-->" & "BATTLESITE: " & OpBattlesite.Name

ShowAttackOpLines 8

End Sub

Private Sub ShowAttackOpLines(ShowLine)
If frmAttack.Visible = False Then Exit Sub

For i = 4 To 8
    lnFrontLine(i).Visible = False
Next i

lnFrontLine(ShowLine).Visible = True

Select Case ShowLine
Case 4
    myattack.DefenderID = imgOpponent(0).Tag
Case 5
    myattack.DefenderID = imgOpponent(1).Tag
Case 6
    myattack.DefenderID = imgOpponent(2).Tag
Case 7
    myattack.DefenderID = imgOppReserve.Tag
Case 8
    myattack.DefenderID = 5
Case Else
    myattack.DefenderID = 0
End Select


End Sub
Private Sub OpponentModifierDetail(heroindex, cardindex)
Dim ccard

Set ccard = cOpponent.Modifiers_GetCard(heroindex, cardindex)

If ccard.LoadImage(ccard.ID) = True Then
    imgCardDetail.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgCardDetail.Picture = LoadPicture(sBlankImagePath)
End If

If ccard.Attack_Frontline_BattleBonus = True Then Me.Caption = "OVERPOWER ONLINE-->MODIFIER [BATTLE]: " & ccard.Title
If ccard.Attack_Frontline_GameBonus = True Then Me.Caption = "OVERPOWER ONLINE-->MODIFIER [GAME]: " & ccard.Title

imgCardDetail.ToolTipText = ccard.Effect
imgCardDetail.Tag = cardindex
HideFrames False, False, False
imgHeroCard.Visible = False
imgCardDetail.Visible = True

End Sub

Private Sub imgOPBuffer1_Click(Index As Integer)

OpponentBufferDetail Index, imgOPBuffer1(Index), imgOpponent(0).Tag

End Sub

Private Sub imgOpBuffer2_Click(Index As Integer)
OpponentBufferDetail Index, imgOpBuffer2(Index), imgOpponent(1).Tag

End Sub

Private Sub imgOpBuffer3_Click(Index As Integer)
OpponentBufferDetail Index, imgOpBuffer3(Index), imgOpponent(2).Tag

End Sub

Private Sub imgOpBuffer4_Click(Index As Integer)
OpponentBufferDetail Index, imgOpBuffer4(Index), imgOppReserve.Tag

End Sub

Private Sub imgOpEffect1_Click(Index As Integer)

OpponentModifierDetail imgOpponent(0).Tag, Index

End Sub

Private Sub imgOpEffect2_Click(Index As Integer)

OpponentModifierDetail imgOpponent(1).Tag, Index

End Sub

Private Sub imgOpEffect3_Click(Index As Integer)

OpponentModifierDetail imgOpponent(2).Tag, Index

End Sub

Private Sub imgOpEffect4_Click(Index As Integer)

OpponentModifierDetail imgOppReserve.Tag, Index

End Sub

Private Sub imgOPHit1_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgOPHit1(Index).Picture
    frmPR.Tag = imgOpponent(0).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HTCB: " & cOpponent.Character_Name(Val(imgOpponent(0).Tag))

    Set ccard = cOpponent.HitsToCurrentBattle_GetCard(Val(imgOpponent(0).Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgOPHit2_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgOPHit2(Index).Picture
    frmPR.Tag = imgOpponent(1).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HTCB: " & cOpponent.Character_Name(Val(imgOpponent(1).Tag))

    Set ccard = cOpponent.HitsToCurrentBattle_GetCard(Val(imgOpponent(1).Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgOPHit3_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgOPHit3(Index).Picture
    frmPR.Tag = imgOpponent(2).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HTCB: " & cOpponent.Character_Name(Val(imgOpponent(2).Tag))

    Set ccard = cOpponent.HitsToCurrentBattle_GetCard(Val(imgOpponent(2).Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgOPHit4_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgOPHit4(Index).Picture
    frmPR.Tag = imgOppReserve.Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HTCB: " & cOpponent.Character_Name(Val(imgOppReserve.Tag))

    Set ccard = cOpponent.HitsToCurrentBattle_GetCard(Val(imgOppReserve.Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgOpHomeBase_Click()
On Error Resume Next

If OpHomebase.ID < 1 Then Exit Sub

HideFrames False, False, False
imgHeroCard.Picture = imgOpHomeBase.Picture
imgHeroCard.Visible = True
frmHomebase.Visible = True
Me.Caption = "OVERPOWER ONLINE-->" & "HOMEBASE: " & OpHomebase.Name
txtHomebaseBonus.Text = OpHomebase.Effect

lstHomeBaseChars.Clear
a$ = OpHomebase.Characters & ","

While a$ <> ""

X = InStr(a$, ",")
lstHomeBaseChars.AddItem Trim(Left(a$, X - 1))
a$ = Right$(a$, Len(a$) - X)

Wend

End Sub

Private Sub imgOpPlace1_Click(Index As Integer)

OpponentPlacedCardDetail Index, imgOpPlace1(Index), imgOpponent(0).Tag

End Sub
Private Sub OpponentPlacedCardDetail(Index, oPicture As Image, sTag)
Dim ccard

    imgCardDetail.Picture = oPicture.Picture
    imgCardDetail.Tag = Index
    frmPlaced.Tag = sTag
    HideFrames False, False, False
    frmPlaced.Visible = False
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    
    If sTag = 5 Then
        Set ccard = OpHomebase.PlacedCard(Index)
        Me.Caption = "OVERPOWER ONLINE-->PLACED: " & ccard.Title
    Else
        Me.Caption = "OVERPOWER ONLINE-->" & "PLACED: " & cOpponent.Placed_Type(sTag, Index)
        Set ccard = cOpponent.PlacedCard(sTag, Index)
    End If
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If
    
    
    HideAllBorders

End Sub

Private Sub imgOpPlace2_Click(Index As Integer)

    OpponentPlacedCardDetail Index, imgOpPlace2(Index), imgOpponent(1).Tag
    
End Sub

Private Sub imgOpPlace3_Click(Index As Integer)

OpponentPlacedCardDetail Index, imgOpPlace3(Index), imgOpponent(2).Tag

End Sub

Private Sub imgOpPlace4_Click(Index As Integer)

OpponentPlacedCardDetail Index, imgOpPlace4(Index), imgOppReserve.Tag

End Sub

Private Sub imgOpPlacedHomebase_Click(Index As Integer)
OpponentPlacedCardDetail Index, imgOpPlacedHomebase(Index), 5

End Sub

Private Sub imgOpponent_Click(Index As Integer)

    OpponentDetail imgOpponent(Index), Index + 4
    
End Sub
Private Sub imgOppReserve_Click()
    
    OpponentDetail imgOppReserve, 7
    
End Sub

Private Sub imgPlaced1_Click(Index As Integer)

PlacedCardDetail Index, imgPlaced1(Index), imgFrontLine(0).Tag


End Sub

Private Sub imgPlaced2_Click(Index As Integer)

PlacedCardDetail Index, imgPlaced2(Index), imgFrontLine(1).Tag

End Sub

Private Sub imgPlaced3_Click(Index As Integer)

PlacedCardDetail Index, imgPlaced3(Index), imgFrontLine(2).Tag

End Sub

Private Sub imgPlaced4_Click(Index As Integer)

PlacedCardDetail Index, imgPlaced4(Index), ImgReserve.Tag

End Sub

Private Sub imgPlacedHomeBase_Click(Index As Integer)
frmHandCard.Tag = "P"

PlacedCardDetail Index, imgPlacedHomeBase(Index), 5

End Sub

Private Sub imgPR1_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPR1(Index).Picture
    frmPR.Tag = imgFrontLine(0).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cFrontLine.Character_Name(Val(imgFrontLine(0).Tag))

    Set ccard = cFrontLine.PermanentRecord_GetCard(Val(imgFrontLine(0).Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPR2_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPR2(Index).Picture
    frmPR.Tag = imgFrontLine(1).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cFrontLine.Character_Name(Val(imgFrontLine(1).Tag))

    Set ccard = cFrontLine.PermanentRecord_GetCard(Val(imgFrontLine(1).Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPR3_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPR3(Index).Picture
    frmPR.Tag = imgFrontLine(2).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cFrontLine.Character_Name(Val(imgFrontLine(2).Tag))

    Set ccard = cFrontLine.PermanentRecord_GetCard(Val(imgFrontLine(2).Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPR4_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPR4(Index).Picture
    frmPR.Tag = ImgReserve.Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cFrontLine.Character_Name(Val(ImgReserve.Tag))

    Set ccard = cFrontLine.PermanentRecord_GetCard(Val(ImgReserve.Tag), Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPRBS_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPRBS(Index).Picture
    frmPR.Tag = 5
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & myBattleSite.Name

    Set ccard = myBattleSite.PermanentRecord_GetCard(Index)
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPROP1_Click(Index As Integer)

Dim ccard

    imgCardDetail.Picture = imgPROP1(Index).Picture
    frmPR.Tag = imgOpponent(0).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cOpponent.Character_Name(Val(imgOpponent(0).Tag))

    Set ccard = cOpponent.PermanentRecord_GetCard(Val(imgOpponent(0).Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPROP2_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPROP2(Index).Picture
    frmPR.Tag = imgOpponent(1).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cOpponent.Character_Name(Val(imgOpponent(1).Tag))

    Set ccard = cOpponent.PermanentRecord_GetCard(Val(imgOpponent(1).Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPROP3_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPROP3(Index).Picture
    frmPR.Tag = imgOpponent(2).Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cOpponent.Character_Name(Val(imgOpponent(2).Tag))

    Set ccard = cOpponent.PermanentRecord_GetCard(Val(imgOpponent(2).Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPROP4_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPROP4(Index).Picture
    frmPR.Tag = imgOppReserve.Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & cOpponent.Character_Name(Val(imgOppReserve.Tag))

    Set ccard = cOpponent.PermanentRecord_GetCard(Val(imgOppReserve.Tag), Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub imgPROPBS_Click(Index As Integer)
Dim ccard

    imgCardDetail.Picture = imgPROPBS(Index).Picture
    frmPR.Tag = imgOpBattlesite.Tag
    imgCardDetail.Tag = Index
    HideFrames False, False, True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    'frmPR.Visible = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "P.R.: " & OpBattlesite.Name

    Set ccard = OpBattlesite.PermanentRecord_GetCard(Index)
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If

    HideAllBorders
End Sub

Private Sub ImgReserve_Click()
    If Val(ImgReserve.Tag) = 0 Then Exit Sub
    
    imgHeroCard.Picture = ImgReserve.Picture
    imgHeroCard.Tag = ImgReserve.Tag
    HideFrames True, False, False
   
    a = Val(imgHeroCard.Tag)
    
    If cFrontLine.Character_HasInherent(a) = True Then
        txtInherent.Text = cFrontLine.Character_Inherent(a)
    Else
        txtInherent.Text = "NO INHERENT ABILITY"
    End If
    
    cmdTakeAction.Enabled = True
    cmdKOCharacter.Enabled = True
    cmdSwitchWithReserve.Enabled = True
    cmdReserveToFrontline.Enabled = True
    

    Me.Caption = "OVERPOWER ONLINE-->" & "HERO: " & cFrontLine.Character_Name(a)
    
    If cFrontLine.isCharacterReserve(a) = True Then
        cmdReserveToFrontline.Visible = True
        cmdSwitchWithReserve.Visible = False
    Else
        cmdReserveToFrontline.Visible = False
        cmdSwitchWithReserve.Visible = True
    End If
    

    HideAllBorders
End Sub




Private Sub ImgReserve_DblClick()

If myPhase <> nPhase_Attack Then Exit Sub

myattack.AttackerID = ImgReserve.Tag

For i = 0 To 3
    lnFrontLine(i).Visible = False
Next i

If frmAttack.Visible = True Then
    lnFrontLine(3).Visible = True
Else
    NewAction
End If

End Sub

Private Sub imgStringAttack_Click()

HideFrames False, False, False
imgCardDetail.Picture = imgStringAttack.Picture
imgCardDetail.ToolTipText = imgStringAttack.ToolTipText
Me.Caption = "OVERPOWER ONLINE-->" & "ATTACK CARD"
imgCardDetail.Visible = True
imgCardDetail.Tag = "-1"

End Sub

Private Sub imgVenture_Click()
If cVenturedMissions.Count = 0 Then Exit Sub

Me.Caption = "OVERPOWER ONLINE-->" & "MISSION: VENTURED (" & cVenturedMissions.Count & ")"
imgMissionCard.Picture = imgVenture.Picture
HideFrames False, False, False
imgMissionCard.Visible = True
frmVentured.Visible = True

End Sub

Private Sub imgVenture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbBar1.Panels(1).Text = "Current Venture (" & cVenturedMissions.Count & ")"

End Sub

Private Sub imgVentureC_Click()
If cVenturedC.Count = 0 Then Exit Sub

Me.Caption = "OVERPOWER ONLINE-->" & "MISSIONS: VENTURED (" & cVenturedC.Count & ")"

imgMissionCard.Picture = imgVentureC.Picture
HideFrames False, False, False
imgMissionCard.Visible = True
frmVentureC.Visible = True

End Sub

Private Sub imgVentureC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbBar1.Panels(1).Text = "Current Venture (" & cVenturedC.Count & ")"

End Sub

Private Sub imgWGF1_DblClick()
Load frmCardDetail
frmCardDetail.imgCard.Picture = imgWGF1.Picture

frmCardDetail.Show 1
End Sub

Private Sub imgWGF2_DblClick()
Load frmCardDetail
frmCardDetail.imgCard.Picture = imgWGF2.Picture

frmCardDetail.Show 1
End Sub

Private Sub lstAvailableActivators_DblClick()

With lstAvailableActivators

frmActInfo.Visible = False

If .ListIndex = -1 Then Exit Sub

Index = Val(frmActInfo.Tag)
cindex = .ItemData(.ListIndex)

cHand.Add myBattleSite.Deck_GetCard(cindex)

SendData "CSC:11:2:" & Trim(Str(cindex)) & ":|"

If cHandTags.Count = 0 Then
    cHandTags.Add "A"
Else
    cHandTags.Add Chr$(Asc(cHandTags.Item(cHandTags.Count)) + 1)
End If

myBattleSite.RemoveDeckCard cindex

cDeadPile.Add cHand.Item(Index)
cHand.Remove Index

SendData "CSC:2:4:" & Trim(Str(Index)) & ":|"

Status "Fetching card from Battlesite deck..."
History_Add "ACTIVATOR USED TO FETCH SPECIAL..."

UpdateDeckDisplay
FetchHandImages
ShowHand

Status ""

End With
End Sub

Private Sub lstGameHistory_Click()
If lstGameHistory.ListIndex = -1 Then Exit Sub
lstGameHistory.ToolTipText = lstGameHistory.List(lstGameHistory.ListIndex)

Status lstGameHistory.List(lstGameHistory.ListIndex)

End Sub

Private Sub lstMessages_Click()
If lstMessages.ListIndex = -1 Then Exit Sub
lstMessages.ToolTipText = lstMessages.List(lstMessages.ListIndex)
Status lstMessages.List(lstMessages.ListIndex)

End Sub

Private Sub mnuConcedeToOpponent_Click()
X = MsgBox("Are you sure you want to concede the battle?", vbYesNoCancel, "Confirm Concession to Opponent")

If X <> 6 Then Exit Sub

SendData "CCB:1:|"
History_Add mySettings.PlayerName & " CONCEDES BATTLE"
bIHaveConceded = True

If OpponentHasConcedeEffect = False Then
    ResolveConcession True
Else
    myPhase = nPhase_Defend
    History_Add "AWAITING RESPONSE..."
    UpdatePhase
End If

End Sub

Private Sub mnuDeck_Click()
If cHand.Count < 2 Then
    mnuShowDuplicates.Enabled = False
Else
    mnuShowDuplicates.Enabled = True
End If

If tcpChannel.State = sckConnected Then
    mnuToolsConnect.Enabled = True
    mnuPlayOpenHanded.Enabled = True
Else
    mnuPlayOpenHanded.Enabled = False
    mnuPlayOpenHanded.Checked = False
End If

If cDrawPile.Count = 0 And cDiscardPile.Count = 0 Then
    mnuDrawCard.Enabled = False
Else
    mnuDrawCard.Enabled = True
End If

On Error Resume Next
If myBattleSite.ID = 0 Then
    mnuViewBattleSiteDEck.Enabled = False
Else
    mnuViewBattleSiteDEck.Enabled = True
End If

If cDrawPile.Count = 0 Then
    mnuViewDrawPile.Enabled = False
Else
    mnuViewDrawPile.Enabled = True
End If

If cDiscardPile.Count = 0 Then
    mnuViewDiscardPile.Enabled = False
Else
    mnuViewDiscardPile.Enabled = True
End If

If cDeadPile.Count = 0 Then
    mnuViewDeadPile.Enabled = False
Else
    mnuViewDeadPile.Enabled = True
End If

If cDefeatedCharactersPile.Count = 0 Then
    mnuViewDefeatedCharacters.Enabled = False
Else
    mnuViewDefeatedCharacters.Enabled = True
End If

If myPhase = nPhase_Attack Or myPhase = nPhase_Defend Then
    mnuConcedeToOpponent.Enabled = True
Else
    mnuConcedeToOpponent.Enabled = False
End If

If myPhase = nPhase_Attack Then
    mnuPassToOpponent.Enabled = True
    mnuEndTurn.Enabled = True
Else
    mnuPassToOpponent.Enabled = False
    mnuEndTurn.Enabled = False
End If

If bHavePassed = True Then
    mnuStopPass.Visible = True
Else
    mnuStopPass.Visible = False
End If

End Sub

Private Sub mnuDeckDrawHand_Click()

DrawNewHand

End Sub
Private Sub DrawNewHand()
Status "Drawing hand..."
X = DrawHand(8)
If X = 0 Then Exit Sub


Status "DREW " & X & " CARDS..."
'ShowHand
FetchHandImages
UpdateDeckDisplay

Status "Hand (" & cHand.Count & ")"

History_Add mySettings.PlayerName & " DREW " & X & " CARDS."

End Sub
Private Sub mnuDeckEditor_Click()
'Status "Loading deck editor..."
'
'frmDeckEditor.Show 1
'
'Status ""

X = Shell(App.Path & "\opdeck.exe", vbNormalFocus)

End Sub

Private Sub mnuDrawCard_Click()

DrawCard

SendData "CDC:1:|"

History_Add mySettings.PlayerName & " DREW 1 CARD."

End Sub
Private Sub DrawCard()
'If cHand.Count = 0 Then Exit Sub

Status "Drawing card..."

If cDrawPile.Count = 0 Then
    'Need to shuffle in power pack
    ShufflePile 1
    
    For i = 1 To cDiscardPile.Count
        cDrawPile.Add cDiscardPile.Item(i)
    Next i
    
    Set cDiscardPile = New Collection
        
    'Now that power pack has been shuffled in, check and see if there are enough cards
    If cDrawPile.Count = 0 Then
    
    X = MsgBox("There are no cards available to draw.", vbCritical, "No More Cards!")
    Exit Sub
    End If
End If

SendData "CSC:1:2:1:|"

cHand.Add cDrawPile.Item(1)
cDrawPile.Remove 1
AddHandImage

Status "Drew 1 card"
FetchHandImages
ShowHand
UpdateDeckDisplay

Status "Hand (" & cHand.Count & ")"
End Sub

Private Sub mnuDrawTestHand_Click()
DrawHand 8
FetchHandImages

ShowHand
End Sub

Private Sub mnuEndTurn_Click()
HideStringAttackFrame
SendData "CAC:1:|"
myPhase = nPhase_Defend
History_Add sOpponentName & " IS PREPARING TO ATTACK"
UpdatePhase
    
End Sub

Private Sub mnuFileExit_Click()
End

End Sub

Private Sub mnuFileOpenDeck_Click()

With cmD1

.InitDir = App.Path & "\Decks"
.FileName = "*.dat"
.Action = 1

If .FileName = "*.dat" Then Exit Sub

OpenDeck .FileName, .FileTitle


End With
End Sub
Private Sub SendOpponentDeck()

a$ = "CHL:"

For i = 1 To 4
    If cFrontLine.isCharacterReserve(i) = True Then
        a$ = a$ & cFrontLine.Character_ID(i) & "R:"
    Else
        a$ = a$ & cFrontLine.Character_ID(i) & ":"
    End If
Next i

SendData a$ & "|"

DoEvents

On Error Resume Next

If myHomebase.ID <> 0 Then
    SendData "CHN:" & Trim(Str(myHomebase.ID)) & ":|"
Else
    SendData "CHN:0:|"
End If

DoEvents
Me.Refresh

If myBattleSite.ID <> 0 Then
    SendData "CBN:" & Trim(Str(myBattleSite.ID)) & ":|"
Else
    SendData "CBN:0:|"
End If

DoEvents
Me.Refresh

Set cMissionsO = New Collection
For i = 1 To 7
    cMissionsO.Add "1"
Next i

DoEvents


'Send Draw Pile
SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"

'Send battlesitedeck
If myBattleSite.Deck_Count > 0 Then
    
    Dim ctemp As Collection
    Set ctemp = New Collection
    
    For i = 1 To myBattleSite.Deck_Count
    
    Set ccard = myBattleSite.Deck_GetCard(i)
    ctemp.Add ccard
    Next i
    
    SendData "CDB:" & GetCode_CardString(ctemp) & "|"

End If

SendData "CEP:1:|"

myPhase = nPhase_WhoGoesFirst
UpdatePhase

'figure out who goes first
If bHost = True Then

    bStopProcessing = True

    Do Until bStopProcessing = False
        DoEvents
    Loop

    WhoGoesFirst
End If

DoEvents


End Sub
Private Sub LoadOpponentHomebase()
'Load HomeBase

If OpHomebase.ID > 0 Then
    If OpHomebase.LoadImage(OpHomebase.ID) = True Then
        imgOpHomeBase.Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgOpHomeBase.ToolTipText = OpHomebase.Effect
        
    Else
        imgOpHomeBase.Picture = LoadPicture(sBlankImagePath)
    End If
End If

End Sub
Sub OpenDeck(sFileName, sfiletitle)

NewGame

X = FreeFile

Open sFileName For Input As #X

'read 4 heroes
For i = 1 To 4

Line Input #X, a$
    cFrontLine.AddCharacter Val(GetVal(a$)), False, False
Next i

Line Input #X, a$
nreserve = Val(GetVal(a$))
cFrontLine.isCharacterReserve(nreserve) = True

Line Input #X, a$
Set myHomebase = New clsHomebase

myHomebase.Load Val(GetVal(a$))

Line Input #X, a$
myBattleSite.Load Val(GetVal(a$))

loadbattlesite

'Get Mission
Line Input #X, a$

b = Val(GetVal(a$))

Set myMission = New clsMission
myMission.Load b

cMissions.Add myMission

For i = 1 To 6
Set myMission = New clsMission
myMission.Load b + i
cMissions.Add myMission
Next i


'Get number of cards in deck
Line Input #X, a$
ncards = Val(GetVal(a$))

For i = 1 To ncards

Line Input #X, a$
scardtype = GetVal(a$)

Line Input #X, a$
ncardid = Val(GetVal(a$))

Select Case scardtype

Case "Activator"
    Set myActivator = New clsActivator
    myActivator.Load ncardid
    cDrawPile.Add myActivator

Case "Artifact"
    Set myArtifact = New clsArtifact
    myArtifact.Load ncardid
    cDrawPile.Add myArtifact
    
Case "Ally Card"
    Set myAlly = New clsAlly
    myAlly.Load ncardid
    cDrawPile.Add myAlly
    
Case "Aspect Card"
    Set myAspect = New clsAspect
    myAspect.Load ncardid
    cDrawPile.Add myAspect
    
Case "Basic Universe"
    Set myBasic = New clsBasicUniverse
    myBasic.Load ncardid
    cDrawPile.Add myBasic
    
Case "Double Shot"
    Set myDoubleShot = New clsDoubleShot
    myDoubleShot.Load ncardid
    cDrawPile.Add myDoubleShot
    
Case "Event"
    Set myEvent = New clsEvent
    myEvent.Load ncardid
    cDrawPile.Add myEvent
    
Case "Power Card"
    Set myPower = New clsPowerCard
    myPower.Load ncardid
    cDrawPile.Add myPower
    
Case "Special Card"
    Set myspecial = New clsSpecial
    myspecial.Load ncardid
    cDrawPile.Add myspecial
    
Case "Teamwork"
    Set myTeamwork = New clsTeamwork
    myTeamwork.Load ncardid
    cDrawPile.Add myTeamwork
    
    
Case "Training"
    Set myTraining = New clsTraining
    myTraining.Load ncardid
    cDrawPile.Add myTraining

Case Else
End Select

Next i


'Load Battlesite deck

Line Input #X, a$
nbd = Val(GetVal(a$))

For i = 1 To nbd

Line Input #X, a$
scardtype = GetVal(a$)

Line Input #X, a$
ncardid = Val(GetVal(a$))

Select Case scardtype

Case "Special Card"
    Set myspecial = New clsSpecial
    myspecial.Load ncardid
    myBattleSite.Deck_AddCard myspecial

Case Else
End Select

Next i

Close #X

UpdateDeckDisplay
ShuffleDrawPile

LoadCharacters
'LoadOpponentCharacters

LoadHomeBase

'CurrentPhase = Draw
imgHand(0).OLEDragMode = 0

End Sub



Private Sub mnuHelpAbout_Click()
frmAbout.Show 1

End Sub

Private Sub mnuHelpRules_Click()
ret& = ShellExecute(Screen.ActiveControl.hWnd, "Open", App.Path & "\Rules.htm", "", App.Path, 1)
End Sub

Private Sub mnuHelpTopics_Click()
App.HelpFile = App.Path & "\oponline.hlp"


SendKeys "{F1}", True
End Sub

Private Sub mnuOpponent_Click()

If tcpChannel.State = sckConnected Then
    mnuToolsConnect.Enabled = False
Else

        mnuToolsConnect.Enabled = True


End If

If cDrawPileO.Count = 0 Then
    mnuOppViewDrawPile.Enabled = False
Else
    mnuOppViewDrawPile.Enabled = True
End If

If cHandO.Count = 0 Then
    mnuOppViewHand.Enabled = False
Else
    mnuOppViewHand.Enabled = True
End If

If cDeadPileO.Count = 0 Then
    mnuViewOPDeadPile.Enabled = False
Else
    mnuViewOPDeadPile.Enabled = True
End If

If cDiscardPileO.Count = 0 Then
    mnuViewOpPowerPack.Enabled = False
Else
    mnuViewOpPowerPack.Enabled = False
End If

End Sub

Private Sub mnuOppViewBattleSiteDeck_Click()
Dim ctemp As Collection
Dim ccard

If OpBattlesite.Deck_Count = 0 Then
    MsgBox sOpponentName & "does not have any cards in his Battlesite deck.", vbCritical, "No Cards"
    Exit Sub
End If

With FrmViewPile

Set ctemp = New Collection
For i = 1 To OpBattlesite.Deck_Count
Set ccard = OpBattlesite.Deck_GetCard(i)
ctemp.Add ccard
Next i

Set .ShowPile = ctemp
.PileType = 6
.Show 1

End With
Unload FrmViewPile

End Sub

Private Sub mnuOppViewDrawPile_Click()
Dim ctemp As Collection
Dim ccard

X = MsgBox("Note: You are not allowed to view your opponent's draw pile unless a special allows you to do so.  Your opponent will be informed that you are doing so.  Would you like to continue?", vbYesNoCancel, "Viewing Opponent Draw Pile")

If X <> 6 Then Exit Sub

frmViewCards.Show 1

With frmViewCards
    If .chkCancel.Value = 1 Then
        Unload frmViewCards
        Exit Sub
    End If

SendData "CVD:1:|"
History_Add "VIEWING OPPONENT DRAW PILE"

If cDrawPileO.Count = 0 Then
    MsgBox sOpponentName & " does not have any cards in hand.", vbCritical, "No Cards in Hand"
    Exit Sub
End If


Set ctemp = New Collection

If .optView(0).Value = True Then
    For i = 1 To cDrawPileO.Count
        Set ccard = cDrawPileO.Item(i)
        ctemp.Add ccard
    Next i
End If

If .optView(1).Value = True Then
    a = Val(.txtTop.Text)
    If a > cDrawPileO.Count Then a = cDrawPileO.Count
    
    For i = 1 To a
        Set ccard = cDrawPileO.Item(i)
        ctemp.Add ccard
    Next i

End If

If .optView(2).Value = True Then
    a = Val(.txtRandom.Text)
    Dim ctemp2 As Collection
    Set ctemp2 = New Collection
    For i = 1 To cDrawPileO.Count
        Set ccard = cDrawPileO.Item(i)
        ctemp2.Add ccard
    Next i

    If a >= cDrawPileO.Count Then
        Set ctemp = ctemp2
    Else
        For z = 1 To a
            cn = Int(Rnd * ctemp2.Count) + 1
            Set ccard = ctemp2.Item(cn)
            ctemp.Add ccard
            ctemp2.Remove cn
        Next z
    End If
    

End If

If .optView(3).Value = True Then
    For i = 1 To cDrawPileO.Count
        Set ccard = cDrawPileO.Item(i)
        If ccard.CardType = "Special Card" Then
            ctemp.Add ccard
        End If
    Next i
End If

If .optView(4).Value = True Then
    For i = 1 To cDrawPileO.Count
        Set ccard = cDrawPileO.Item(i)
        If ccard.CardType = "Ally Card" Or ccard.CardType = "Teamwork" Or ccard.CardType = "Basic Universe" Then
            ctemp.Add ccard
        End If
    Next i
End If

If .optView(5).Value = True Then

    If .lstPowerTypes.ListIndex = 0 Then
    
    For i = 1 To cDrawPileO.Count
        Set ccard = cDrawPileO.Item(i)
        If ccard.CardType = "Power Card" Then
            ctemp.Add ccard
        End If
    Next i
    
    
    Else

    Select Case .lstPowerTypes.ListIndex
    Case 1
        pt$ = "Energy"
    Case 2
        pt$ = "Fighting"
    Case 3
        pt$ = "Strength"
    Case 4
        pt$ = "Intellect"
    End Select
    
    For i = 1 To cDrawPileO.Count
        Set ccard = cDrawPileO.Item(i)
        If ccard.CardType = "Power Card" Then
        
            If ccard.PowerType = pt$ Or ccard.PowerType = "Multi-Power" Then
            
            ctemp.Add ccard
            
            End If
            
        End If
    Next i
    

    End If
    
End If

End With

Unload frmViewCards


With FrmViewPile

Set .ShowPile = ctemp
.PileType = 6
.Show 1

End With

End Sub

Private Sub mnuOppViewHand_Click()
Dim ctemp As Collection
Dim ccard

If cHandO.Count = 0 Then
    MsgBox sOpponentName & " does not have any cards in hand.", vbCritical, "No Cards in Hand"
    Exit Sub
End If

If bOppOpenHanded = True Then
    
    With FrmViewPile
    
    Set ctemp = New Collection
    For i = 1 To cHandO.Count
    Set ccard = cHandO.Item(i)
    ctemp.Add ccard
    Next i
    
    Set .ShowPile = ctemp
    .PileType = 6
    .Show 1
    
    End With

    SendData "CV4:1:|"

Else

    X = MsgBox("Note: You are not allowed to view your opponent's hand unless a special allows you to do so.  Your opponent will be informed that you are doing so.  Would you like to continue?", vbYesNoCancel, "Viewing Opponent Draw Pile")

    If X <> 6 Then Exit Sub

    SendData "CV4:1:|"


    frmViewCards.Show 1

    With frmViewCards
        If .chkCancel.Value = 1 Then
            Unload frmViewCards
            Exit Sub
        End If

    Set ctemp = New Collection

    If .optView(0).Value = True Then
        For i = 1 To cHandO.Count
            Set ccard = cHandO.Item(i)
            ctemp.Add ccard
        Next i
    End If

    If .optView(1).Value = True Then
        a = Val(.txtTop.Text)
        If a > cHandO.Count Then a = cHandO.Count
    
        For i = 1 To a
            Set ccard = cHandO.Item(i)
            ctemp.Add ccard
        Next i

    End If

    If .optView(2).Value = True Then
        a = Val(.txtRandom.Text)
        Dim ctemp2 As Collection
        Set ctemp2 = New Collection
        For i = 1 To cHandO.Count
            Set ccard = cHandO.Item(i)
            ctemp2.Add ccard
        Next i

        If a >= cHandO.Count Then
            Set ctemp = ctemp2
        Else
        For z = 1 To a
            cn = Int(Rnd * ctemp2.Count) + 1
            Set ccard = ctemp2.Item(cn)
            ctemp.Add ccard
            ctemp2.Remove cn
        Next z
    End If
    

    End If

If .optView(3).Value = True Then
    For i = 1 To cHandO.Count
        Set ccard = cHandO.Item(i)
        If ccard.CardType = "Special Card" Then
            ctemp.Add ccard
        End If
    Next i
End If

If .optView(4).Value = True Then
    For i = 1 To cHandO.Count
        Set ccard = cHandO.Item(i)
        If ccard.CardType = "Ally Card" Or ccard.CardType = "Teamwork" Or ccard.CardType = "Basic Universe" Then
            ctemp.Add ccard
        End If
    Next i
End If

If .optView(5).Value = True Then

    If .lstPowerTypes.ListIndex = 0 Then
    
    For i = 1 To cHandO.Count
        Set ccard = cHandO.Item(i)
        If ccard.CardType = "Power Card" Then
            ctemp.Add ccard
        End If
    Next i
    
    
    Else

    Select Case .lstPowerTypes.ListIndex
    Case 1
        pt$ = "Energy"
    Case 2
        pt$ = "Fighting"
    Case 3
        pt$ = "Strength"
    Case 4
        pt$ = "Intellect"
    End Select
    
    For i = 1 To cHandO.Count
        Set ccard = cHandO.Item(i)
        If ccard.CardType = "Power Card" Then
        
            If ccard.PowerType = pt$ Or ccard.PowerType = "Multi-Power" Then
            
            ctemp.Add ccard
            
            End If
            
        End If
    Next i
    

    End If
    
End If
    End With
    
    Unload frmViewCards
   
    
    With FrmViewPile
    
    Set .ShowPile = ctemp
    .PileType = 6
    .Show 1
    
    End With

End If
    


End Sub

Private Sub mnuPassToOpponent_Click()
X = MsgBox("Are you sure you want to pass to your opponent?  If you do, you will only be able to defend for the remainder of the round.", vbYesNoCancel, "Confirm Pass to Opponent")

If X <> 6 Then Exit Sub

bHavePassed = True

SendData "CPB:1:|"

If bOppPassed = True Then
    myPhase = nPhase_Resolve
    ShowResolveVentureFrame
    UpdatePhase
    History_Add "RESOLVE VENTURE"
    
Else
    myPhase = nPhase_Defend
    History_Add mySettings.PlayerName & " HAS PASSED TO OPPONENT"
    History_Add sOpponentName & " IS PREPARING TO ATTACK"
    UpdatePhase

End If

End Sub

Private Sub mnuPlayOpenHanded_Click()

If mnuPlayOpenHanded.Checked = True Then
    mnuPlayOpenHanded.Checked = False
    SendData "CPO:0:|"
Else
    mnuPlayOpenHanded.Checked = True
    SendData "CPO:1:|"
End If

End Sub



Private Sub mnuShowDuplicates_Click()

CheckForDupes
End Sub
Private Sub CheckForDupes()
Dim ccard
Dim ccard2
Dim bdupefound As Boolean

On Error Resume Next

For i = 1 To cHand.Count

Set ccard = cHand.Item(i)

'loop through rest of hand

    For k = (i + 1) To cHand.Count
      Set ccard2 = cHand.Item(k)
      
      If ccard2.CardType = ccard.CardType Then
        If ccard.ID = ccard2.ID Then
        
            msg$ = "Duplicate " & ccard.CardType & "s found." & vbCrLf
            msg$ = msg$ & Trim(Str(i)) & ". " & ccard.Title
            msg$ = msg$ & vbCrLf & Trim(Str(k)) & ". " & ccard2.Title
            bdupefound = True
            
            
        GoTo dupefound
        
        End If
        
        If ccard.Title = ccard2.Title Then
            msg$ = "Duplicate " & ccard.CardType & "s found." & vbCrLf
            msg$ = msg$ & Trim(Str(i)) & ". " & ccard.Title
            msg$ = msg$ & vbCrLf & Trim(Str(k)) & ". " & ccard2.Title
            bdupefound = True
            
            
        GoTo dupefound
        
        End If
        
        If ccard.CardType = "Power Card" Then
        If ccard.Power = ccard2.Power Then
            msg$ = "Duplicate Power Cards (" & ccard.Power & ") found." & vbCrLf
            msg$ = msg$ & Trim(Str(i)) & ". " & ccard.Title
            msg$ = msg$ & vbCrLf & Trim(Str(k)) & ". " & ccard2.Title
            bdupefound = True
           
        GoTo dupefound
        
        End If
        End If
        
    End If

    Next k
    
    For k = 1 To i - 1
    
    
    Next k
    

Next i

'Check for unusables

For k = 1 To 4

    If cFrontLine.isCharacterDead(k) = True Then
        For i = 1 To cHand.Count
            Set ccard = cHand.Item(i)
            If ccard.CardType = "Special Card" Then
            
                If ccard.Character = cFrontLine.Character_Name(k) Then
                                
                GoTo unusablefound
            End If
            End If

        Next i
    End If
    
Next k

Exit Sub

unusablefound:

msg$ = "Unusable card: " & vbCrLf
msg$ = msg$ & Trim(Str(i)) & ". " & ccard.Title
MsgBox msg$, vbInformation, "Unusable Found"

Exit Sub

MsgBox "No Duplicates or Unusables Found", vbInformation, "No Dupes Found"
Exit Sub

dupefound:
    MsgBox msg$, vbInformation, "Duplicates Found"
    
End Sub
Private Sub mnuShuffleDeadPile_Click()
Status "Shuffling Dead pile."

ShufflePile 2
UpdateDeckDisplay
Status "Dead pile shuffled."

SendData "CDD:" & GetCode_CardString(cDeadPile) & "|"

End Sub

Private Sub mnuShuffleDiscardPile_Click()
Status "Shuffling Discard pile."

ShufflePile 1
UpdateDeckDisplay
Status "Discard pile shuffled."

SendData "CDI:" & GetCode_CardString(cDiscardPile) & "|"

End Sub

Private Sub mnuShuffleDeadIntoDraw_Click()
For i = 1 To cDeadPile.Count
    cDrawPile.Add cDeadPile.Item(i)
Next i

Set cDeadPile = New Collection
ShufflePile 0
UpdateDeckDisplay

SendData "CDD:|"
SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"

End Sub

Private Sub mnuShuffleDiscardsIntoDraw_Click()

For i = 1 To cDiscardPile.Count
    cDrawPile.Add cDiscardPile.Item(i)
Next i

Set cDiscardPile = New Collection
ShufflePile 0
UpdateDeckDisplay

SendData "CDX:1:|"
SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"

End Sub

Private Sub mnuShuffleDrawPile_Click()

Status "Shuffling draw pile."

ShufflePile 0
UpdateDeckDisplay
Status "Draw pile shuffled."

'Send Draw Pile
SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"


End Sub
Private Sub Status(sString)
sbBar1.Panels(1).Text = sString

End Sub
Private Sub FetchHandImages()
'Clear current hand
For i = 1 To imgTemp.Count - 1
    Unload imgTemp(i)
Next i

For i = 1 To cHand.Count
   
    Load imgTemp(i)
    imgTemp(i).Left = imgTemp(i - 1).Left + 400
    imgTemp(i).Visible = True
    imgTemp(i).Tag = cHandTags.Item(i)
    
    If cHand.Item(i).LoadImage(cHand.Item(i).ID) = True Then
        imgTemp(i).Picture = LoadPicture(App.Path & "\temppic.jpg")
    Else
        imgTemp(i).Picture = LoadPicture(sBlankImagePath)
    End If
    
    'Set tool tip texts
    Set ccard = cHand.Item(i)
    
    Select Case ccard.CardType
    Case "Special Card", "Artifact", "Aspect Card"
        imgTemp(i).ToolTipText = ccard.Title & "-->" & ccard.Effect
    Case "Power Card", "Ally Card", "Training Card", "Teamwork", "Double Shot"
        imgTemp(i).ToolTipText = ccard.Title
    Case "Event"
        imgTemp(i).ToolTipText = ccard.Title & "-->" & ccard.Description
    Case "Activator"
        Counter = 0
        a$ = "XXA"
        
        For z = 1 To myBattleSite.Deck_Count
            Set myspecial = myBattleSite.Deck_GetCard(z)
            If myspecial.Character = ccard.Name Then
                Counter = Counter + 1
                a$ = a$ & Trim(Str(Counter)) & ". " & myspecial.Name & ": " & myspecial.Effect & "~" & Trim(Str(z)) & "^"
            End If
        Next z
        
        imgTemp(i).ToolTipText = a$
        
    Case Else
        imgTemp(i).ToolTipText = ""
        
    End Select
    
Next i

ShowHand

End Sub
Private Sub ShowHand()
'Clear current hand
For i = 1 To imgHand.Count - 1
    Unload imgHand(i)
Next i

Select Case cHand.Count
Case 0 To 8
    cspacer = 450
Case 9
    cspacer = -150
Case 10
    cspacer = -300
Case 11
    cspacer = -400
Case 12
    cspacer = -550
Case 13
    cspacer = -700
Case 14 To 100
    cspacer = -800
Case Else
End Select

lstHandTips.Clear

For i = 1 To cHand.Count
    Load imgHand(i)
    'imgHand(i).Visible = True
    imgHand(i).ZOrder (0)
    
    If i = 1 Then
        imgHand(i).Top = 0
        'imgHand(i).Left = 120
    Else
        imgHand(i).Top = imgHand(i - 1).Top + 600 + cspacer
        'imgHand(i).Left = imgHand(i - 1).Left + 1455 + cspacer
    End If
    
    For k = 1 To imgTemp.Count - 1
        If imgTemp(k).Tag = cHandTags.Item(i) Then
            imgHand(i).Picture = imgTemp(k).Picture
            lstHandTips.AddItem imgTemp(k).ToolTipText
            
        End If
    Next k
    
    imgHand(i).Tag = cHandTags.Item(i)
    imgHand(i).Refresh
    DoEvents
      
Next i

If cHand.Count > 0 Then
    HandClick 1

Else
    imgCardDetail.Picture = LoadPicture(sBlankImagePath)
    Me.Caption = "OVERPOWER ONLINE-->" & "NO CARD SELECTED"
    imgCardDetail.Tag = ""
    HideFrames False, False, True
    HideAllBorders
End If

For i = 1 To cHand.Count
    imgHand(i).Visible = True
Next i

End Sub
Private Sub NewGame()
ClearTable
ClearCollections
ClearOpponentCollections

UpdateDeckDisplay
UpdateOpponentDeckDisplay


End Sub
Private Sub LoadHomeBase()
'Load HomeBase

If myHomebase.ID > 0 Then
    If myHomebase.LoadImage(myHomebase.ID) = True Then
        imgHomebase.Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgHomebase.ToolTipText = myHomebase.Effect
        
    Else
        imgHomebase.Picture = LoadPicture(sBlankImagePath)
    End If
End If


End Sub
Private Sub loadbattlesite()

'load battlesite
If myBattleSite.ID > 0 Then
'MsgBox myBattleSite.LoadImage(myBattleSite.ID)
'End

    If myBattleSite.LoadImage(myBattleSite.ID) = True Then
        imgBattlesite.Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgBattlesite.ToolTipText = myBattleSite.Name & ": " & myBattleSite.Effect
        
    Else
        imgBattlesite.Picture = LoadPicture(sBlankImagePath)
    End If
Else
    imgBattlesite.Picture = Nothing
    imgBattlesite.ToolTipText = "No current Battlesite"
    HeroClick 0
End If

End Sub
Private Sub LoadOpBattlesite()
'load battlesite
If OpBattlesite.ID > 0 Then

    If OpBattlesite.LoadImage(OpBattlesite.ID) = True Then
        imgOpBattlesite.Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgOpBattlesite.ToolTipText = OpBattlesite.Name & ": " & OpBattlesite.Effect
        
    Else
        imgOpBattlesite.Picture = LoadPicture(sBlankImagePath)
    End If
Else
    imgOpBattlesite.Picture = Nothing
    imgOpBattlesite.ToolTipText = "No current Battlesite"
End If

End Sub
Private Sub LoadCharacters()

imgFrontLine(0).Picture = Nothing
imgFrontLine(1).Picture = Nothing
imgFrontLine(2).Picture = Nothing
imgFrontLine(0).Tag = "0"
imgFrontLine(1).Tag = "0"
imgFrontLine(2).Tag = "0"
ImgReserve.Tag = "0"
ImgReserve.Picture = Nothing

cfc = 0

'Do frontline first

For i = 1 To 4

    If cFrontLine.isCharacterDead(i) = False Then
    
        If cFrontLine.isCharacterReserve(i) = False Then
        
            If cFrontLine.LoadImage(i) = True Then
                imgFrontLine(cfc).Picture = LoadPicture(App.Path & "\temppic.jpg")
            Else
                imgFrontLine(cfc).Picture = LoadPicture(sBlankImagePath)
            End If
        
        imgFrontLine(cfc).Tag = i
        cfc = cfc + 1
        
        Else
        
            If cFrontLine.LoadImage(i) = True Then
                ImgReserve.Picture = LoadPicture(App.Path & "\temppic.jpg")
            Else
                ImgReserve.Picture = LoadPicture(sBlankImagePath)
            End If
            
        ImgReserve.Tag = i
        
        End If

    
    End If

Next i


End Sub
Private Sub LoadOpponentCharacters()
imgOpponent(0).Picture = Nothing
imgOpponent(1).Picture = Nothing
imgOpponent(2).Picture = Nothing
imgOpponent(0).Tag = "0"
imgOpponent(1).Tag = "0"
imgOpponent(2).Tag = "0"
imgOppReserve.Tag = "0"
imgOppReserve.Picture = Nothing

cfc = 0

'Do frontline first

For i = 1 To 4

    If cOpponent.isCharacterDead(i) = False Then
    
        If cOpponent.isCharacterReserve(i) = False Then
        
            If cOpponent.LoadImage(i) = True Then
                imgOpponent(cfc).Picture = LoadPicture(App.Path & "\temppic.jpg")
            Else
                imgOpponent(cfc).Picture = LoadPicture(sBlankImagePath)
            End If
        
        imgOpponent(cfc).Tag = i
        cfc = cfc + 1
        
        Else
        
            If cOpponent.LoadImage(i) = True Then
                imgOppReserve.Picture = LoadPicture(App.Path & "\temppic.jpg")
            Else
                imgOppReserve.Picture = LoadPicture(sBlankImagePath)
            End If
            
        imgOppReserve.Tag = i
        
        End If

    
    End If

Next i


End Sub
Private Sub ShowHitsToCurrentBattle()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgHit1.Count - 1
    Unload imgHit1(z)
Next z

For z = 1 To imgHit2.Count - 1
    Unload imgHit2(z)
Next z

For z = 1 To imgHit3.Count - 1
    Unload imgHit3(z)
Next z

For z = 1 To imgHit4.Count - 1
    Unload imgHit4(z)
Next z

For z = 1 To imgHitBS.Count - 1
    Unload imgHitBS(z)
Next z

For i = 1 To 4

'If cFrontLine.isCharacterDead(i) = False Then

    Select Case CharPic(i)
    Case 0
        Set pCard = imgHit1
    Case 1
        Set pCard = imgHit2
    Case 2
        Set pCard = imgHit3
    Case 3
        Set pCard = imgHit4
    End Select
    
    For k = 1 To cFrontLine.HitsToCurrentBattle_Count(i)
        Load pCard(k)
        Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(i, k)
                                
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(App.Path & "\NotFound.jpg")
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
        pCard(k).ToolTipText = ccard.Effect
        Else
        pCard(k).ToolTipText = ccard.Title
        End If
        
        
    
    Next k

'End If

Next i

For k = 1 To myBattleSite.HitsToCurrentBattle_Count
Load imgHitBS(k)
Set ccard = myBattleSite.HitsToCurrentBattle_GetCard(k)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgHitBS(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
End If

With imgHitBS(k)
.Left = imgHitBS(k - 1).Left + 200
.ZOrder 0
.Visible = True

If ccard.CardType = "Special Card" Or ccard.CardType = "Aspect Card" Then
    .ToolTipText = ccard.Effect
Else
    .ToolTipText = ccard.Title
End If

End With

Next k

End Sub
Private Sub ShowPermanentRecord()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgPR1.Count - 1
    Unload imgPR1(z)
Next z

For z = 1 To imgPR2.Count - 1
    Unload imgPR2(z)
Next z

For z = 1 To imgPR3.Count - 1
    Unload imgPR3(z)
Next z

For z = 1 To imgPR4.Count - 1
    Unload imgPR4(z)
Next z

For z = 1 To imgPRBS.Count - 1
    Unload imgPRBS(z)
Next z

For i = 1 To 4

If cFrontLine.isCharacterDead(i) = False Then

    Select Case CharPic(i)
    Case 0
        Set pCard = imgPR1
    Case 1
        Set pCard = imgPR2
    Case 2
        Set pCard = imgPR3
    Case 3
        Set pCard = imgPR4
    End Select
    
    For k = 1 To cFrontLine.PermanentRecord_Count(i)
        Load pCard(k)
        Set ccard = cFrontLine.PermanentRecord_GetCard(i, k)
                                
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
        pCard(k).ToolTipText = ccard.Effect
        Else
        pCard(k).ToolTipText = ccard.Title
        End If
      
    
    Next k

End If

    Select Case CharPic(i)
    Case 0
        imgFrontLine(0).ZOrder 0
    Case 1
        imgFrontLine(1).ZOrder 0
    Case 2
        imgFrontLine(2).ZOrder 0
    Case 3
        ImgReserve.ZOrder 0
    End Select

Next i

For k = 1 To myBattleSite.PermanentRecord_Count
Load imgPRBS(k)
Set ccard = myBattleSite.PermanentRecord_GetCard(k)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgPRBS(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgPRBS(k).Picture = LoadPicture(sBlankImagePath)
End If

With imgPRBS(k)
.Left = imgPRBS(k - 1).Left + 200
.ZOrder 0
.Visible = True

If ccard.CardType = "Special Card" Or ccard.CardType = "Aspect Card" Then
    .ToolTipText = ccard.Effect
Else
    .ToolTipText = ccard.Title
End If

End With

Next k

imgBattlesite.ZOrder 0

End Sub
Private Sub ShowOpponentPlacedCards()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgOpPlace1.Count - 1
    Unload imgOpPlace1(z)
Next z

For z = 1 To imgOpPlace2.Count - 1
    Unload imgOpPlace2(z)
Next z

For z = 1 To imgOpPlace3.Count - 1
    Unload imgOpPlace3(z)
Next z

For z = 1 To imgOpPlace4.Count - 1
    Unload imgOpPlace4(z)
Next z

For z = 1 To imgOpPlacedHomebase.Count - 1
    Unload imgOpPlacedHomebase(z)
Next z

For i = 1 To 4

If cOpponent.isCharacterDead(i) = False Then
    
    Select Case OppCharPic(i)
    Case 0
        Set pCard = imgOpPlace1
    Case 1
        Set pCard = imgOpPlace2
    Case 2
        Set pCard = imgOpPlace3
    Case 3
        Set pCard = imgOpPlace4
    End Select
    
    For k = 1 To cOpponent.Placed_Count(i)
        Load pCard(k)
        
        
        Set ccard = cOpponent.PlacedCard(i, k)
                        
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
            pCard(k).ToolTipText = ccard.Effect
        Else
            pCard(k).ToolTipText = ccard.Title
        End If
    
    Next k

End If

Next i

On Error Resume Next
For k = 1 To OpHomebase.Placed_Count
Load imgOpPlacedHomebase(k)
Set ccard = OpHomebase.PlacedCard(k)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgOpPlacedHomebase(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgOpPlacedHomebase(k).Picture = LoadPicture(sBlankImagePath)
End If

With imgOpPlacedHomebase(k)
.Left = imgOpPlacedHomebase(k - 1).Left + 200
.ZOrder 0
.Visible = True

If ccard.CardType = "Special Card" Or ccard.CardType = "Aspect Card" Then
    .ToolTipText = ccard.Effect
Else
    .ToolTipText = ccard.Title
End If

End With

Next k


End Sub
Private Sub ShowPlacedCards()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgPlaced1.Count - 1
    Unload imgPlaced1(z)
Next z

For z = 1 To imgPlaced2.Count - 1
    Unload imgPlaced2(z)
Next z

For z = 1 To imgPlaced3.Count - 1
    Unload imgPlaced3(z)
Next z

For z = 1 To imgPlaced4.Count - 1
    Unload imgPlaced4(z)
Next z

For z = 1 To imgPlacedHomeBase.Count - 1
    Unload imgPlacedHomeBase(z)
Next z

For i = 1 To 4

If cFrontLine.isCharacterDead(i) = False Then
    
    Select Case CharPic(i)
    Case 0
        Set pCard = imgPlaced1
    Case 1
        Set pCard = imgPlaced2
    Case 2
        Set pCard = imgPlaced3
    Case 3
        Set pCard = imgPlaced4
    End Select
    
    For k = 1 To cFrontLine.Placed_Count(i)
        Load pCard(k)
        Set ccard = cFrontLine.PlacedCard(i, k)
                        
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
            pCard(k).ToolTipText = ccard.Effect
        Else
            pCard(k).ToolTipText = ccard.Title
        End If
        
        
    
    Next k

End If

Next i

On Error Resume Next
For k = 1 To myHomebase.Placed_Count
Load imgPlacedHomeBase(k)
Set ccard = myHomebase.PlacedCard(k)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgPlacedHomeBase(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgPlacedHomeBase(k).Picture = LoadPicture(sBlankImagePath)
End If

With imgPlacedHomeBase(k)
.Left = imgPlacedHomeBase(k - 1).Left + 200
.ZOrder 0
.Visible = True

If ccard.CardType = "Special Card" Or ccard.CardType = "Aspect Card" Then
    .ToolTipText = ccard.Effect
Else
    .ToolTipText = ccard.Title
End If

End With

Next k

End Sub
Private Sub ShowOpponentModifiers()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgOpEffect1.Count - 1
    Unload imgOpEffect1(z)
Next z

For z = 1 To imgOpEffect2.Count - 1
    Unload imgOpEffect2(z)
Next z

For z = 1 To imgOpEffect3.Count - 1
    Unload imgOpEffect3(z)
Next z

For z = 1 To imgOpEffect4.Count - 1
    Unload imgOpEffect4(z)
Next z


For i = 1 To 4

If cOpponent.isCharacterDead(i) = False Then
    
    Select Case OppCharPic(i)
    Case 0
        Set pCard = imgOpEffect1
    Case 1
        Set pCard = imgOpEffect2
    Case 2
        Set pCard = imgOpEffect3
    Case 3
        Set pCard = imgOpEffect4
    End Select
    
    For k = 1 To cOpponent.Modifiers_Count(i)
        Load pCard(k)
        Set ccard = cOpponent.Modifiers_GetCard(i, k)
                                
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
            pCard(k).ToolTipText = ccard.Effect
        Else
            pCard(k).ToolTipText = ccard.Title
        End If
    
    Next k

End If

Next i
End Sub
Private Sub ShowModifiers()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgBGameEffect1.Count - 1
    Unload imgBGameEffect1(z)
Next z

For z = 1 To imgBGameEffect2.Count - 1
    Unload imgBGameEffect2(z)
Next z

For z = 1 To imgBGameEffect3.Count - 1
    Unload imgBGameEffect3(z)
Next z

For z = 1 To imgBGameEffect4.Count - 1
    Unload imgBGameEffect4(z)
Next z


For i = 1 To 4

If cFrontLine.isCharacterDead(i) = False Then
    
    Select Case CharPic(i)
    Case 0
        Set pCard = imgBGameEffect1
    Case 1
        Set pCard = imgBGameEffect2
    Case 2
        Set pCard = imgBGameEffect3
    Case 3
        Set pCard = imgBGameEffect4
    End Select
    
    For k = 1 To cFrontLine.Modifiers_Count(i)
        Load pCard(k)
        Set ccard = cFrontLine.Modifiers_GetCard(i, k)
                                
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
            pCard(k).ToolTipText = ccard.Effect
        Else
            pCard(k).ToolTipText = ccard.Title
        End If
    
    Next k

End If

Next i

End Sub
Private Sub ShowOpponentHTCB()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgOPHit1.Count - 1
    Unload imgOPHit1(z)
Next z

For z = 1 To imgOPHit2.Count - 1
    Unload imgOPHit2(z)
Next z

For z = 1 To imgOPHit3.Count - 1
    Unload imgOPHit3(z)
Next z

For z = 1 To imgOPHit4.Count - 1
    Unload imgOPHit4(z)
Next z

For z = 1 To imgHitOpBS.Count - 1
    Unload imgHitOpBS(z)
Next z

For i = 1 To 4

If cOpponent.isCharacterDead(i) = False Then
    
    Select Case OppCharPic(i)
    Case 0
        Set pCard = imgOPHit1
    Case 1
        Set pCard = imgOPHit2
    Case 2
        Set pCard = imgOPHit3
    Case 3
        Set pCard = imgOPHit4
    End Select
    
    For k = 1 To cOpponent.HitsToCurrentBattle_Count(i)
        Load pCard(k)
        Set ccard = cOpponent.HitsToCurrentBattle_GetCard(i, k)
                        
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
        pCard(k).ToolTipText = ccard.Effect
        Else
        pCard(k).ToolTipText = ccard.Title
        End If
        
        
    
    Next k

End If

Next i

On Error Resume Next
For k = 1 To OpBattlesite.HitsToCurrentBattle_Count
Load imgHitOpBS(k)
Set ccard = OpBattlesite.HitsToCurrentBattle_GetCard(k)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgHitOpBS(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgHitOpBS(k).Picture = LoadPicture(sBlankImagePath)
End If

With imgHitOpBS(k)
.Left = imgHitOpBS(k - 1).Left + 200
.ZOrder 0
.Visible = True

If ccard.CardType = "Special Card" Or ccard.CardType = "Aspect Card" Then
    .ToolTipText = ccard.Effect
Else
    .ToolTipText = ccard.Title
End If

End With

Next k

End Sub
Private Sub ShowOpponentPermanentRecord()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgPROP1.Count - 1
    Unload imgPROP1(z)
Next z

For z = 1 To imgPROP2.Count - 1
    Unload imgPROP2(z)
Next z

For z = 1 To imgPROP3.Count - 1
    Unload imgPROP3(z)
Next z

For z = 1 To imgPROP4.Count - 1
    Unload imgPROP4(z)
Next z

For z = 1 To imgPROPBS.Count - 1
    Unload imgPROPBS(z)
Next z

For i = 1 To 4

If cOpponent.isCharacterDead(i) = False Then
    
    Select Case OppCharPic(i)
    Case 0
        Set pCard = imgPROP1
    Case 1
        Set pCard = imgPROP2
    Case 2
        Set pCard = imgPROP3
    Case 3
        Set pCard = imgPROP4
    End Select
    
    For k = 1 To cOpponent.PermanentRecord_Count(i)
        Load pCard(k)
        Set ccard = cOpponent.PermanentRecord_GetCard(i, k)
                                
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
'        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
        pCard(k).ToolTipText = ccard.Effect
        Else
        pCard(k).ToolTipText = ccard.Title
        End If
        
        
    
    Next k

End If

Next i

For k = 1 To OpBattlesite.PermanentRecord_Count
Load imgPROPBS(k)

Set ccard = OpBattlesite.PermanentRecord_GetCard(k)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgPROPBS(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgPROPBS(k).Picture = LoadPicture(sBlankImagePath)
End If

With imgPROPBS(k)
.Left = imgPROPBS(k - 1).Left + 200
.Visible = True

If ccard.CardType = "Special Card" Or ccard.CardType = "Aspect Card" Then
    .ToolTipText = ccard.Effect
Else
    .ToolTipText = ccard.Title
End If

End With

Next k

End Sub
Private Function CharPic(Index) As Integer

If Val(imgFrontLine(0).Tag) = Index Then CharPic = 0
If Val(imgFrontLine(1).Tag) = Index Then CharPic = 1
If Val(imgFrontLine(2).Tag) = Index Then CharPic = 2
If Val(ImgReserve.Tag) = Index Then CharPic = 3


End Function
Private Function OppCharPic(Index) As Integer

If Val(imgOpponent(0).Tag) = Index Then OppCharPic = 0
If Val(imgOpponent(1).Tag) = Index Then OppCharPic = 1
If Val(imgOpponent(2).Tag) = Index Then OppCharPic = 2
If Val(imgOppReserve.Tag) = Index Then OppCharPic = 3

End Function

Private Sub mnuStopPass_Click()

bHavePassed = False
SendData "CSP:1:|"

End Sub

Private Sub mnuTools_Click()

If cDrawPileO.Count = 0 Then
    mnuToolsWhoGoesFirst.Enabled = False
Else
    mnuToolsWhoGoesFirst.Enabled = True
End If

If cDrawPile.Count = 0 Then
    mnuDrawTestHand.Enabled = False
Else
    mnuDrawTestHand.Enabled = True
End If

If myPhase <> nPhase_WhoGoesFirst Then
    mnuToolsWhoGoesFirst.Enabled = False
End If
End Sub

Private Sub mnuToolsConnect_Click()

With frmConnect
.Show 1


Select Case Val(.lblResult.Caption)

Case 0
'canceled
    Unload frmConnect
    Exit Sub

Case 1
'standard connect - new game

If .optCType(0).Value = True Then
    bHost = True

    tcpChannel.Close
    tcpChannel.LocalPort = .txtPort.Text
    tcpChannel.Listen
Else
    bHost = False

    tcpChannel.RemoteHost = .cbIP.Text
    tcpChannel.RemotePort = .txtPort2.Text
    tcpChannel.Connect

End If

Unload frmConnect

History_Add "LISTENING FOR OPPONENT"

Me.Refresh

mnuToolsConnect.Enabled = False

nTurn = 1

Case 2
'Resume Game

bResuming = True

Status "Preparing to resume game"

LoadResumeInfo .lstGames.List(.lstGames.ListIndex)

If .optCType2(0).Value = True Then
    bHost = True

    tcpChannel.Close
    tcpChannel.LocalPort = .txtPort2.Text
    tcpChannel.Listen
Else
    bHost = False

    tcpChannel.RemoteHost = .cbIP2.Text
    tcpChannel.RemotePort = .txtPort4.Text
    tcpChannel.Connect

'    tcpChannel.RemoteHost = mySettings.IP_Address
'    tcpChannel.RemotePort = mySettings.Port
'    tcpChannel.Connect

End If

Unload frmConnect

History_Add "LISTENING FOR OPPONENT"

Me.Refresh

mnuToolsConnect.Enabled = False

Case 3
'Auto Connect
ip$ = ""
pt = 1544
bHost = False

X = FreeFile
Open "c:\autogame.ini" For Input As #X

Do Until EOF(X)

Line Input #X, a$

z = InStr(a$, "=")

If z > 0 Then

n$ = LCase(Left(a$, z - 1))

If n$ = "port" Then pt = Val(Right(a$, Len(a$) - z))

If n$ = "ip" Then ip$ = Right(a$, Len(a$) - z)

If n$ = "host" Then

    If Right(a$, Len(a$) - z) = "TRUE" Then
        bHost = True
    Else
        bHost = False
   
    End If

End If
End If

Loop

Close #X

If ip$ = "" Then
    MsgBox "Invalid connection information.  Connection canceled.", vbCritical, "Connection Canceled."
    Exit Sub
End If

If bHost = True Then
    tcpChannel.Close
    tcpChannel.LocalPort = pt
    tcpChannel.Listen
    
Else
    
    tcpChannel.RemoteHost = ip$
    tcpChannel.RemotePort = pt
    tcpChannel.Connect

End If

Unload frmConnect

History_Add "LISTENING FOR OPPONENT"

Case Else
End Select

End With

End Sub

Private Sub mnuToolsSettings_Click()

frmSettings.Show 1

End Sub

Private Sub mnuToolsSpecialEditor_Click()
frmSpecialEditor.Show 1

End Sub

Private Sub mnuToolsWhoGoesFirst_Click()
If nTurn > 1 Then Exit Sub

WhoGoesFirst

End Sub

Private Sub mnuViewBattleSiteDeck_Click()
ViewBattleSiteDeck

End Sub
Private Sub ViewBattleSiteDeck()
Dim ctemp As Collection
Dim ccard

If myBattleSite.Deck_Count = 0 Then
    MsgBox "You do not currently have any cards in your Battlesite deck.", vbCritical, "No Cards"
    Exit Sub
End If

With FrmViewPile

Set ctemp = New Collection
For i = 1 To myBattleSite.Deck_Count
Set ccard = myBattleSite.Deck_GetCard(i)
ctemp.Add ccard
Next i

Set .ShowPile = ctemp
.PileType = 3
.Show 1

If .AddedToMyHand = True Then
    FetchHandImages
End If

End With
Unload FrmViewPile

ShowHand
UpdateDeckDisplay
End Sub
Private Sub mnuViewDeadPile_Click()

SendData "CV7:1:|"

With FrmViewPile
Set .ShowPile = cDeadPile
.PileType = 2
.Show 1

If .AddedToMyHand = True Then
    FetchHandImages
End If

End With
Unload FrmViewPile

ShowHand
UpdateDeckDisplay
End Sub

Private Sub mnuViewDefeatedCharacters_Click()
a = cFrontLine.LiveCharacterCount


With FrmViewPile
Set .ShowPile = cDefeatedCharactersPile
.PileType = 4
.Show 1

If .AddedToMyHand = True Then
    FetchHandImages
End If

End With

If a <> cFrontLine.LiveCharacterCount Then
        SendData "CRC:" & Trim(Str(Val(FrmViewPile.cmdResurrectChar.Tag))) & ":|"
        History_Add "RESURRECTED: " & cFrontLine.Character_Name(Val(FrmViewPile.cmdResurrectChar.Tag))
End If
Unload FrmViewPile

ShowHand
LoadCharacters
ShowPermanentRecord
ShowHitsToCurrentBattle
ShowPlacedCards
ShowModifiers
ShowBuffers

UpdateDeckDisplay
End Sub

Private Sub mnuViewDiscardPile_Click()

SendData "CV6:1:|"

With FrmViewPile
Set .ShowPile = cDiscardPile
.PileType = 1
.Show 1

If .AddedToMyHand = True Then
    FetchHandImages
End If

End With
Unload FrmViewPile

ShowHand
UpdateDeckDisplay
End Sub

Private Sub mnuViewDrawPile_Click()

SendData "CV5:1:|"

With FrmViewPile
Set .ShowPile = cDrawPile
.PileType = 0
.Show 1

If .AddedToMyHand = True Then
    FetchHandImages
End If

End With
Unload FrmViewPile

ShowHand
UpdateDeckDisplay
End Sub
Private Sub MoveVentureCards(cCollectionFrom As Collection, cCollectionTo As Collection, bMoveAll As Boolean, imgSource As Control, sCaption As String)

If bMoveAll = True Then

For i = cCollectionFrom.Count To 1 Step -1
cCollectionTo.Add cCollectionFrom.Item(i)
cCollectionFrom.Remove i
Next i

UpdateDeckDisplay
HeroClick 0

Else

cCollectionTo.Add cCollectionFrom.Item(cCollectionFrom.Count)
cCollectionFrom.Remove cCollectionFrom.Count
UpdateDeckDisplay

If cCollectionFrom.Count > 0 Then
    imgMissionCard.Picture = imgSource.Picture
    Me.Caption = "OVERPOWER ONLINE-->" & sCaption & " (" & cCollectionFrom.Count & ")"
Else
    imgMissionCard.Picture = Nothing
    HeroClick 0
End If
End If

'Send venture totals to opponent
SendData "CV0:" & Trim(Str(cMissions.Count)) & ":" & Trim(Str(cCompletedMissions.Count)) & ":" & Trim(Str(cDeadMissions.Count)) & ":" & Trim(Str(cVenturedMissions.Count)) & ":" & Trim(Str(cVenturedC.Count)) & ":|"

End Sub
Private Sub Code_Received(Scode As Variant)
Dim cd$
Dim cdt$
Dim ncardfrom()
Dim ncardid()

'Figure out the type of code

If Scode = "" Then Exit Sub

cdt$ = Right(Scode, Len(Scode) - 3)
cd$ = Left(Scode, 2)

Select Case UCase(cd$)

Case "DC"
    History_Add sOpponentName & " DREW A CARD"
    
Case "ON" 'Opponent Name

    sOpponentName = cdt$
    History_Add "OPPONENT: " & sOpponentName

Case "SP"
    bOppPassed = False
    
Case "BK"
    'Opponent has KOD battlesite
    lblKO(9).Visible = True
    History_Add "OP BATTLESITE KO'D"
    
Case "VD"
    X = MsgBox(sOpponentName & " is viewing your draw pile.", vbInformation, "Opponent Action")
    History_Add sOpponentName & " VIEWING MY DRAW PILE"

Case "V2"
    X = MsgBox(sOpponentName & " is viewing your Power Pack.", vbInformation, "Opponent Action")
    History_Add sOpponentName & " VIEWING MY POWER PACK"

Case "V3"
    X = MsgBox(sOpponentName & " is viewing your Dead Pile.", vbInformation, "Opponent Action")
    History_Add sOpponentName & " VIEWING MY DEAD PILE"

Case "V4"
    X = MsgBox(sOpponentName & " is viewing your hand.", vbInformation, "Opponent Action")
    History_Add sOpponentName & " VIEWING MY HAND"

Case "V5"
    History_Add sOpponentName & " VIEWING HIS DRAW PILE"

Case "V6"
    History_Add sOpponentName & " VIEWING HIS POWER PACK"

Case "V7"
    History_Add sOpponentName & " VIEWING HIS DEAD PILE"

Case "HL" 'Herolist
    Set cOpponent = New clsOpponent
    
    History_Add "======================================"
    History_Add "ENEMIES:"
    History_Add "======================================"
    
    For i = 1 To 4
        Dim bReserve As Boolean
        
        X = InStr(cdt$, ":")
        a$ = Left(cdt$, X - 1)
        
        If Right(a$, 1) = "R" Then
            bReserve = True
            a$ = Left(a$, Len(a$) - 1)
        Else
            bReserve = False
        End If
        
        nId = Val(a$)
        cdt$ = Right(cdt$, Len(cdt$) - X)
        
        cOpponent.AddCharacter nId, bReserve, False
        
    Next i
    LoadOpponentCharacters
    
    For i = 1 To cOpponent.LiveCharacterCount
        History_Add Trim(Str(i)) & ". " & cOpponent.Character_Name(i) & " [E" & cOpponent.Character_Energy(i) & "/F" & cOpponent.Character_Fighting(i) & "/S" & cOpponent.Character_Strength(i) & "/I" & cOpponent.Character_Intellect(i) & "]"
    Next i
    
    History_Add "======================================"
    
    
Case "BN"  'Battlesite ID

    nId = Val(cdt$)
    
    Set OpBattlesite = New clsBattlesite
    OpBattlesite.Load nId
    LoadOpBattlesite
    
    History_Add "OP BATTLESITE: " & OpBattlesite.Name
    
Case "HN" ' Homebase ID

    nId = Val(cdt$)
    Set OpHomebase = New clsHomebase
    OpHomebase.Load nId
    LoadOpponentHomebase
    
    History_Add "OP HOMEBASE: " & OpHomebase.Name
    
Case "DX"
    'Discard pile has been shuffled into draw pile.  Clear discard pile
    Set cDiscardPileO = New Collection
    UpdateOpponentDeckDisplay
    
Case "DP" 'Receive draw pile from opponent
'Format DP:[CARDCODE][CARDID]:[CARDCODE][CARDID]:
'i.e. DP:S33:T44:
'Codes - [A]ctivator,A[L]ly,As[P]ect,[B]asic Universe, [D]oubleshot
'[E]vent, [P]ower Card, [S]pecial, [T]eamwork, T[R]aining
        
        Status "Retrieving Opponent Draw Pile..."
        
        X = Code_ImportPileString(cDrawPileO, cdt$)
        
        UpdateOpponentDeckDisplay
        
        History_Add "OP DRAW PILE UPDATED [" & cDrawPileO.Count & "]"

        Status ""
        
Case "DB" 'Receive Battlesite deck from opponent

        Dim cTempBS As Collection
        
        Set cTempBS = New Collection
        
        X = Code_ImportPileString(cTempBS, cdt$)
        
        For i = 1 To cTempBS.Count
            OpBattlesite.Deck_AddCard cTempBS.Item(i)
        Next i
        
        Set cTempBS = Nothing
        
        UpdateOpponentDeckDisplay
        
        History_Add "OP BATTLESITE DECK UPDATED [" & OpBattlesite.Deck_Count & "]"
    
Case "DD" 'Receive Opponent Dead Pile

    X = Code_ImportPileString(cDeadPileO, cdt$)
    
    UpdateOpponentDeckDisplay
    
    History_Add "OP DEAD PILE UPDATED [" & cDeadPileO.Count & "]"
    
Case "DI"  'Receive Opponent Discard Pile

    X = Code_ImportPileString(cDiscardPileO, cdt$)
    
    UpdateOpponentDeckDisplay
    
    History_Add "OP DISCARD PILE UPDATED [" & cDiscardPileO.Count & "]"

Case "DH" ' Receive Opponent Hand

    X = Code_ImportPileString(cHandO, cdt$)
    
    UpdateOpponentDeckDisplay
    
    History_Add "OP HAND UPDATED [" & cHandO.Count & "]"

Case "SW"
'Attack has been shifted to another character
    X = InStr(cdt$, ":")
    sc = Val(Left(cdt$, X - 1))

    History_Add "ATTACK SHIFTED TO: " & cOpponent.Character_Name(sc)

    For i = 4 To 8
        lnFrontLine(i).Visible = False
    Next i
    
    For i = 0 To 2
        If imgOpponent(i).Tag = sc Then
            lnFrontLine(i + 4).Visible = True
        End If
    Next i
    
    If imgOppReserve.Tag = sc Then lnFrontLine(7).Visible = True
    
    myattack.DefenderID = sc
    

    
Case "SC" 'Cards being switch between collections
' Use - SC:[FROM]:[TO]:[# WITHIN COLLECTION]:
'1=Draw; 2= Hand; 3= Discard; 4 = Dead; 5 = Defeated; 6 = Reserve missions
' 7=Dead missions; 8= Completed missions; 9 = Ventured from reserve
'10= ventured from completed; 11 = Battlesite deck
'12 = Hero 1 Placed; '13 = Hero 2 Placed; '14 = Hero 3 Placed; '15 = Hero 4 Placed
'16 = Attack; 17 = Defense
'18 = Hero 1 Modifier; 19 = Hero2 Modifier; 20 = Hero3 Modifier; 21 = Hero4 Modifier
'22 = Hero 1 Buffer; 23 = Hero2 Buffer; 24 = Hero3 Buffer; 25 = Hero4 Buffer

    X = InStr(cdt$, ":")
    fp = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    If fp = 11 Then History_Add "ACTIVATOR EXCHANGED FOR SPECIAL"
    
    X = InStr(cdt$, ":")
    tp = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    nId = Val(Left(cdt$, X - 1))
    
    Code_MoveCard fp, tp, nId
    
    If (fp = 16 Or tp = 16) And frmAttack.Visible = True Then
        ShowIncomingAttackCards
    End If
    
    If fp = 17 And frmDefense.Visible = True Then ShowOpponentDefense
    
    If fp = 1 And tp = 2 Then
        UpdateOpponentDeckDisplay
    Else
        ShowOpponentModifiers
        ShowOpponentBuffers
        ShowOpponentPlacedCards
        UpdateOpponentDeckDisplay
    End If
    
Case "PC" 'place a card to hero or battlesite
'Use - PC:[FROM PILE]:[TO ID]:[# WITHIN COLLECTION]
' Use same numbers as for SC code to signify pile
' TO ID should be 9999 for homebase, otherwise HERO ID (not number)

    X = InStr(cdt$, ":")
    fp = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    tp = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    nId = Val(Left(cdt$, X - 1))

    Code_PlaceCard fp, tp, nId
    
    ShowOpponentPlacedCards
    UpdateOpponentDeckDisplay

Case "PE" 'Play an event
'PE:[# within hand]:"

    nId = Val(Left(cdt$, Len(cdt$) - 1))
    
    ShowOpponentEvent nId
    

Case "CH" 'My defense has been challenged
    m$ = Left(cdt$, Len(cdt$) - 1)
    
    If sSounds(14) <> "" Then PlaySound sSounds(14)
    
    Load frmChallenge
    frmChallenge.txtDispute.Text = m$
    frmChallenge.Show 1
    
    cmdOKDefense.Enabled = True
    cmdNoDefense.Enabled = True
    
Case "AT" 'An attack coming in from opponent
'Usage "AT:[FROMCHARDID:[TO CHARID]:x|"
' If attack is facedown, x = 1, otherwise 0

    X = InStr(cdt$, ":")
    fp = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    tp = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    fd = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    If fd = 1 Then
        bIncomingAttackFaceDown = True
    Else
        bIncomingAttackFaceDown = False
    End If
    
    Set OpAttack = New clsAttack
    
    OpAttack.NewAttack
    OpAttack.AttackerID = fp
    OpAttack.DefenderID = tp
   
    ShowIncomingAttack
    
    If tp = 5 Then
        History_Add "===================================================================="
        History_Add cOpponent.Character_Name(fp) & " ATTACKS " & OpBattlesite.Name
        History_Add "-------------------------------------------------------------------------------------------------"
    Else
        History_Add "===================================================================="
        History_Add cOpponent.Character_Name(fp) & " ATTACKS " & cFrontLine.Character_Name(tp)
        History_Add "-------------------------------------------------------------------------------------------------"
    End If
    
    If bIncomingAttackFaceDown = True Then
    
        History_Add "ATTACK PLAYED FACE DOWN"
        
    Else
    
        For i = 1 To cIncomingAttack.Count
            Set ccard = cIncomingAttack.Item(i)
            History_Add Trim(Str(i)) & ". " & ccard.Title
        Next i

    End If
    History_Add "===================================================================="
    
    If cIncomingAttack.Count > 0 Then ActionClick 0
    
    If frmDefense.Visible = False Then
    
        Set myDefense = New clsDefense
        myDefense.NewDefense
        myPhase = nPhase_Defend
        ShowDefenseFrame
        
    End If
    
    

Case "ND" 'Opponent will not defend attack

Dim bCheck As Boolean

    If sSounds(5) <> "" Then PlaySound sSounds(5)
    
    History_Add sOpponentName & " HAS NO DEFENSE"
    
    nId = myattack.DefenderID
    
    If bOpponentConceded = True Then
        Set ccard = myattack.GetCard(1)
        If ccard.CardType = "Special Card" Then
            If ccard.Attack_PostConcessionAttack = True Then
                bCheck = True
            End If
        End If
    End If
    
    If cOpponent.Buffers_Count(nId) > 0 Then
        
        Set ccard = cOpponent.Buffers_GetCard(nId, 1)
        
        cOpponent.Buffers_RemoveCard nId, 1
        ShowOpponentBuffers
        
        For i = 1 To myattack.Card_Count
        
            cOpponent.BufferHits_AddCard myattack.GetCard(1)
            
        Next i
        
        ShowOpponentBufferHits
        ShowVentureTotals
        
        History_Add ccard.Title & " KO'D"
        
    Else
    
        For i = 1 To myattack.Card_Count
        
        If nId = 5 Then
            OpBattlesite.HitsToCurrentBattle_AddCard myattack.GetCard(i)
        Else
            cOpponent.HitsToCurrentBattle_AddCard nId, myattack.GetCard(i)
        End If
        
        Next i
    
    End If
    
    ShowOpponentHTCB
    myattack.NewAttack
    UpdateDeckDisplay
    ShowVentureTotals
    
    frmAttack.Visible = False
    
    For i = 0 To lnFrontLine.Count - 1
        lnFrontLine(i).Visible = False
    Next i
    
    frmDefense.Visible = False
    
    If bCheck = True Then
        ResolveConcession False
        Exit Sub
    End If
    
    CheckForAdditionalAttack

Case "DF" ' Incoming Defense
'Use DF:1:|

    ShowOpponentPlacedCards
    UpdateOpponentDeckDisplay
    
    History_Add "===================================================================="
    History_Add "OPPONENT DEFENSE"
    History_Add "-------------------------------------------------------------------------------------------------"
    
    For i = 1 To cIncomingDefense.Count
        Set ccard = cIncomingDefense.Item(i)
        History_Add Trim(Str(i)) & ". " & ccard.Title
    Next i
    
    History_Add "===================================================================="
   
    ShowOpponentDefense
    
    If cIncomingDefense.Count > 0 Then
        DefenseCardDetail 0
    End If


Case "DA" 'Defense has been accepted by opponent

    History_Add sOpponentName & " DEFENSE ACCEPTED"
    
    frmAttack.Visible = False
    frmDefense.Visible = False
    
    For i = 0 To lnFrontLine.Count - 1
        lnFrontLine(i).Visible = False
    Next i
    
    For i = 1 To cIncomingAttack.Count
        Set ccard = cIncomingAttack.Item(i)
        
        Select Case ccard.CardType
        Case "Power Card"
            cDiscardPileO.Add ccard
        Case Else
            cDeadPileO.Add ccard
        End Select
        
    Next i
    
    For i = 1 To myDefense.Card_Count
        Set ccard = myDefense.GetCard(i)
        
        Select Case ccard.CardType
        Case "Power Card"
            cDiscardPile.Add ccard
        Case Else
            cDeadPile.Add ccard
        End Select
    
    Next i
    
    OpAttack.NewAttack
    Set cIncomingAttack = New Collection
    myDefense.NewDefense
    
    UpdateOpponentDeckDisplay
    UpdateDeckDisplay

Case "KO" 'Opponent has Ko'd a character

    X = InStr(cdt$, ":")
    nId = Val(Left(cdt$, X - 1))

    If sSounds(6) <> "" Then PlaySound sSounds(6)
    
    History_Add cOpponent.Character_Name(nId) & " HAD BEEN KO'D!"
    
    For i = 0 To 2
        If Val(imgOpponent(i).Tag) = nId Then
            lblKO(i + 5).Visible = True
        End If
    Next i
    
    If Val(imgOppReserve.Tag) = nId Then
        lblKO(8).Visible = True
    End If

    If cOpponent.LiveCharacterCount = 1 Then
        X = MsgBox(sOpponentName & " has lost!", vbInformation, "You win!")
    End If
    
Case "RF" 'opponent reserve to frontline
' usage RF:[CHAR ID]:|

    X = InStr(cdt$, ":")
    nId = Val(Left(cdt$, X - 1))
    
    cOpponent.isCharacterReserve(nId) = False
    
    LoadOpponentCharacters
    ShowOpponentHTCB
    ShowOpponentPermanentRecord
    ShowOpponentPlacedCards
    ShowOpponentModifiers
    ShowOpponentBuffers
    
Case "FR" 'opponent frontline to reserve
    X = InStr(cdt$, ":")
    nId = Val(Left(cdt$, X - 1))
    
    cOpponent.isCharacterReserve(nId) = True
    
    LoadOpponentCharacters
    ShowOpponentPlacedCards
    ShowOpponentHTCB
    ShowOpponentPermanentRecord

Case "GF" 'Received code indicating who will go first
'Usage "GF:[0/1]:|" '0 = opponent goes first, 1 = I go first

    X = InStr(cdt$, ":")
    c1 = Val(Left(cdt$, X - 1))

    If c1 = 1 Then
        bIGoFirst = True
        History_Add mySettings.PlayerName & " GOES FIRST"
    Else
        bIGoFirst = False
        History_Add sOpponentName & " GOES FIRST"
    End If
    
    frmWhoGoesFirst.Visible = False
    myPhase = nPhase_Draw
    UpdatePhase
    DrawNewHand

    ShowDiscardFrame
    
Case "WG" 'received who goes first information from opponent
'Usage WG:[CARD1 ID]:[CARD 2 ID]:|

    X = InStr(cdt$, ":")
    c1 = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    c2 = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)

    If cDrawPileO.Item(c1).LoadImage(cDrawPileO.Item(c1).ID) = True Then
        imgWGF1.Picture = LoadPicture(App.Path & "\temppic.jpg")
    Else
        imgWGF1.Picture = LoadPicture(sBlankImagePath)
    End If
    
    If cDrawPile.Item(c2).LoadImage(cDrawPile.Item(c2).ID) = True Then
        imgWGF2.Picture = LoadPicture(App.Path & "\temppic.jpg")
    Else
        imgWGF2.Picture = LoadPicture(sBlankImagePath)
    End If

    cmdWGF1.Enabled = False
    cmdWGF2.Enabled = False
    cmdWGFDraw1.Enabled = False
    cmdWGFDraw2.Enabled = False
    
    lblWGF2.Caption = mySettings.PlayerName
    lblWGF1.Caption = sOpponentName
    cmdWGFDraw2.Caption = "Redraw (" & mySettings.PlayerName & ")"
    cmdWGFDraw1.Caption = "Redraw (" & sOpponentName & ")"
    cmdWGF2.Caption = mySettings.PlayerName & " Goes First"
    cmdWGF1.Caption = sOpponentName & " Goes First"

    frmWhoGoesFirst.Visible = True

Case "FD" 'Opponent has finished discarding
'Usage "FD:1:|"
    lblDiscard1(1).Caption = sOpponentName & ": Done"
    
    If lblDiscard1(0).Caption = mySettings.PlayerName & ": Done" Then
        frmDiscardPhase.Visible = False
        
        myPhase = nPhase_Place
        UpdatePhase
        ShowPlacingFrame
        
    End If
    
Case "FP" 'Opponent has finished placing
    lblDiscard1(3).Caption = sOpponentName & ": Done"
    
    If lblDiscard1(2).Caption = mySettings.PlayerName & ": Done" Then
        frmPlacingPhase.Visible = False
        
        myPhase = nPhase_Venture
        UpdatePhase
        ShowVentureFrame
    End If

Case "FV" ' Opponent has finished venture
    lblDiscard1(5).Caption = sOpponentName & ": Done"

    If lblDiscard1(4).Caption = mySettings.PlayerName & ": Done" Then
    frmVenturePhase.Visible = False
    
    If bIGoFirst = True Then
        myPhase = nPhase_Attack
        History_Add mySettings.PlayerName & " PREPARE YOUR ATTACK"
        If sSounds(18) <> "" Then PlaySound sSounds(18)
    Else
        myPhase = nPhase_Defend
        History_Add sOpponentName & " IS PREPARING ATTACK"
      End If
    
    'set pass flag to false
    bHavePassed = False
    bOppPassed = False
    
    UpdatePhase
    
    End If

Case "EP" 'End Pause
    'Send this code 'EP:1:| to continue processing codes after something has paused
    bStopProcessing = False

Case "AC" 'Attack complete.  Opponent may now attack.
    'Send AC:1:| to indicate that it is opponent's turn to attack
    
    HideStringAttackFrame
    
    If bHavePassed = True Then
        X = MsgBox("You have passed and may not make any more attacks.  To pass again, choose 'Pass to...' from the Opponent menu.  To concede, choose 'Concede to...' from the Opponent menu.", vbInformation, "Pass or Concede")
        myPhase = nPhase_Attack
        UpdatePhase
    Else
        myPhase = nPhase_Attack
        History_Add mySettings.PlayerName & " PREPARE YOUR ATTACK"
        If sSounds(18) <> "" Then PlaySound sSounds(18)
        UpdatePhase
    End If
    
Case "AA"
    'Send AA:1:| to indicate to opponent that an additional attack is coming
    History_Add sOpponentName & " IS PREPARING AN ADDT'L ATTACK"
    UpdatePhase

Case "RC" 'Opponent ressurects character
    
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    
    History_Add "RESSURECTED: " & cOpponent.Character_Name(hindex)
    cOpponent.RessurrectCharacter hindex
    
    LoadOpponentCharacters
    ShowOpponentHTCB
    ShowOpponentPermanentRecord
    ShowOpponentBuffers
    ShowOpponentModifiers
    ShowOpponentPlacedCards
        
Case "V0"
SendData "CVT:" & Trim(Str(cMissions.Count)) & ":" & Trim(Str(cCompletedMissions.Count)) & ":" & Trim(Str(cDeadMissions.Count)) & ":" & Trim(Str(cVenturedMissions.Count)) & ":" & Trim(Str(cVenturedC.Count)) & ":|"
    X = InStr(cdt$, ":")
    vt = Val(Left(cdt$, X - 1))
    
    Set cMissionsO = New Collection
    For i = 1 To vt
        cMissionsO.Add "1"
    Next i
    
    cdt$ = Right(cdt$, Len(cdt$) - X)
    X = InStr(cdt$, ":")
    vt = Val(Left(cdt$, X - 1))
    
    Set cCompletedMissionsO = New Collection
    For i = 1 To vt
        cCompletedMissionsO.Add "1"
    Next i
    
    cdt$ = Right(cdt$, Len(cdt$) - X)
    X = InStr(cdt$, ":")
    vt = Val(Left(cdt$, X - 1))
    
    Set cDeadMissionsO = New Collection
    For i = 1 To vt
        cDeadMissionsO.Add "1"
    Next i

    cdt$ = Right(cdt$, Len(cdt$) - X)
    X = InStr(cdt$, ":")
    vt = Val(Left(cdt$, X - 1))
    
    Set cVenturedMissionsO = New Collection
    For i = 1 To vt
        cVenturedMissionsO.Add "1"
    Next i

    cdt$ = Right(cdt$, Len(cdt$) - X)
    X = InStr(cdt$, ":")
    vt = Val(Left(cdt$, X - 1))
    
    Set cVenturedCO = New Collection
    For i = 1 To vt
        cVenturedCO.Add "1"
    Next i

    UpdateOpponentDeckDisplay
    
Case "VT"
    'Venture total received from opponent
    txtOppVentureTotal.Text = cdt$
    cmdAcceptOppVenTotal.Enabled = True

Case "VA"
    'Opponent has accepted my venture total
    chkVTTotalAccepted.Value = 1
    
    If cmdAcceptOppVenTotal.Tag = "1" Then
    'both totals have been accepted.  Resolve win
    ShowMoveVentureFrame
    
    End If
   
Case "M1"
    'Opponent has finished moving venture cards
    
    lblMoveVenture(1).Caption = sOpponentName & ": Finished"
    
    If lblMoveVenture(0).Caption = mySettings.PlayerName & ": Finished" Then
        'Both players finished
        CheckMissionMessages
        EndTurn
        
    End If

Case "YA"
    'Concession accepted
    ResolveConcession True
    
Case "CB"
    'Opponent has conceded battle
    
    If sSounds(11) <> "" Then PlaySound sSounds(11)
    
    History_Add sOpponentName & " CONCEDES BATTLE"
    txtMyVentureTotal.Text = "100"
    txtOppVentureTotal.Text = "-100"

    bOpponentConceded = True
    If HaveConcedeEffect = True Then
        frmAcceptConcede.Show 1
        If frmAcceptConcede.chkAccepted.Value = 1 Then
            ResolveConcession False
            SendData "CYA:1:|"
            
            Exit Sub
        End If
        
        bOpponentConceded = False
        
        myPhase = nPhase_Attack
        History_Add mySettings.PlayerName & " PREPARE YOUR ATTACK"
        If sSounds(18) <> "" Then PlaySound sSounds(18)
        UpdatePhase
        
    Else
        ResolveConcession False
    End If
    

Case "PO"
'Will opponent play open handed this battle?
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    
    If hindex = 1 Then
        bOppOpenHanded = True
        History_Add sOpponentName & ": PLAYING OPEN-HANDED"
    Else
        bOppOpenHanded = False
    End If
    
Case "PX"
'Permanent Record hit removed by opponent
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    cindex = Val(Left(cdt$, Len(cdt$) - 1))

    If hindex < 5 Then
        cOpponent.PermanentRecord_RemoveCard hindex, cindex, False
    Else
        OpBattlesite.PermanentRecord_RemoveCard cindex, False
    End If
    
    ShowOpponentPermanentRecord
    ShowOpponentPlacedCards
    
    UpdateDeckDisplay
    
Case "PM"
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    hindex2 = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    cindex = Val(Left(cdt$, Len(cdt$) - 1))

    If hindex < 5 Then
        Set ccard = cOpponent.PermanentRecord_GetCard(hindex, cindex)
    Else
        Set ccard = OpBattlesite.PermanentRecord_GetCard(cindex)
    End If
    
    If hindex2 < 5 Then
        cOpponent.PermanentRecord_AddCard hindex2, ccard
    Else
        OpBattlesite.PermanentRecord_AddCard ccard
    End If
    
    If hindex < 5 Then
        cOpponent.PermanentRecord_RemoveCard hindex, cindex, True
    Else
        OpBattlesite.PermanentRecord_RemoveCard cindex, True
    End If

    ShowOpponentPermanentRecord
    UpdateDeckDisplay
    
Case "RH"
'HTCB Hit removed by opponent

    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    cindex = Val(Left(cdt$, Len(cdt$) - 1))
    
    If hindex < 5 Then
        cOpponent.HitsToCurrentBattle_RemoveCard hindex, cindex, False
    Else
        OpBattlesite.HitsToCurrentBattle_RemoveCard cindex, False
    End If
    
    ShowOpponentHTCB
    UpdateDeckDisplay

Case "HR"
    'Removing aspect from homebase
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    
    Set ccard = OpHomebase.PlacedCard(hindex)
    
    If sSounds(12) <> "" Then PlaySound sSounds(12)
    History_Add "OP DISCARD: " & ccard.Title
    
    cDeadPileO.Add ccard
    OpHomebase.RemovePlacedCard hindex
    ShowOpponentPlacedCards
    UpdateOpponentDeckDisplay
    
Case "HP"
'placing aspect to homebase
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    cindex = Val(Left(cdt$, Len(cdt$) - 1))
    
    Set ccard = cHandO.Item(hindex)
    cHandO.Remove hindex
    
    UpdateOpponentDeckDisplay
    
    If cindex = 1 Then
        OpHomebase.PlaceCard ccard, True
    Else
        OpHomebase.PlaceCard ccard, False
    End If
    
    ShowOpponentPlacedCards
    
Case "PH"
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    cindex = Val(Left(cdt$, Len(cdt$) - 1))

    If hindex < 5 Then
        Set ccard = cOpponent.HitsToCurrentBattle_GetCard(hindex, cindex)
        cOpponent.PermanentRecord_AddCard hindex, ccard
        cOpponent.HitsToCurrentBattle_RemoveCard hindex, cindex, True
    Else
        Set ccard = OpBattlesite.HitsToCurrentBattle_GetCard(cindex)
        opattleSite.PermanentRecord_AddCard ccard
        OpBattlesite.HitsToCurrentBattle_RemoveCard cindex, True
    End If

    ShowOpponentPermanentRecord
    ShowOpponentHTCB
    UpdateDeckDisplay
    
Case "MH"
    X = InStr(cdt$, ":")
    hindex = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    X = InStr(cdt$, ":")
    hindex2 = Val(Left(cdt$, X - 1))
    cdt$ = Right(cdt$, Len(cdt$) - X)
    
    cindex = Val(Left(cdt$, Len(cdt$) - 1))
    
    If hindex < 5 Then
        Set ccard = cOpponent.HitsToCurrentBattle_GetCard(hindex, cindex)
    Else
        Set ccard = OpBattlesite.HitsToCurrentBattle_GetCard(cindex)
    End If

    If hindex2 < 5 Then
        cOpponent.HitsToCurrentBattle_AddCard hindex2, ccard
    Else
        OpBattlesite.HitsToCurrentBattle_AddCard ccard
    End If

    If hindex < 5 Then
        cOpponent.HitsToCurrentBattle_RemoveCard hindex, cindex, True
    Else
        OpBattlesite.HitsToCurrentBattle_RemoveCard cindex, True
    End If

    ShowOpponentHTCB
    UpdateDeckDisplay

Case "PB"
    'Opponent has passed for remainder of battle
    bOppPassed = True
    
    If bHavePassed = True Then
    
        myPhase = nPhase_Resolve
        ShowResolveVentureFrame
        UpdatePhase
        History_Add "RESOLVE VENTURE"
        
    Else
    
        History_Add sOpponentName & " HAS PASSED."
        myPhase = nPhase_Attack
        History_Add mySettings.PlayerName & " PREPARE YOUR ATTACK"
        If sSounds(18) <> "" Then PlaySound sSounds(18)
        UpdatePhase
    End If
    
Case Else
End Select

End Sub
Private Function GetNextValue(cdt$) As Variant

X = InStr(cdt$, ":")

If X = 0 Then
    GetNextValue = "ERROR"
    Exit Function
End If

GetNextValue = Left(cdt$, X - 1)
cdt$ = Right(cdt$, Len(cdt$) - X)


End Function


Private Sub OpponentDetail(oPicture As Image, nAttackLine)

    If Val(oPicture.Tag) = 0 Then Exit Sub

    imgHeroCard.Picture = oPicture.Picture
    imgHeroCard.Tag = oPicture.Tag
    HideFrames True, False, False

    a = Val(imgHeroCard.Tag)

    If cOpponent.Character_HasInherent(a) = True Then
        txtInherent.Text = cOpponent.Character_Inherent(a)
    Else
        txtInherent.Text = "NO INHERENT ABILITY"
    End If

    cmdTakeAction.Enabled = False
    cmdKOCharacter.Enabled = False
    cmdSwitchWithReserve.Enabled = False
    cmdReserveToFrontline.Enabled = False
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HERO: " & cOpponent.Character_Name(a)

    If cmdOKAction.Enabled = True Then
        
        If cOpponent.Buffers_Count(a) > 0 Then
            Set ccard = cOpponent.Buffers_GetCard(a, 1)
            MsgBox "NOTE: " & cOpponent.Character_Name(a) & " has " & ccard.Title & " placed.  Your attack would be on this Buffer, not the character.", vbInformation, "Character has Buffer placed"
        End If
        ShowAttackOpLines nAttackLine
    End If
    
    HideAllBorders

End Sub
Private Sub PlayCard()
Dim ccard

If myPhase = nPhase_Defend Then
    
    If frmDefense.Visible = False Then Exit Sub
    
    If cmdOKDefense.Enabled = False Then
        X = MsgBox("Cards cannot be added to a defense once it has been submitted to your opponent.  Message your opponent and ask him to challenge your defense.", vbCritical, "Cannot add Defense Card")
        Exit Sub
    End If
    
    Index = Val(imgCardDetail.Tag)
    
    SendData "CSC:2:17:" & Trim(Str(Index)) & ":|"
    
    Set ccard = cHand.Item(Index)
    myDefense.AddCard ccard, "H", Index
    
    ShowDefenseCards

Else
    
    If frmAttack.Visible = False Then Exit Sub
    
    If myattack.AttackerID < 1 Then
        X = MsgBox("Please start an Action with a character before playing cards.", vbCritical, "No Current Action")
        Exit Sub
    End If
    
    Index = Val(imgCardDetail.Tag)
    
    SendData "CSC:2:16:" & Trim(Str(Index)) & ":|"
    
    Set ccard = cHand.Item(Index)
    myattack.AddCard ccard, "H", Index
    ShowAttackCards

End If

ShowHand

End Sub
Private Sub ShowIncomingAttackCards()
Dim ccard

For i = 0 To 3
imgAction(i).Picture = Nothing
imgAction(i).Tag = -1

Next i

For i = 1 To cIncomingAttack.Count
Set ccard = cIncomingAttack.Item(i)

a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgAction(i - 1).Picture = LoadPicture(App.Path & "\temppic.jpg")
    imgAction(i - 1).Tag = 1
Else
    imgAction(i - 1).Picture = LoadPicture(sBlankImagePath)
    imgAction(i - 1).Tag = 1
End If

Next i

End Sub
Private Sub ShowAttackCards()
Dim ccard

For i = 0 To 3
imgAction(i).Picture = Nothing
imgAction(i).Tag = -1

Next i

For i = 1 To myattack.Card_Count

Set ccard = myattack.GetCard(i)

a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgAction(i - 1).Picture = LoadPicture(App.Path & "\temppic.jpg")
    imgAction(i - 1).Tag = 1
Else
    imgAction(i - 1).Picture = LoadPicture(sBlankImagePath)
    imgAction(i - 1).Tag = 1
End If
   

Next i

'figure out if attack has already been sent
    If myattack.Card_Count > 0 Then
        cmdOKAction.Enabled = True
        ActionClick 0
    Else
        cmdOKAction.Enabled = False
        shpActionBorder.Visible = False
    End If

End Sub
Private Sub ShowIncomingAttack()
On Error Resume Next

Dim ccard

'Show frame

For i = 0 To 3
    imgAction(i).Picture = Nothing
    imgAction(i).Tag = -1
Next i

shpAction.Visible = True
frmAttack.Visible = True

'Play sounds if app
If sSounds(2) <> "" Then
    PlaySound sSounds(2)
End If

For i = 1 To cIncomingAttack.Count

Set ccard = cIncomingAttack.Item(i)

a = ccard.ID

chkPlayFaceDown.Enabled = False

If bIncomingAttackFaceDown = True Then

    imgAction(i - 1).Picture = LoadPicture(sBlankImagePath)
    imgAction(i - 1).Tag = 1
    chkPlayFaceDown.Value = 1
    
Else

    If ccard.LoadImage(a) = True Then
        imgAction(i - 1).Picture = LoadPicture(App.Path & "\temppic.jpg")
        imgAction(i - 1).Tag = 1
    Else
        imgAction(i - 1).Picture = LoadPicture(sBlankImagePath)
        imgAction(i - 1).Tag = 1
    End If

    chkPlayFaceDown.Value = 0
    
End If

Next i

    
If bIncomingAttackFaceDown = False Then

    Set ccard = cIncomingAttack.Item(1)
    If ccard.Attack_isStringAttack = True Then
        imgStringAttack.Picture = imgAction(0).Picture
        imgStringAttack.ToolTipText = ccard.Description
        ShowStringAttackFrame
    End If
    
End If

cmdCancelAction.Enabled = False
cmdOKAction.Enabled = False

For i = 4 To 8
    lnFrontLine(i).Visible = False
Next i

If OpAttack.AttackerID = 5 Then lnFrontLine(8).Visible = True
If imgOpponent(0).Tag = OpAttack.AttackerID Then lnFrontLine(4).Visible = True
If imgOpponent(1).Tag = OpAttack.AttackerID Then lnFrontLine(5).Visible = True
If imgOpponent(2).Tag = OpAttack.AttackerID Then lnFrontLine(6).Visible = True
If imgOppReserve.Tag = OpAttack.AttackerID Then lnFrontLine(7).Visible = True

If imgFrontLine(0).Tag = OpAttack.DefenderID Then lnFrontLine(0).Visible = True
If imgFrontLine(1).Tag = OpAttack.DefenderID Then lnFrontLine(1).Visible = True
If imgFrontLine(2).Tag = OpAttack.DefenderID Then lnFrontLine(2).Visible = True
If ImgReserve.Tag = OpAttack.DefenderID Then lnFrontLine(3).Visible = True
If OpAttack.DefenderID = 5 Then lnFrontLine(9).Visible = True

ActionClick 0


End Sub
Private Sub ClearTable()

imgCompletedMissions.Picture = Nothing
imgVenture.Picture = Nothing
imgDeadMissions.Picture = Nothing
imgFrontLine(0).Picture = Nothing
imgFrontLine(1).Picture = Nothing
imgFrontLine(2).Picture = Nothing
ImgReserve.Picture = Nothing


End Sub
Private Sub UpdateDeckDisplay()

lblPile(0).Caption = cDrawPile.Count
lblPile(2).Caption = cDeadPile.Count
lblPile(1).Caption = cDiscardPile.Count

If cMissions.Count = 0 Then
    imgMissions.Picture = Nothing
Else

If cMissions.Item(1).LoadImage(cMissions.Item(1).ID) = True Then
    imgMissions.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgMissions.Picture = LoadPicture(sBlankImagePath)
End If
End If

a = cVenturedMissions.Count

If cVenturedMissions.Count = 0 Then
    imgVenture.Picture = Nothing
    imgVenture.Visible = False
End If

If cVenturedMissions.Count > 0 Then
If cVenturedMissions.Item(a).LoadImage(cVenturedMissions.Item(a).ID) = True Then
    imgVenture.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgVenture.Picture = LoadPicture(sBlankImagePath)
End If

End If

a = cCompletedMissions.Count

If cCompletedMissions.Count = 0 Then imgCompletedMissions.Picture = Nothing

If cCompletedMissions.Count > 0 Then
If cCompletedMissions.Item(a).LoadImage(cCompletedMissions.Item(a).ID) = True Then
    imgCompletedMissions.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgCompletedMissions.Picture = LoadPicture(sBlankImagePath)
End If

End If

a = cDeadMissions.Count

If cDeadMissions.Count = 0 Then imgDeadMissions.Picture = Nothing

If cDeadMissions.Count > 0 Then
If cDeadMissions.Item(a).LoadImage(cDeadMissions.Item(a).ID) = True Then
    imgDeadMissions.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgDeadMissions.Picture = LoadPicture(sBlankImagePath)
End If

End If

a = cVenturedC.Count

If cVenturedC.Count = 0 Then
    imgVentureC.Picture = Nothing
    imgVentureC.Visible = False
End If

If cVenturedC.Count > 0 Then
If cVenturedC.Item(a).LoadImage(cVenturedC.Item(a).ID) = True Then
    imgVentureC.Picture = LoadPicture(App.Path & "\temppic.jpg")
    imgVentureC.Visible = True
    If cCompletedMissions.Count = 0 Then
        imgVentureC.ZOrder 0
    Else
        imgVentureC.ZOrder 1
    End If
        
End If
End If

'defeated characters
'If cDefeatedCharactersPile.Count = 0 Then
'    imgDefeated.Picture = Nothing
'Else
'    imgDefeated.Picture = LoadPicture(sBlankImagePath)
'End If

lblPile(6).Caption = cMissions.Count & " (" & cVenturedMissions.Count & ")"
lblPile(5).Caption = cCompletedMissions.Count & " (" & cVenturedC.Count & ")"
lblPile(4).Caption = cDeadMissions.Count
lblPile(3).Caption = myBattleSite.Deck_Count
lblPile(7).Caption = cDefeatedCharactersPile.Count

End Sub
Private Sub UpdateOpponentDeckDisplay()
On Error Resume Next
lblPile(11).Caption = cDrawPileO.Count
lblPile(10).Caption = cDeadPileO.Count
lblPile(12).Caption = cDiscardPileO.Count
lblPile(13).Caption = cMissionsO.Count & " (" & cVenturedMissionsO.Count & ")"
lblPile(14).Caption = cCompletedMissionsO.Count & " (" & cVenturedCO.Count & ")"
lblPile(15).Caption = cDeadMissionsO.Count
lblPile(9).Caption = OpBattlesite.Deck_Count
lblPile(8).Caption = cDefeatedCharactersPileO.Count
lblPile(16).Caption = cHandO.Count

End Sub

Private Sub PlaceCard()
a = Val(imgCardDetail.Tag)
If a = 0 Then Exit Sub

Set ccard = cHand.Item(a)

If ccard.CardType = "Special Card" Then frmChooseCharacter.SpecialCharacter = ccard.Character

frmChooseCharacter.Show 1

If frmChooseCharacter.SelectedCharacter = -1 Then
    Unload frmChooseCharacter
    Exit Sub
End If

b = frmChooseCharacter.SelectedCharacter
Unload frmChooseCharacter

Set ccard = cHand.Item(a)

If b < 5 Then
    cFrontLine.PlaceCard b, ccard
Else
    myHomebase.PlaceCard ccard, True
End If

History_Add "PLACED: " & cHand.Item(a).Title

cHand.Remove a
cHandTags.Remove a

If sSounds(16) <> "" Then PlaySound sSounds(16)
SendData "CPC:2:" & Trim(Str(b)) & ":" & Trim(Str(a)) & ":|"

ShowHand
ShowPlacedCards
HideAllBorders
End Sub

Private Sub NewAction()
cmdRemoveAttackCard.Tag = "-1"
chkPlayFaceDown.Value = 0
chkPlayFaceDown.Enabled = True

If myPhase = nPhase_Defend Then
    MsgBox "You are currently engaged in a defense.  You may not attack.", vbCritical, "Incorrect Phase"
    Exit Sub
End If

If frmAttack.Visible = True Then
'If myattack.AttackerID > 0 Then
    X = MsgBox("Please cancel or resolve current attack before starting another.", vbCritical, "Attack Already Underway.")
    Exit Sub
End If

a = Val(imgHeroCard.Tag)
myattack.NewAttack
myattack.AttackerID = a

frmAttack.Tag = imgHeroCard.Tag

For i = 0 To 3
imgAction(i).Picture = Nothing
imgAction(i).Tag = -1
Next i

shpAction.Visible = True
frmAttack.Visible = True

cmdOKAction.Enabled = True
cmdCancelAction.Enabled = True

lc = -1

For i = 0 To 2
    If imgFrontLine(i).Tag = a Then
        lc = i
    End If
Next i

If lc = -1 Then
    If ImgReserve.Tag = a Then
        lc = 3
    End If
End If

lnFrontLine(lc).Visible = True


End Sub

Private Sub ActionClick(Index)
On Error Resume Next

If imgAction(Index).Tag = -1 Then Exit Sub

If myPhase = nPhase_Attack Then
    Set ccard = myattack.GetCard(Index + 1)
Else
    Set ccard = cIncomingAttack.Item(Index + 1)
End If

cmdRemoveAttackCard.Tag = imgAction(Index).Tag

shpActionBorder.Visible = False
shpActionBorder.Left = imgAction(Index).Left
shpActionBorder.Visible = True

HideFrames False, False, False
imgCardDetail.Picture = imgAction(Index).Picture

If bIncomingAttackFaceDown = True Then

    imgCardDetail.ToolTipText = "Description Not Available.  Attack played face down."
Else

If ccard.CardType = "Special Card" Then
    imgCardDetail.ToolTipText = ccard.Effect
Else
    imgCardDetail.ToolTipText = ccard.Title
End If

End If

Me.Caption = "OVERPOWER ONLINE-->" & ccard.Title
imgCardDetail.Visible = True
frmAttackCard.Visible = True

If myPhase <> nPhase_Attack Then
    cmdRemoveAttackCard.Enabled = False
Else
    cmdRemoveAttackCard.Enabled = True
End If

End Sub

Private Sub HeroClick(Index As Integer)
    If Val(imgFrontLine(Index).Tag) = 0 Then Exit Sub
    
    
    imgHeroCard.Picture = imgFrontLine(Index).Picture
    imgHeroCard.Tag = imgFrontLine(Index).Tag
    HideFrames True, False, False
   
    a = Val(imgHeroCard.Tag)
    
    If cFrontLine.Character_HasInherent(a) = True Then
        txtInherent.Text = cFrontLine.Character_Inherent(a)
    Else
        txtInherent.Text = "NO INHERENT ABILITY"
    End If
    
    cmdTakeAction.Enabled = True
    cmdKOCharacter.Enabled = True
    cmdSwitchWithReserve.Enabled = True
    cmdReserveToFrontline.Enabled = True
    
    Me.Caption = "OVERPOWER ONLINE-->" & "HERO: " & cFrontLine.Character_Name(a)
    
    If cFrontLine.isCharacterReserve(a) = True Then
        cmdReserveToFrontline.Visible = True
        cmdSwitchWithReserve.Visible = False
    Else
        cmdReserveToFrontline.Visible = False
        cmdSwitchWithReserve.Visible = True
    End If
    
    If lblKO(Index).Visible = True Then
        cmdTakeAction.Enabled = False
        cmdKOCharacter.Enabled = False
        cmdSwitchWithReserve.Enabled = False
        cmdReserveToFrontline.Enabled = False
        Me.Caption = "OVERPOWER ONLINE-->K.O.'D HERO: " & cFrontLine.Character_Name(a)
    End If
    
    HideAllBorders
End Sub
Private Sub PlacedCardDetail(Index, oPicture As Image, sTag)
Dim ccard

    imgCardDetail.Picture = oPicture.Picture
    imgCardDetail.Tag = Index
    frmPlaced.Tag = sTag
    HideFrames False, False, True
    frmPlaced.Visible = True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    
    If sTag = 5 Then
       Set ccard = myHomebase.PlacedCard(Index)
       Me.Caption = "OVERPOWER ONLINE-->" & "PLACED: " & ccard.Title
        frmPlaced.Tag = 5
    Else
        Set ccard = cFrontLine.PlacedCard(sTag, Index)
        Me.Caption = "OVERPOWER ONLINE-->" & "PLACED: " & cFrontLine.Placed_Type(sTag, Index)
    End If
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If
    
    HideAllBorders
    
End Sub
Private Sub ShowDefenseFrame()

For i = 0 To imgDefense.Count - 1
    imgDefense(i).Picture = Nothing
Next i

frmDefense.Top = 5400
frmDefense.Visible = True

cmdAcceptDefense.Visible = False
cmdOKDefense.Visible = True
cmdNoDefense.Visible = True
cmdChallengeDefense.Visible = False
cmdShiftAttack.Visible = True
cmdShiftAttack.Enabled = True

cmdOKDefense.Enabled = True
cmdNoDefense.Enabled = True

End Sub
Private Sub ShowOpponentDefense()
Dim ccard

frmDefense.Top = 5400
frmDefense.Visible = True

If sSounds(3) <> "" Then PlaySound sSounds(3)

For i = 0 To 4
imgDefense(i).Picture = Nothing
imgDefense(i).Tag = -1

Next i

For i = 1 To cIncomingDefense.Count

Set ccard = cIncomingDefense.Item(i)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgDefense(i - 1).Picture = LoadPicture(App.Path & "\temppic.jpg")
    imgDefense(i - 1).Tag = 1
Else
    imgDefense(i - 1).Picture = LoadPicture(sBlankImagePath)
    imgDefense(i - 1).Tag = 1
End If

Next i

cmdAcceptDefense.Visible = True
cmdOKDefense.Visible = False
cmdNoDefense.Visible = False
cmdChallengeDefense.Visible = True
cmdAcceptDefense.Enabled = True
cmdChallengeDefense.Enabled = True
cmdShiftAttack.Enabled = False
cmdShiftAttack.Visible = False

End Sub
Private Sub ShowDefenseCards()
Dim ccard

For i = 0 To 4
imgDefense(i).Picture = Nothing
imgDefense(i).Tag = -1

Next i

For i = 1 To myDefense.Card_Count

Set ccard = myDefense.GetCard(i)
a = ccard.ID

If ccard.LoadImage(a) = True Then
    imgDefense(i - 1).Picture = LoadPicture(App.Path & "\temppic.jpg")
    imgDefense(i - 1).Tag = 1
Else
    imgDefense(i - 1).Picture = LoadPicture(sBlankImagePath)
    imgDefense(i - 1).Tag = 1
End If

Next i

End Sub
Private Sub BeginGame()

UpdatePhase

End Sub
Private Sub UpdatePhase()

Select Case myPhase

Case -1
    lblPhase.Caption = "PHASE: WAITING..."
Case 1
    lblPhase.Caption = "PHASE: WHO GOES FIRST?"
Case 1
    lblPhase.Caption = "PHASE: DRAW"
Case 2
    lblPhase.Caption = "PHASE: DISCARD"
Case 4
    lblPhase.Caption = "PHASE: PLACE"
Case 6
    lblPhase.Caption = "PHASE: VENTURE"
Case 7
    lblPhase.Caption = "PHASE: ATTACK"
Case 8
    lblPhase.Caption = "PHASE: DEFEND"
Case 9
    lblPhase.Caption = "PHASE: RESOLVE VENTURE"
Case Else
End Select

'If myPhase = 7 Then
'    cmdEndAttack.Enabled = True
'Else
'    cmdEndAttack.Enabled = False
'End If

End Sub

Private Sub mnuViewOPDeadPile_Click()
Dim ctemp As Collection
Dim ccard

X = MsgBox("Note: You are not allowed to view your opponent's Dead Pile unless a special allows you to do so.  Your opponent will be informed that you are doing so.  Would you like to continue?", vbYesNoCancel, "Viewing Opponent Dead Pile")

If X <> 6 Then Exit Sub

SendData "CV3:1:|"

History_Add "VIEWING OPPONENT DEAD PILE"

If X <> 6 Then Exit Sub

If cDrawPileO.Count = 0 Then
    MsgBox sOpponentName & " does not have any cards in the Dead Pile.", vbCritical, "No Cards in Dead Pile"
    Exit Sub
End If

With FrmViewPile

Set ctemp = New Collection
For i = 1 To cDeadPileO.Count
Set ccard = cDeadPileO.Item(i)
ctemp.Add ccard
Next i

Set .ShowPile = ctemp
.PileType = 6
.Show 1

End With
End Sub

Private Sub mnuViewOpPowerPack_Click()
Dim ctemp As Collection
Dim ccard

X = MsgBox("Note: You are not allowed to view your opponent's Power Pack unless a special allows you to do so.  Your opponent will be informed that you are doing so.  Would you like to continue?", vbYesNoCancel, "Viewing Opponent Draw Pile")

If X <> 6 Then Exit Sub

SendData "CV1:1:|"

History_Add "VIEWING OPPONENT POWER PACK"

If X <> 6 Then Exit Sub

If cDrawPileO.Count = 0 Then
    MsgBox sOpponentName & " does not have any cards in the Power Pack.", vbCritical, "No Cards in Power Pack"
    Exit Sub
End If

With FrmViewPile

Set ctemp = New Collection
For i = 1 To cDiscardPileO.Count
Set ccard = cDiscardPileO.Item(i)
ctemp.Add ccard
Next i

Set .ShowPile = ctemp
.PileType = 6
.Show 1

End With
End Sub

Private Sub tcpChannel_Connect()
On Error Resume Next

If bHost = False Then

   If isConnectedFlag = False Then

      If tcpChannel.State = sckConnected Then
         SendData "CON:" & mySettings.PlayerName & "|"
      End If

    isConnectedFlag = True
        
    'check to see if game is resumed

    If bResuming = True Then
    
        ResumeGame
    
    Else
        SendOpponentDeck
    End If
    
    
   End If

End If
End Sub

Private Sub tcpChannel_ConnectionRequest(ByVal requestID As Long)
If bHost = True Then

   If tcpChannel.State <> sckClosed Then
      tcpChannel.Close
   End If

   tcpChannel.Accept requestID
End If

End Sub

Private Sub tcpChannel_DataArrival(ByVal bytesTotal As Long)
   
On Error Resume Next

   Dim strData As String   'holds incoming data
   Dim buff As String
   
   If isConnectedFlag = True Then
   
    tcpChannel.GetData strData

    a$ = strData
    
looper:
    
    X = InStr(a$, "|")
    
    If X > 0 Then
        b$ = Left(a$, X - 1)
        a$ = Right(a$, Len(a$) - X)
    
    Select Case Left(b$, 1)
    Case "C"
        lstCodes.AddItem Right(b$, Len(b$) - 1)
    Case "M"
        ReceivedMessage sOpponentName & ": " & Right(b$, Len(b$) - 1)
    Case Else
    End Select
    
    If a$ <> 0 Then GoTo looper
    
    End If
    
   Else
   
      isConnectedFlag = True
        
      tcpChannel.GetData strData
    
    a$ = strData
    
looper2:
    
    X = InStr(a$, "|")
    
    If X > 0 Then
        b$ = Left(a$, X - 1)
        a$ = Right(a$, Len(a$) - X)
    
    Select Case Left(b$, 1)
    Case "C"
        lstCodes.AddItem Right(b$, Len(b$) - 1)
    Case "M"
        ReceivedMessage sOpponentName & ": " & Right(b$, Len(b$) - 1)
    Case Else
    End Select
    
    If a$ <> 0 Then GoTo looper2
    
    End If
    
      
        SendData "CON:" & mySettings.PlayerName & "|"
            
        If bResuming = True Then
        
            ResumeGame
        
        Else
        
            SendOpponentDeck
        
        End If
        
   End If

End Sub

Private Sub tcpChannel_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

CancelDisplay = True

End Sub

Private Sub Timer1_Timer()

If lstCodes.ListCount = 0 Then Exit Sub

CodeProcessLoop:
a = lstCodes.ListIndex + 1

If a <= (lstCodes.ListCount - 1) Then
    lstCodes.ListIndex = a
    Code_Received (lstCodes.List(a))
    DoEvents
    GoTo CodeProcessLoop
    
End If


End Sub
Private Function DrawHand(nNumberCards) As Integer
Dim ccard

On Error Resume Next

If cHand.Count > 0 Then
    X = MsgBox("You still have cards in your hand.  Are you sure you want to draw a new hand?", vbYesNoCancel, "Draw New Hand")

    If X <> 6 Then Exit Function
    
    For i = 1 To cHand.Count
    
    Set ccard = cHand.Item(i)
    
    If ccard.CardType = "Power Card" Then
        cDiscardPile.Add ccard
    
    Else
        cDeadPile.Add ccard
    
    End If
    
        
    Next i
    

End If

Set cHand = New Collection
Set cHandTags = New Collection

ncounter = 1

DrawLoop:

If cDrawPile.Count = 0 Then
    'Need to shuffle in power pack
    ShufflePile 1
        
    For i = 1 To cDiscardPile.Count
        cDrawPile.Add cDiscardPile.Item(i)
    Next i
    
    Set cDiscardPile = New Collection
        
    SendData "CDX:1:|"
    SendData "CDP:" & GetCode_CardString(cDrawPile) & "|"
    SendData "CEP:1:|"
    
    'Now that power pack has been shuffled in, check and see if there are enough cards
    If cDrawPile.Count = 0 Then GoTo nomorecards
    
End If

'add top card in draw pile to hand

cHand.Add cDrawPile.Item(1)
cHandTags.Add Chr$(ncounter + 64)

SendData "CSC:1:2:1:|"

cDrawPile.Remove 1

ncounter = ncounter + 1
If ncounter > nNumberCards Then GoTo nomorecards

GoTo DrawLoop

nomorecards:

DrawHand = ncounter - 1

End Function
Private Sub ReceivedMessage(sMessage)

If mySettings.MessageBeep = True Then

    If sSounds(1) = "" Then
        Beep
    Else
        PlaySound sSounds(1)
    End If
End If


If mySettings.PopupMessages = True Then
    X = MsgBox(sMessage, vbOKOnly, "Message received from " & sOpponentName)
Else
    Status sMessage
End If

lstMessages.AddItem sMessage
lstMessages.ListIndex = lstMessages.ListCount - 1


End Sub
Private Sub EndTurn()
Dim bKO As Boolean

bKO = False
bOppOpenHanded = False

Set imgEffect(0).Picture = Nothing
Set imgEffect(1).Picture = Nothing
imgEffect(0).Visible = False
imgEffect(1).Visible = False

'Permanently KO any characters
For i = 0 To 4
    If lblKO(i).Visible = True Then
    bKO = True
    Exit For
    End If
Next i

If bKO = True Then
    If lblKO(0).Visible = True Then cFrontLine.KillCharacter (Val(imgFrontLine(0).Tag))
    If lblKO(1).Visible = True Then cFrontLine.KillCharacter (Val(imgFrontLine(1).Tag))
    If lblKO(2).Visible = True Then cFrontLine.KillCharacter (Val(imgFrontLine(2).Tag))
    If lblKO(3).Visible = True Then cFrontLine.KillCharacter (Val(ImgReserve.Tag))
    LoadCharacters
End If

bKO = False

'check to see if battlesite has been KOd
If lblKO(4).Visible = True Then
    For i = 1 To myBattleSite.Deck_Count
        cDefeatedCharactersPile.Add myBattleSite.Deck_GetCard(1)
        myBattleSite.RemoveDeckCard 1
    Next i

    Set myBattleSite = New clsBattlesite
    myBattleSite.NewBattlesite
    
    UpdateDeckDisplay
    loadbattlesite
    lblKO(4).Visible = False
    
End If


For i = 5 To 9
    If lblKO(i).Visible = True Then
        bKO = True
        Exit For
    End If
Next i

If bKO = True Then
    If lblKO(5).Visible = True Then cOpponent.KillCharacter Val(imgOpponent(0).Tag)
    If lblKO(6).Visible = True Then cOpponent.KillCharacter Val(imgOpponent(1).Tag)
    If lblKO(7).Visible = True Then cOpponent.KillCharacter Val(imgOpponent(2).Tag)
    If lblKO(8).Visible = True Then cOpponent.KillCharacter Val(imgOppReserve.Tag)
    LoadOpponentCharacters
End If

'check to see if opponents battlesite has been ko'd
'check to see if battlesite has been KOd
If lblKO(9).Visible = True Then
    For i = 1 To OpBattlesite.Deck_Count
        cDefeatedCharactersPileO.Add OpBattlesite.Deck_GetCard(1)
        OpBattlesite.RemoveDeckCard 1
    Next i

    Set OpBattlesite = New clsBattlesite
    OpBattlesite.NewBattlesite
    
    UpdateOpponentDeckDisplay
    LoadOpBattlesite
    lblKO(9).Visible = False
End If


For i = 0 To 9
    lblKO(i).Visible = False
Next i

For i = 1 To 4
    If cFrontLine.isCharacterDead(i) = False Then
        For k = 1 To cFrontLine.HitsToCurrentBattle_Count(i)
            Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(i, 1)
            
            If ccard.Attack_isPlaced = True Then
                cFrontLine.PermanentRecord_AddCard i, ccard
                cFrontLine.HitsToCurrentBattle_RemoveCard i, 1, True
            Else
                cDeadPileO.Add ccard
                cFrontLine.HitsToCurrentBattle_RemoveCard i, 1, True
            End If
        Next k
    End If

Next i

On Error Resume Next
    For i = 1 To myBattleSite.HitsToCurrentBattle_Count
        Set ccard = myBattleSite.HitsToCurrentBattle_GetCard(1)
        
        If ccard.Attack_isPlaced = True Then
            myBattleSite.PermanentRecord_AddCard ccard
            myBattleSite.HitsToCurrentBattle_RemoveCard 1, True
        Else
            cDeadPileO.Add ccard
            myBattleSite.HitsToCurrentBattle_RemoveCard 1, True
        End If
    Next i
    
For i = 1 To 4
    If cOpponent.isCharacterDead(i) = False Then
    
        For k = 1 To cOpponent.HitsToCurrentBattle_Count(i)
            Set ccard = cOpponent.HitsToCurrentBattle_GetCard(i, 1)
            
            If ccard.Attack_isPlaced = True Then
                cOpponent.PermanentRecord_AddCard i, ccard
                cOpponent.HitsToCurrentBattle_RemoveCard i, 1, True
            Else
                cDeadPile.Add ccard
                cOpponent.HitsToCurrentBattle_RemoveCard i, 1, True
            End If
            
        Next k
    
    
    End If

Next i

'Move Opponent Battlesite HTKB to permanent record
    For i = 1 To OpBattlesite.HitsToCurrentBattle_Count
        Set ccard = OpBattlesite.HitsToCurrentBattle_GetCard(1)
        
        If ccard.Attack_isPlaced = True Then
            OpBattlesite.PermanentRecord_AddCard ccard
            OpBattlesite.HitsToCurrentBattle_RemoveCard 1, True
        Else
            cDeadPile.Add ccard
            OpBattlesite.HitsToCurrentBattle_RemoveCard 1, True
        End If
    Next i
    
'Clear buffer hits, if necessary
For i = 1 To cFrontLine.BufferHits_Count
    Set ccard = cFrontLine.BufferHits_GetCard(i)
    If ccard.CardType = "Power Card" Then
        cDiscardPile.Add ccard
    Else
        cDeadPile.Add ccard
    End If

Next i

cFrontLine.BufferHits_Clear

For i = 1 To cOpponent.BufferHits_Count
    Set ccard = cOpponent.BufferHits_GetCard(i)
    If ccard.CardType = "Power Card" Then
        cDiscardPileO.Add ccard
    Else
        cDeadPileO.Add ccard
    End If

Next i

cOpponent.BufferHits_Clear

ShowBufferHits
ShowOpponentBufferHits

    
'Clear hands

For i = 1 To cHand.Count
    Set ccard = cHand.Item(i)
    
    If ccard.CardType = "Power Card" Then
        cDiscardPile.Add ccard
    Else
        cDeadPile.Add ccard
    End If

Next i

Set cHand = New Collection

For i = 1 To cHandO.Count
    Set ccard = cHandO.Item(i)
    
    If ccard.CardType = "Power Card" Then
        cDiscardPileO.Add ccard
    Else
        cDeadPileO.Add ccard
    End If

Next i

Set cHandO = New Collection

frmMoveVentureCards.Visible = False
ShowHand
UpdateDeckDisplay
UpdateOpponentDeckDisplay
ShowHitsToCurrentBattle
ShowOpponentHTCB
ShowPermanentRecord
ShowOpponentPermanentRecord
ShowPlacedCards
ShowOpponentPlacedCards
ShowModifiers
ShowOpponentModifiers
ShowBuffers
ShowOpponentBuffers
ShowVentureTotals

txtMyVentureTotal.Text = "--"
txtOppVentureTotal.Text = "--"

nTurn = nTurn + 1
History_Add "*********************************************************"
History_Add "ROUND " & Trim(Str(nTurn))
History_Add "*********************************************************"

If bIGoFirst = True Then
    History_Add mySettings.PlayerName & " GOES FIRST"
Else
    History_Add sOpponentName & " GOES FIRST"
End If

'Save Resume Info
Status "Saving Resume Information..."
Timer1.Enabled = False

'==============RESUME INFO==================='
Dim ctemp As Collection

X = Dir(App.Path & "\Resume", vbDirectory)
If X = "" Then MkDir App.Path & "\Resume"


a$ = sOpponentName
a$ = ReplaceAllInString(a$, "\", "")
a$ = ReplaceAllInString(a$, "/", "")
a$ = ReplaceAllInString(a$, ".", "")
a$ = ReplaceAllInString(a$, "?", "")
a$ = ReplaceAllInString(a$, "*", "")
a$ = ReplaceAllInString(a$, "<", "")
a$ = ReplaceAllInString(a$, ">", "")
a$ = ReplaceAllInString(a$, "|", "")

'write out messages
X = FreeFile
m$ = a$ & " " & Format(Now(), "mm-dd-yyyy") & ".Rem"
Open App.Path & "\Resume\" & m$ For Output As #X
For i = 0 To frmTable.lstMessages.ListCount - 1
    Print #X, frmTable.lstMessages.List(i)
Next i
Close #X



'Write out Game History
X = FreeFile
h$ = a$ & " " & Format(Now(), "mm-dd-yyyy") & ".Reh"
Open App.Path & "\Resume\" & h$ For Output As #X
For i = 0 To frmTable.lstGameHistory.ListCount - 1
    Print #X, frmTable.lstGameHistory.List(i)
Next i
Close #X



a$ = a$ & " " & Format(Now(), "mm-dd-yyyy") & ".res"

X = FreeFile
Open App.Path & "\Resume\" & a$ For Output As #X

'Write Heroes
For i = 1 To 4

'Print character ID
Print #X, cFrontLine.Character_ID(i)

'Is Alive?
Print #X, cFrontLine.isCharacterDead(i)

'In reserve?
Print #X, cFrontLine.isCharacterReserve(i)

Next i



'Homebase ID
Print #X, myHomebase.ID


'Battlesite ID
Print #X, myBattleSite.ID



If myBattleSite.ID = 0 Then
    Print #X, ""
Else
    Set ctemp = New Collection
     
    For i = 1 To myBattleSite.Deck_Count
        Set ccard = myBattleSite.Deck_GetCard(i)
        ctemp.Add ccard
    Next i
    
    Print #X, GetCode_CardString(ctemp)
    
End If


Print #X, GetCode_CardString(cHand)

Print #X, GetCode_CardString(cDrawPile)

Print #X, GetCode_CardString(cDiscardPile)

Print #X, GetCode_CardString(cDeadPile)

Print #X, GetCode_CardString(cDefeatedCharactersPile)

'missions
Print #X, GetCode_CardString(cCompletedMissions)

Print #X, GetCode_CardString(cMissions)

Print #X, GetCode_CardString(cDeadMissions)

Print #X, GetCode_CardString(cVenturedMissions)

Print #X, GetCode_CardString(cVenturedC)


'Placed on Characters
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cFrontLine.Placed_Count(i)

ctemp.Add cFrontLine.PlacedCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Placed on Homebase
If myHomebase.ID = 0 Then
    Print #X, ""
Else
    
Set ctemp = New Collection

For k = 1 To myHomebase.Placed_Count

ctemp.Add myHomebase.PlacedCard(k)

Next k

Print #X, GetCode_CardString(ctemp)
    

End If


'Permanent Records
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cFrontLine.PermanentRecord_Count(i)

ctemp.Add cFrontLine.PermanentRecord_GetCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Buffers
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cFrontLine.Buffers_Count(i)

ctemp.Add cFrontLine.Buffers_GetCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Buffers
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cFrontLine.Modifiers_Count(i)

ctemp.Add cFrontLine.Modifiers_GetCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Permanent Record of Battlesite
If myBattleSite.ID = 0 Then
    Print #X, ""
Else
    Set ctemp = New Collection
    
    For k = 1 To myBattleSite.PermanentRecord_Count
        ctemp.Add myBattleSite.PermanentRecord_GetCard(k)
    Next k
    
    Print #X, GetCode_CardString(ctemp)

End If



Print #X, nTurn
Print #X, bIGoFirst
Close #X


'===========================================================

'=======OPPONENT RESUME INFO================================

X = Dir(App.Path & "\Resume", vbDirectory)
If X = "" Then MkDir App.Path & "\Resume"


a$ = sOpponentName
a$ = ReplaceAllInString(a$, "\", "")
a$ = ReplaceAllInString(a$, "/", "")
a$ = ReplaceAllInString(a$, ".", "")
a$ = ReplaceAllInString(a$, "?", "")
a$ = ReplaceAllInString(a$, "*", "")
a$ = ReplaceAllInString(a$, "<", "")
a$ = ReplaceAllInString(a$, ">", "")
a$ = ReplaceAllInString(a$, "|", "")

a$ = a$ & " " & Format(Now(), "mm-dd-yyyy") & ".re2"

X = FreeFile
Open App.Path & "\Resume\" & a$ For Output As #X

'Write Heroes
For i = 1 To 4

'Print character ID
Print #X, cOpponent.Character_ID(i)

'Is Alive?
Print #X, cOpponent.isCharacterDead(i)

'In reserve?
Print #X, cOpponent.isCharacterReserve(i)

Next i


'Homebase ID
Print #X, OpHomebase.ID

'Battlesite ID
Print #X, OpBattlesite.ID

If OpBattlesite.ID = 0 Then
    Print #X, ""
Else
    Set ctemp = New Collection
     
    For i = 1 To OpBattlesite.Deck_Count
        Set ccard = OpBattlesite.Deck_GetCard(i)
        ctemp.Add ccard
    Next i
    
    Print #X, GetCode_CardString(ctemp)
    
End If

Print #X, GetCode_CardString(cHandO)
Print #X, GetCode_CardString(cDrawPileO)


Print #X, GetCode_CardString(cDiscardPileO)


Print #X, GetCode_CardString(cDeadPileO)


Print #X, GetCode_CardString(cDefeatedCharactersPileO)


'missions
Print #X, cCompletedMissionsO.Count
Print #X, cMissionsO.Count
Print #X, cDeadMissionsO.Count
Print #X, cVenturedMissionsO.Count
Print #X, cVenturedCO.Count



'Placed on Characters
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cOpponent.Placed_Count(i)

ctemp.Add cOpponent.PlacedCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Placed on Homebase
If OpHomebase.ID = 0 Then
    Print #X, ""
Else
    
Set ctemp = New Collection

For k = 1 To OpHomebase.Placed_Count

ctemp.Add OpHomebase.PlacedCard(k)

Next k

Print #X, GetCode_CardString(ctemp)
    

End If





'Permanent Records
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cOpponent.PermanentRecord_Count(i)

ctemp.Add cOpponent.PermanentRecord_GetCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Buffers
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cOpponent.Buffers_Count(i)

ctemp.Add cOpponent.Buffers_GetCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Buffers
For i = 1 To 4

Set ctemp = New Collection

For k = 1 To cOpponent.Modifiers_Count(i)

ctemp.Add cOpponent.Modifiers_GetCard(i, k)

Next k

Print #X, GetCode_CardString(ctemp)

Next i



'Permanent Record of Battlesite
If OpBattlesite.ID = 0 Then
    Print #X, ""
Else
    Set ctemp = New Collection
    
    For k = 1 To OpBattlesite.PermanentRecord_Count
        ctemp.Add OpBattlesite.PermanentRecord_GetCard(k)
    Next k
    
    Print #X, GetCode_CardString(ctemp)

End If


Close #X

'======================================================================

frmTable.Refresh

Timer1.Enabled = True

Status ""

myPhase = nPhase_Draw
UpdatePhase
DrawNewHand

bIHaveConceded = False
bOpponentConceded = False

ShowDiscardFrame

End Sub

Private Sub WhoGoesFirst()
If nTurn > 1 Then Exit Sub
If cDeadPile.Count > 0 Then Exit Sub

MsgBox "The computer will now draw random cards from each draw pile so that the Host can determine who will go first.", vbOKOnly, "Who Goes First?"


lblWGF1.Caption = mySettings.PlayerName
lblWGF2.Caption = sOpponentName
cmdWGFDraw1.Caption = "Redraw (" & mySettings.PlayerName & ")"
cmdWGFDraw2.Caption = "Redraw (" & sOpponentName & ")"
cmdWGF1.Caption = mySettings.PlayerName & " Goes First"
cmdWGF2.Caption = sOpponentName & " Goes First"
frmWhoGoesFirst.Visible = True

WGFDrawCards


End Sub
Private Sub WGFDrawCards()
Dim ccard
Dim ccard2

Randomize

p1 = Int((cDrawPile.Count * Rnd) + 1)
p2 = Int((cDrawPileO.Count * Rnd) + 1)
t1$ = cDrawPile.Item(p1).Title
t2$ = cDrawPileO.Item(p2).Title

imgWGF1.ToolTipText = t1$
imgWGF2.ToolTipText = t2$

If cDrawPile.Item(p1).LoadImage(cDrawPile.Item(p1).ID) = True Then
    imgWGF1.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgWGF1.Picture = LoadPicture(sBlankImagePath)
End If

If cDrawPileO.Item(p2).LoadImage(cDrawPileO.Item(p2).ID) = True Then
    imgWGF2.Picture = LoadPicture(App.Path & "\temppic.jpg")
Else
    imgWGF2.Picture = LoadPicture(sBlankImagePath)
End If

imgWGF1.Tag = p1
imgWGF2.Tag = p2

SendData "CWG:" & Trim(Str(p1)) & ":" & Trim(Str(p2)) & ":|"

End Sub
Private Sub ShowDiscardFrame()

lblDiscard1(0).Caption = mySettings.PlayerName & ": Discarding"
lblDiscard1(1).Caption = sOpponentName & ": Discarding"
frmDiscardPhase.Visible = True
cmdFinishedDiscarding.Enabled = True

End Sub
Private Sub ShowPlacingFrame()

lblDiscard1(2).Caption = mySettings.PlayerName & ": Placing"
lblDiscard1(3).Caption = sOpponentName & ": Placing"

If bIGoFirst = True Then
    Label6.Caption = mySettings.PlayerName & " should place the first card."
Else
    Label6.Caption = sOpponentName & " should place the first card."
End If

frmPlacingPhase.Visible = True
cmdFinishedPlacing.Enabled = True

End Sub
Private Sub ShowResolveVentureFrame()
Dim nMyVenture As Integer


txtOppVentureTotal.Text = OppVentureTotal
txtMyVentureTotal.Text = MyVentureTotal
cmdSendMyVenTotal.Enabled = True
cmdAcceptOppVenTotal.Tag = ""
chkVTTotalAccepted.Value = 0

frmResolveVenturePhase.Visible = True

End Sub
Private Function MyVentureTotal()
nMyVenture = 0

For i = 1 To 4
        For k = 1 To cOpponent.HitsToCurrentBattle_Count(i)
            Set ccard = cOpponent.HitsToCurrentBattle_GetCard(i, k)
            nMyVenture = nMyVenture + ccard.Attack_VentureValue
       Next k
Next i

For i = 1 To cOpponent.BufferHits_Count
    Set ccard = cOpponent.BufferHits_GetCard(i)
    nMyVenture = nMyVenture + ccard.Attack_VentureValue
Next i

On Error Resume Next

For i = 1 To OpBattlesite.HitsToCurrentBattle_Count
    Set ccard = OpBattlesite.HitsToCurrentBattle_GetCard(i)
    nMyVenture = nMyVenture + ccard.Attack_VentureValue
Next i

MyVentureTotal = nMyVenture

End Function
Private Function OppVentureTotal()
nMyVenture = 0

For i = 1 To 4
        For k = 1 To cFrontLine.HitsToCurrentBattle_Count(i)
            Set ccard = cFrontLine.HitsToCurrentBattle_GetCard(i, k)
            nMyVenture = nMyVenture + ccard.Attack_VentureValue
       Next k
Next i

For i = 1 To cFrontLine.BufferHits_Count
    Set ccard = cFrontLine.BufferHits_GetCard(i)
    nMyVenture = nMyVenture + ccard.Attack_VentureValue
Next i

On Error Resume Next

For i = 1 To myBattleSite.HitsToCurrentBattle_Count
    Set ccard = myBattleSite.HitsToCurrentBattle_GetCard(i)
    nMyVenture = nMyVenture + ccard.Attack_VentureValue
Next i

OppVentureTotal = nMyVenture
End Function
Private Sub ShowVentureTotals()

sbBar1.Panels(2).Text = "Me: " & MyVentureTotal
sbBar1.Panels(3).Text = "Opp: " & OppVentureTotal

End Sub
Private Sub ShowVentureFrame()
lblDiscard1(4).Caption = mySettings.PlayerName & ": Venturing"
lblDiscard1(5).Caption = sOpponentName & ": Venturing"

frmVenturePhase.Visible = True
cmdFinishedVenture.Enabled = True

Me.Caption = "OVERPOWER ONLINE-->" & "MISSIONS: (" & cMissions.Count & ")"
imgMissionCard.Picture = imgMissions.Picture
HideFrames False, True, False
End Sub
Private Sub CheckForAdditionalAttack()

frmDlgAttack.Show 1

If frmDlgAttack.Check1.Value = 1 Then
    HideStringAttackFrame
    SendData "CAC:1:|"
    myPhase = nPhase_Defend
    History_Add sOpponentName & " IS PREPARING TO ATTACK"
    
    For i = 0 To 8
        lnFrontLine(i).Visible = False
    Next i

    shpAction.Visible = False
    frmAttack.Visible = False

    UpdatePhase
End If

Unload frmDlgAttack

End Sub

Private Sub txtMyVentureTotal_Change()
a$ = Trim(txtMyVentureTotal.Text)

If a$ = "" Then
    cmdSendMyVenTotal.Enabled = False
Else
    cmdSendMyVenTotal.Enabled = True
End If


End Sub
Private Sub ShowMoveVentureFrame()

mv = Val(txtMyVentureTotal.Text)
ov = Val(txtOppVentureTotal.Text)

cmdFinishedMovingVenture.Enabled = True

frmResolveVenturePhase.Visible = False

If mv > ov Then
    If sSounds(10) <> "" Then PlaySound sSounds(10)
    
    Label11.Caption = mySettings.PlayerName & " has won the venture."
    History_Add mySettings.PlayerName & " WINS VENTURE"
    bIGoFirst = True
    
End If

If ov > mv Then
    If sSounds(7) <> "" Then PlaySound sSounds(7)
    Label11.Caption = sOpponentName & " has won the venture."
    History_Add sOpponentName & " WINS VENTURE"
    bIGoFirst = False
End If

If ov = mv Then
    Label11.Caption = mySettings.PlayerName & " && " & sOpponentName & " have tied in the venture."
    History_Add "OPPONENTS TIE IN VENTURE"
    
End If

lblMoveVenture(0).Caption = mySettings.PlayerName & ": Moving..."
lblMoveVenture(1).Caption = sOpponentName & ": Moving..."

frmMoveVentureCards.Visible = True

End Sub

Private Sub ShowStringAttackFrame()

frmString.Left = 2640
frmString.ZOrder 1
frmString.Visible = True

For i = 2640 To 4320 Step 400
    frmString.Left = i
    frmString.Refresh
Next i


End Sub
Private Sub HideStringAttackFrame()

frmString.Left = 2640
frmString.Visible = False

End Sub
Private Sub ShowBuffers()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgBuffer1.Count - 1
    Unload imgBuffer1(z)
Next z

For z = 1 To imgBuffer2.Count - 1
    Unload imgBuffer2(z)
Next z

For z = 1 To imgBuffer3.Count - 1
    Unload imgBuffer3(z)
Next z

For z = 1 To imgBuffer4.Count - 1
    Unload imgBuffer4(z)
Next z


For i = 1 To 4

If cFrontLine.isCharacterDead(i) = False Then
    
    Select Case CharPic(i)
    Case 0
        Set pCard = imgBuffer1
    Case 1
        Set pCard = imgBuffer2
    Case 2
        Set pCard = imgBuffer3
    Case 3
        Set pCard = imgBuffer4
    End Select
    
    For k = 1 To cFrontLine.Buffers_Count(i)
        Load pCard(k)
        Set ccard = cFrontLine.Buffers_GetCard(i, k)
                                
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
            pCard(k).ToolTipText = ccard.Effect
        Else
            pCard(k).ToolTipText = ccard.Title
        End If
    
    Next k

End If

Next i


End Sub
Private Sub ShowOpponentBuffers()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgOPBuffer1.Count - 1
    Unload imgOPBuffer1(z)
Next z

For z = 1 To imgOpBuffer2.Count - 1
    Unload imgOpBuffer2(z)
Next z

For z = 1 To imgOpBuffer3.Count - 1
    Unload imgOpBuffer3(z)
Next z

For z = 1 To imgOpBuffer4.Count - 1
    Unload imgOpBuffer4(z)
Next z


For i = 1 To 4

If cOpponent.isCharacterDead(i) = False Then
    
    Select Case OppCharPic(i)
    Case 0
        Set pCard = imgOPBuffer1
    Case 1
        Set pCard = imgOpBuffer2
    Case 2
        Set pCard = imgOpBuffer3
    Case 3
        Set pCard = imgOpBuffer4
    End Select
    
    For k = 1 To cOpponent.Buffers_Count(i)
        Load pCard(k)
        Set ccard = cOpponent.Buffers_GetCard(i, k)
                                
        a = ccard.ID
        
        If ccard.LoadImage(a) = True Then
            pCard(k).Picture = LoadPicture(App.Path & "\temppic.jpg")
        Else
            pCard(k).Picture = LoadPicture(sBlankImagePath)
        End If
        
        pCard(k).Left = pCard(k - 1).Left + 200
        pCard(k).ZOrder (0)
        pCard(k).Visible = True
        
        If ccard.CardType = "Special Card" Then
            pCard(k).ToolTipText = ccard.Effect
        Else
            pCard(k).ToolTipText = ccard.Title
        End If
    
    Next k

End If

Next i


End Sub
Private Sub BufferDetail(Index, oPicture As Image, sTag, bCardisPlaced As Boolean)
Dim ccard

    imgCardDetail.Picture = oPicture.Picture
    imgCardDetail.Tag = Index
    frmModifier.Tag = sTag
    HideFrames False, False, True
    frmModifier.Visible = True
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    
    cmdDiscardModifier.Tag = "BUFFER"
    
    If bCardisPlaced = True Then
        cmdPlayModifier.Enabled = False
        cmdPlaceModifier.Enabled = False
        cmdDiscardModifier.Enabled = True
    Else
        cmdPlayModifier.Enabled = True
        cmdPlaceModifier.Enabled = True
        cmdDiscardModifier.Enabled = True
    End If
    
    sTag = Val(sTag)
    Set ccard = cFrontLine.Buffers_GetCard(sTag, Index)
            
    Me.Caption = "OVERPOWER ONLINE-->" & "BUFFER: " & ccard.Title
    
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If
    
    HideAllBorders
    
End Sub
Private Sub OpponentBufferDetail(Index, oPicture As Image, sTag)
Dim ccard

    imgCardDetail.Picture = oPicture.Picture
    imgCardDetail.Tag = Index
    HideFrames False, False, False
    imgHeroCard.Visible = False
    imgCardDetail.Visible = True
    
    sTag = Val(sTag)
    Set ccard = cOpponent.Buffers_GetCard(sTag, Index)
            
    Me.Caption = "OVERPOWER ONLINE-->" & "BUFFER: " & ccard.Title
    
    
    If ccard.CardType = "Special Card" Then
        imgCardDetail.ToolTipText = ccard.Effect
    Else
        imgCardDetail.ToolTipText = ccard.Title
    End If
    
    HideAllBorders
    
End Sub
Private Sub ResolveConcession(bIConceded As Boolean)

If bIConceded = True Then
    bIGoFirst = False
    txtMyVentureTotal.Text = "-100"
    txtOppVentureTotal.Text = "100"

Else
    txtMyVentureTotal.Text = "100"
    txtOppVentureTotal.Text = "-100"
    bIGoFirst = True
End If

myPhase = nPhase_Resolve
ShowMoveVentureFrame
UpdatePhase

End Sub
Private Sub ShowOpponentBufferHits()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgHitBufferOP.Count - 1
    Unload imgHitBufferOP(z)
Next z

For i = 1 To cOpponent.BufferHits_Count
    Load imgHitBufferOP(i)
    Set ccard = cOpponent.BufferHits_GetCard(i)
    
    a = ccard.ID
        
    If ccard.LoadImage(a) = True Then
        imgHitBufferOP(i).Picture = LoadPicture(App.Path & "\temppic.jpg")
    Else
        imgHitBufferOP(i).Picture = LoadPicture(sBlankImagePath)
    End If
        
        imgHitBufferOP(i).Left = imgHitBufferOP(i - 1).Left + 200
        imgHitBufferOP(i).ZOrder (0)
        imgHitBufferOP(i).Visible = True
        
        If ccard.CardType = "Special Card" Then
            imgHitBufferOP(i).ToolTipText = ccard.Effect
        Else
            imgHitBufferOP(i).ToolTipText = ccard.Title
        End If

Next i


End Sub
Private Sub ShowBufferHits()
Dim ccard
Dim pCard

'Loop through characters

For z = 1 To imgHitBuffer.Count - 1
    Unload imgHitBuffer(z)
Next z

For i = 1 To cFrontLine.BufferHits_Count
    Load imgHitBuffer(i)
    Set ccard = cFrontLine.BufferHits_GetCard(i)
    
    a = ccard.ID
        
    If ccard.LoadImage(a) = True Then
        imgHitBuffer(i).Picture = LoadPicture(App.Path & "\temppic.jpg")
    Else
        imgHitBuffer(i).Picture = LoadPicture(sBlankImagePath)
    End If
        
        imgHitBuffer(i).Left = imgHitBuffer(i - 1).Left + 200
        imgHitBuffer(i).ZOrder (0)
        imgHitBuffer(i).Visible = True
        
        If ccard.CardType = "Special Card" Then
            imgHitBuffer(i).ToolTipText = ccard.Effect
        Else
            imgHitBuffer(i).ToolTipText = ccard.Title
        End If

Next i


End Sub
Private Sub PlayAspect(a)
Dim ccard
a = Val(imgCardDetail.Tag)
Set ccard = cHand.Item(a)

Set myAspect = New clsAspect
myAspect.Load ccard.ID

'Figure out type of aspect card

If myAspect.Attack_isGameEffect = True Or myAspect.Attack_isBattleEffect = True Then
    myHomebase.PlaceCard ccard, myAspect.Attack_isGameEffect
    cHand.Remove a
    cHandTags.Remove a
    ShowPlacedCards
    ShowHand
    
    History_Add ccard.Title
    
    If myAspect.Attack_isGameEffect = True Then
        SendData "CHP:" & Trim(Str(a)) & ":1:|"
    Else
        SendData "CHP:" & Trim(Str(a)) & ":0:|"
    End If
    
    CheckForAdditionalAttack
    Exit Sub
    
End If

If myAspect.Attack_isCharacterModifier = True Then
    
    Load frmChooseCharacter
    frmChooseCharacter.HideHomebase = True
    frmChooseCharacter.Show 1
    
    If frmChooseCharacter.SelectedCharacter = -1 Then
        Unload frmChooseCharacter
        Exit Sub
    End If
    
    b = frmChooseCharacter.SelectedCharacter
    Unload frmChooseCharacter
    
    Set ccard = cHand.Item(a)
    cFrontLine.Modifiers_AddCard b, ccard, modifies_battle
        
    cHand.Remove a
    cHandTags.Remove a
    
    SendData "CSC:2:" & Trim(Str(b + 17)) & ":" & Trim(Str(a)) & ":|"
        
    ShowHand
    ShowModifiers
    HideAllBorders
    
    CheckForAdditionalAttack
    Exit Sub
End If

PlayCard
End Sub
Private Sub LoadResumeInfo(sfiletitle)
'On Error GoTo badopen

Status "Retrieving game information..."

sresdir = App.Path & "\Resume\"

'clear images, set collections
NewGame

X = FreeFile
Open sresdir & sfiletitle & ".res" For Input As #X


'Load myCharacters
Set cFrontLine = New clsFrontLine

For i = 1 To 4
    Line Input #X, a$
    cFrontLine.AddCharacter Val(a$), False, False
    
    Line Input #X, a$
    If CBool(a$) = True Then
        cFrontLine.KillCharacter i
        
    End If
    
    Line Input #X, a$
    If CBool(a$) = True Then
        cFrontLine.isCharacterReserve(i) = True
    End If

Next i

LoadCharacters
     
'Load homebase
Line Input #X, a$
Set myHomebase = New clsHomebase

If a$ <> "" Then
    myHomebase.Load Val(a$)
    LoadHomeBase
End If

'Load Battlesite
Line Input #X, a$
Set myBattleSite = New clsBattlesite

If a$ <> "" Then
    myBattleSite.Load Val(a$)
    loadbattlesite
End If

'load battlesitedeck
Dim ctemp2 As Collection
Set ctemp2 = New Collection

Line Input #X, a$
If a$ <> "" Then
    Code_ImportPileString ctemp2, a$
    
    For i = 1 To ctemp2.Count
    
    myBattleSite.Deck_AddCard ctemp2.Item(i)
    
    Next i
    
    UpdateDeckDisplay
End If

Status "Loading piles..."

Line Input #X, a$
Code_ImportPileString cHand, a$

Line Input #X, a$
Code_ImportPileString cDrawPile, a$

Line Input #X, a$
Code_ImportPileString cDiscardPile, a$

Line Input #X, a$
Code_ImportPileString cDeadPile, a$

Line Input #X, a$
Code_ImportPileString cDefeatedCharactersPile, a$

Line Input #X, a$
Code_ImportPileString cCompletedMissions, a$

Line Input #X, a$
Code_ImportPileString cMissions, a$

Line Input #X, a$
Code_ImportPileString cDeadMissions, a$

Line Input #X, a$
Code_ImportPileString cVenturedMissions, a$

Line Input #X, a$
Code_ImportPileString cVenturedC, a$

'cards placed on characters
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cFrontLine.PlaceCard i, ctemp2.Item(k)
Next k

Next i

Line Input #X, a$
Set ctemp2 = New Collection

Code_ImportPileString ctemp2, a$

For i = 1 To ctemp2.Count
    myHomebase.PlaceCard ctemp2.Item(i), False
Next i

'permanent records
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cFrontLine.PermanentRecord_AddCard i, ctemp2.Item(k)
Next k

Next i

'Buffers
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cFrontLine.Buffers_AddCard i, ctemp2.Item(k)
Next k

Next i

'Modifiers
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cFrontLine.Modifiers_AddCard i, ctemp2.Item(k), Modifies_Game
Next k

Next i

'battlesite deck permanent record
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    myBattleSite.PermanentRecord_AddCard ctemp2.Item(k)
Next k

Line Input #X, a$
nTurn = Val(a$)

Line Input #X, a$

bIGoFirst = CBool(a$)

Close #X

UpdateDeckDisplay
ShowPermanentRecord
ShowPlacedCards
ShowBuffers
ShowModifiers

LoadOpponentResumeInfo sfiletitle

'load messages
X = FreeFile
Open sresdir & sfiletitle & ".rem" For Input As #X

Do Until EOF(X)
Line Input #X, a$
lstMessages.AddItem a$
Loop

Close #X

If lstMessages.ListCount > 0 Then lstMessages.ListIndex = lstMessages.ListCount - 1

'load History
X = FreeFile
Open sresdir & sfiletitle & ".reh" For Input As #X

Do Until EOF(X)
Line Input #X, a$
lstGameHistory.AddItem a$
Loop

Close #X

If lstGameHistory.ListCount > 0 Then lstGameHistory.ListIndex = lstGameHistory.ListCount - 1

'UpdatePhase

Status ""

Exit Sub


badopen:
X = MsgBox("Error in save files.  Unable to resume game.", vbCritical, "Resume Error")

End Sub
Private Sub LoadOpponentResumeInfo(sfiletitle)
Status "Retrieving opponent information..."

sresdir = App.Path & "\Resume\"

'load messages
X = FreeFile
Open sresdir & sfiletitle & ".re2" For Input As #X

'Load myCharacters
Set cOpponent = New clsOpponent

For i = 1 To 4
    Line Input #X, a$
    cOpponent.AddCharacter Val(a$), False, False
    
    Line Input #X, a$
    If CBool(a$) = True Then
        cOpponent.KillCharacter i
        
    End If
    
    Line Input #X, a$
    If CBool(a$) = True Then
        cOpponent.isCharacterReserve(i) = True
    End If

Next i

LoadOpponentCharacters
     
'Load homebase
Line Input #X, a$
Set OpHomebase = New clsHomebase

If a$ <> "" Then
    OpHomebase.Load Val(a$)
    LoadOpponentHomebase
End If

'Load Battlesite
Line Input #X, a$
Set OpBattlesite = New clsBattlesite

If a$ <> "" Then
    OpBattlesite.Load Val(a$)
    LoadOpBattlesite
End If

'load battlesitedeck
Dim ctemp2 As Collection
Set ctemp2 = New Collection

Line Input #X, a$
If a$ <> "" Then
    Code_ImportPileString ctemp2, a$
    
    For i = 1 To ctemp2.Count
    
    OpBattlesite.Deck_AddCard ctemp2.Item(i)
    
    Next i
    
    UpdateOpponentDeckDisplay
End If

Status "Loading piles..."

Line Input #X, a$
Code_ImportPileString cHandO, a$

Line Input #X, a$
Code_ImportPileString cDrawPileO, a$

Line Input #X, a$
Code_ImportPileString cDiscardPileO, a$

Line Input #X, a$
Code_ImportPileString cDeadPileO, a$

Line Input #X, a$
Code_ImportPileString cDefeatedCharactersPileO, a$

Line Input #X, a$
Set cCompletedMissionsO = New Collection
For i = 1 To Val(a$)
cCompletedMissionsO.Add "1"
Next i

Line Input #X, a$
Set cMissionsO = New Collection
For i = 1 To Val(a$)
cMissionsO.Add "1"
Next i

Line Input #X, a$
Set cDeadMissionsO = New Collection
For i = 1 To Val(a$)
cDeadMissionsO.Add "1"
Next i

Line Input #X, a$
Set cVenturedMissionsO = New Collection
For i = 1 To Val(a$)
cVenturedMissionsO.Add "1"
Next i

Line Input #X, a$
Set cVenturedCO = New Collection
For i = 1 To Val(a$)
cVenturedCO.Add "1"
Next i

'cards placed on characters
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cOpponent.PlaceCard i, ctemp2.Item(k)
Next k

Next i

Line Input #X, a$
Set ctemp2 = New Collection

Code_ImportPileString ctemp2, a$

For i = 1 To ctemp2.Count
    OpHomebase.PlaceCard ctemp2.Item(i), False
Next i

'permanent records
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cOpponent.PermanentRecord_AddCard i, ctemp2.Item(k)
Next k

Next i

'Buffers
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cOpponent.Buffers_AddCard i, ctemp2.Item(k)
Next k

Next i

'Modifiers
For i = 1 To 4
Set ctemp2 = New Collection

Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    cOpponent.Modifiers_AddCard i, ctemp2.Item(k), Modifies_Game
Next k

Next i

'perm record of battlesite
Set ctemp2 = New Collection
Line Input #X, a$
Code_ImportPileString ctemp2, a$

For k = 1 To ctemp2.Count
    OpBattlesite.PermanentRecord_AddCard ctemp2.Item(k)
Next k

UpdateOpponentDeckDisplay
ShowOpponentPermanentRecord
ShowOpponentPlacedCards
ShowOpponentBuffers
ShowOpponentModifiers


Status ""

Close #X

Exit Sub


End Sub
Private Sub ResumeGame()

myPhase = nPhase_Draw
UpdatePhase
DrawNewHand

bIHaveConceded = False
bOpponentConceded = False

ShowDiscardFrame

End Sub
Sub SaveOpponentResumeInfo()


End Sub

Sub SaveResumeInfo()


End Sub

