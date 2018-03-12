VERSION 5.00
Begin VB.Form frmViewCards 
   Caption         =   "View Opponent Pile"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   2940
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox lstPowerTypes 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmViewCards.frx":0000
      Left            =   1920
      List            =   "frmViewCards.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2010
      Width           =   735
   End
   Begin VB.OptionButton optView 
      Caption         =   "Only Power Cards"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton optView 
      Caption         =   "Only Universe"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.OptionButton optView 
      Caption         =   "Only Specials"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CheckBox chkCancel 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtRandom 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "2"
      Top             =   960
      Width           =   375
   End
   Begin VB.OptionButton optView 
      Caption         =   "Random"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   975
      Width           =   1095
   End
   Begin VB.TextBox txtTop 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "3"
      Top             =   590
      Width           =   375
   End
   Begin VB.OptionButton optView 
      Caption         =   "Top"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton optView 
      Caption         =   "Entire pile (all cards)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "cards"
      Height          =   255
      Index           =   1
      Left            =   1725
      TabIndex        =   6
      Top             =   1005
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "cards"
      Height          =   255
      Index           =   0
      Left            =   1360
      TabIndex        =   3
      Top             =   630
      Width           =   495
   End
End
Attribute VB_Name = "frmViewCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
chkCancel.Value = 1
Me.Hide

End Sub

Private Sub cmdOK_Click()
chkCancel.Value = 0

Me.Hide

End Sub

Private Sub Form_Load()
lstPowerTypes.AddItem "ALL"
lstPowerTypes.AddItem "E"
lstPowerTypes.AddItem "F"
lstPowerTypes.AddItem "S"
lstPowerTypes.AddItem "I"

lstPowerTypes.ListIndex = 0

End Sub

Private Sub optView_Click(Index As Integer)
If Index = 1 Then
    txtTop.Enabled = True
    txtTop.SetFocus
Else
    txtTop.Enabled = False
End If

If Index = 2 Then
    txtRandom.Enabled = True
    txtRandom.SetFocus
Else
    txtRandom.Enabled = False
End If

End Sub

Private Sub txtRandom_GotFocus()
txtRandom.SelStart = 0
txtRandom.SelLength = Len(txtRandom.Text)

End Sub

Private Sub txtRandom_LostFocus()
If Val(txtRandom.Text) = 0 Then txtRandom.Text = "2"

End Sub

Private Sub txtTop_GotFocus()
txtTop.SelStart = 0
txtTop.SelLength = Len(txtTop.Text)

End Sub

Private Sub txtTop_LostFocus()
If Val(txtTop.Text) = 0 Then txtTop.Text = "3"

End Sub
