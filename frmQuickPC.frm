VERSION 5.00
Begin VB.Form frmQuickPC 
   Caption         =   "Power Cards Quick-Add Tool"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Tag             =   " "
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   67
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   66
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   7
      Left            =   2925
      TabIndex        =   65
      Tag             =   "153"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   4
      Left            =   1710
      TabIndex        =   64
      Tag             =   "152"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   60
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   495
      TabIndex        =   59
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   930
      TabIndex        =   58
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   57
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   1710
      TabIndex        =   56
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   2130
      TabIndex        =   55
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   2535
      TabIndex        =   54
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   2925
      TabIndex        =   53
      Text            =   "0"
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Tag             =   "35"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   45
      Tag             =   "36"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   2
      Left            =   930
      TabIndex        =   46
      Tag             =   "37"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   47
      Tag             =   "38"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   5
      Left            =   2130
      TabIndex        =   48
      Tag             =   "33"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtMulti 
      Height          =   375
      Index           =   6
      Left            =   2535
      TabIndex        =   49
      Tag             =   "34"
      Text            =   "0"
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   7
      Left            =   2925
      TabIndex        =   42
      Tag             =   "32"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   6
      Left            =   2535
      TabIndex        =   41
      Tag             =   "31"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   5
      Left            =   2130
      TabIndex        =   40
      Tag             =   "30"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   4
      Left            =   1710
      TabIndex        =   39
      Tag             =   "29"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   38
      Tag             =   "28"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   2
      Left            =   930
      TabIndex        =   37
      Tag             =   "27"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   1
      Left            =   495
      TabIndex        =   36
      Tag             =   "26"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtIntellect 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Tag             =   "25"
      Text            =   "0"
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   7
      Left            =   2925
      TabIndex        =   31
      Tag             =   "24"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   6
      Left            =   2535
      TabIndex        =   30
      Tag             =   "23"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   5
      Left            =   2130
      TabIndex        =   29
      Tag             =   "22"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   4
      Left            =   1710
      TabIndex        =   28
      Tag             =   "21"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   27
      Tag             =   "20"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   2
      Left            =   930
      TabIndex        =   26
      Tag             =   "19"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   1
      Left            =   495
      TabIndex        =   25
      Tag             =   "18"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Tag             =   "17"
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   7
      Left            =   2925
      TabIndex        =   20
      Tag             =   "16"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   6
      Left            =   2535
      TabIndex        =   19
      Tag             =   "15"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   5
      Left            =   2130
      TabIndex        =   18
      Tag             =   "14"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   4
      Left            =   1710
      TabIndex        =   17
      Tag             =   "13"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   16
      Tag             =   "12"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   2
      Left            =   930
      TabIndex        =   15
      Tag             =   "11"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   1
      Left            =   495
      TabIndex        =   14
      Tag             =   "10"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFighting 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Tag             =   "9"
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   7
      Left            =   2920
      TabIndex        =   9
      Tag             =   "8"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   6
      Left            =   2540
      TabIndex        =   8
      Tag             =   "7"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   5
      Left            =   2130
      TabIndex        =   7
      Tag             =   "6"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   4
      Left            =   1710
      TabIndex        =   6
      Tag             =   "5"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Tag             =   "4"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   2
      Left            =   930
      TabIndex        =   4
      Tag             =   "3"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   1
      Left            =   500
      TabIndex        =   3
      Tag             =   "2"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtEnergy 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "1"
      Text            =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgHero 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1650
      Index           =   3
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   2370
   End
   Begin VB.Image imgHero 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1650
      Index           =   2
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   2370
   End
   Begin VB.Image imgHero 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1650
      Index           =   0
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2370
   End
   Begin VB.Image imgHero 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1650
      Index           =   1
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "TOTALS:"
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
      Index           =   5
      Left            =   1440
      TabIndex        =   63
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "1     2     3     4     5     6     7     8"
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
      Left            =   120
      TabIndex        =   62
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label lblBigTotal 
      Caption         =   "(0)"
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
      Left            =   3360
      TabIndex        =   61
      Top             =   6045
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Multi/Anypower"
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
      Index           =   4
      Left            =   120
      TabIndex        =   52
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "1     2     3     4     5     6     7     8"
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
      Left            =   120
      TabIndex        =   51
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label lblMultiTotal 
      Caption         =   "(0)"
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
      Left            =   3360
      TabIndex        =   50
      Top             =   4965
      Width           =   495
   End
   Begin VB.Label lblIntellectTotal 
      Caption         =   "(0)"
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
      Left            =   3360
      TabIndex        =   43
      Top             =   3885
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "1     2     3     4     5     6     7     8"
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
      Left            =   120
      TabIndex        =   34
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Intellect:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblStrengthTotal 
      Caption         =   "(0)"
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
      Left            =   3360
      TabIndex        =   32
      Top             =   2805
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "1     2     3     4     5     6     7     8"
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
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Strength:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblFightingTotal 
      Caption         =   "(0)"
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
      Left            =   3360
      TabIndex        =   21
      Top             =   1725
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "1     2     3     4     5     6     7     8"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Fighting:"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblEnergyTotal 
      Caption         =   "(0)"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   640
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "1     2     3     4     5     6     7     8"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Energy:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmQuickPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide

End Sub

Private Sub Command2_Click()
For i = 0 To 7
    txtIntellect(i).Text = "0"
    txtStrength(i).Text = "0"
    txtFighting(i).Text = "0"
    txtEnergy(i).Text = "0"
    txtMulti(i).Text = "0"
Next i

Me.Hide

End Sub

Private Sub txtEnergy_GotFocus(Index As Integer)
With txtEnergy(Index)
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub

Private Sub txtEnergy_LostFocus(Index As Integer)
With txtEnergy(Index)
    .Text = Val(.Text)
    If .Text = "" Then .Text = "0"
End With

UpdateTotals

End Sub
Private Sub txtFighting_GotFocus(Index As Integer)
With txtFighting(Index)
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub
Private Sub txtFighting_LostFocus(Index As Integer)
With txtFighting(Index)
    .Text = Val(.Text)
    If .Text = "" Then .Text = "0"
End With

UpdateTotals

End Sub
Private Sub txtIntellect_GotFocus(Index As Integer)
With txtIntellect(Index)
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub
Private Sub txtIntellect_LostFocus(Index As Integer)
With txtIntellect(Index)
    .Text = Val(.Text)
    If .Text = "" Then .Text = "0"
End With

UpdateTotals
End Sub



Private Sub txtMulti_GotFocus(Index As Integer)
With txtMulti(Index)
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub

Private Sub txtMulti_LostFocus(Index As Integer)
With txtMulti(Index)
    .Text = Val(.Text)
    If .Text = "" Then .Text = "0"
End With

UpdateTotals
End Sub

Private Sub txtStrength_GotFocus(Index As Integer)
With txtStrength(Index)
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub

Private Sub txtStrength_LostFocus(Index As Integer)
With txtStrength(Index)
    .Text = Val(.Text)
    If .Text = "" Then .Text = "0"
End With

UpdateTotals

End Sub
Private Sub UpdateTotals()

For i = 0 To 7

txtTotal(i).Text = Trim(Str(Val(txtIntellect(i).Text) + Val(txtStrength(i).Text) + Val(txtFighting(i).Text) + Val(txtEnergy(i).Text) + Val(txtMulti(i).Text)))

Next i

ntot = 0
For i = 0 To 7
ntot = ntot + Val(txtIntellect(i).Text)
Next i

lblIntellectTotal.Caption = Trim(Str(ntot))

'energy
ntot = 0
For i = 0 To 7
ntot = ntot + Val(txtEnergy(i).Text)
Next i

lblEnergyTotal.Caption = Trim(Str(ntot))


'Fighting
ntot = 0
For i = 0 To 7
ntot = ntot + Val(txtFighting(i).Text)
Next i

lblFightingTotal.Caption = Trim(Str(ntot))

'Strtength
ntot = 0
For i = 0 To 7
ntot = ntot + Val(txtStrength(i).Text)
Next i

lblStrengthTotal.Caption = Trim(Str(ntot))

'Multipower
ntot = 0
For i = 0 To 7
ntot = ntot + Val(txtMulti(i).Text)
Next i

lblMultiTotal.Caption = Trim(Str(ntot))

ntot = 0
For i = 0 To 7
ntot = ntot + Val(txtTotal(i).Text)
Next i

lblBigTotal.Caption = Trim(Str(ntot))


End Sub
