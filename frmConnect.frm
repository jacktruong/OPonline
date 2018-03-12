VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConnect 
   Caption         =   "Connect to Opponent"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&New Game"
      TabPicture(0)   =   "frmConnect.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblWarning"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "optCType(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtIPAddress"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPort"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optCType(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbIP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPort2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdGo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCancel"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdCheckIP"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAutoConnect"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "&Resume Game"
      TabPicture(1)   =   "frmConnect.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(5)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(6)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2(7)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblWarning2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdCancel2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdGO2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtPort4"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cbIP2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "optCType2(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtPort3"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtIP2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "optCType2(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lstGames"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmdDelete"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68400
         TabIndex        =   35
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdAutoConnect 
         Caption         =   "Auto Connect"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "If you have set up a game via IRC, click here."
         Top             =   4080
         Width           =   1455
      End
      Begin VB.ComboBox lstGames 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2880
         Width           =   4695
      End
      Begin VB.OptionButton optCType2 
         Caption         =   "&Host"
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
         Left            =   -74760
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtIP2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73320
         TabIndex        =   10
         Top             =   1395
         Width           =   2775
      End
      Begin VB.TextBox txtPort3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -69840
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optCType2 
         Caption         =   "&Guest"
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
         Left            =   -74760
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cbIP2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73080
         TabIndex        =   14
         Text            =   "[Enter or Select IP Address]"
         Top             =   2235
         Width           =   2895
      End
      Begin VB.TextBox txtPort4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -69480
         TabIndex        =   15
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdGO2 
         Caption         =   "Go!"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69600
         TabIndex        =   16
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -68520
         TabIndex        =   17
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Check IP"
         Height          =   375
         Left            =   -68760
         TabIndex        =   12
         Top             =   1395
         Width           =   855
      End
      Begin VB.CommandButton cmdCheckIP 
         Caption         =   "Check IP"
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   1400
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go!"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtPort2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox cbIP 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Text            =   "[Enter or Select IP Address]"
         Top             =   2240
         Width           =   2895
      End
      Begin VB.OptionButton optCType 
         Caption         =   "&Guest"
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
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   2
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtIPAddress 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   1400
         Width           =   2775
      End
      Begin VB.OptionButton optCType 
         Caption         =   "&Host"
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
         Left            =   240
         TabIndex        =   0
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Resume Game:"
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
         Left            =   -74640
         TabIndex        =   32
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblWarning2 
         Caption         =   "**You already have a deck open.  Please restart the program and select Opponent-->Connect to Opponent without opening a deck.**"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   -74760
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.Label Label3 
         Caption         =   "Connect as:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "My IP Address:"
         Height          =   255
         Index           =   7
         Left            =   -74520
         TabIndex        =   29
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Index           =   6
         Left            =   -70320
         TabIndex        =   28
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Host's IP Address:"
         Height          =   255
         Index           =   5
         Left            =   -74520
         TabIndex        =   27
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Index           =   4
         Left            =   -69960
         TabIndex        =   26
         Top             =   2325
         Width           =   375
      End
      Begin VB.Label lblWarning 
         Caption         =   "**Please open a deck (File-->Open Deck) before connecting to an opponent for a new game.**"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   3120
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   23
         Top             =   2325
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Host's IP Address:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   22
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   21
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "My IP Address:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Connect as:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Label lblResult 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbIP_Change()

CheckGo2

End Sub

Private Sub cbIP_Click()

CheckGo2

p$ = txtPort2.Text

If cbIP.ListIndex = -1 Then Exit Sub

p2 = cbIP.ItemData(cbIP.ListIndex)

If p2 > 0 Then txtPort2.Text = Trim(Str(p2))

End Sub

Private Sub cbIP_GotFocus()
On Error Resume Next

With cbIP
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub

Private Sub cbIP2_Change()
CheckGo4

End Sub

Private Sub cmdAutoConnect_Click()
X = Dir("c\autogame.ini")
If X = "" Then
    MsgBox "Autogame.ini not found.  Please set up your game in the IRC channel to use Auto-Connect.", vbCritical, "Connect Information Not Found"
    Exit Sub
End If

If lblWarning.Visible = True Then
    X = MsgBox(lblWarning.Caption, vbCritical, "Cannot Begin New Game")
    Exit Sub
End If

lblResult.Caption = "3"
Me.Hide

End Sub

Private Sub cmdCancel_Click()
SaveIPs

lblResult.Caption = "0"
Me.Hide

End Sub

Private Sub cmdCancel2_Click()
SaveIPs

lblResult.Caption = "0"
Me.Hide
End Sub

Private Sub cmdCheckIP_Click()
txtIPAddress.Text = frmTable.tcpChannel.LocalIP

End Sub

Private Sub cmdDelete_Click()
X = MsgBox("Are you sure you want to delete resume information for this game: " & a$ & "?", vbYesNoCancel, "Delete Resume Info?")

If X <> 6 Then Exit Sub

a$ = lstGames.List(lstGames.ListIndex)
di$ = App.Path & "\resume\" & a$

X = Dir(di$ & ".res", vbNormal)
If X <> "" Then Kill di$ & ".res"

X = Dir(di$ & ".re2", vbNormal)
If X <> "" Then Kill di$ & ".re2"

X = Dir(di$ & ".reh", vbNormal)
If X <> "" Then Kill di$ & ".reh"

X = Dir(di$ & ".rem", vbNormal)
If X <> "" Then Kill di$ & ".rem"

lstGames.RemoveItem lstGames.ListIndex

If lstGames.ListCount > 0 Then lstGames.ListIndex = 0

End Sub

Private Sub cmdGo_Click()
If lblWarning.Visible = True Then
    X = MsgBox(lblWarning.Caption, vbCritical, "Cannot Begin New Game")
    Exit Sub
End If

SaveIPs

lblResult.Caption = "1"
Me.Hide

End Sub
Private Sub SaveIPs()
Dim ctemp As Collection
Dim ctemp2 As Collection
Dim bWrite1 As Boolean
Dim bWrite2 As Boolean

Set ctemp = New Collection
Set ctemp2 = New Collection

For i = 0 To cbIP.ListCount - 1
    ctemp.Add cbIP.List(i)
    ctemp2.Add cbIP.ItemData(i)
Next i

ip$ = cbIP.Text
pt$ = txtPort2.Text

If ip$ = "" Then Exit Sub

bWrite1 = True

For i = 1 To ctemp.Count
    If ctemp.Item(i) = ip$ Then
        bWrite1 = False
        Exit For
    End If
Next i

bWrite2 = True

ip2$ = cbIP2.Text
pt2$ = txtPort4.Text

For i = 1 To ctemp.Count
    If ctemp.Item(i) = ip2$ Then
        bWrite2 = False
        Exit For
    End If
Next i

X = FreeFile
Open App.Path & "\ips.txt" For Output As #X

If bWrite1 = True Then
    Print #X, ip$
    Print #X, pt$
End If

If bWrite2 = True Then
    Print #X, ip2$
    Print #X, pt2$
End If

For i = 1 To ctemp.Count
Print #X, ctemp.Item(i)
Print #X, ctemp2.Item(i)
Next i

Close #X

End Sub

Private Sub cmdGO2_Click()
If lblWarning2.Visible = True Then
    X = MsgBox(lblWarning2.Caption, vbCritical, "Cannot Resume Game")
    Exit Sub
End If

SaveIPs

lblResult.Caption = "2"
Me.Hide
End Sub

Private Sub Command1_Click()
txtIP2.Text = frmTable.tcpChannel.LocalIP

End Sub

Private Sub Form_Load()

LoadConnectSettings

End Sub
Private Sub LoadConnectSettings()
On Error Resume Next

X = Dir(App.Path & "\resume\*.res")
If X = "" Then
    SSTab1.TabEnabled(1) = False
Else
    SSTab1.TabEnabled(1) = True
    
    
    'loadresume games
    
    
X = Dir(App.Path & "\resume\*.res")

looper:

    If X <> "" Then

    a$ = Left(X, Len(X) - 4)
    lstGames.AddItem a$
    X = Dir()
    GoTo looper
    End If
    
    If lstGames.ListCount > 0 Then lstGames.ListIndex = 0
End If

txtIPAddress.Text = mySettings.IP_Address
txtPort.Text = mySettings.Port
txtPort2.Text = mySettings.Port
txtPort4.Text = mySettings.Port
txtPort3.Text = mySettings.Port

txtIP2.Text = mySettings.IP_Address

'optCType(0).Value = True

If cDrawPile.Count = 0 Then lblWarning.Visible = True
If cDrawPile.Count > 0 Then lblWarning2.Visible = True

X = Dir(App.Path & "\ips.txt")
If X <> "" Then

X = FreeFile

Open App.Path & "\ips.txt" For Input As #X

While Not EOF(X)

Line Input #X, a$
Line Input #X, p$

cbIP.AddItem a$
cbIP.ItemData(cbIP.NewIndex) = Val(p$)
cbIP2.AddItem a$
cbIP2.ItemData(cbIP2.NewIndex) = Val(p$)

Wend

Close #X

End If

If cbIP.ListCount > 0 Then cbIP.ListIndex = 0

End Sub

Private Sub lstGames_Click()
If lstGames.ListIndex = -1 Then
    cmdDelete.Enabled = False
Else
    cmdDelete.Enabled = True
End If

End Sub

Private Sub optCType_Click(Index As Integer)
On Error Resume Next

Select Case Index
Case 0

    txtIPAddress.Enabled = True
    txtPort.Enabled = True
    
    txtIPAddress.SetFocus
    
    cbIP.Enabled = False
    txtPort2.Enabled = False

    CheckGo1
    
Case 1
    txtIPAddress.Enabled = False
    txtPort.Enabled = False
    
    cbIP.Enabled = True
    txtPort2.Enabled = True
    
    cbIP.SetFocus
    CheckGo2
        
End Select

End Sub

Private Sub optCType2_Click(Index As Integer)
On Error Resume Next

Select Case Index
Case 0

    txtIP2.Enabled = True
    txtPort3.Enabled = True
    
    txtIP2.SetFocus
    
    cbIP2.Enabled = False
    txtPort4.Enabled = False
    lstGames.Enabled = True
    
    CheckGo3
    
Case 1
    txtIP2.Enabled = False
    txtPort3.Enabled = False
    
    cbIP2.Enabled = True
    txtPort4.Enabled = True
    lstGames.Enabled = True
    cbIP2.SetFocus
    CheckGo4
        
End Select
End Sub

Private Sub txtIP2_Change()
CheckGo3

End Sub

Private Sub txtIP2_GotFocus()
With txtIP2
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub

Private Sub txtIPAddress_Change()

CheckGo1

End Sub

Private Sub txtIPAddress_GotFocus()
With txtIPAddress
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub

Private Sub txtPort_Change()

CheckGo1

End Sub

Private Sub txtPort_GotFocus()
With txtPort
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub

Private Sub txtPort2_Change()
CheckGo2

End Sub

Private Sub txtPort2_GotFocus()
With txtPort2
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub
Private Sub CheckGo1()

If optCType(0).Value = True And Trim(txtIPAddress.Text) <> "" And Val(txtPort.Text) > 0 Then
    cmdGo.Enabled = True
Else
    cmdGo.Enabled = False
End If
End Sub
Private Sub CheckGo2()

If optCType(1).Value = True And Trim(cbIP.Text) <> "" And Val(txtPort2.Text) > 0 Then
    cmdGo.Enabled = True
Else
    cmdGo.Enabled = False
End If

End Sub
Private Sub CheckGo3()

If optCType2(0).Value = True And Trim(txtIP2.Text) <> "" And Val(txtPort3.Text) > 0 Then
    cmdGO2.Enabled = True
Else
    cmdGO2.Enabled = False
End If
End Sub
Private Sub CheckGo4()

If optCType2(1).Value = True And Trim(cbIP2.Text) <> "" And Val(txtPort4.Text) > 0 Then
    cmdGO2.Enabled = True
Else
    cmdGO2.Enabled = False
End If

End Sub

Private Sub txtPort3_Change()
CheckGo3

End Sub

Private Sub txtPort3_GotFocus()
With txtPort3
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub

Private Sub txtPort4_Change()
CheckGo4

End Sub

Private Sub txtPort4_GotFocus()
With txtPort4
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
