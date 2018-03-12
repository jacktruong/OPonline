VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Settings"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Top             =   5280
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmSettings.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblName(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblName(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkBeep"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkPopupMessage"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCheckIP"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkAutoIPCheck"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPort"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtIP"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Sounds"
      TabPicture(1)   =   "frmSettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstSounds"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdPreview"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lstSelected"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdPreview2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdSet"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdClear"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70560
         TabIndex        =   22
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71640
         TabIndex        =   21
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdPreview2 
         Caption         =   "Preview"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68760
         TabIndex        =   20
         Top             =   4200
         Width           =   975
      End
      Begin VB.ListBox lstSelected 
         Height          =   2790
         Left            =   -71640
         TabIndex        =   18
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73080
         TabIndex        =   17
         Top             =   4200
         Width           =   975
      End
      Begin VB.ListBox lstSounds 
         Height          =   2790
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1020
         Width           =   3375
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkAutoIPCheck 
         Caption         =   "Automatically verify IP"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton cmdCheckIP 
         Caption         =   "Check"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkPopupMessage 
         Caption         =   "Show incoming messages in popup window"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CheckBox chkBeep 
         Caption         =   "Beep when message received"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   2760
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Current Sound Settings:"
         Height          =   255
         Index           =   1
         Left            =   -71640
         TabIndex        =   19
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Available sounds:"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   15
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Place .wav files in ...."
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label lblName 
         Caption         =   "User Name/ID:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "(8 Characters Max)"
         Height          =   255
         Index           =   1
         Left            =   3525
         TabIndex        =   10
         Top             =   645
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Host IP Address:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1500
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAutoIPCheck_Click()
If chkAutoIPCheck.Value = 0 Then
    mySettings.AutoCheckIP = False
Else
    mySettings.AutoCheckIP = True
End If

End Sub

Private Sub chkBeep_Click()

If chkBeep.Value = 0 Then
    mySettings.MessageBeep = False
Else
    mySettings.MessageBeep = True
End If

End Sub

Private Sub chkPopupMessage_Click()

If chkPopupMessage.Value = 0 Then
    mySettings.PopupMessages = False
Else
    mySettings.PopupMessages = True
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCheckIP_Click()
txtIP.Text = frmTable.tcpChannel.LocalIP
End Sub

Private Sub cmdClear_Click()
a = lstSelected.ListIndex

sSounds(a + 1) = ""
LoadSoundInfo
lstSelected.ListIndex = a

End Sub

Private Sub cmdOK_Click()
If txtIP.Text = "" Then txtIP.Text = frmTable.tcpChannel.LocalIP
If txtPort.Text = "" Then txtPort.Text = "1564"
If txtName.Text = "" Then txtName.Text = "[NONE]"

mySettings.IP_Address = txtIP.Text
mySettings.Port = Val(txtPort.Text)
mySettings.PlayerName = txtName.Text

SaveSoundSettings

Unload Me


End Sub

Private Sub cmdPreview_Click()
PlaySound lstSounds.List(lstSounds.ListIndex)

End Sub

Private Sub cmdPreview2_Click()

PlaySound sSounds(lstSelected.ListIndex + 1)

End Sub

Private Sub cmdSet_Click()

a = lstSelected.ListIndex + 1

sSounds(a) = lstSounds.List(lstSounds.ListIndex)

LoadSoundInfo

lstSelected.ListIndex = a - 1

End Sub

Private Sub Form_Load()

Label2.Caption = "Place .wav files in " & App.Path & "\Sounds\"

X = Dir(App.Path & "\sounds\*.wav")

If X <> "" Then

looper:
lstSounds.AddItem X

X = Dir()
If X <> "" Then GoTo looper

End If

LoadSoundInfo

mySettings.Load
txtName.Text = mySettings.PlayerName

txtIP.Text = mySettings.IP_Address
txtPort.Text = mySettings.Port

If txtIP.Text = "" Then txtIP.Text = frmTable.tcpChannel.LocalIP
If txtPort.Text = "" Then txtPort.Text = "1564"

chkAutoIPCheck.Value = CInt(mySettings.AutoCheckIP) * -1
chkPopupMessage.Value = CInt(mySettings.PopupMessages) * -1
chkBeep.Value = CInt(mySettings.MessageBeep) * -1

If chkAutoIPCheck.Value <> 0 Then
    txtIP.Text = frmTable.tcpChannel.LocalIP
    mySettings.IP_Address = txtIP.Text
End If

End Sub

Private Sub lstSelected_Click()
If lstSelected.ListIndex = -1 Then
    cmdSet.Enabled = False
    cmdClear.Enabled = False
    cmdPreview.Enabled = False
    Exit Sub
End If

If sSounds(lstSelected.ListIndex + 1) = "" Then
    cmdPreview2.Enabled = False
Else
    cmdPreview2.Enabled = True
End If

cmdClear.Enabled = True

If lstSounds.ListIndex = -1 Then
    cmdSet.Enabled = False
Else
    cmdSet.Enabled = True
End If

End Sub

Private Sub lstSounds_Click()
If lstSounds.ListIndex = -1 Then
    cmdPreview.Enabled = False
    cmdSet.Enabled = False
    Exit Sub
End If

cmdPreview.Enabled = True

If lstSelected.ListIndex = -1 Then
    cmdSet.Enabled = False
    cmdClear.Enabled = False
    cmdPreview2.Enabled = False
Else
    cmdSet.Enabled = True
    cmdClear.Enabled = True
    cmdPreview2.Enabled = True
End If

End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)

End Sub
Private Sub LoadSoundInfo()

lstSelected.Clear

For i = 1 To UBound(sSoundDes())

a$ = sSoundDes(i) & ": "

If sSounds(i) = "" Then
    a$ = a$ & "[NONE]"
Else
    a$ = a$ & StripDirectory(sSounds(i))
End If

lstSelected.AddItem a$


Next i

End Sub
Private Sub SaveSoundSettings()

X = Dir(App.Path & "\sounds\sound.ini")
If X = "" Then Exit Sub

X = FreeFile
Open App.Path & "\sounds\sound.ini" For Output As #X

For i = 1 To 18

Print #X, sSounds(i)
Next i

Close #X

End Sub
