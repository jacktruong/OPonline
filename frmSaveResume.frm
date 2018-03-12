VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSaveResume 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saving Resume Information"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Saving information necessary for resuming game..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmSaveResume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
pBar1.Min = 0
pBar1.Max = 40
pBar1.Value = 0
Me.Caption = "Saving Resume Information...Turn " & Trim(Str(nTurn - 1))
Me.Visible = True
Me.Refresh

MsgBox "about to save"
Save

End Sub
Private Sub Save()

MsgBox "saving"
SaveResumeInfo

MsgBox "saving2"
SaveOpponentResumeInfo

pBar1.Value = 40

MsgBox "unloading"

Unload Me


End Sub
