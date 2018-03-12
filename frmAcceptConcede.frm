VERSION 5.00
Begin VB.Form frmAcceptConcede 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opponent Concedes"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAccepted 
      Caption         =   "Accepted"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdDontAccept 
      Caption         =   "Don't Accept"
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept Concession"
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAcceptConcede.frx":0000
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "2.  You have a special that forces your opponent to continue the battle"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "1.  You have a special that allows additional attack(s) after your opponent has conceded"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Your opponent has conceded the battle.  In certain cases you do not have to let the battle end at this point:  "
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmAcceptConcede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccept_Click()
chkAccepted.Value = 1
Me.Hide

End Sub


Private Sub cmdDontAccept_Click()
chkAccepted.Value = 0
Me.Hide
End Sub
