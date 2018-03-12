VERSION 5.00
Begin VB.Form frmChallenge 
   Caption         =   "Defense Challenged!"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDispute 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   $"frmChallenge.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Your opponent has challenged your defense.  The following is his objection:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmChallenge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me

End Sub
