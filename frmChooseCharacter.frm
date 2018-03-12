VERSION 5.00
Begin VB.Form frmChooseCharacter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a Character"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4635
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optCharacter 
      Caption         =   "Homebase"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.OptionButton optCharacter 
      Caption         =   "Name"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.OptionButton optCharacter 
      Caption         =   "Name"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.OptionButton optCharacter 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.OptionButton optCharacter 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "frmChooseCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nSelectedChar As Integer
Dim bShowBS As Boolean
Dim sCharacter As String
Dim bHideHomebase As Boolean

Public Property Get SelectedCharacter() As Integer

SelectedCharacter = -1
For i = 0 To 4
    If optCharacter(i).Value = True Then
        SelectedCharacter = Val(optCharacter(i).Tag)
    End If
Next i

End Property

Private Sub cmdCancel_Click()
For i = 0 To 3
    optCharacter(i).Value = False
Next i

Me.Hide

End Sub

Private Sub cmdOK_Click()
Me.Hide

End Sub

Private Sub Form_Load()

cc = -1

For i = 1 To 4

If cFrontLine.isCharacterDead(i) = False Then
    cc = cc + 1
    optCharacter(cc).Caption = cFrontLine.Character_Name(i)
    optCharacter(cc).Tag = i
    optCharacter(cc).Visible = True
End If

Next i

optCharacter(0).Value = True

If bShowBS = True Then

If myBattleSite.ID > 0 Then
    optCharacter(4).Caption = "BATTLESITE: " & myBattleSite.Name
    optCharacter(4).Visible = True
    optCharacter(4).Tag = 5
Else
    optCharacter(4).Visible = False
End If


Else
On Error Resume Next

If bHideHomebase = False Then
    
    If myHomebase.ID > 0 Then
        optCharacter(4).Caption = "HOMEBASE: " & myHomebase.Name
        optCharacter(4).Visible = True
        optCharacter(4).Tag = 5
    Else
        optCharacter(4).Visible = False
    End If

Else
    optCharacter(4).Visible = False
    
End If

End If


For i = 0 To 3
    If optCharacter(i).Caption = sCharacter Then
        optCharacter(i).Value = True
        Exit For
    End If
Next i


End Sub
Public Property Let ShowBattlesite(ByVal vnewValue As Boolean)
bShowBS = vnewValue

End Property

Public Property Let SpecialCharacter(ByVal vnewValue)

sCharacter = vnewValue

End Property
Public Property Let HideHomebase(ByVal vnewValue As Boolean)
bHideHomebase = vnewValue

End Property
