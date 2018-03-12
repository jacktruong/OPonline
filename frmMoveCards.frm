VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMoveCards 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move Cards"
   ClientHeight    =   3210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNumberCards 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "1"
      Top             =   1760
      Width           =   495
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   2800
      TabIndex        =   6
      Top             =   1720
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      Enabled         =   -1  'True
   End
   Begin VB.OptionButton optDropType 
      Caption         =   "Add and then shuffle Draw Pile"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.OptionButton optDropType 
      Caption         =   "Add to BOTTOM of Draw Pile"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.OptionButton optDropType 
      Caption         =   "Add to TOP of Draw Pile"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblCount 
      Caption         =   "There are currently 3 cards in the Discard Pile."
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Number of cards to move:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblMoveType 
      Caption         =   "Move cards from Discard Pile to Draw Pile:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmMoveCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSource As String
Dim sTarget As String
Public Property Let Source(ByVal vNewValue As Variant)

sSource = vNewValue
End Property
Public Property Let Target(ByVal vNewValue As Variant)
sTarget = vNewValue

End Property

Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub Form_Load()
lblMoveType.Caption = "Move cards from " & sSource & " Pile to " & sTarget & " Pile."

optDropType(0).Caption = "Add to &TOP of " & sTarget & " Pile"
optDropType(1).Caption = "Add to &BOTTOM of " & sTarget & " Pile"
optDropType(2).Caption = "&SHUFFLE into " & sTarget & " Pile"

Select Case sSource
Case "Discard"
    lblCount.Caption = "There are currently " & cDiscardPile.Count & " cards in the Discard Pile."
Case "Dead"
    lblCount.Caption = "There are currently " & cDeadPile.Count & " cards in the Dead Pile."
Case "Draw"
    lblCount.Caption = "There are currently " & cDrawPile.Count & " cards in the Draw Pile."
End Select

End Sub
Private Sub OKButton_Click()
Dim ctempSource As Collection
Dim ctempTarget As Collection
Dim ctemp2 As Collection

Select Case sSource
Case "Draw"
    Set ctempSource = cDrawPile
Case "Discard"
    Set ctempSource = cDiscardPile
Case "Dead"
    Set ctempSource = cDeadPile
End Select

Select Case sTarget
Case "Draw"
    Set ctempTarget = cDrawPile
Case "Discard"
    Set ctempTarget = cDiscardPile
Case "Dead"
    Set ctempTarget = cDeadPile
End Select


If optDropType(0).Value = True Then
'Add to top of pile
Set ctemp2 = New Collection

    For i = 1 To Val(txtNumberCards.Text)
        ctemp2.Add ctempSource.Item(ctempSource.Count)
        ctempSource.Remove ctempSource.Count
    Next i
    
    For i = 1 To ctempTarget.Count
        ctemp2.Add ctempTarget.Item(i)
    Next i
    
    For i = 1 To ctempTarget.Count
        ctempTarget.Remove ctempTarget.Count
    Next i
    
    For i = 1 To ctemp2.Count
        ctempTarget.Add ctemp2.Item(i)
    Next i
    
End If


If optDropType(1).Value = True Then
'Add to bottom of pile
    For i = 1 To Val(txtNumberCards.Text)
        ctempTarget.Add ctempSource.Item(ctempSource.Count)
        ctempSource.Remove ctempSource.Count
    Next i


End If


If optDropType(2).Value = True Then
'Add to top of draw pile then shuffle
    For i = 1 To Val(txtNumberCards.Text)
        ctempTarget.Add ctempSource.Item(ctempSource.Count)
        ctempSource.Remove ctempSource.Count
    Next i
    
    ShuffleDrawPile

End If

Unload Me

End Sub

Private Sub UpDown1_Change()

txtNumberCards.Text = UpDown1.Value

End Sub

Private Sub UpDown1_DownClick()
If UpDown1.Value < 1 Then UpDown1.Value = 1

End Sub

Private Sub UpDown1_UpClick()
Select Case sSource
Case "Draw"
    If UpDown1.Value > cDrawPile.Count Then UpDown1.Value = cDrawPile.Count
Case "Discard"
    If UpDown1.Value > cDiscardPile.Count Then UpDown1.Value = cDiscardPile.Count
Case "Dead"
    If UpDown1.Value > cDeadPile.Count Then UpDown1.Value = cDeadPile.Count
End Select
End Sub
