VERSION 5.00
Begin VB.Form frmMessaging 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overpower Online - Messages"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   ControlBox      =   0   'False
   Icon            =   "frmMessaging.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   10815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBackToGame 
      Caption         =   "Back to &Game"
      Height          =   375
      Left            =   9480
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtNewMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   7815
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send &Message"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtMessages 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   10455
   End
End
Attribute VB_Name = "frmMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Declare Function PutFocus Lib "user32" _
   Alias "SetFocus" _
  (ByVal hwnd As Long) As Long

Private Const EM_LINESCROLL = &HB6
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Dim sMessage As Variant

Function ScrollText(TextBox As Control, vLines As Integer) As Long
    
    Dim Success As Long
    Dim SavedWnd As Long
    Dim moveLines As Long
        
   'save the window handle of the control that currently has focus
    SavedWnd = Screen.ActiveControl.hwnd
    moveLines = vLines
      
   'Set the focus to the passed control (text control)
    TextBox.SetFocus
  
   'Scroll the lines.
    Success = SendMessage(TextBox.hwnd, EM_LINESCROLL, 0, ByVal moveLines)
      
   'Restore the focus to the original control
    Call PutFocus(SavedWnd)
      
   'Return the number of lines actually scrolled
    ScrollText = Success
    
End Function

Function AddText(textcontrol As Object, text2add As String)
    On Error GoTo errhandlr
    tmptxt$ = textcontrol.Text 'just In Case of an accident
    textcontrol.SelStart = Len(textcontrol.Text) ' move the "cursor" To the End of the text file
    textcontrol.SelLength = 0 ' highlight nothing (this becomes the selected text)
    textcontrol.SelText = text2add ' Set the selected text ot text2add
    AddText = 1
    GoTo quitt ' goto the End of the Sub
    'error handlers
errhandlr:


    If Err.Number <> 438 Then 'check the Error number and restore the
        textcontrol.Text = tmptxt$ 'original text If the control supports it
    End If
    AddText = 0
    GoTo quitt
quitt:
    tmptxt$ = ""
End Function

Private Sub cmdBackToGame_Click()
frmTable.Show

End Sub

Private Sub cmdSend_Click()
If txtNewMessage.Text = "" Then Exit Sub

CURRLINE = SendMessage(txtMessages.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1

AddText txtMessages, mySettings.PlayerName & ": " & txtNewMessage.Text & vbCrLf & vbCrLf

SendData "M" & txtNewMessage.Text & "|"

CURRLINE2 = SendMessage(txtMessages.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1

ScrollText txtMessages, (CURRLINE2 - CURRLINE)

txtNewMessage.Text = ""
txtNewMessage.SetFocus

End Sub

Private Sub Form_Activate()
On Error Resume Next
txtNewMessage.SetFocus

End Sub

Private Sub txtNewMessage_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then cmdSend = True

End Sub
Public Property Let NewMessage(ByVal vNewValue As Variant)

sMessage = vNewValue
Me.Show
CURRLINE = SendMessage(txtMessages.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1

AddText txtMessages, sMessage & vbCrLf & vbCrLf

CURRLINE2 = SendMessage(txtMessages.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1

ScrollText txtMessages, (CURRLINE2 - CURRLINE)



End Property
