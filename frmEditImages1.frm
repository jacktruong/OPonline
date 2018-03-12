VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditImages1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Character/Special Card Images"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chk3Grid 
      Caption         =   "3 Grid"
      Height          =   195
      Left            =   1680
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4845
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox lstPics 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   6615
   End
   Begin VB.OptionButton optSpecial 
      Caption         =   "Special"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton optCharacter 
      Caption         =   "Character"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox lstSpecials 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.ComboBox lstCharacters 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Images:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Set Picture For:"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Special:"
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
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.Image imgCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   3840
      Picture         =   "frmEditImages1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Character:"
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
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuNextSpecial 
         Caption         =   "Next Special"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFindUnattached 
         Caption         =   "Find Next Unattached Image"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmEditImages1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curPath As String

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdSave_Click()
If lstPics.Text = "" Or lstPics.Text = "[None]" Then
    MsgBox "Please select an image.", vbInformation, "Image Needed"
    Exit Sub
End If

If Me.optCharacter.Value = True Then
    x = OpenRecordSet("SELECT * FROM Characters WHERE (((Characters.ID)=" + Trim$(Str$(lstCharacters.ItemData(lstCharacters.ListIndex))) + "));")
Else
    x = OpenRecordSet("SELECT * FROM Specials WHERE (((Specials.ID)=" + Trim$(Str$(lstSpecials.ItemData(lstSpecials.ListIndex))) + "));")
End If

If x <> 0 Then

If optCharacter.Value = True And chk3Grid.Value <> 0 Then

dbRec.Edit
dbRec.Fields("3ImagePath").Value = lstPics.Text

If IsNull(dbRec.Fields("3E").Value) = True Then
dbRec.Fields("3E").Value = InputBox$("Enter 3 Grid Energy Rating:", "Energy", "0")
dbRec.Fields("3F").Value = InputBox$("Enter 3 Grid Fighting Rating:", "Fighting", "0")
dbRec.Fields("3S").Value = InputBox$("Enter 3 Grid Strength Rating:", "Strength", "0")

End If

dbRec.Update


Else

dbRec.Edit
dbRec.Fields("ImagePath").Value = lstPics.Text
dbRec.Update

End If

End If

CloseRecordSet

    
If optSpecial.Value = True Then
    If lstSpecials.ListIndex < (lstSpecials.ListCount - 1) And lstPics.ListCount > 0 Then
        lstSpecials.ListIndex = lstSpecials.ListIndex + 1
        lstPics.SetFocus
    End If
End If


'If optCharacter.Value = True Then
'    If lstSpecials.ListCount > 0 Then lstSpecials.ListIndex = 0
'    optSpecial.Value = True
'End If

End Sub

Private Sub Form_Load()
x = OpenRecordSet("Characters")
FillFromRecordSet Me.lstCharacters, "Character", "ID"
CloseRecordSet

End Sub

Private Sub lstCharacters_Click()

x = OpenRecordSet("SELECT * FROM Specials WHERE (((Specials.Character)=" + Chr(34) + lstCharacters.List(lstCharacters.ListIndex) + Chr(34) + "));")
FillFromRecordSet Me.lstSpecials, "Description", "ID"
CloseRecordSet

optCharacter.Value = True
CheckForImage
If lstPics.ListCount > 0 Then lstPics.SetFocus
chk3Grid.Value = 0

End Sub

Private Sub CheckForImage()

If Me.optCharacter.Value = True Then
    x = OpenRecordSet("SELECT * FROM Characters WHERE (((Characters.ID)=" + Trim$(Str$(lstCharacters.ItemData(lstCharacters.ListIndex))) + "));")
Else
    x = OpenRecordSet("SELECT * FROM Specials WHERE (((Specials.ID)=" + Trim$(Str$(lstSpecials.ItemData(lstSpecials.ListIndex))) + "));")
End If

If x > 0 Then

    If IsNull(dbRec.Fields("ImagePath").Value) = False Then
    
        curPath = dbRec.Fields("ImagePath").Value
        
        x = Dir(curPath, vbNormal)
        
            If x = "" Then
                curPath = pBlankPic
            End If
    
    Else
    
        curPath = pBlankPic
    
    End If
    
    imgCard.Picture = LoadPicture(curPath)
        
Else
    
    imgCard.Picture = LoadPicture(pBlankPic)

End If

If curPath = pBlankPic Then
    lstPics.Text = "[None]"
Else
    lstPics.Text = curPath
End If

GetAllPicsFromDir

End Sub
Private Sub GetAllPicsFromDir()
a$ = lstPics.Text
lstPics.Clear

If a$ = "[None]" Then a$ = pPath + "\Heroes\" + lstCharacters.List(lstCharacters.ListIndex) + "\temp"


lastSlash = 0

Looper:
x = InStr(x + 1, a$, "\")

If x <> 0 Then
    lastSlash = x
    GoTo Looper
End If

If lastSlash <> 0 Then

b$ = Left$(a$, lastSlash)

x = Dir(b$ + "*.*", vbNormal)

looper2:
If x <> "" Then
    lstPics.AddItem b$ + x
    x = Dir()
    GoTo looper2
End If


End If

lstPics.Text = a$

If Right$(a$, 5) = "\temp" Then lstPics.Text = "[None]"

If lstPics.ListCount > 0 Then
    StatusBar1.Panels(1).Text = Trim(Str(lstPics.ListCount)) + " images available."
Else
    StatusBar1.Panels(1).Text = "No images available"
End If

End Sub
Private Sub lstPics_Click()
imgCard.Picture = LoadPicture(lstPics.List(lstPics.ListIndex))

End Sub

Private Sub lstSpecials_Click()
optSpecial.Value = True
CheckForImage
lstPics.SetFocus

End Sub

Private Sub mnuExit_Click()
Unload Me

End Sub

Private Sub mnuFindUnattached_Click()

x = OpenRecordSet("SELECT * FROM Characters WHERE (((Characters.ImagePath) Is Null));")

If x > 0 Then

dbRec.MoveLast
dbRec.MoveFirst

For i = 1 To dbRec.RecordCount

a$ = dbRec.Fields("Character").Value
p$ = pPath + "\Heroes\" + a$ + "\*.*"

On Error Resume Next
x = Dir(p$, vbNormal)

If x <> "" Then

    For k = 0 To lstCharacters.ListCount - 1
        If lstCharacters.List(k) = a$ Then
            lstCharacters.ListIndex = k
            GoTo foundone
        End If
    Next k
    
End If

dbRec.MoveNext

Next i
x = MsgBox("All available images have been associated with characters/specials!", vbInformation, "No Unassociated Images")

Else
x = MsgBox("All Characters have associated images!", vbInformation, "No Unassociated Images")

End If

foundone:
CloseRecordSet

End Sub

Private Sub mnuNextSpecial_Click()
If lstSpecials.ListIndex < lstSpecials.ListCount - 1 Then
    lstSpecials.ListIndex = lstSpecials.ListIndex + 1
End If

End Sub

Private Sub optCharacter_Click()
CheckForImage

End Sub
