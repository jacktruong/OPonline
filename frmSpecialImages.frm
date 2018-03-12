VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSpecialImages 
   Caption         =   "Special Images"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDeleteImage 
      Caption         =   "Delete Original Image"
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.ListBox lstSpecials 
      Height          =   4545
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.ComboBox lstCharacters 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Overpower Image"
      Filter          =   "*.*"
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5205
      Left            =   3840
      OLEDragMode     =   1  'Automatic
      Picture         =   "frmSpecialImages.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   4440
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmSpecialImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim dbrec As Recordset

Private Sub cmdLoadImage_Click()

With CommonDialog1
.Action = 1

If .FileName <> "" Then
    Me.Image1.Picture = LoadPicture(.FileName)
    Label1.Caption = .FileName
End If


End With



End Sub

Private Sub cmdSave_Click()
Dim strMsg As String
Dim bytBLOB() As Byte
Dim strImageTitle As String
Dim strImagePath As String
Dim intNum As Integer

x = MsgBox("Would you like to save this image to " & lstSpecials.List(lstSpecials.ListIndex) & "?", vbYesNoCancel, "Save Image?")

If x <> 6 Then Exit Sub

strImagePath = Label1.Caption

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbrec = db.OpenRecordset("SELECT * FROM Specials WHERE specials.ID=" & Trim(Str(lstSpecials.ItemData(lstSpecials.ListIndex))) & ";", dbOpenDynaset)

dbrec.MoveFirst

intNum = FreeFile
Open strImagePath For Binary As #intNum
ReDim bytBLOB(FileLen(strImagePath))

'Read the data and close the file
Get #intNum, , bytBLOB
Close #intNum

dbrec.Edit
dbrec.Fields("Image").AppendChunk bytBLOB
dbrec.Update

dbrec.Close
db.Close

If Me.chkDeleteImage.Value = 1 Then Kill Label1.Caption


End Sub

Private Sub Form_Load()

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbrec = db.OpenRecordset("SELECT Specials.Character FROM Specials GROUP BY Specials.Character;", dbOpenDynaset)

dbrec.MoveLast
dbrec.MoveFirst


For i = 1 To dbrec.RecordCount

lstCharacters.AddItem dbrec.Fields("Character").Value

dbrec.MoveNext
Next i

lstCharacters.ListIndex = 0

End Sub

Private Sub lstCharacters_Click()
If lstCharacters.ListIndex = -1 Then Exit Sub

a$ = lstCharacters.List(lstCharacters.ListIndex)
lstSpecials.Clear

strQry = "SELECT Specials.ID, Specials.Character, Specials.Description From Specials WHERE (((Specials.Character)=" & Chr(34) & a$ & Chr(34) & "));"


Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbrec = db.OpenRecordset(strQry, dbOpenDynaset)

dbrec.MoveLast
dbrec.MoveFirst


For i = 1 To dbrec.RecordCount

lstSpecials.AddItem dbrec.Fields("Description").Value
lstSpecials.ItemData(lstSpecials.NewIndex) = dbrec.Fields("ID").Value

dbrec.MoveNext
Next i

cmdSave.Enabled = False

End Sub

Private Sub lstSpecials_Click()
If lstSpecials.ListIndex = -1 Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If

Dim b As Boolean

x = FreeFile
Open App.Path & "\sql.txt" For Output As #x
Print #x, "SELECT * FROM Specials WHERE Specials.id=" & Trim(Str(lstSpecials.ItemData(lstSpecials.ListIndex))) & ";"
Close #x

x = Shell(App.Path & "\openimage.exe", vbHide)

Image1.Picture = LoadPicture(App.Path & "\temppic.jpg")

End Sub
