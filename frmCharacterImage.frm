VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCharacterImage 
   Caption         =   "Add Character Images"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowAll 
      Caption         =   "Show All"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
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
   Begin VB.ListBox lstCharacters 
      Height          =   3180
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmCharacterImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim dbrec As Recordset

Private Sub chkShowAll_Click()

If chkShowAll.Value = 1 Then

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbrec = db.OpenRecordset("SELECT * FROM Characters;", dbOpenDynaset)

lstCharacters.Clear

dbrec.MoveLast
dbrec.MoveFirst


For i = 1 To dbrec.RecordCount

lstCharacters.AddItem dbrec.Fields("Character").Value
lstCharacters.ItemData(lstCharacters.NewIndex) = dbrec.Fields("ID").Value


dbrec.MoveNext
Next i

End If

End Sub

Private Sub cmdLoadImage_Click()
Dim strMsg As String
Dim bytBLOB() As Byte
Dim strImageTitle As String
Dim strImagePath As String
Dim intNum As Integer

If lstCharacters.ListIndex = -1 Then Exit Sub

With CommonDialog1
.Action = 1

If .FileName <> "" Then
    Me.Image1.Picture = LoadPicture(.FileName)
    strImagePath = .FileName
End If

x = MsgBox("Would you like to save this image to " & lstCharacters.List(lstCharacters.ListIndex) & "?", vbYesNoCancel, "Save Image?")

If x <> 6 Then Exit Sub

End With

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbrec = db.OpenRecordset("SELECT * FROM Characters WHERE Characters.ID=" & Trim(Str(lstCharacters.ItemData(lstCharacters.ListIndex))) & ";", dbOpenDynaset)

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

End Sub

Private Sub Form_Load()

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbrec = db.OpenRecordset("SELECT * FROM Characters;", dbOpenDynaset)

dbrec.MoveLast
dbrec.MoveFirst


For i = 1 To dbrec.RecordCount

lstCharacters.AddItem dbrec.Fields("Character").Value
lstCharacters.ItemData(lstCharacters.NewIndex) = dbrec.Fields("ID").Value


dbrec.MoveNext
Next i

End Sub

Private Sub lstCharacters_Click()
Dim b As Boolean

x = FreeFile
Open App.Path & "\sql.txt" For Output As #x
Print #x, "SELECT * FROM Characters WHERE Characters.id=" & Trim(Str(lstCharacters.ItemData(lstCharacters.ListIndex))) & ";"
Close #x

x = Shell(App.Path & "\openimage.exe", vbHide)

Image1.Picture = LoadPicture(App.Path & "\temppic.jpg")
End Sub
