VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrainingImages 
   Caption         =   "Training Images"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ListBox lstCharacters 
      Height          =   3180
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Overpower Image"
      Filter          =   "*.*"
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmTrainingImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim dbRec As Recordset

Private Sub cmdLoadImage_Click()


If lstCharacters.ListIndex = -1 Then Exit Sub

With CommonDialog1
.Action = 1

If .FileName <> "" Then
    Me.Image1.Picture = LoadPicture(.FileName)
    Label1.Caption = .FileName
End If

End With



End Sub

Private Sub cmdSaveImage_Click()
Dim strMsg As String
Dim bytBLOB() As Byte
Dim strImageTitle As String
Dim strImagePath As String
Dim intNum As Integer

x = MsgBox("Would you like to save this image to " & lstCharacters.List(lstCharacters.ListIndex) & "?", vbYesNoCancel, "Save Image?")

If x <> 6 Then Exit Sub

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbRec = db.OpenRecordset("SELECT * FROM training WHERE training.ID=" & Trim(Str(lstCharacters.ItemData(lstCharacters.ListIndex))) & ";", dbOpenDynaset)

dbRec.MoveFirst

strImagePath = Label1.Caption

intNum = FreeFile
Open strImagePath For Binary As #intNum
ReDim bytBLOB(FileLen(strImagePath))

'Read the data and close the file
Get #intNum, , bytBLOB
Close #intNum

dbRec.Edit
dbRec.Fields("Image").AppendChunk bytBLOB
dbRec.Update

dbRec.Close
db.Close
End Sub

Private Sub Form_Load()

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbRec = db.OpenRecordset("SELECT * FROM training Order by training.ID;", dbOpenDynaset)

dbRec.MoveLast
dbRec.MoveFirst


For i = 1 To dbRec.RecordCount

nBonus = dbRec.Fields("Bonus").Value
sType1 = dbRec.Fields("PWR1").Value
sType2 = dbRec.Fields("PWR2").Value
sPictured = dbRec.Fields("Pictured").Value
Title = "TRAINING: " & sType1 & sType2 & " + " & Trim(Str(nBonus)) & " [" & sPictured & "]"
lstCharacters.AddItem Title
lstCharacters.ItemData(lstCharacters.NewIndex) = dbRec.Fields("ID").Value


dbRec.MoveNext
Next i

End Sub


Private Sub lstCharacters_Click()
Dim b As Boolean

x = FreeFile
Open App.Path & "\sql.txt" For Output As #x
Print #x, "SELECT * FROM training WHERE training.id=" & Trim(Str(lstCharacters.ItemData(lstCharacters.ListIndex))) & ";"
Close #x

x = Shell(App.Path & "\openimage.exe", vbHide)

Image1.Picture = LoadPicture(App.Path & "\temppic.jpg")


End Sub




