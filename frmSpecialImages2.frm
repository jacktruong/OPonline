VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSpecialImages2 
   Caption         =   "Specials Without Images"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ListBox lstSpecials 
      Height          =   2985
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.CheckBox chkDeleteImage 
      Caption         =   "Delete Original Image"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
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
      Height          =   3255
      Left            =   4920
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   5520
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmSpecialImages2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim dbRec As Recordset

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
Set dbRec = db.OpenRecordset("SELECT * FROM Specials WHERE specials.ID=" & Trim(Str(lstSpecials.ItemData(lstSpecials.ListIndex))) & ";", dbOpenDynaset)

dbRec.MoveFirst

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

If Me.chkDeleteImage.Value = 1 Then Kill Label1.Caption
lstSpecials.RemoveItem lstSpecials.ListIndex

'lstSpecials.ListIndex = 0

Me.Caption = "Specials without Images (" & Trim(Str(lstSpecials.ListCount)) & ")"

End Sub

Private Sub Form_Load()

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbRec = db.OpenRecordset("SELECT * FROM Specials WHERE (Specials.Image) IS NULL;", dbOpenDynaset)

dbRec.MoveLast
dbRec.MoveFirst


For i = 1 To dbRec.RecordCount

lstSpecials.AddItem dbRec.Fields("Character").Value & "-->" & dbRec.Fields("Description").Value
lstSpecials.ItemData(lstSpecials.NewIndex) = dbRec.Fields("ID").Value


dbRec.MoveNext
Next i

lstSpecials.ListIndex = 0

End Sub


Private Sub lstSpecials_Click()
If lstSpecials.ListIndex = -1 Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If

'Dim b As Boolean
'
'x = FreeFile
'Open App.Path & "\sql.txt" For Output As #x
'Print #x, "SELECT * FROM Specials WHERE Specials.id=" & Trim(Str(lstSpecials.ItemData(lstSpecials.ListIndex))) & ";"
'Close #x
'
'x = Shell(App.Path & "\openimage.exe", vbHide)

Image1.Picture = LoadPicture(App.Path & "\temppic.jpg")

End Sub

