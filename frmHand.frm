VERSION 5.00
Begin VB.Form frmHand 
   Caption         =   "Add Character Images"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCharacters 
      Height          =   3180
      Left            =   120
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
Attribute VB_Name = "frmHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim dbRec As Recordset

Private Sub Command2_Click()
Dim strMsg As String
Dim bytBLOB() As Byte
Dim strImageTitle As String
Dim strImagePath As String
Dim intNum As Integer

'Add da0 3.51 and then run

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbRec = db.OpenRecordset("Characters", dbOpenDynaset)

dbRec.MoveLast
dbRec.MoveFirst

For i = 1 To dbRec.RecordCount

Me.Caption = "Record: " & i

If IsNull(dbRec.Fields("ImagePath").Value) = True Then
    strImagePath = ""
Else
    strImagePath = dbRec.Fields("ImagePath").Value
End If

If strImagePath <> "" Then

Me.Caption = "Processing image..."

intNum = FreeFile
Open strImagePath For Binary As #intNum
ReDim bytBLOB(FileLen(strImagePath))

'Read the data and close the file
Get #intNum, , bytBLOB
Close #intNum

dbRec.Edit
dbRec.Fields("Image").AppendChunk bytBLOB
dbRec.Update

End If


dbRec.MoveNext
Next i

dbRec.Close
db.Close
End

End Sub
Private Sub Form_Load()

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbRec = db.OpenRecordset("SELECT * FROM Characters WHERE (((Characters.Image) Is Null));", dbOpenDynaset)

dbRec.MoveLast
dbRec.MoveFirst


For i = 1 To dbRec.RecordCount

lstCharacters.AddItem dbRec.Fields("Character").Value


dbRec.MoveNext
Next i

End Sub
