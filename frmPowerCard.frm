VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPowerCard 
   Caption         =   "Power Card Images"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Only No Images"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ListBox lstCharacters 
      Height          =   3180
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
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
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPowerCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim dbrec As Recordset

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
Set dbrec = db.OpenRecordset("SELECT * FROM power WHERE Power.ID=" & Trim(Str(lstCharacters.ItemData(lstCharacters.ListIndex))) & ";", dbOpenDynaset)

dbrec.MoveFirst

strImagePath = Label1.Caption

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

loadimages


End Sub
Private Sub loadimages()
Set db = OpenDatabase(App.Path & "\Overpower.mdb")

lstCharacters.Clear

If Check1.Value = 1 Then
    Set dbrec = db.OpenRecordset("SELECT * FROM POWER WHERE (Power.Image) IS NULL Order by Power.ID;", dbOpenDynaset)
Else
    Set dbrec = db.OpenRecordset("SELECT * FROM Power Order by Power.ID;", dbOpenDynaset)
End If


dbrec.MoveLast
dbrec.MoveFirst


For i = 1 To dbrec.RecordCount

If dbrec.Fields("E").Value = True Then
    lstCharacters.AddItem "ENERGY: " & dbrec.Fields("Value").Value
End If

If dbrec.Fields("F").Value = True Then
    lstCharacters.AddItem "FIGHTING: " & dbrec.Fields("Value").Value
End If

If dbrec.Fields("I").Value = True Then
    lstCharacters.AddItem "INTELLECT: " & dbrec.Fields("Value").Value
End If

If dbrec.Fields("S").Value = True Then
    lstCharacters.AddItem "STRENGTH: " & dbrec.Fields("Value").Value
End If

If dbrec.Fields("A").Value = True Then
    lstCharacters.AddItem "ANYPOWER: " & dbrec.Fields("Value").Value
End If

If dbrec.Fields("M").Value = True Then
    lstCharacters.AddItem "MULTI-POWER: " & dbrec.Fields("Value").Value
End If

lstCharacters.ItemData(lstCharacters.NewIndex) = dbrec.Fields("ID").Value


dbrec.MoveNext
Next i

End Sub

Private Sub lstCharacters_Click()
Dim b As Boolean

x = FreeFile
Open App.Path & "\sql.txt" For Output As #x
Print #x, "SELECT * FROM Power WHERE Power.id=" & Trim(Str(lstCharacters.ItemData(lstCharacters.ListIndex))) & ";"
Close #x

x = Shell(App.Path & "\openimage.exe", vbHide)

Image1.Picture = LoadPicture(App.Path & "\temppic.jpg")


End Sub
