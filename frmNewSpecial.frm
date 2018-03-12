VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewSpecial 
   Caption         =   "Add New Special"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "One Per Deck?"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ComboBox lstCharacters 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Overpower Image"
      Filter          =   "*.*"
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   6120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Select Character:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5205
      Left            =   4320
      OLEDragMode     =   1  'Automatic
      Picture         =   "frmNewSpecial.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3720
   End
End
Attribute VB_Name = "frmNewSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoadImage_Click()
With CommonDialog1
.Action = 1

If .FileName <> "" Then
    Me.Image1.Picture = LoadPicture(.FileName)
    Label4.Caption = .FileName
    
End If


End With
End Sub

Private Sub Command1_Click()
Dim db As Database
Dim dbrec As Recordset
Dim strMsg As String
Dim bytBLOB() As Byte
Dim strImageTitle As String
Dim strImagePath As String
Dim intNum As Integer

Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbrec = db.OpenRecordset("Specials", dbOpenDynaset)
dbrec.AddNew

strImagePath = Label4.Caption

intNum = FreeFile
Open strImagePath For Binary As #intNum
ReDim bytBLOB(FileLen(strImagePath))

'Read the data and close the file
Get #intNum, , bytBLOB
Close #intNum

dbrec.Fields("Image").AppendChunk bytBLOB
dbrec.Fields("CharID").Value = lstCharacters.ItemData(lstCharacters.ListIndex)
dbrec.Fields("Character").Value = lstCharacters.List(lstCharacters.ListIndex)
dbrec.Fields("Description").Value = Text3.Text
dbrec.Fields("Effect").Value = Text1.Text
dbrec.Fields("Custom").Value = True
dbrec.Fields("Code").Value = Text2.Text
dbrec.Fields("OPD").Value = CBool(Check1.Value)


dbrec.Update

dbrec.Close
db.Close

Command2 = True

End Sub

Private Sub Command2_Click()
lstCharacters.ListIndex = -1
Text1.Text = ""
Text2.Text = ""
Check1.Value = 0
Text3.Text = ""

End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\Overpower.mdb")
strQry = "SELECT Characters.Character, First(Characters.ID) AS ID From Characters GROUP BY Characters.Character;"
Set dbrec = db.OpenRecordset(strQry, dbOpenDynaset)

dbrec.MoveLast
dbrec.MoveFirst


For i = 1 To dbrec.RecordCount

lstCharacters.AddItem dbrec.Fields("Character").Value
lstCharacters.ItemData(lstCharacters.NewIndex) = dbrec.Fields("ID").Value

dbrec.MoveNext
Next i

lstCharacters.ListIndex = 0
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)

End Sub
