VERSION 5.00
Begin VB.Form frmCheckSpecialImages 
   Caption         =   "Check Special Images"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Fix"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.Image imgCardDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5205
      Left            =   5400
      OLEDragMode     =   1  'Automatic
      Picture         =   "frmCheckSpecialImages.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3720
   End
End
Attribute VB_Name = "frmCheckSpecialImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.RemoveItem List1.ListIndex
List1.ListIndex = 0

End Sub

Private Sub Command2_Click()
x = FreeFile
Open "c:\overpower\fixspec.txt" For Append As #x
Print #x, List1.List(List1.ListIndex)
List1.RemoveItem List1.ListIndex
Close #x
List1.ListIndex = 0

End Sub

Private Sub Form_Load()
'x = FreeFile
'Open "c:\overpower\specimage.txt" For Input As #x
'
'Do Until EOF(x)
'Line Input #x, a$
'List1.AddItem a$
'Loop
'Close #x
'
'x = FreeFile
'Open "c:\overpower\specnum.txt" For Input As #x
'
'For i = 0 To List1.ListCount - 1
'Line Input #x, a$
'List1.ItemData(i) = Val(a$)
'Next i
'
'Close #x


Set db = OpenDatabase(App.Path & "\Overpower.mdb")
Set dbRec = db.OpenRecordset("SELECT * FROM Specials;", dbOpenDynaset)

dbRec.MoveLast
dbRec.MoveFirst


For i = 1 To dbRec.RecordCount

List1.AddItem dbRec.Fields("Character").Value & "-->" & dbRec.Fields("Description").Value
List1.ItemData(List1.NewIndex) = dbRec.Fields("ID").Value


dbRec.MoveNext
Next i

List1.ListIndex = 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
x = FreeFile
Open "c:\overpower\specimage.txt" For Output As #x
For i = 0 To List1.ListCount - 1
Print #x, List1.List(i)
Next i
Close #x


x = FreeFile
Open "c:\overpower\specnum.txt" For Output As #x
For i = 0 To List1.ListCount - 1
Print #x, List1.ItemData(i)
Next i
Close #x

End
End Sub

Private Sub List1_Click()

x = FreeFile
Open App.Path & "\sql.txt" For Output As #x
Print #x, "SELECT * FROM Specials WHERE Specials.id=" & Trim(Str(List1.ItemData(List1.ListIndex))) & ";"
Close #x

x = Shell(App.Path & "\openimage.exe", vbHide)

imgCardDetail.Picture = LoadPicture(App.Path & "\temppic.jpg")

End Sub
