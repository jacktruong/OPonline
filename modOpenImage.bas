Attribute VB_Name = "modOpenImage"
Sub main()
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

dbName = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\Overpower.mdb"

x = FreeFile
Open App.Path & "\sql.txt" For Input As #x
Line Input #x, strSQL
Close #x

On Error GoTo LoadError

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.open

Set dbRec = New ADODB.Recordset

dbRec.open strSQL, db

Dim lngImageSize As Long
Dim lngOffset As Long
Dim bytChunk() As Byte
Dim intFile As Integer
Dim strTempPic As String
Const conChunkSize = 100

'Make sure the temporary file does not already exist
strTempPic = App.Path & "\TempPic.jpg"
If Len(Dir(strTempPic)) > 0 Then
    Kill strTempPic
End If

If IsNull(dbRec.Fields("Image").Value) = True Then
    dbRec.Close
    db.Close
    FileCopy App.Path & "\notfound.jpg", App.Path & "\temppic.jpg"
    End
End If

'Open the temporary file to save the BLOB to
intFile = FreeFile
Open strTempPic For Binary As #intFile

'Read the binary data into the byte variable array

lngImageSize = dbRec("Image").ActualSize
Do While lngOffset < lngImageSize
   bytChunk() = dbRec("Image").GetChunk(conChunkSize)
   Put #intFile, , bytChunk()
   lngOffset = lngOffset + conChunkSize
Loop

Close #intFile


dbRec.Close
db.Close
End


LoadError:
End

End Sub
