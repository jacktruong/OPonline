Attribute VB_Name = "opsedit"
Global dbName
Public Function LoadImageFromDatabase(strSQL) As Boolean
Dim db As ADODB.Connection
Dim dbRec As ADODB.Recordset

On Error GoTo LoadError

Set db = New ADODB.Connection
db.ConnectionString = dbName
      
db.Open

Set dbRec = New ADODB.Recordset

dbRec.Open strSQL, db

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
    LoadImageFromDatabase = False
    Exit Function
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

LoadImageFromDatabase = True

dbRec.Close
db.Close

Exit Function

LoadError:
LoadImageFromDatabase = False

End Function
