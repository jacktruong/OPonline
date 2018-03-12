VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Communications channel between players
Public Event ReceivedMessage(sMessage)
Public Event ReceivedCode(sCode)

Dim sMyChannelLocation
Dim sOpChannelLocation
Dim nRole As MyRole

Private MyChannelMessage As ChannelMessage
Private SendChannelMessage As ChannelMessage

Dim RecLength As Integer
Dim numrecs

Public Enum MyRole
   Guest = 0
   Host = 1
End Enum
    
Dim nMCL 'My Channel file code
Dim nOCL 'Opponent Channel file code
Public Property Get opponentChannelLocation()

    ChannelLocation = sOpChannelLocation

End Property
Public Property Let opponentChannelLocation(vnewValue)

    sOpChannelLocation = vnewValue

End Property
Public Property Get MyChannelLocation()

    ChannelLocation = sMyChannelLocation

End Property
Public Property Let MyChannelLocation(vnewValue)

    sMyChannelLocation = vnewValue

End Property
Public Sub OpenChannels()

RecLength = Len(MyChannelMessage)

nMCL = FreeFile
Open sMyChannelLocation For Output As #nMCL
Close #nMCL

Open sMyChannelLocation For Random As #nMCL Len = RecLength

nOCL = FreeFile
Open sOpChannelLocation For Output As #nOCL
Close #nOCL

Open sOpChannelLocation For Random As #nOCL Len = RecLength

numrecs = NumberOfRecords(nMCL)

End Sub
Public Function CheckChannel() As Boolean
'Returns false if there are no new messages on the channel

numrecs2 = NumberOfRecords(nMCL)

If numrecs <> numrecs2 Then

    CheckChannel = True
    
For i = 1 To (numrecs2 - numrecs)

    Get #nMCL, numrecs + 1, MyChannelMessage
        
    'figure out type of message
    If MyChannelMessage.Code = "M" Then
        'This is a message
        RaiseEvent ReceivedMessage(Trim(MyChannelMessage.Message))
    End If
    
    If MyChannelMessage.Code = "C" Then
        'receivedcode
        RaiseEvent ReceivedCode(Trim(MyChannelMessage.Message))
    End If
    
Next i
End If

numrecs = numrecs2

End Function
Public Sub ForceEvent_Code(sString)

RaiseEvent ReceivedCode(sString)


End Sub
Public Function SendMessage(sMessage As String)
sp$ = "                                                                                                                                                                                    "
sMessage = sMessage & Left(sp$, 300 - Len(sMessage))

SendChannelMessage.Code = "M"
SendChannelMessage.Message = sMessage

newrecord = NumberOfRecords(nOCL) + 1

Put #nOCL, newrecord, SendChannelMessage

End Function
Public Function SendCode(sCode As String)

SendChannelMessage.Code = "C"
SendChannelMessage.Message = sMessage

newrecord = NumberOfRecords(nOCL) + 1

Put #nOCL, newrecord, SendChannelMessage

End Function
Private Function NumberOfRecords(nFileNum)

If FileLen(sMyChannelLocation) = 0 Then NumberOfRecords = 0: Exit Function
If RecLength = 0 Then NumberOfRecords = 0: Exit Function

If nFileNum = nMCL Then
    NumberOfRecords = Int(FileLen(sMyChannelLocation) / RecLength)
Else
    NumberOfRecords = Int(FileLen(sOpChannelLocation) / RecLength)
End If

End Function
Public Function SendMyselfMessage(sMessage As String)

SendChannelMessage.Code = "M"
SendChannelMessage.Message = sMessage

newrecord = NumberOfRecords(nMCL) + 1

Put #nMCL, newrecord, SendChannelMessage

End Function
Public Function SendMyselfCode(sCode As String)
SendChannelMessage.Code = "C"
SendChannelMessage.Message = sCode
newrecord = NumberOfRecords(nMCL) + 1

Put #nMCL, newrecord, SendChannelMessage

End Function
Public Property Let Role(vnewValue As MyRole)

nRole = vnewValue

End Property
Public Property Get Role() As MyRole

Role = nRole

End Property
Public Function IsOpponentConnected() As Boolean

If NumberOfRecords(nMCL) = 0 Then
    IsOpponentConnected = False
Else
    IsOpponentConnected = True
End If

End Function
Public Function WaitForCode(sCode) As Variant
WaitForCode = "ERROR"

numrecs2 = NumberOfRecords(nMCL)
If numrecs2 = 0 Then
    WaitForCode = "ERROR"
    Exit Function
End If

Get #nMCL, numrecs2, MyChannelMessage

a$ = MyChannelMessage.Message
a$ = UCase$(Left$(a$, 2))

If a$ = sCode Then
        WaitForCode = Right(MyChannelMessage.Message, Len(MyChannelMessage.Message) - 3)
        numrecs = numrecs2
    Else
        WaitForCode = "ERROR"
End If

End Function
