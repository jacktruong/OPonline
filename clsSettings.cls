VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim nPort
Dim sIpAddress
Dim bCheckIP As Boolean
Dim bShowAttackMessage As Boolean
Dim bpopup As Boolean
Dim bBeep As Boolean

Dim sPlayerName As String
Public Sub Load()

sPlayerName = "[NONE]"
nPort = 1564
sIpAddress = ""
bShowAttackMessage = True
bpopup = False
bBeep = False

X = Dir(App.Path & "\settings.ini", vbNormal)

If X = "" Then

X = FreeFile
Open App.Path & "\settings.ini" For Output As #X
Print #X, "IP=" & frmTable.tcpChannel.LocalIP
Print #X, "PORT=1544"
Print #X, , "PlayerName=UNKNOWN"
Print #X, "CheckIP=FALSE"
Print #X, "ShowAttackMessage=FALSE"
Print #X, "Popup=TRUE"
Print #X, "Beep=FALSE"
Close #X

End If

X = FreeFile

Open App.Path & "\settings.ini" For Input As #X

While Not EOF(X)

Line Input #X, a$

Select Case GetItem(a$)

Case "Beep"
    bBeep = CBool(GetValue(a$))
Case "Popup"
    bpopup = CBool(GetValue(a$))
Case "IP"
    sIpAddress = GetValue(a$)
Case "PORT"
    nPort = Val(GetValue(a$))
Case "PlayerName"
    sPlayerName = GetValue(a$)
Case "CheckIP"
    bCheckIP = CBool(GetValue(a$))
Case "ShowAttackMessage"
    bShowAttackMessage = CBool(GetVal(a$))
    
End Select

Wend

Close #X

End Sub
Private Sub SaveSettings()

X = FreeFile
Open App.Path & "\settings.ini" For Output As #X
Print #X, "PlayerName=" & PlayerName
Print #X, "PORT=" & Trim(Str(nPort))
Print #X, "IP=" & Trim(sIpAddress)

If bpopup = True Then
    Print #X, "Popup=TRUE"
Else
    Print #X, "Popup=FALSE"
End If

If bBeep = True Then
    Print #X, "Beep=TRUE"
Else
    Print #X, "Beep=FALSE"
End If

If bShowAttackMessage = True Then
    Print #X, "ShowAttackMessage=TRUE"
Else
    Print #X, "ShowAttackMessage=FALSE"
End If

If bCheckIP = True Then
    Print #X, "CheckIP=TRUE"
Else
    Print #X, "CheckIP=FALSE"
End If

Close #X

End Sub
Private Function GetItem(sString) As Variant

z = InStr(sString, "=")
GetItem = Left(sString, z - 1)

End Function

Private Function GetValue(sString) As Variant

z = InStr(sString, "=")
GetValue = Right(sString, Len(sString) - z)

End Function
Public Property Get PlayerName() As Variant
PlayerName = sPlayerName

End Property
Public Property Let PlayerName(vnewValue As Variant)
sPlayerName = vnewValue
SaveSettings

End Property
Public Property Get IP_Address() As Variant
IP_Address = sIpAddress

End Property
Public Property Let IP_Address(vnewValue As Variant)

sIpAddress = vnewValue
SaveSettings
End Property
Public Property Get PopupMessages() As Boolean

PopupMessages = bpopup

End Property
Public Property Let PopupMessages(vnewValue As Boolean)

bpopup = vnewValue
SaveSettings

End Property
Public Property Get Port() As Variant
Port = nPort

End Property
Public Property Let Port(vnewValue As Variant)

nPort = vnewValue
SaveSettings
End Property

Public Property Get AutoCheckIP() As Boolean
AutoCheckIP = bCheckIP

End Property
Public Property Let AutoCheckIP(vnewValue As Boolean)

bCheckIP = vnewValue
SaveSettings
End Property

Public Property Get ShowAttackMessage() As Boolean
ShowAttackMessage = bShowAttackMessage

End Property
Public Property Let ShowAttackMessage(vnewValue As Boolean)

bShowAttackMessage = vnewValue
SaveSettings

End Property

Public Property Get MessageBeep() As Boolean

MessageBeep = bBeep

End Property
Public Property Let MessageBeep(vnewValue As Boolean)

bBeep = vnewValue
SaveSettings

End Property

