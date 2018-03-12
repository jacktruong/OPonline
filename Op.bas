Attribute VB_Name = "modOp"
Global cbattlesitedeck As Collection
Global cIncomingAttack As Collection
Global cIncomingDefense As Collection

Global bHost As Boolean
Global isConnectedFlag As Boolean
Global bIGoFirst As Boolean
Global bResuming As Boolean

Global pPath As String
Global dbName As String
Global sBlankImagePath As String
Global mySettings As clsSettings

Global myattack As clsAttack
Global OpAttack As clsAttack

Global myDefense As clsDefense
Global OpDefense As clsDefense

'Various Deck Variables
Global cDrawPile As Collection 'This will contain all deck cards at the beginning
Global cDeadPile As Collection
Global cDiscardPile As Collection
Global cDefeatedCharactersPile As Collection

'Opponent Deck Variables
Global cDrawPileO As Collection
Global cDeadPileO As Collection
Global cDiscardPileO As Collection
Global cDefeatedCharactersPileO As Collection

'Characters
Global cFrontLine As clsFrontLine
Global cOpponent As clsOpponent

'Missions
Global cMissions As Collection 'Will contain all seven mission cards at the beginning
Global cDeadMissions As Collection
Global cVenturedMissions As Collection
Global cCompletedMissions As Collection
Global cVenturedC As Collection

'Opponent Missions
Global cMissionsO As Collection 'Will contain all seven mission cards at the beginning
Global cDeadMissionsO As Collection
Global cVenturedMissionsO As Collection
Global cCompletedMissionsO As Collection
Global cVenturedCO As Collection

'My Hand
Global cHand As Collection 'the current hand
Global cHandTags As Collection 'Tags for each card in the current hand

'Opponent Hand
Global cHandO As Collection

'Various card types
Global myHero As clsHero
Global myspecial As clsSpecial
Global myPower As clsPowerCard
Global myTeamwork As clsTeamwork
Global myEvent As clsEvent
Global myAlly As clsAlly
Global myMission As clsMission
Global myTraining As clsTraining
Global myHomebase As clsHomebase
Global myActivator As clsActivator
Global myAspect As clsAspect
Global myBasic As clsBasicUniverse
Global myDoubleShot As clsDoubleShot
Global myBattleSite As clsBattlesite
Global myArtifact As clsArtifact

Global OpHomebase As clsHomebase
Global OpBattlesite As clsBattlesite

Public Const nPhase_WhoGoesFirst = 1
Public Const nPhase_Draw = 2
Public Const nPhase_Discard = 3
Public Const nPhase_Place = 4
Public Const nPhase_Venture = 6
Public Const nPhase_Attack = 7
Public Const nPhase_Defend = 8
Public Const nPhase_Resolve = 9

Global bStopProcessing As Boolean

Public Enum OpPile
    Draw = 1
    Hand = 2
    DISCARD = 3
    Dead = 4
    Defeated = 5
    [Reserve Missions] = 6
    [Dead Missions] = 7
    [Completed Missions] = 8
    [Reserve Ventured] = 9
    [Completed Ventured] = 10
    Battlesite = 11
    [Hero 1 Placed] = 12
    [Hero 2 Placed] = 13
    [Hero 3 Placed] = 14
    [Hero 4 Placed] = 15
    Attack = 16
    Defense = 17
End Enum

Public Enum ModifierType
    modifies_battle = 1
    Modifies_Game = 2
    Modifies_Artifact = 3
End Enum

Global sOpponentName As String

Global nTurn  'Current turn number
Global myPhase As Integer
Global nPile As OpPile
Global bHavePassed As Boolean
Global bOppPassed As Boolean
Global bOppOpenHanded As Boolean
Global bIncomingAttackFaceDown As Boolean
Global bIHaveConceded As Boolean
Global bOpponentConceded As Boolean

'Sound variables
Global sSounds(18) As String
Global sSoundDes(18) As String

Const HW = 1.395759
'HW is what you multiple the Width by to get correct height

Public Declare Function sndPlaySound Lib "winmm.dll" Alias _
       "sndPlaySoundA" (ByVal lpszSoundName As String, _
       ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10


Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub Main()
dbName = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Overpower.mdb"
sBlankImagePath = App.Path & "\NotFound.jpg"

sOpponentName = "ENEMY"
sOpponentLocation = "c:\OpTest\"

Set mySettings = New clsSettings

mySettings.Load
LoadSounds

Set myattack = New clsAttack

'Startup
frmTable.Show


End Sub
Sub ShowErrorMessage(ErrorNumber As Integer, ErrorLocation As String, ErrorCaption As String, ErrorText As String, critical As Boolean)
X = MsgBox("ERROR #" & Trim(Str$(ErrorNumber)) & vbCrLf & vbCrLf & ErrorText & vbCrLf & vbCrLf & "LOCATION: " & ErrorLocation, vbCritical, ErrorCaption)

If critical = True Then
Close
End
End If

End Sub
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
Public Function ConvertPowerCode(Scode) As String

Select Case Scode
Case "S"
ConvertPowerCode = "Strength"
Case "F"
ConvertPowerCode = "Fighting"
Case "E"
ConvertPowerCode = "Energy"
Case "I"
ConvertPowerCode = "Intellect"
Case "A"
ConvertPowerCode = "Any-Power"
Case "M"
ConvertPowerCode = "Multi-Power"

Case Else
ConvertPowerCode = "Unknown"
End Select

End Function
Public Sub LoadOpponentDeck()

Set cOpponent = New clsOpponent

cOpponent.AddCharacter 30, False, False
cOpponent.AddCharacter 12, True, False
cOpponent.AddCharacter 22, False, False
cOpponent.AddCharacter 46, False, False

'Clear Various Deck Variables
Set cDrawPileO = New Collection
Set cDeadPileO = New Collection
Set cDiscardPileO = New Collection
Set cDefeatedCharactersPileO = New Collection

For i = 1 To 10
Set myspecial = New clsSpecial

myspecial.Load 513 + i
cDrawPileO.Add myspecial
Next i

End Sub
Public Sub ClearCollections()
'Clear the frontline
Set cFrontLine = New clsFrontLine

'Clear the various deck variables
Set cDrawPile = New Collection
Set cDeadPile = New Collection
Set cDiscardPile = New Collection
Set cDefeatedCharactersPile = New Collection

'Missions
Set cMissions = New Collection
Set cDeadMissions = New Collection
Set cVenturedMissions = New Collection
Set cCompletedMissions = New Collection
Set cVenturedC = New Collection

'Load homebase
Set myHomebase = New clsHomebase

'Load Battlesite
Set myBattleSite = New clsBattlesite
myBattleSite.NewBattlesite

Set cHand = New Collection

End Sub
Public Sub ClearOpponentCollections()
Set cOpponent = New clsOpponent

'Clear Various Deck Variables
Set cDrawPileO = New Collection
Set cDeadPileO = New Collection
Set cDiscardPileO = New Collection
Set cDefeatedCharactersPileO = New Collection

'Clear Missions Variables
Set cMissionsO = New Collection
Set cDeadMissionsO = New Collection
Set cVenturedMissionsO = New Collection
Set cCompletedMissionsO = New Collection
Set cVenturedCO = New Collection

Set cIncomingAttack = New Collection
Set cIncomingDefense = New Collection

'Clear Opponents Hand
Set cHandO = New Collection

End Sub
Public Sub ShuffleDrawPile()
Dim ctemp As Collection

Set ctemp = New Collection
Randomize

While cDrawPile.Count > 0

X = Int(Rnd * cDrawPile.Count) + 1

ctemp.Add cDrawPile.Item(X)
cDrawPile.Remove X

Wend

For i = 1 To ctemp.Count
cDrawPile.Add ctemp.Item(i)
Next i

Set ctemp = Nothing


End Sub
Public Sub ShufflePile(nPile As Integer)
'0 = Draw pile, 1 = Discard, 2 = Dead
Dim ctemp As Collection
Dim ctemp2 As Collection

Set ctemp2 = New Collection

Select Case nPile
Case 0
    Set ctemp = cDrawPile
Case 1
    Set ctemp = cDiscardPile
Case 2
    Set ctemp = cDeadPile
End Select

Randomize

While ctemp.Count > 0
X = Int(Rnd * ctemp.Count) + 1
ctemp2.Add ctemp.Item(X)
ctemp.Remove X
    
Wend

Select Case nPile
Case 0
    Set cDrawPile = ctemp2
Case 1
    Set cDiscardPile = ctemp2
Case 2
    Set cDeadPile = ctemp2
End Select


Set ctemp2 = Nothing

End Sub

Public Function ReplaceAllInString(String1, FindString, ReplaceString)
a$ = ""
b$ = String1

looper:
X = InStr(b$, FindString)

If X = 0 Then
    ReplaceAllInString = a$ + b$
    Exit Function
End If

a$ = a$ + Left$(b$, X - 1) + ReplaceString
b$ = Right$(b$, Len(b$) - ((X + Len(FindString)) - 1))
GoTo looper

End Function
Public Function GetVal(sString) As String
z = InStr(sString, "=")
If z > 0 Then
    sString = Right(sString, Len(sString) - z)
Else
    sString = ""
End If

GetVal = sString

End Function

Public Sub Code_ImportMissionString(ctemp As Collection, sCards As String)

cdt$ = sCards

If Right(cdt$, 1) = ":" Then cdt$ = Left(cdt$, Len(cdt$) - 1)

Set ctemp = New Collection

For i = 1 To Val(cdt$)
ctemp.Add "1"
Next i

End Sub
Public Function Code_ReturnCollectionFromNumber(pID) As Collection

Select Case pID
    Case 1
        Set Code_ReturnCollectionFromNumber = cDrawPileO
    Case 2
        Set Code_ReturnCollectionFromNumber = cHandO
    Case 3
        Set Code_ReturnCollectionFromNumber = cDiscardPileO
    Case 4
        Set Code_ReturnCollectionFromNumber = cDeadPileO
    Case 5
        Set Code_ReturnCollectionFromNumber = cDefeatedCharactersPileO
    Case 6
        Set Code_ReturnCollectionFromNumber = cMissionsO
    Case 7
        Set Code_ReturnCollectionFromNumber = cDeadMissionsO
    Case 8
        Set Code_ReturnCollectionFromNumber = cCompletedMissionsO
    Case 9
        Set Code_ReturnCollectionFromNumber = cVenturedMissionsO
    Case 10
        Set Code_ReturnCollectionFromNumber = cVenturedCO
    Case 11
        Set Code_ReturnCollectionFromNumber = temp
    Case 16
        Set Code_ReturnCollectionFromNumber = cIncomingAttack
    Case 17
        Set Code_ReturnCollectionFromNumber = cIncomingDefense
        
    Case Else
End Select

End Function
Public Sub Code_PlaceCard(nFromID, nToID, ncardid)
Dim ctemp As Collection
Dim ctemp2 As Collection
Dim ccard

Set ctemp = Code_ReturnCollectionFromNumber(nFromID)
Set ccard = ctemp.Item(ncardid)

If nToID = 5 Then
    OpHomebase.PlaceCard ccard, True
    ctemp.Remove ncardid
    Exit Sub
End If

If sSounds(13) <> "" Then PlaySound sSounds(13)

frmTable.lstGameHistory.AddItem "OP PLACED: " & ccard.Title
frmTable.lstGameHistory.ListIndex = frmTable.lstGameHistory.ListCount - 1

cOpponent.PlaceCard nToID, ccard
ctemp.Remove ncardid

End Sub
Public Sub Code_IncomingDefense(ncardfrom(), ncardid())
Dim ctemp As Collection
Dim ccard

Set OpDefense = New clsDefense

OpDefense.NewDefense

For i = 1 To 5

    If ncardfrom(i) <> 0 Then
    
    Select Case ncardfrom(i)
    
    Case 1 To 4
    
    Set ccard = cOpponent.PlacedCard(nFromID, ncardid(i))
    OpDefense.AddCardO ccard
    cOpponent.RemovePlacedCard nFromID, ncardid(i)
                
    Case 5
        
    Set ccard = cHandO.Item(ncardid(i))
    OpDefense.AddCardO ccard
    cHandO.Remove ncardid(i)
    
    Case 6

    Set ccard = OpHomebase.PlacedCard(ncardid(i))
    OpDefense.AddCardO ccard
    OpHomebase.RemovePlacedCard ncardid(i)
    
    End Select
    
    End If
    
Next i

    
End Sub
Public Sub Code_MoveCard(nFromID, nToID, ncardid)
'1=Draw; 2= Hand; 3= Discard; 4 = Dead; 5 = Defeated; 6 = Reserve missions
' 7=Dead missions; 8= Completed missions; 9 = Ventured from reserve
'10= ventured from completed; 11 = Battlesite deck
'12 = Hero 1 Placed; '13 = Hero 2 Placed; '14 = Hero 3 Placed; '15 = Hero 4 Placed
'16 = Attack; 17 = Defense
'18 = Hero 1 Modifier; 19 = Hero2 Modifier; 20 = Hero3 Modifier; 21 = Hero4 Modifier
'22 = Hero 1 Buffer; 23 = Hero2 Buffer; 24 = Hero3 Buffer; 25 = Hero4 Buffer
'26 = Homebase

Dim ctemp As Collection
Dim ctemp2 As Collection
Dim ccard

On Error Resume Next

If nToID = 16 Then
'Card is being to an attack

    If nFromID = 26 Then

     Set ccard = OpHomebase.PlacedCard(ncardid)
     
     cIncomingAttack.Add ccard
                 
     OpHomebase.RemovePlacedCard ncardid

    Exit Sub
    
    End If
    
    If (nFromID >= 12 And nFromID <= 15) Then
    'Card is placed
    
        nId = nFromID - 11
    
        Set ccard = cOpponent.PlacedCard(nId, ncardid)
        
        cIncomingAttack.Add ccard
                    
        cOpponent.RemovePlacedCard nId, ncardid
    
    Else
    
    'Card is in hand
            Set ccard = cHandO.Item(ncardid)
            
            cIncomingAttack.Add ccard
        
            cHandO.Remove ncardid
    
    End If

Exit Sub

End If

If nFromID = 16 Then
'Card is being moved from an attack

    If nToID = 26 Then
    'Card is being returned to a homebase
    
            Set ccard = cIncomingAttack.Item(ncardid)
            
            OpHomebase.PlaceCard ccard, False
            
            cIncomingAttack.Remove ncardid
               
        

    Exit Sub
        
    End If
    
    If (nToID >= 12 And nToID <= 15) Then
    'Card is placed
    
        nId = nToID - 11
        

            Set ccard = cIncomingAttack.Item(ncardid)
            
            cOpponent.PlaceCard nId, ccard
            
            cIncomingAttack.Remove ncardid
        
        
    Else
    
    'Card is in hand
            Set ccard = cIncomingAttack.Item(ncardid)
            
            cHandO.Add ccard
        
            cIncomingAttack.Remove ncardid
    
    End If

Exit Sub

End If


If nFromID = 26 Then
'Move card from Homebase
    Set ctemp = Code_ReturnCollectionFromNumber(nToID)
    Set ccard = OpHomebase.PlacedCard(ncardid)
    OpHomebase.RemovePlacedCard ncardid
    ctemp.Add ccard
    Exit Sub
End If

    
If (nFromID >= 12 And nFromID <= 15) And (nToID = 3 Or nToID = 4) Then
    If sSounds(12) <> "" Then PlaySound sSounds(12)

    Set ccard = cOpponent.PlacedCard(nFromID - 11, ncardid)
    frmTable.lstGameHistory.AddItem "OP DISCARD: " & ccard.Title
    frmTable.lstGameHistory.ListIndex = frmTable.lstGameHistory.ListCount - 1

    Set ccard = Nothing
End If

If (nFromID >= 22 And nFromID <= 25) And (nToID = 3 Or nToID = 4) Then
    If sSounds(12) <> "" Then PlaySound sSounds(12)

    Set ccard = cOpponent.Modifiers_GetCard(nFromID - 21, ncardid)
    frmTable.lstGameHistory.AddItem "OP DISCARD: " & ccard.Title
    frmTable.lstGameHistory.ListIndex = frmTable.lstGameHistory.ListCount - 1
    Set ccard = Nothing
End If

If (nFromID >= 18 And nFromID <= 21) And (nToID = 3 Or nToID = 4) Then
    If sSounds(12) <> "" Then PlaySound sSounds(12)

    Set ccard = cOpponent.Buffers_GetCard(nFromID - 17, ncardid)
    frmTable.lstGameHistory.AddItem "OP DISCARD: " & ccard.Title
    frmTable.lstGameHistory.ListIndex = frmTable.lstGameHistory.ListCount - 1
    Set ccard = Nothing
End If


If (nFromID = 2 And nToID = 3) Or (nFromID = 2 And nToID = 4) Then
'Discarding card.  Add to history
    If sSounds(12) <> "" Then PlaySound sSounds(12)
    Set ccard = cHandO.Item(ncardid)

    frmTable.lstGameHistory.AddItem "OP DISCARD: " & ccard.Title
    frmTable.lstGameHistory.ListIndex = frmTable.lstGameHistory.ListCount - 1

    Set ccard = Nothing

End If



'plaing a buffer from placed
If (nFromID >= 12 And nFromID <= 15) And (nToID >= 22 And nToID <= 25) Then
    nId = nToID - 21
    
    Set ccard = cOpponent.PlacedCard(nId, ncardid)
    
    cOpponent.RemovePlacedCard nId, ncardid
    
    cOpponent.Buffers_AddCard nId, ccard

    History_Add "BUFFER PLAYED TO: " & cOpponent.Character_Name(nId)

    Exit Sub

End If


'Add Buffer
If (nToID >= 22 And nToID <= 25) Then

    nId = nToID - 21
    
    Set ctemp = Code_ReturnCollectionFromNumber(nFromID)
    
    Set ccard = ctemp.Item(ncardid)
    
    cOpponent.Buffers_AddCard nId, ccard
    
    ctemp.Remove ncardid
    
    History_Add "==============================================================================================================="
    History_Add sOpponentName & ": PLAYS BUFFER TO " & cOpponent.Character_Name(nId)
    History_Add "---------------------------------------------------------------------------------------------------------------"
    History_Add ccard.Title
    History_Add "==============================================================================================================="
    
    Exit Sub

End If

'Remove Buffer
If (nFromID >= 22 And nToID <= 25) Then
    nId = nToID - 21
    
    Set ctemp = Code_ReturnCollectionFromNumber(nToID)
    
    Set ccard = cOpponent.Buffers_GetCard(nId, ncardid)
        
    ctemp.Add ccard
    
    cOpponent.Buffers_RemoveCard nId, ncardid
    
    Exit Sub
    
End If

If (nFromID >= 12 And nFromID <= 15) And (nToID >= 18 And nToID <= 21) Then
    nId = nToID - 17
    
    Set ccard = cOpponent.PlacedCard(nId, ncardid)
    
    cOpponent.RemovePlacedCard nId, ncardid
    
    If ccard.Attack_Frontline_BattleBonus = True Then
        cOpponent.Modifiers_AddCard nId, ccard, modifies_battle
    End If
    
    If ccard.Attack_Frontline_GameBonus = True Then
        cOpponent.Modifiers_AddCard nId, ccard, Modifies_Game
    End If

    History_Add "MODIFIER PLAYED TO: " & cOpponent.Character_Name(nId)

    Exit Sub

End If

'Remove Modifier
If (nFromID >= 18 And nFromID <= 21) Then
    nId = nFromID - 17
    
'    Set ctemp = Code_ReturnCollectionFromNumber(nFromID)
    
    Set ccard = cOpponent.Modifiers_GetCard(nId, ncardid)
    
    cDeadPileO.Add ccard
    
    cOpponent.Modifiers_RemoveCard nId, ncardid
        
    Exit Sub
    
End If

'Add Modifier
If (nToID >= 18 And nToID <= 21) Then

    nId = nToID - 17
    
    Set ctemp = Code_ReturnCollectionFromNumber(nFromID)
    
    Set ccard = ctemp.Item(ncardid)
    
    If ccard.Attack_Frontline_BattleBonus = True Then cOpponent.Modifiers_AddCard nId, ccard, modifies_battle
    If ccard.Attack_Frontline_GameBonus = True Then cOpponent.Modifiers_AddCard nId, ccard, Modifies_Game
    
    ctemp.Remove ncardid
    
    History_Add "==============================================================================================================="
    History_Add sOpponentName & ": PLAYS MODIFIER [" & cOpponent.Modifiers_TypeText(nId, ncardid) & "] TO " & cOpponent.Character_Name(nId)
    History_Add "---------------------------------------------------------------------------------------------------------------"
    History_Add ccard.Title
    History_Add "==============================================================================================================="
    
    Exit Sub

End If



If (nToID >= 12 And nToID <= 15) Then
'Move from a collection to a placed card

nId = nToID - 11

Set ctemp = Code_ReturnCollectionFromNumber(nFromID)

Set ccard = ctemp.Item(ncardid)

If sSounds(13) <> "" Then PlaySound sSounds(13)
frmTable.lstGameHistory.AddItem "OP PLACE: " & ccard.Title
frmTable.lstGameHistory.ListIndex = frmTable.lstGameHistory.ListCount - 1

cOpponent.PlaceCard nId, ccard

ctemp.Remove ncardid

Exit Sub

End If


'Moving a placed card to a pile or hand
If (nFromID >= 12 And nFromID <= 15) And (nToID < 6 Or nToID > 15) Then

    nId = nFromID - 11
    
    Set ccard = cOpponent.PlacedCard(nId, ncardid)
    
    Set ctemp2 = Code_ReturnCollectionFromNumber(nToID)
    
    ctemp2.Add ccard
    
    cOpponent.RemovePlacedCard nId, ncardid
    
    Exit Sub

End If

'Movements to or from Battlesite
If nFromID = 11 Then

    Set ctemp2 = Code_ReturnCollectionFromNumber(nToID)
    
    Set ccard = OpBattlesite.Deck_GetCard(ncardid)
    
    ctemp2.Add ccard
    
    OpBattlesite.RemoveDeckCard ncardid

Exit Sub

End If

If nToID = 11 Then
    Set ctemp = Code_ReturnCollectionFromNumber(nFromID)
    
    Set ccard = ctemp.Item(ncardid)
    
    OpBattlesite.Deck_AddCard ccard
    
    ctemp.Remove ncardid

    Exit Sub
End If

Set ctemp = Code_ReturnCollectionFromNumber(nFromID)
Set ctemp2 = Code_ReturnCollectionFromNumber(nToID)


If nFromID >= 6 And nFromID <= 10 Then
'This is a movement of venture cards
'In this case, ncardID = the number of cards to move

On Error Resume Next
For i = 1 To ncardid
    ctemp.Remove 1
    ctemp2.Add "1"
Next i

Exit Sub
End If


MoveOppCard ctemp, ctemp2, ncardid

End Sub
Public Sub MoveOppCard(cTempFrom As Collection, cTempTo As Collection, nId)
Dim ccard

Set ccard = cTempFrom.Item(nId)

cTempTo.Add ccard

cTempFrom.Remove nId


End Sub
Public Function Code_ImportPileString(ctemp As Collection, sCards As String)

Set ctemp = New Collection

If sCards = "" Then Exit Function

cdt$ = sCards

ProcessCards:

        X = InStr(cdt$, ":")
        a$ = Left(cdt$, X - 1)
        nId = Val(Right(a$, Len(a$) - 1))
        a$ = Left(a$, 1)
           
        Select Case Trim(UCase(a$))
        
        Case "M"
            Set myMission = New clsMission
            myMission.Load nId
            
            ctemp.Add myMission
            
        Case "A"
            
            Set myActivator = New clsActivator
            
            myActivator.Load nId
            
            ctemp.Add myActivator
        
        Case "I"
            Set myArtifact = New clsArtifact
            
            myArtifact.Load nId
            ctemp.Add myArtifact
            
        Case "L"
        
            Set myAlly = New clsAlly
            
            myAlly.Load nId
            
            ctemp.Add myAlly
            
        Case "X"
        
            Set myAspect = New clsAspect
            
            myAspect.Load nId
            
            ctemp.Add myAspect
        
        Case "B"
        
            Set myBasic = New clsBasicUniverse
            
            myBasic.Load nId
            
            ctemp.Add myBasic
            
        Case "D"
        
            Set myDoubleShot = New clsDoubleShot
            
            myDoubleShot.Load nId
            
            ctemp.Add myDoubleShot
            
        Case "E"
        
            Set myEvent = New clsEvent
            
            myEvent.Load nId
            
            ctemp.Add myEvent
            
        Case "P"
        
            Set myPower = New clsPowerCard
            
            myPower.Load nId
            
            ctemp.Add myPower
            
        Case "S"
        
            Set myspecial = New clsSpecial
            
            myspecial.Load nId
            
            ctemp.Add myspecial
            
        Case "T"
            
            Set myTeamwork = New clsTeamwork
            
            myTeamwork.Load nId
            
            ctemp.Add myTeamwork
            
        Case "R"
                    
            Set myTraining = New clsTraining
            
            myTraining.Load nId
            
            ctemp.Add myTraining

        Case Else
        
        End Select
        
        
        cdt$ = Right(cdt$, Len(cdt$) - X)
        
        If cdt$ <> "" Then GoTo ProcessCards
        
End Function
Public Function GetCode_CardString(ctemp As Collection) As String

Dim ccard

a$ = ""

For i = 1 To ctemp.Count

Set ccard = ctemp.Item(i)

Select Case ccard.CardType

Case "Mission"
    ccode = "M"
    
Case "Activator" 'A
    ccode = "A"

Case "Ally Card" 'Y
    ccode = "L"
    
Case "Aspect Card" 'P
    ccode = "X"
    
Case "Basic Universe"
    ccode = "B"
    
Case "Double Shot"
    ccode = "D"
    
Case "Event"
    ccode = "E"
    
Case "Power Card"
    ccode = "P"
    
Case "Special Card"
    ccode = "S"
        
Case "Teamwork"
    ccode = "T"
    
Case "Training"
    ccode = "R"

Case "Artifact"
    ccode = "I"
Case Else
End Select

a$ = a$ & ccode & ccard.ID & ":"

Next i

GetCode_CardString = a$


End Function
Public Sub History_Add(sItem)
frmTable.lstGameHistory.AddItem sItem
frmTable.lstGameHistory.ListIndex = frmTable.lstGameHistory.ListCount - 1

End Sub
Public Function HaveConcedeEffect() As Boolean
Dim ccard

HaveConcedeEffect = False

For i = 1 To cHand.Count
    Set ccard = cHand.Item(i)
    If ccard.CardType = "Special Card" Then
        If ccard.Attack_StopsConcede = True Or ccard.Attack_PostConcessionAttack = True Then
            HaveConcedeEffect = True
            Exit Function
        End If
    End If
Next i

For i = 1 To 4

    For k = 1 To cFrontLine.Placed_Count(i)
    Set ccard = cFrontLine.PlacedCard(i, k)
    
    If ccard.CardType = "Special Card" Then
        If ccard.Attack_StopsConcede = True Or ccard.Attack_PostConcessionAttack = True Then
            HaveConcedeEffect = True
            Exit Function
        End If
    End If
    Next k
    
Next i


End Function
Public Function OpponentHasConcedeEffect() As Boolean
Dim ccard

OpponentHasConcedeEffect = False

For i = 1 To cHandO.Count
    Set ccard = cHandO.Item(i)
    If ccard.CardType = "Special Card" Then
        If ccard.Attack_StopsConcede = True Or ccard.Attack_PostConcessionAttack = True Then
            OpponentHasConcedeEffect = True
            Exit Function
        End If
    End If
Next i

For i = 1 To 4

    For k = 1 To cOpponent.Placed_Count(i)
    Set ccard = cOpponent.PlacedCard(i, k)
    
    If ccard.CardType = "Special Card" Then
        If ccard.Attack_StopsConcede = True Or ccard.Attack_PostConcessionAttack = True Then
            OpponentHasConcedeEffect = True
            Exit Function
        End If
    End If
    Next k
    
Next i


End Function
Public Sub SendData(sData)

If frmTable.tcpChannel.State = 0 Then Exit Sub

frmTable.tcpChannel.SendData sData


End Sub
Public Sub PlaySound(ByVal sSoundFile)
Dim sFlags As Long

On Error Resume Next

sSoundFile = App.Path & "\Sounds\" & sSoundFile

X = Dir(sSoundFile)
If X = "" Then Exit Sub

   sFlags = SND_ASYNC Or SND_NODEFAULT
   sndPlaySound sSoundFile, sFlags

End Sub

Private Sub LoadSounds()
'1 = Incoming Message
'2 = Incoming Attack
'3 = Incoming Defense
'4 = Outgoing Attack
'5 = Attack Hits Op
'6 = Opponent KO Hero
'7 = Opp Win Venture
'8 = Me no Defense
'9 = Me K.O. Hero
'10 = Me Win Venture

sSoundDes(1) = "Message Rec'd"
sSoundDes(2) = "Incoming Attack"
sSoundDes(3) = "Opp Defense"
sSoundDes(4) = "Attacking Opp"
sSoundDes(5) = "Attack Hits Opp"
sSoundDes(6) = "Opp K.O."
sSoundDes(7) = "Opp Win Venture"
sSoundDes(8) = "Attack Hits"
sSoundDes(9) = "K.O."
sSoundDes(10) = "Win Venture"
sSoundDes(11) = "Opp Concedes"
sSoundDes(12) = "Opp Discards"
sSoundDes(13) = "Opp Places Card"
sSoundDes(14) = "Opp Challenges Defense"
sSoundDes(15) = "I Discard"
sSoundDes(16) = "I Place Card"
sSoundDes(17) = "I Challenge Defense"
sSoundDes(18) = "Prepare Attack"

X = Dir(App.Path & "\Sounds\sound.ini")

On Error Resume Next

If X = "" Then

For i = 1 To 18
    sSounds(i) = ""
Next i

Else

X = FreeFile

Open App.Path & "\Sounds\sound.ini" For Input As #X
Counter = 0

Do Until EOF(X)

Counter = Counter + 1

Line Input #X, a$

If Counter < 19 Then sSounds(Counter) = a$
Loop

End If


Close #X


End Sub

Public Function StripDirectory(ByVal sFileName) As String

looper:
X = InStr(sFileName, "\")

If X < 1 Then
    StripDirectory = sFileName
    Exit Function
End If

sFileName = Right(sFileName, Len(sFileName) - X)

GoTo looper

End Function
