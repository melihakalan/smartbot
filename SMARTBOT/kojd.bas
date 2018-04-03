Attribute VB_Name = "kojd"
Option Explicit
Public buy_LastX As Long, buy_LastY As Long, buy_LastZ As Long
Public buy_LastHeal As Boolean, buy_LastBuff As Boolean, buy_LastRestore As Boolean, buy_LastGroup As Boolean, buy_LastSTR As Boolean, buy_LastSit As Boolean, buy_LastCure As Boolean, buy_LastMalice As Boolean, buy_LastBot As Boolean, buy_LastNoExp As Boolean, buy_LastFar As Boolean, buy_LastRPR As Boolean
Public rpr_LastHeal As Boolean, rpr_LastBuff As Boolean, rpr_LastRestore As Boolean, rpr_LastGroup As Boolean, rpr_LastSTR As Boolean, rpr_LastSit As Boolean, rpr_LastCure As Boolean, rpr_LastMalice As Boolean, rpr_LastBot As Boolean, rpr_LastNoExp As Boolean, rpr_LastFar As Boolean, rpr_LastBUY As Boolean
Public NeededHP As Boolean, NeededMP As Boolean
Public onGoing As Boolean
Public InvSlot1 As Long, InvSlot2 As Long, InvQuantity1 As Long, InvQuantity2 As Long
Public DisableHeal As Boolean
Public GroupHealInterval As Integer
Public CureLatency As Integer
Public PTCount As Long
Public MaxHP1 As Long, MaxHP2 As Long, MaxHP3 As Long, MaxHP4 As Long, MaxHP5 As Long, MaxHP6 As Long, MaxHP7 As Long, MaxHP8, MaxHPME As Long
Public PTUserID1 As Long, PTUserID2 As Long, PTUserID3 As Long, PTUserID4 As Long, PTUserID5 As Long, PTUserID6 As Long, PTUserID7 As Long, PTUserID8 As Long
Public PTUserHP1 As Long, PTUserHP2 As Long, PTUserHP3 As Long, PTUserHP4 As Long, PTUserHP5 As Long, PTUserHP6 As Long, PTUserHP7 As Long, PTUserHP8 As Long
Public PTUserMP1 As Long, PTUserMP2 As Long, PTUserMP3 As Long, PTUserMP4 As Long, PTUserMP5 As Long, PTUserMP6 As Long, PTUserMP7 As Long, PTUserMP8 As Long
Public PTUserMAXHP1 As Long, PTUserMAXHP2 As Long, PTUserMAXHP3 As Long, PTUserMAXHP4 As Long, PTUserMAXHP5 As Long, PTUserMAXHP6 As Long, PTUserMAXHP7 As Long, PTUserMAXHP8 As Long
Public PTUserMAXMP1 As Long, PTUserMAXMP2 As Long, PTUserMAXMP3 As Long, PTUserMAXMP4 As Long, PTUserMAXMP5 As Long, PTUserMAXMP6 As Long, PTUserMAXMP7 As Long, PTUserMAXMP8 As Long
Public PTUserLVL1 As Long, PTUserLVL2 As Long, PTUserLVL3 As Long, PTUserLVL4 As Long, PTUserLVL5 As Long, PTUserLVL6 As Long, PTUserLVL7 As Long, PTUserLVL8 As Long

Public WaitingUserLeave As Boolean, WaitingUser As Long
Public ExpBefore As Long, ExpAfter As Long
Public tmpCurMobHP As Long, MobHPOk As Boolean, WaitingForHP As Boolean
Public BindingInterval As Integer, RockInterval As Integer
Public MobChanged As Boolean
Public GotTimedAction As Boolean
Public conNtnTowned As Boolean, conDCOff As Boolean, conDeadOff As Boolean, conNoPtTowned As Boolean, conNoManaTowned As Boolean, conReturnedBack As Boolean
Public TestRpr As Boolean
Public LastExp As Long
Public sSid As Long, Lvl As Long
Public GMID As String, UserInAuth As Integer
Public LeftItem As Long, durLeftItem As Long, RightItem As Long, durRightItem As Long
Public tmpMob As Long
Public GetSlot As Integer
Public BoxIDNew As String
Public DropItem1 As String
Public DropItem2 As String
Public DropItem3 As String
Public DropItem4 As String

Public KO_TITLE As String
Public KO_PARTY As Long
Public KO_PTR_CHR As Long
Public KO_PTR_CHR1 As Long
Public KO_PTR_DLG As Long
Public KO_PTR_PKT As Long
Public KO_SND_FNC As Long
Public KO_RECVHK As Long
Public KO_RCVHKB As Long
Public KO_HOOKFIX As Long
Public KO_SENDFIX As Long
Public KO_SENDPTR As Long
Public KO_ADR_DLG As Long
Public KO_OFF_RTM As Long
Public KO_OFF_NT As Long
Public KO_OFF_CLASS As Long
Public KO_OFF_SWIFT As Long
Public KO_OFF_HP As Long
Public KO_OFF_MAXHP As Long
Public KO_OFF_MP As Long
Public KO_OFF_MAXMP As Long
Public KO_OFF_MOB As Long
Public KO_OFF_WH As Long
Public KO_OFF_SIT As Long
Public KO_OFF_LUP As Long
Public KO_OFF_LUP2 As Long
Public KO_OFF_Y As Long
Public KO_OFF_ID As Long
Public KO_OFF_HD As Long
Public KO_OFF_X As Long
Public KO_OFF_SEL As Long
Public KO_OFF_LUPINE As Long
Public KO_OFF_EXP As Long
Public KO_OFF_MAXEXP As Long
Public KO_OFF_GOLD As Long
Public KO_OFF_Z As Long
Public KO_OFF_LVL As Long
Public KO_OFF_STAT_MP As Long
Public KO_OFF_STAT_INT As Long
Public KO_OFF_STAT_HP As Long
Public KO_OFF_STAT_DEX As Long
Public KO_OFF_STAT_STR As Long
Public KO_OFF_AP As Long
Public KO_OFF_AC As Long
Public KO_DLGBMA As Long
Public KO_KEYBPTR As Long
Public nation As Long

Public Function HexToStr(ByVal Data As String) As String
Dim Buffer As String
Dim i As Integer
If Len(Data) Mod 2 <> 0 Then
HexToStr = vbNullString
Else
For i = 1 To Len(Data) - 1 Step 2
Buffer = Buffer & Chr("&H" & Mid(Data, i, 2))
Next i
HexToStr = Buffer
End If
End Function

Public Function SendPacket(packet As String)
Dim pStr As String
Dim pBytes() As Byte
pStr = packet
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Function

Public Function ReadAByte(Addr As Long) As Long 'read a byte value
Dim Value As Long
Dim pBytes() As Byte
Dim pSize As Long
    pSize = &H1
ReDim pBytes(1 To pSize)
    ReadByteArray Addr, pBytes, pSize
    Value = hex(pBytes(1))
    ReadAByte = Value
End Function

Function OpenBox()
SendPacket "24" + BoxIDNew
End Function
Function LootBox()
If Form2.opLootAll.Value = True Then
If DropItem1 <> "00000000" Then SendPacket "26" & BoxIDNew & DropItem1
If DropItem2 <> "00000000" Then SendPacket "26" & BoxIDNew & DropItem2
If DropItem3 <> "00000000" Then SendPacket "26" & BoxIDNew & DropItem3
If DropItem4 <> "00000000" Then SendPacket "26" & BoxIDNew & DropItem4
Else
SendPacket "26" & BoxIDNew & "00E9A435"
End If
End Function

Function RevHex8(hex As String)
'//* Function by KoJD
Dim newHex, hexLen, byte1, byte2, byte3, byte4 As String
hexLen = Len(hex)
Select Case hexLen
Case 1 '0000000A
newHex = "0" & hex & "000000" '0A000000
Case 2 '000000AB
newHex = hex & "000000" 'AB000000
Case 3 '00000ABC
byte1 = Left(hex, 1) 'A
byte2 = Right(hex, 2) 'BC
newHex = byte2 & "0" & byte1 & "0000" 'BC0A0000
Case 4 '0000ABCD
byte1 = Left(hex, 2) 'AB
byte2 = Right(hex, 2) 'CD
newHex = byte2 & byte1 & "0000" 'CDAB0000
Case 5 '000ABCDE
byte1 = Left(hex, 1) 'A
byte2 = Mid(hex, 2, 2) 'BC
byte3 = Right(hex, 2) 'DE
newHex = byte3 & byte2 & "0" & byte1 & "00" 'DEBC0A00
Case 6 '00ABCDEF
byte1 = Left(hex, 2) 'AB
byte2 = Mid(hex, 2, 2) 'CD
byte3 = Right(hex, 2) 'EF
newHex = byte3 & byte2 & byte1 & "00"
Case 7 '0ABCDEFG
byte1 = Left(hex, 1) 'A
byte2 = Mid(hex, 2, 2) 'BC
byte3 = Mid(hex, 4, 2) 'DE
byte4 = Right(hex, 2) 'FG
newHex = byte4 & byte3 & byte2 & "0" & byte1 'FGDEBC0A
Case 8 'ABCDEFGH
byte1 = Left(hex, 2) 'AB
byte2 = Mid(hex, 3, 2) 'CD
byte3 = Mid(hex, 5, 2) 'EF
byte4 = Right(hex, 2) 'GH
newHex = byte4 & byte3 & byte2 & byte1 'GHEFCDAB
Case Else
End Select
RevHex8 = newHex
End Function

Function FormatHex(strHex As String, inLength As Integer)
'// Function by KoJD
'//-Repair and Reverse Hex String for Knight Online-
Dim newHex As String, byte1 As String, byte2 As String, byte3 As String, byte4 As String
Dim ZeroSpaces As Integer
'ABC,4
ZeroSpaces = inLength - Len(strHex) '1
newHex = String(ZeroSpaces, "0") + strHex '0ABC
byte1 = Left(newHex, 2)
byte2 = Mid(newHex, 3, 2)
byte3 = Mid(newHex, 5, 2)
byte4 = Right(newHex, 2)
Select Case Len(newHex)
Case 2 '0A
newHex = byte1
Case 4 '0ABC
newHex = byte4 & byte1
Case 6 '000ABC
newHex = byte4 & byte2 & byte1
Case 8 '00000ABC
newHex = byte4 & byte3 & byte2 & byte1
Case Else
End Select
FormatHex = newHex
'\\
End Function

Function GetMobID()
Dim mob As String
mob = hex(ReadLong(KO_ADR_CHR + KO_OFF_MOB))
If mob = "FFFFFFFF" Then
mob = "FFFF"
Else
mob = FormatHex(mob, 4)
End If
GetMobID = mob
End Function

Function GetCharID()
Dim char As String
char = hex(ReadLong(KO_ADR_CHR + KO_OFF_ID))
char = FormatHex(char, 6)
char = Right(char, 4)
GetCharID = char
End Function

Function GetMobHP()
Dim MobBase As Long
Dim MobBase2 As Long
Dim tmpPtr As Long
Dim mobhp As Long
tmpPtr = ReadLong(KO_PTR_DLG)
MobBase = ReadLong(tmpPtr + &H1B8)
MobBase2 = ReadLong(MobBase + &HC4)
mobhp = ReadLong(MobBase2 + &HEC)
GetMobHP = mobhp
End Function

Function GetMobX()
Dim tmpPtr As Long
Dim MobBase As Long
Dim MobX As Single
tmpPtr = ReadLong(KO_PTR_DLG)
MobBase = ReadLong(tmpPtr + &H3D4)
MobX = ReadFloat(MobBase + &H48)
GetMobX = MobX
End Function

Function GetMobY()
Dim tmpPtr As Long
Dim MobBase As Long
Dim MobY As Single
tmpPtr = ReadLong(KO_PTR_DLG)
MobBase = ReadLong(tmpPtr + &H3D4)
MobY = ReadFloat(MobBase + &H50)
GetMobY = MobY
End Function

Function GetMobDistance()
Dim tmpPtr, a, b, mX, mY, cx, cy As Long
Dim frkx, frky, uz As Single
tmpPtr = ReadLong(KO_PTR_CHR)
cx = ReadFloat(tmpPtr + &HB4)
cy = ReadFloat(tmpPtr + &HBC)
frkx = (GetMobX - cx) * (GetMobX - cx)
frky = (GetMobY - cy) * (GetMobY - cy)
uz = Fix(((frkx + frky) ^ 0.5) / 4)
GetMobDistance = uz
End Function

Function Chat(strType As Integer)
Dim pStr As String
Dim pBytes() As Byte

If strType = 0 Then
pStr = "1001FF00" & hexword
ElseIf strType = 1 Then
pStr = "1005FF00" & hexword
ElseIf strType = 2 Then
pStr = "1003FF00" & hexword
ElseIf strType = 3 Then
pStr = "1006FF00" & hexword
ElseIf strType = 4 Then
pStr = "1004FF00" & hexword
ElseIf strType = 5 Then
pStr = "100EFF00" & hexword
End If

ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Function

Function GetSkillNum(LastNum As String)
On Error Resume Next
Dim SkillNum
SkillNum = ReadLong(KO_ADR_CHR + KO_OFF_CLASS) & LastNum
GetSkillNum = FormatHex(hex(SkillNum), 6)
End Function

Function FormatDec(strHex As String, inLength As Integer)
'//Function by KoJD
'//Convert reversed hex to it's real deciminal value
'500->1F4->F401

Dim NewDec As String, byte1 As String, byte2 As String, byte3 As String, byte4 As String
Dim ZeroSpaces As Integer
ZeroSpaces = inLength - Len(strHex)
NewDec = String(ZeroSpaces, "0") + strHex

byte1 = Left(NewDec, 2)
byte2 = Mid(NewDec, 3, 2)
byte3 = Mid(NewDec, 5, 2)
byte4 = Right(NewDec, 2)

Select Case Len(NewDec)
Case 2
NewDec = CDec("&H" & byte1)
Case 4
NewDec = CDec("&H" & byte4 & byte1)
Case 6
NewDec = CDec("&H" & byte4 & byte2 & byte1)
Case 8
NewDec = CDec("&H" & byte4 & byte3 & byte2 & byte1)
End Select

FormatDec = NewDec
End Function

Public Sub GetInventory()
      Dim tmpBase As Long, tmpLng1 As Long, tmpLng2 As Long, tmpLng3 As Long, tmpLng4 As Long
      Dim lngItemID As Long, lngItemID_Ext As Long, lngItemNameLen As Long, AdrItemName As Long, ItemQuantity As Long
      Dim ItemNameB() As Byte
      Dim ItemName As String
      Dim i As Integer
      
      tmpBase = ReadLong(KO_PTR_DLG) 'read KO_DLGBMA adress
      tmpLng1 = ReadLong(tmpBase + &H1A0) 'first pointer
      
For i = 0 To 41 'read 0 to 41 inventory slots (0=earring, 1=helmet, 2=earring, 3=necklace, 4=pauldron ....14=first inventory slot)
          tmpLng2 = ReadLong(tmpLng1 + (&H134 + (4 * i))) 'inventory slot
          tmpLng3 = ReadLong(tmpLng2 + &H38) 'item id adress
          tmpLng4 = ReadLong(tmpLng2 + &H3C) 'item id_ext adress
          ItemQuantity = ReadLong(tmpLng2 + &H40)
          
            lngItemID = ReadLong(tmpLng3) 'item id value
          lngItemID_Ext = ReadLong(tmpLng4) 'item id_ext value
          lngItemID = lngItemID + lngItemID_Ext 'real item id
          lngItemNameLen = ReadLong(tmpLng3 + &H10) 'n° characters in item name
          AdrItemName = ReadLong(tmpLng3 + &HC) 'item name adress
          
           ItemName = "" 'reset ItemName variable
          If lngItemNameLen > 0 Then
              ReadByteArray AdrItemName, ItemNameB, lngItemNameLen 'get item name (byte array)
              ItemName = StrConv(ItemNameB, vbUnicode) 'convert it to string
          End If
          
Debug.Print Format$(i, "00") & "- " & lngItemID & " " & ItemName & "Quantity:" & ItemQuantity 'print item id and item name in debug window (add ur code here)
      Next
  End
  End Sub
  
Function InviteParty(User As String)
Dim UserLen As Integer
UserLen = FormatHex(hex(Len(User)), 2)
HexString User
SendPacket "2F01" & UserLen & "00" & hexword
End Function

Function GoTown()
If ReadLong(KO_ADR_CHR + KO_OFF_HP) > 0 Then
SendPacket "4800"
End If
End Function

Function KickParty(UserID As Long)
SendPacket "2F04" & UserID
End Function

Function HealUser(UserID As String)
Dim LastNum As String
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

Select Case Form2.comAutoHeal.ListIndex
Case 0
LastNum = "002"
Case 1
LastNum = "005"
Case 2
LastNum = "500"
Case 3
LastNum = "509"
Case 4
LastNum = "518"
Case 5
LastNum = "527"
Case 6
LastNum = "536"
Case 7
LastNum = "545"
Case Else
End Select

SendPacket "3101" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "000000000000000000000000"

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function BuffUser(UserID As String)
Dim LastNum As String
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

Select Case Form2.comAutoBuff.ListIndex
Case 1
LastNum = "606"
Case 2
LastNum = "615"
Case 3
LastNum = "624"
Case 4
LastNum = "633"
Case 5
LastNum = "642"
Case 6
LastNum = "654"
Case 7
LastNum = "655"
Case 8
LastNum = "657"
Case 9
LastNum = "670"
Case 10
LastNum = "675"
Case Else
Exit Function
End Select

SendPacket "3101" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "000000000000000000000000"

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function ACUser(UserID As String)
Dim LastNum As String
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

Select Case Form2.comAutoAc.ListIndex
Case 1
LastNum = "603"
Case 2
LastNum = "612"
Case 3
LastNum = "621"
Case 4
LastNum = "630"
Case 5
LastNum = "639"
Case 6
LastNum = "651"
Case 7
LastNum = "660"
Case 8
LastNum = "674"
Case Else
Exit Function
End Select

SendPacket "3101" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "000000000000000000000000"

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function CureCurse(UserID As String)
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

SendPacket "3101" & GetSkillNum("525") & "00" & GetCharID & UserID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("525") & "00" & GetCharID & UserID & "000000000000000000000000"

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function CureDisease(UserID As String)
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

SendPacket "3101" & GetSkillNum("535") & "00" & GetCharID & UserID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("535") & "00" & GetCharID & UserID & "000000000000000000000000"

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function BuffByChat(UserName As String)
On Error Resume Next
Dim tmpUser As String
Dim UserID
Dim NameLen As Long
Dim pBytes() As Byte
tmpUser = ""
Dim i As Integer

If PTCount > 0 Then
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H34)
Debug.Print "BuffByChat-NameLen:" & NameLen
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H30) + &H0, pBytes, NameLen
Debug.Print hex(pBytes(1)) & "," & hex(pBytes(2)) & "," & hex(pBytes(3)) & "," & hex(pBytes(4)) & "," & hex(pBytes(5)) & "," & hex(pBytes(6)) & "," & hex(pBytes(7)) & "," & hex(pBytes(8))
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes(i))
Next
Debug.Print "BuffByChat-Name:" & HexToStr(tmpUser)
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 1 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H34)
Dim pBytes2() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H30) + &H0, pBytes2, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes2(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 2 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H34)
Dim pBytes3() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes3, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes3(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 3 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes4() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes4, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes4(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 4 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes5() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes5, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes5(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 5 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes6() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes6, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes6(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 6 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes7() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes7, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes7(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 7 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes8() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes8, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes8(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
BuffUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If
End Function
Function ACByChat(UserName As String)
On Error Resume Next
Dim tmpUser As String
Dim UserID
Dim NameLen As Long
Dim pBytes() As Byte
tmpUser = ""
Dim i As Integer

If PTCount > 0 Then
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H34)
Debug.Print "BuffByChat-NameLen:" & NameLen
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H30) + &H0, pBytes, NameLen
Debug.Print hex(pBytes(1)) & "," & hex(pBytes(2)) & "," & hex(pBytes(3)) & "," & hex(pBytes(4)) & "," & hex(pBytes(5)) & "," & hex(pBytes(6)) & "," & hex(pBytes(7)) & "," & hex(pBytes(8))
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes(i))
Next
Debug.Print "BuffByChat-Name:" & HexToStr(tmpUser)
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 1 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H34)
Dim pBytes2() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H30) + &H0, pBytes2, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes2(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 2 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H34)
Dim pBytes3() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes3, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes3(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 3 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes4() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes4, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes4(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 4 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes5() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes5, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes5(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 5 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes6() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes6, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes6(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 6 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes7() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes7, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes7(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 7 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes8() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes8, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes8(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ACUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If
End Function
Function CureByChat(UserName As String)
On Error Resume Next
Dim tmpUser As String
Dim UserID
Dim NameLen As Long
Dim pBytes() As Byte
tmpUser = ""
Dim i As Integer

If PTCount > 0 Then
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H34)
Debug.Print "BuffByChat-NameLen:" & NameLen
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H30) + &H0, pBytes, NameLen
Debug.Print hex(pBytes(1)) & "," & hex(pBytes(2)) & "," & hex(pBytes(3)) & "," & hex(pBytes(4)) & "," & hex(pBytes(5)) & "," & hex(pBytes(6)) & "," & hex(pBytes(7)) & "," & hex(pBytes(8))
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes(i))
Next
Debug.Print "BuffByChat-Name:" & HexToStr(tmpUser)
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 1 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H34)
Dim pBytes2() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H30) + &H0, pBytes2, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes2(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 2 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H34)
Dim pBytes3() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes3, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes3(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 3 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes4() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes4, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes4(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 4 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes5() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes5, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes5(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 5 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes6() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes6, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes6(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 6 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes7() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes7, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes7(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 7 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes8() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes8, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes8(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
CureCurse (FormatHex(hex(UserID), 4))
Exit Function
End If
End If
End Function

Function ResistByChat(UserName As String)
On Error Resume Next
Dim tmpUser As String
Dim UserID
Dim NameLen As Long
Dim pBytes() As Byte
tmpUser = ""
Dim i As Integer

If PTCount > 0 Then
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H34)
Debug.Print "BuffByChat-NameLen:" & NameLen
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H30) + &H0, pBytes, NameLen
Debug.Print hex(pBytes(1)) & "," & hex(pBytes(2)) & "," & hex(pBytes(3)) & "," & hex(pBytes(4)) & "," & hex(pBytes(5)) & "," & hex(pBytes(6)) & "," & hex(pBytes(7)) & "," & hex(pBytes(8))
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes(i))
Next
Debug.Print "BuffByChat-Name:" & HexToStr(tmpUser)
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 1 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H34)
Dim pBytes2() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H30) + &H0, pBytes2, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes2(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 2 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H34)
Dim pBytes3() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes3, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes3(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 3 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes4() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes4, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes4(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 4 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes5() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes5, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes5(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 5 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes6() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes6, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes6(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 6 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes7() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes7, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes7(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 7 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes8() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes8, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes8(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
ResistUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If
End Function

Function STRByChat(UserName As String)
On Error Resume Next
Dim tmpUser As String
Dim UserID
Dim NameLen As Long
Dim pBytes() As Byte
tmpUser = ""
Dim i As Integer

If PTCount > 0 Then
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H34)
Debug.Print "BuffByChat-NameLen:" & NameLen
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H30) + &H0, pBytes, NameLen
Debug.Print hex(pBytes(1)) & "," & hex(pBytes(2)) & "," & hex(pBytes(3)) & "," & hex(pBytes(4)) & "," & hex(pBytes(5)) & "," & hex(pBytes(6)) & "," & hex(pBytes(7)) & "," & hex(pBytes(8))
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes(i))
Next
Debug.Print "BuffByChat-Name:" & HexToStr(tmpUser)
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 1 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H34)
Dim pBytes2() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H30) + &H0, pBytes2, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes2(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 2 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H34)
Dim pBytes3() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes3, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes3(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 3 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes4() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes4, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes4(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 4 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes5() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes5, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes5(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 5 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes6() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes6, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes6(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 6 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes7() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes7, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes7(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If

If PTCount > 7 Then
tmpUser = ""
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H34)
Dim pBytes8() As Byte
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H30) + &H0, pBytes8, NameLen
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes8(i))
Next
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
STRUser (FormatHex(hex(UserID), 4))
Exit Function
End If
End If
End Function

Function CalcAutoHeal(UserNo As Integer)
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

Dim Fark As Long, UserID As Long, SkillID As String
Fark = 0
UserID = 0
Select Case UserNo
Case 1
Fark = PTUserMAXHP1 - PTUserHP1
UserID = PTUserID1
Case 2
Fark = PTUserMAXHP2 - PTUserHP2
UserID = PTUserID2
Case 3
Fark = PTUserMAXHP3 - PTUserHP3
UserID = PTUserID3
Case 4
Fark = PTUserMAXHP4 - PTUserHP4
UserID = PTUserID4
Case 5
Fark = PTUserMAXHP5 - PTUserHP5
UserID = PTUserID5
Case 6
Fark = PTUserMAXHP6 - PTUserHP6
UserID = PTUserID6
Case 7
Fark = PTUserMAXHP7 - PTUserHP7
UserID = PTUserID7
Case 8
Fark = PTUserMAXHP8 - PTUserHP8
UserID = PTUserID8
Case Else
Fark = 0
UserID = 0
End Select

Select Case Fark
Case 1 To 15
SkillID = "002"
Case 16 To 30
SkillID = "005"
Case 31 To 60
SkillID = "500"
Case 61 To 240
SkillID = "509"
Case 241 To 360
SkillID = "518"
Case 361 To 720
SkillID = "527"
Case 721 To 960
SkillID = "536"
Case Is > 960
SkillID = "545"
Case Else
SkillID = ""
End Select

If Fark > 0 And UserID > 0 And SkillID <> "" Then
SendPacket "3101" & GetSkillNum(SkillID) & "00" & GetCharID & FormatHex(hex(UserID), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum(SkillID) & "00" & GetCharID & FormatHex(hex(UserID), 4) & "000000000000000000000000"
End If

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function ResistUser(UserID As String)
Dim LastNum As String
If Form2.chAutoSit.Value = 1 Then
StandUp
End If
Select Case Form2.comAutoResist.ListIndex
Case 1
LastNum = "609"
Case 2
LastNum = "627"
Case 3
LastNum = "636"
Case 4
LastNum = "645"
Case Else
Exit Function
End Select

SendPacket "3101" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum(LastNum) & "00" & GetCharID & UserID & "000000000000000000000000"

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function Sit()
SendPacket "290102"
End Function

Function StandUp()
SendPacket "290101"
End Function

Function STRUser(UserID As String)

If Form2.chAutoSit.Value = 1 Then
StandUp
End If

SendPacket "3101" & GetSkillNum("004") & "00" & GetCharID & UserID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("004") & "00" & GetCharID & UserID & "000000000000000000000000"

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function SaveGroup()
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

If Form2.chSaveParty.Value = 1 Then
If Form2.opGroupMassive.Value = True Then
SendPacket "3101" & GetSkillNum("557") & "00" & GetCharID & "FFFFDB020400D3010000000000000F00"
SendPacket "3103" & GetSkillNum("557") & "00" & GetCharID & "FFFFDB020400D301000000000000"
Else
SendPacket "3101" & GetSkillNum("560") & "00" & GetCharID & "FFFFDB020400D3010000000000000F00"
SendPacket "3103" & GetSkillNum("560") & "00" & GetCharID & "FFFFDB020400D301000000000000"
End If
End If

If Form2.chAutoSit.Value = 1 Then
Sit
End If
End Function

Function ReadString2(Addr As Long, Size As Long)
'//Function ~KoJD
On Error Resume Next
Dim bKoJD() As Byte
Dim sKoJD As String
ReadByteArray Addr, bKoJD, Size
Dim i As Integer
For i = 1 To Size
sKoJD = sKoJD & hex(bKoJD(i))
Next
sKoJD = HexToStr(sKoJD)
ReadString = sKoJD
End Function

Function LoadOffsets()
KO_TITLE = Form1.txtTitle
KO_PTR_CHR = &HC08020
KO_PTR_DLG = &HC08310
KO_PTR_PKT = &HC082DC
KO_SND_FNC = &H475390
KO_RECVHK = &HC07E94
KO_RCVHKB = &HC07E98
KO_KEYBPTR = &HC082D8
KO_SENDPTR = &HBFD7E8
KO_PARTY = &HC082FC

KO_OFF_LVL = &H9A0 'yada &H5DC
KO_OFF_CLASS = &H5D8
KO_OFF_ID = &H5B3
KO_OFF_SWIFT = &H680
KO_OFF_NT = &H584
KO_OFF_HP = &H5E4
KO_OFF_MAXHP = &H5E0
KO_OFF_MP = &H9A8
KO_OFF_MAXMP = &H9A4
KO_OFF_SIT = &HAE6
KO_OFF_WH = &H598
KO_OFF_Y = &HBC
KO_OFF_X = &HB4
KO_OFF_Z = &HB8
KO_OFF_LUP = &HC84
KO_OFF_LUP2 = &HAF0
KO_OFF_EXP = &H9C0
KO_OFF_MAXEXP = &H9B8
KO_OFF_GOLD = &H810
KO_OFF_HD = &H514
KO_OFF_MOB = &H580

KO_OFF_STAT_MP = &H9F8
KO_OFF_STAT_INT = &H9F0
KO_OFF_STAT_HP = &H9E0
KO_OFF_STAT_DEX = &H9E8
KO_OFF_STAT_STR = &H9D8
KO_OFF_AP = &HA00
KO_OFF_AC = &HA08

nation = &H5D0
x_cursor = &HCD4
y_cursor = &HCDC
End Function

Public Function ReadLong(Addr As Long) As Long 'read a 4 byte value
    Dim Value As Long
    ReadProcessMem Pencere, Addr, Value, 4, 0&
    ReadLong = Value
End Function

Public Function ReadFloat(Addr As Long) As Long 'read a float value
    Dim Value As Single
    ReadProcessMem Pencere, Addr, Value, 4, 0&
    ReadFloat = Value
End Function

Public Function WriteFloat(Addr As Long, val As Single) 'write a float value
    WriteProcessMem Pencere, Addr, val, 4, 0&
End Function

Public Function WriteLong(Addr As Long, val As Long) ' write a 4 byte value
    WriteProcessMem Pencere, Addr, val, 4, 0&
End Function
