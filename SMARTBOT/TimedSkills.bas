Attribute VB_Name = "TimedSkills"
Option Explicit
Dim pStr As String
Dim SkillID As String
Dim tsid As String
Dim pBytes() As Byte
Dim selectedskill As String
Dim selectskill As String
Public Sub solowolf()
WolfSil
Dim pStr As String
Dim pBytes() As Byte
selectedskill = "030"
SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + "FFFF" + "FFFFB5010E005F060000000000001100"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + "FFFF" + "FFFFB5010E005F06000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub caos1000()
pStr = "3103" + "EE7A0700" + GetCharID + GetCharID + "0000000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub NewLF()
If Form2.roglist2.Selected(7) Then
selectedskill = "725"

SkillID = GetSkillNum(selectedskill)
pStr = "3106" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub
Public Sub sw()
If Form2.roglist2.Selected(8) Then
selectedskill = "010"

SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub
Public Sub warriorsw()
selectedskill = "002"
SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub PartySw1()
selectedskill = "010"
SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub Def()

If Form2.roglist2.Selected(1) Then

selectedskill = "710"

SkillID = GetSkillNum(selectedskill)
pStr = "3106" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "4901120003070000000000001100"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "490112000307000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub



Public Sub Def2()
If Form2.roglist2.Selected(2) Then
selectedskill = "730"

SkillID = GetSkillNum(selectedskill)
pStr = "3106" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "4901120003070000000000001100"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "490112000307000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub

Public Sub Def3()
If Form2.roglist2.Selected(3) Then
selectedskill = "760"

SkillID = GetSkillNum(selectedskill)
pStr = "3106" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "4901120003070000000000001100"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "490112000307000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub








Public Sub testsw()
selectedskill = "001"
SkillID = GetSkillNum(selectedskill)
'pStr = "3101" + skillid + "00" + CharID + CharID + "000000000000000000000000000000"
'ConvHEX2ByteArray pStr, pBytes
'SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub mastersc()
If Form2.roglist2.Selected(6) Then
selectedskill = "802"

SkillID = GetSkillNum(selectedskill)
pStr = "3106" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000A00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000A00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub
Public Sub CatEyes()
If Form2.roglist2.Selected(4) Then
selectedskill = "715"

SkillID = GetSkillNum(selectedskill)
pStr = "3106" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub
Public Sub Lupin()
If Form2.roglist2.Selected(5) Then
selectedskill = "735"

SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000001400"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Sub
Public Sub GizleBeni()
selectedskill = "700"
SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub sprint()

selectedskill = "002"
SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub Wolf()
    If Form2.roglist2.Selected(0) Then
WolfSil
    Dim pBytes() As Byte
    Dim pStr As String
    Dim m_SkillID As String
    Dim mXZY As String
    Dim mX As Single, mY As Single, mZ As Single

    mX = ReadFloat(ReadLong(KO_PTR_CHR) + &HB4)
    mY = ReadFloat(ReadLong(KO_PTR_CHR) + &HBC)
    mZ = ReadFloat(ReadLong(KO_PTR_CHR) + &HB8)
    

    m_SkillID = "030"

    mXZY = Left$(AlignDWORD(CLng(Fix(mX))), 4) & Left$(AlignDWORD(CLng(Fix(mZ))), 4) & Left$(AlignDWORD(CLng(Fix(mY))), 4)
    m_SkillID = AlignDWORD(CLng(ClassOku & "030"))
    
    pStr = "3101" & m_SkillID & GetCharID & "FFFF" & mXZY & "0000000000001100"
    ConvHEX2ByteArray pStr, pBytes
    SendPackets pBytes
    
    pStr = "3103" & m_SkillID & GetCharID & "FFFF" & mXZY & "000000000000"
    ConvHEX2ByteArray pStr, pBytes
    SendPackets pBytes
        End If
End Sub
Public Sub WolfSil()
    selectedskill = "030"
    SkillID = GetSkillNum(selectedskill)
    pStr = "3106" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
    ConvHEX2ByteArray pStr, pBytes
    SendPackets pBytes
End Sub
Public Sub LF()
 selectedskill = "725"
 '3103FAA101006A006A000000000000000000000000000000
'3103CDA40100760B760B0000000000000000000000000000
SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub sprint2()
 '3103FAA101006A006A000000000000000000000000000000
'3103CDA40100760B760B0000000000000000000000000000
selectedskill = "001"
SkillID = GetSkillNum(selectedskill)
pStr = "3101" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub Minor()
If ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 108 Then
SkillID = "a1a801"
ElseIf ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 107 Then
SkillID = "b9a401"
ElseIf ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 208 Then
SkillID = "412f03"
ElseIf ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 207 Then
SkillID = "592b03"
End If
pStr = "3103" + SkillID + "00" + GetCharID + GetCharID + "0000000000000000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub HpChecked()
If HpOku <= ((MaxHpOku * val(Form2.txtMinor)) / 100) Then
    Minor
End If
End Sub
Public Function HpOku()
Dim a As Long
a = ReadLong(KO_ADR_CHR + KO_OFF_HP)
HpOku = a
End Function
Public Function MaxHpOku()
Dim a As Long
a = ReadLong(KO_ADR_CHR + KO_OFF_MAXHP)
MaxHpOku = a
End Function

Public Function KrowazTS()
Dim pBase3 As String
Dim pBase2 As String
Dim pBytes() As Byte

pBase3 = "3103" + "434B" + "0700" + GetCharID + GetCharID + "00000100000064050000C800"
ConvHEX2ByteArray pBase3, pBytes
SendPackets pBytes

pBase2 = "3104" + "434B" + "0700" + GetCharID + GetCharID + "00000000000099FF00000000"
ConvHEX2ByteArray pBase2, pBytes
SendPackets pBytes
End Function

Public Function DeleteKrowazTS()
Dim pBase3 As String
Dim pBase2 As String
Dim pBytes() As Byte

pBase3 = "3106" + "434B" + "0700" + GetCharID + GetCharID + "00000100000064050000C800"
ConvHEX2ByteArray pBase3, pBytes
SendPackets pBytes

pBase2 = "3106" + "434B" + "0700" + GetCharID + GetCharID + "00000000000099FF00000000"
ConvHEX2ByteArray pBase2, pBytes
SendPackets pBytes
End Function
Public Function CaosAttack()
If GetMobID <> "FFFF" Then
    If Form5.List1.Selected(0) = True Then
    selectskill = "490219"
    End If
     If Form5.List1.Selected(1) = True Then
    selectskill = "490220"
    End If
     If Form5.List1.Selected(2) = True Then
    selectskill = "490221"
    End If
     If Form5.List1.Selected(3) = True Then
    selectskill = "490224"
    End If
     If Form5.List1.Selected(4) = True Then
    selectskill = "490226"
    End If
     If Form5.List1.Selected(5) = True Then
    selectskill = "490228"
    End If
    'If Form5.List1.Selected(6) = True Then
    'selectskill = "490232"
    'End If



Dim pStr As String, pBytes() As Byte
SkillID = FormatHex(hex(val(selectskill)), 6)
pStr = "3101" + SkillID + "00" + GetCharID + GetMobID + "0000000000000000000000000D00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3102" + SkillID + "00" + "FFFF" + GetCharID + "0D020600B7019BFF0000F0000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SkillID + "00" + GetCharID + GetMobID + "0D020600B7019BFF0000F0000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End If
End Function




