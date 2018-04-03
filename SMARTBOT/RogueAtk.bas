Attribute VB_Name = "RogueAtk"
Dim pStr As String
Dim SK As String
Dim SK1 As String
Dim pBytes() As Byte

Public Sub Skill52()

SK1 = "552"

SK = GetSkillNum(SK1)
pStr = "3101" + SK + "00" + GetCharID + GetMobID + FormatHex(hex(GetMobX), 4) + FormatHex(hex(GetMobZ), 4) + FormatHex(hex(GetMobY), 4) + "9BFF0000F0000A00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "3102" + SK + "00" + GetCharID + GetMobID + FormatHex(hex(GetMobX), 4) + FormatHex(hex(GetMobZ), 4) + FormatHex(hex(GetMobY), 4) + "9BFF0000F0000A00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "3103" + SK + "00" + GetCharID + GetMobID + FormatHex(hex(GetMobX), 4) + FormatHex(hex(GetMobZ), 4) + FormatHex(hex(GetMobY), 4) + "9BFF0000F0000A00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub Skill57()
If Form2.List2.Text = "Shadow Shot" Then
SK1 = "557"
End If
SK = GetSkillNum(SK1)
pStr = "3101" + SK + "00" + GetCharID + GetMobID + "0B0300007F010000000000000D00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
'''
pStr = "3102" + SK + "00" + "FFFF" + GetCharID + "00D020600B7019BFF0000F0000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
''
pStr = "3103" + SK + "00" + GetCharID + GetMobID + "0B0300007F010000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub Skill60()
If Form2.List2.Text = "Shadow Hunter" Then
SK1 = "560"
End If
SK = GetSkillNum(SK1)
pStr = "3101" + SK + "00" + GetCharID + GetMobID + "0B0300007F010000000000000D00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
'''
pStr = "3102" + SK + "00" + "FFFF" + GetCharID + "00D020600B7019BFF0000F0000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
''
pStr = "3103" + SK + "00" + GetCharID + GetMobID + "0B0300007F010000000000000000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub Skill1to50()
    If Form2.List2.Text = "Archery" Then
SK1 = "003"
End If
If Form2.List2.Text = "Through Shot" Then
SK1 = "500"
End If
If Form2.List2.Text = "Fire Arrow" Then
SK1 = "505"
End If
If Form2.List2.Text = "Poison Arrow" Then
SK1 = "510"
End If
If Form2.List2.Text = "Multiple Shot" Then
SK1 = "515"
End If
If Form2.List2.Text = "Guided Arrow" Then
SK1 = "520"
End If
If Form2.List2.Text = "Perfect Shot" Then
SK1 = "525"
End If
If Form2.List2.Text = "Fire Shot" Then
SK1 = "530"
End If
If Form2.List2.Text = "Poison Shot" Then
SK1 = "535"
End If
If Form2.List2.Text = "Arc Shot" Then
SK1 = "540"
End If
If Form2.List2.Text = "Explosive Shot" Then
SK1 = "545"
End If
If Form2.List2.Text = "Arrow Shower" Then
SK1 = "555"
End If
If Form2.List2.Text = "Shadow Shot" Then
SK1 = "557"
End If
If Form2.List2.Text = "Shadow Hunter" Then
SK1 = "560"
End If
If Form2.List2.Text = "Ice Shot" Then
SK1 = "562"
End If
If Form2.List2.Text = "Lightning Shot" Then
SK1 = "566"
End If
If Form2.List2.Text = "Dark Pursuer" Then
SK1 = "570"
End If
If Form2.List2.Text = "Blow Arrow" Then
SK1 = "572"
End If
If Form2.List2.Text = "Blinding Strafe" Then
SK1 = "580"
End If
If Form2.List2.Text = "Power Shot" Then
SK1 = "585"
End If

SK = GetSkillNum(SK1)
'''
pStr = "3101" + SK + "00" + GetCharID + GetMobID + "0000000000000000000000000D00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3102" + SK + "00" + "FFFF" + GetCharID + "0D020600B7019BFF0000F0000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
pStr = "3103" + SK + "00" + GetCharID + GetMobID + "0D020600B7019BFF0000F0000F00"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub


