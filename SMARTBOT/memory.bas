Attribute VB_Name = "memory"
Dim pStr As String
Dim pPtr As Long
Dim pBytes() As Byte
Dim FuckThisGame As Long
' okuma fonksiyonlarý
Public Function MobName()
pPtr = LongOku(KO_DLG)
Dim FuckThisGame As Long
FuckThisGame = LongOku(pPtr + &H1B8)
FuckThisGame = LongOku(FuckThisGame + &HD4)
FuckThisGame = LongOku(FuckThisGame + &H150)
pStr = ReadString(FuckThisGame, False, 25)
MobName = pStr
End Function
Public Function CharName()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Name)
pStr = ReadString(FuckThisGame, False, 20)
CharName = pStr
End Function
Public Function ClanName()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Clan)
pStr = ReadString(FuckThisGame, False, 20)
ClanName = pStr
End Function
Public Function PartyKisi()
Dim CureBase As Long
Dim of1a11 As Long
Dim pointer As Long
pointer = LongOku(KO_DLG)
CureBase = LongOku(pointer + &H1C8)
of1a11 = LongOku(CureBase + &H360)
PartyKisi = of1a11
End Function
Public Function HpOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_HP)
HpOku = FuckThisGame
End Function
Public Function mobhp()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Mob_HP)
mobhp = FuckThisGame
End Function
Public Function mpoku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_MP)
mpoku = FuckThisGame
End Function
Public Function MaxHPOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_MaxHP)
MaxHPOku = FuckThisGame
End Function
Public Function MaxMPOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_MaxMP)
MaxMPOku = FuckThisGame
End Function
Public Function LevelOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Level)
LevelOku = FuckThisGame
End Function
Public Function CharIDOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_ID)
CharIDOku = FuckThisGame
End Function
Public Function ExpOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Exp)
ExpOku = FuckThisGame
End Function
Public Function MaxExpOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_MaxExp)
MaxExpOku = FuckThisGame
End Function
Public Function ParaOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Para)
ParaOku = FuckThisGame
End Function
Public Function ClassOku()
FuckThisGame = LongOku(KO_ADR_CHR + KO_OFF_CLASS)
ClassOku = FuckThisGame
End Function
Public Function MobIDOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_MobID)
MobIDOku = FuckThisGame
End Function
Public Function CharXOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = ReadFloat(pPtr + KO_OFF_X)
CharXOku = FuckThisGame
End Function
Public Function CharYOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = ReadFloat(pPtr + KO_OFF_Y)
CharYOku = FuckThisGame
End Function
Public Function AtakOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Atak)
AtakOku = FuckThisGame
End Function
Public Function DefansOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Defans)
DefansOku = FuckThisGame
End Function
Public Function DexOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_SDex)
DexOku = FuckThisGame
End Function
Public Function StrOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_SStr)
StrOku = FuckThisGame
End Function
Public Function SHpOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_SHp)
SHpOku = FuckThisGame
End Function
Public Function SIntOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_SInt)
SIntOku = FuckThisGame
End Function
Public Function SMpOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_SMp)
SMpOku = FuckThisGame
End Function
Public Function NpOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_NP)
NpOku = FuckThisGame
End Function
Public Function PartyOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Party)
PartyOku = FuckThisGame
End Function
Public Function OturOku()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = LongOku(pPtr + Char_Otur)
PartyOku = FuckThisGame
End Function
Public Function HIZOKU()
pPtr = LongOku(KO_CHRDMA)
FuckThisGame = FloatOku(pPtr + Char_Hýz)
HIZOKU = FuckThisGame
End Function
'''''''''''Yazma Fonksiyonlarým'''''''''''''

Public Function NpYaz(Value As Long)
pPtr = LongOku(KO_CHRDMA)
LongYaz (pPtr + Char_NP), Value
End Function

Public Function SIntYaz(Value As Long)
pPtr = LongOku(KO_CHRDMA)
LongYaz (pPtr + Char_SInt), Value
End Function
Public Function SDexYaz(Value As Long)
pPtr = LongOku(KO_CHRDMA)
LongYaz (pPtr + Char_SDex), Value
End Function
Public Function SStrYaz(Value As Long)
pPtr = LongOku(KO_CHRDMA)
LongYaz (pPtr + Char_SStr), Value
End Function
Public Function SMpYaz(Value As Long)
pPtr = LongOku(KO_CHRDMA)
LongYaz (pPtr + Char_SMp), Value
End Function
Public Function CharHD(Value As Long)
pPtr = LongOku(KO_CHRDMA)
LongYaz (pPtr + Char_Gorunum), Value
End Function
Public Function CharIRK(Value As Long)
pPtr = LongOku(KO_CHRDMA)
LongYaz (pPtr + Char_Irk), Value
End Function
Function GetServer() As String
Dim pPtr As Long
Dim a As String
Dim ServerName() As Byte
Dim Readpacket2 As String
pPtr = LongOku(KO_Paket - &H70)
ReadByteArray pPtr, ServerName, 8
a = StrConv(ServerName, vbUnicode)
GetServer = a
End Function
Function SolElDurabiltyOku() As Long
SolElDurabiltyOku = LongOku(LongOku(LongOku(KO_DLG) + &H2D4) + &HF0)
End Function
Function SaðElDurabiltyOku() As Long
SaðElDurabiltyOku = LongOku(LongOku(LongOku(KO_DLG) + &H2D4) + &HEC)
End Function
Public Function MobHP2()
Dim CureBase As Long
Dim of1a11 As Long
Dim pointer As Long
Dim of1a12 As Long
pointer = LongOku(KO_DLG)
CureBase = LongOku(pointer + &H1B8)
of1a11 = LongOku(CureBase + &HC4)
of1a12 = LongOku(of1a11 + &HEC)
MobHP2 = of1a12
End Function
Function MobX()
Dim pPtr As Long
Dim mb1 As Long
pPtr = LongOku(KO_DLG)
mb1 = LongOku(pPtr + &H3D4)
MobX = FloatOku(mb1 + &H48)
End Function
Function MobY()
Dim pPtr As Long
Dim mb2 As Long
pPtr = LongOku(KO_DLG)
mb2 = LongOku(pPtr + &H3D4)
MobY = FloatOku(mb2 + &H50)
End Function
Function MobUzaklýK()
Dim pPtr, pPtr1, a, b, mX, mY, cx, cy As Long
Dim frkx, frky, uz As Single
pPtr1 = LongOku(KO_CHRDMA)
cx = FloatOku(pPtr1 + &HB4)
cy = FloatOku(pPtr1 + &HBC)
frkx = (MobX - cx) * (MobX - cx)
frky = (MobY - cy) * (MobY - cy)
uz = Fix(((frkx + frky) ^ 0.5) / 4)
MobUzaklýK = uz
End Function
Function SolElRpr() As String
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As String
a = ReadDoublePointer(KO_CHRDMA, &H338, &H0)
b = ReadDoublePointer(KO_CHRDMA, &H348, &H0)
c = a + b
d = AlignDWORD(c)
SolElRpr = d
End Function
Function saðelrpr() As String
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As String
a = ReadDoublePointer(KO_CHRDMA, &H33C, &H0)
b = ReadDoublePointer(KO_CHRDMA, &H34C, &H0)
c = a + b
d = AlignDWORD(c)
saðelrpr = d
End Function
Function OTORPR()
'Paket "3B01" + "06" + Form1.Label28.Caption + saðelrpr
'Paket "3B01" + "08" + Form1.Label28.Caption + saðelrpr
'Paket "3B01" + "06" + Form1.Label28.Caption + SolElRpr
'Paket "3B01"
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PetOpen()
pStr = "760105010100FD0600000800654D6F5374794C65650100002300230014001400022105"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub PetClose()
pStr = "760105020100FD060000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub


