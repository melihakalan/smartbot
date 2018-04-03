Attribute VB_Name = "Module1"
Option Explicit

Private Type send_packet
    Value As String
End Type

Public packet As send_packet

Public Box_Target As Long, Box_Target_X As Single, Box_Target_Y As Single, Box_Target_Dis As Long
Public BytesAddr_MobZ As Long, BytesAddr_MobX As Long, BytesAddr_MobB As Long
Public BytesAddr6 As Long
Public BytesAddr5 As Long
Public BytesAddr4 As Long
Public BytesAddr3 As Long
Public BytesAddr2 As Long
Public ZMob As Long
Public LastPatrolStatus As Integer
Public NeededRPR As Boolean
Public onPatrol As Integer
Public Chest() As Long, ChestMob() As Long
Public ENB_HOOK As Boolean, ENB_DI8 As Boolean
Public CurTarget As String
Public WolfSlot As Long
Public ItemNum(26 To 53) As Long, ItemQua(26 To 53) As Long
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
Public SndFnc(0 To 9) As Long
Public Keyboard As Long
Public Const DIK_0 As Long = 11
Public Const DIK_1 As Long = 2
Public Const DIK_2 As Long = 3
Public Const DIK_3 As Long = 4
Public Const DIK_4 As Long = 5
Public Const DIK_5 As Long = 6
Public Const DIK_6 As Long = 7
Public Const DIK_7 As Long = 8
Public Const DIK_8 As Long = 9
Public Const DIK_9 As Long = 10
Public Const DIK_F1 As Long = &H3B
Public Const DIK_F2 As Long = 60
Public Const DIK_F3 As Long = &H3D
Public Const DIK_F4 As Long = &H3E
Public Const DIK_F5 As Long = &H3F
Public Const DIK_F6 As Long = &H40
Public Const DIK_F7 As Long = &H41
Public Const DIK_F8 As Long = &H42
Public Const DIK_Z As Long = &H2C
Public Const DIK_C As Long = &H2E
Public Const DIK_B As Long = &H30
Public Const DIK_R As Long = &H13
Public Const DIK_TAB As Long = 15
Public Const DIK_E As Long = &H12
Public Const KeybPtr As Long = &HB6FC5C

Private BytesAddr As Long
Private FuncPtr As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Const MEM_COMMIT = &H1000
Private Const MEM_RELEASE = &H8000&
Private Const PAGE_READWRITE = &H4&
Private Const INFINITE = &HFFFF
Public Const MAILSLOT_NO_MESSAGE  As Long = (-1)
Public KOPHandle As Long, KO_CHRDMA As Long, KO_WLF As Long, KOPid As Long
Public Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long 'apidir bunu ekle modulde yukarýya
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMem Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Private Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As Any) As Long
Private Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailSlot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetModuleInformation Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, lpmodinfo As MODULEINFO, ByVal cb As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Integer
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public hexword As String
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const SND_ASYNC = &H1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20

Public MSName As String
Public MSHandle

Public KO_FNCZ As Long
Public KO_FNCX As Long
Public KO_FNCB As Long
Public KO_STMB As Long
Public KO_FPOB As Long
Public KO_FPOX As Long
Public KO_FPOZ As Long
Public KO_OFF_MX As Long
Public KO_OFF_MY As Long
Public KO_OFF_MZ As Long
Public KO_OFF_MOVE As Long
Public KO_OFF_NICK As Long
Public KO_RCVFNC As Long
Public KO_TITLE As String
Public KO_HANDLE As Long
Public KO_PID As Long
Public KO_NODC As Long
Public KO_PARTY As Long
Public KO_FLDB As Long
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
Public KO_RCVHKB1 As Long
Public KO_RCVHKB2 As Long
Public KO_RCVHKB3 As Long
Public KO_RECVHK1 As Long
Public KO_RECVHK2 As Long
Public KO_RECVHK3 As Long
Public KO_ADR_CHR As Long
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
Public KO_TITLE1 As String

Public DINPUT_Handle As Long
Public DINPUT_lpBaseOfDLL As Long
Public DINPUT_SizeOfImage As Long
Public DINPUT_EntryPoint As Long
Public DINPUT_KEYDMA As Long
Public DINPUT_K_1 As Long
Public DINPUT_K_2 As Long
Public DINPUT_K_3 As Long
Public DINPUT_K_4 As Long
Public DINPUT_K_5 As Long
Public DINPUT_K_6 As Long
Public DINPUT_K_7 As Long
Public DINPUT_K_8 As Long
Public DINPUT_K_Z As Long
Public DINPUT_K_C As Long
Public DINPUT_K_S As Long
Public DINPUT_K_R As Long
Public DINPUT_K_E As Long
Public Const KO_DIKKEY As Long = &H26C
Public KO_CHR As Long
Public MP1 As Long
Public Type MODULEINFO
lpBaseOfDLL As Long
SizeOfImage As Long
EntryPoint As Long
End Type
Public Function Txt2Code(txtCode As String)
Dim Kodlar(0 To 35) As String
Dim Ayirma
Dim i, ii, ab As Integer
Dim Yazi
Dim YaziSekli As String
Kodlar(0) = "q": Kodlar(1) = "w": Kodlar(2) = "e": Kodlar(3) = "r": Kodlar(4) = "t"
Kodlar(5) = "y": Kodlar(6) = "u": Kodlar(7) = "ý": Kodlar(8) = "o": Kodlar(9) = "p": Kodlar(10) = "ð"
Kodlar(11) = "ü": Kodlar(12) = "a": Kodlar(13) = "s": Kodlar(14) = "d": Kodlar(15) = "f": Kodlar(16) = "g"
Kodlar(17) = "h": Kodlar(18) = "j": Kodlar(19) = "k": Kodlar(20) = "l": Kodlar(21) = "þ": Kodlar(22) = "i"
Kodlar(23) = "z": Kodlar(24) = "x": Kodlar(25) = "c": Kodlar(26) = "v": Kodlar(27) = "b": Kodlar(27) = "n"
Kodlar(28) = "m": Kodlar(29) = "ö": Kodlar(30) = "ç"
ii = Len(txtCode) * 2
For i = 1 To ii
If Mid(txtCode, i, 1) = "#" Then
YaziSekli = "Büyük"
GoTo gecbunu
End If
If Mid(txtCode, i, 1) = "," Then
ab = val(Ayirma)
Ayirma = ""
If YaziSekli = "Büyük" Then
YaziSekli = ""
Yazi = Yazi & UCase(Kodlar(ab))
Else
Yazi = Yazi & Kodlar(ab)
End If
Else
Ayirma = Ayirma & Mid(txtCode, i, 1)
gecbunu:
End If
Next
Txt2Code = LTrim(RTrim(Yazi))
End Function

'Kullanýþý .... yazdir = Txt2Code("#13,2,#19,12,")
Public Function ko1()
KO_TITLE1 = "Knight OnLine Client"
    GetWindowThreadProcessId FindWindow(vbNullString, KO_TITLE1), KO_PID
    KO_HANDLE = OpenProcess(&H1F0FFF, False, KO_PID)
KO_PTR_CHR1 = ReadLong(&HB6E3BC)
MP1 = 2364
End Function
' dll inject komutlarý
Public Function HookDI8() As Boolean
Dim ret As Long
Dim lmodinfo As MODULEINFO
DINPUT_Handle = 0

DINPUT_Handle = FindModuleHandle("dinput8.dll")


ret = GetModuleInformation(KO_HANDLE, DINPUT_Handle, lmodinfo, Len(lmodinfo))
If ret <> 0 Then
With lmodinfo
DINPUT_EntryPoint = .EntryPoint
DINPUT_lpBaseOfDLL = .lpBaseOfDLL
DINPUT_SizeOfImage = .SizeOfImage
End With
Else
Exit Function
End If
SetupDInput
HookDI8 = True
End Function

Public Function FindModuleHandle(ModuleName As String) As Long
Dim hModules(1 To 256) As Long
Dim BytesReturned As Long
Dim ModuleNumber As Byte
Dim TotalModules As Byte
Dim FileName As String * 128
Dim ModName As String
EnumProcessModules KO_HANDLE, hModules(1), 1024, BytesReturned
TotalModules = BytesReturned / 4
For ModuleNumber = 1 To TotalModules
GetModuleFileNameExA KO_HANDLE, hModules(ModuleNumber), FileName, 128
ModName = Left(FileName, InStr(FileName, Chr(0)) - 1)
If UCase(Right(ModName, Len(ModuleName))) = UCase(ModuleName) Then
FindModuleHandle = hModules(ModuleNumber)
End If
Next
End Function

Sub SetupDInput()
DINPUT_KEYDMA = FindDInputKeyPtr
If DINPUT_KEYDMA <> 0 Then
DINPUT_K_1 = DINPUT_KEYDMA + 2
DINPUT_K_2 = DINPUT_KEYDMA + 3
DINPUT_K_3 = DINPUT_KEYDMA + 4
DINPUT_K_4 = DINPUT_KEYDMA + 5
DINPUT_K_5 = DINPUT_KEYDMA + 6
DINPUT_K_6 = DINPUT_KEYDMA + 7
DINPUT_K_7 = DINPUT_KEYDMA + 8
DINPUT_K_8 = DINPUT_KEYDMA + 9
DINPUT_K_Z = DINPUT_KEYDMA + 44
DINPUT_K_C = DINPUT_KEYDMA + 46
DINPUT_K_S = DINPUT_KEYDMA + 31
DINPUT_K_R = DINPUT_KEYDMA + 19
DINPUT_K_E = DINPUT_KEYDMA + 18
End If
End Sub

Function FindDInputKeyPtr() As Long
Dim pBytes() As Byte
Dim pSize As Long
Dim X As Long
pSize = DINPUT_SizeOfImage
ReDim pBytes(1 To pSize)
ReadByteArray DINPUT_lpBaseOfDLL, pBytes, pSize
For X = 1 To pSize - 10
If pBytes(X) = &H57 And pBytes(X + 1) = &H6A And pBytes(X + 2) = &H40 And pBytes(X + 3) = &H33 And pBytes(X + 4) = &HC0 And pBytes(X + 5) = &H59 And pBytes(X + 6) = &HBF Then
FindDInputKeyPtr = val("&H" & IIf(Len(hex(pBytes(X + 10))) = 1, "0" & hex(pBytes(X + 10)), hex(pBytes(X + 10))) & IIf(Len(hex(pBytes(X + 9))) = 1, "0" & hex(pBytes(X + 9)), hex(pBytes(X + 9))) & IIf(Len(hex(pBytes(X + 8))) = 1, "0" & hex(pBytes(X + 8)), hex(pBytes(X + 8))) & IIf(Len(hex(pBytes(X + 7))) = 1, "0" & hex(pBytes(X + 7)), hex(pBytes(X + 7))))
Exit For
End If
Next
End Function
' Buraya ben yolla yazdým sizde istediðinizi yaza bilir siniz.
'ama prejedeki Bütün Yolla yazan yerleri deðiþtirmelisiniz.
Function yolla(pKey As String) As Long
pKey = Strings.UCase(pKey)
Select Case pKey
Case "S"
yolla = DINPUT_K_S
Case "Z"
yolla = DINPUT_K_Z
Case "1"
yolla = DINPUT_K_1
Case "2"
yolla = DINPUT_K_2
Case "3"
yolla = DINPUT_K_3
Case "4"
yolla = DINPUT_K_4
Case "5"
yolla = DINPUT_K_5
Case "6"
yolla = DINPUT_K_6
Case "7"
yolla = DINPUT_K_7
Case "8"
yolla = DINPUT_K_8
Case "C"
yolla = DINPUT_K_C
Case "R"
yolla = DINPUT_K_R
Case "E"
yolla = DINPUT_K_E
End Select
End Function
Sub WriteByte(Addr As Long, pVal As Byte)
Dim pbw As Long
WriteProcessMem KO_HANDLE, Addr, pVal, 1, pbw
End Sub

Sub ReadByteArray(Addr As Long, pmem() As Byte, pSize As Long)
Dim Value As Byte
ReDim pmem(1 To pSize) As Byte
ReadProcessMem KO_HANDLE, Addr, pmem(1), pSize, 0&
End Sub
' Buraya ben TUS yazdým sizde istediðinizi yaza bilir siniz.
'ama prejedeki Bütün TUS yazan yerleri deðiþtirmelisiniz.
Sub Tuþ(pKey As Long, Optional pTimeMS As Long = 50)
WriteByte pKey, 128
f_Sleep pTimeMS, True
WriteByte pKey, 0
End Sub

Sub f_Sleep(pMS As Long, Optional pDoevents As Boolean = False)
Dim pTime As Long
pTime = GetTickCount
Do While pMS + pTime > GetTickCount
If pDoevents = True Then DoEvents
Loop
End Sub




Function ReadByte(pAddy As Long, Optional pHandle As Long) As Byte
    Dim Value As Byte
    If pHandle <> 0 Then
        ReadProcessMem pHandle, pAddy, Value, 1, 0&
    Else
        ReadProcessMem KO_HANDLE, pAddy, Value, 1, 0&
    End If
    ReadByte = Value
End Function
Public Function ReadLong(Addr As Long) As Long 'read a 4 byte value
    Dim Value As Long
    ReadProcessMem KO_HANDLE, Addr, Value, 4, 0&
    ReadLong = Value
End Function

Public Function ReadFloat(Addr As Long) As Long 'read a float value
    Dim Value As Single
    ReadProcessMem KO_HANDLE, Addr, Value, 4, 0&
    ReadFloat = Value
End Function

Public Function WriteFloat(Addr As Long, val As Single) 'write a float value
    WriteProcessMem KO_HANDLE, Addr, val, 4, 0&
End Function

Public Function WriteLong(Addr As Long, val As Long) ' write a 4 byte value
    WriteProcessMem KO_HANDLE, Addr, val, 4, 0&
End Function

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Else
        SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        SetTopMostWindow = False
    End If
End Function

Public Function ReadIni(FileName As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, "", RetVal, 255, FileName)
ReadIni = Left(RetVal, v)
End Function

'writes an Ini string
Public Function WriteIni(FileName As String, Section As String, Key As String, Value As String)
WritePrivateProfileString Section, Key, Value, FileName
End Function
Public Function ConvHEX2ByteArray(pStr As String, pByte() As Byte)
On Error Resume Next
Dim i As Long
Dim j As Long
ReDim pByte(1 To Len(pStr) / 2)

j = LBound(pByte) - 1
For i = 1 To Len(pStr) Step 2
    j = j + 1
    pByte(j) = CByte("&H" & Mid(pStr, i, 2))
Next
End Function
Public Function WriteByteArray(pA As Long, pS() As Byte, pSize As Long)
    WriteProcessMem KO_HANDLE, pA, pS(LBound(pS)), pSize, 0&
End Function
'start packet handling
Function ExecuteRemoteCode(pCode() As Byte, Optional WaitExecution As Boolean = False) As Long
Dim hthread As Long, ThreadID As Long, ret As Long
Dim SE As SECURITY_ATTRIBUTES

SE.nLength = Len(SE)
SE.bInheritHandle = False

ExecuteRemoteCode = 0
If FuncPtr = 0 Then
FuncPtr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If
If FuncPtr <> 0 Then
    WriteByteArray FuncPtr, pCode, UBound(pCode) - LBound(pCode) + 1
   'MsgBox hex(FuncPtr)
   hthread = CreateRemoteThread(ByVal KO_HANDLE, SE, 0, ByVal FuncPtr, 0&, 0&, ThreadID)
   If hthread Then
      ret = WaitForSingleObject(hthread, INFINITE)
      ExecuteRemoteCode = ThreadID
   End If
   CloseHandle hthread
   'MsgBox hex(FuncPtr)
   ret = VirtualFreeEx(KO_HANDLE, FuncPtr, 0, MEM_RELEASE)
    
End If

End Function
Function SendPackets(pPacket() As Byte)
Dim pSize As Long
Dim pCode() As Byte

pSize = UBound(pPacket) - LBound(pPacket) + 1
If BytesAddr = 0 Then
BytesAddr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If
If BytesAddr <> 0 Then
    WriteByteArray BytesAddr, pPacket, pSize
    'MsgBox hex(BytesAddr)
        ConvHEX2ByteArray "608B0D" & AlignDWORD(KO_PTR_PKT) & "68" & AlignDWORD(pSize) & "68" & AlignDWORD(BytesAddr) & "BF" & AlignDWORD(KO_SND_FNC) & "FFD7C605" & AlignDWORD(KO_PTR_PKT + &HC1) & "0061C3", pCode 'ko_sendfix = ko_ptr_pkt+&HC0
    ExecuteRemoteCode pCode, True
End If
VirtualFreeEx KO_HANDLE, BytesAddr, 0, MEM_RELEASE&
End Function
':D alttaymýs zate:d
Function AlignDWORD(pParam As Long) As String
Dim HiW As Integer
Dim LoW As Integer

Dim HiBHiW As Byte
Dim HiBLoW As Byte

Dim LoBHiW As Byte
Dim LoBLoW As Byte

HiW = HiWord(pParam)
LoW = LoWord(pParam)

HiBHiW = HiByte(HiW)
HiBLoW = HiByte(LoW)

LoBHiW = LoByte(HiW)
LoBLoW = LoByte(LoW)

AlignDWORD = IIf(Len(hex(LoBLoW)) = 1, "0" & hex(LoBLoW), hex(LoBLoW)) & _
         IIf(Len(hex(HiBLoW)) = 1, "0" & hex(HiBLoW), hex(HiBLoW)) & _
         IIf(Len(hex(LoBHiW)) = 1, "0" & hex(LoBHiW), hex(LoBHiW)) & _
         IIf(Len(hex(HiBHiW)) = 1, "0" & hex(HiBHiW), hex(HiBHiW))
End Function 'sansa bak 8 miþ zate:D1:D
Function AlignDWORD8(pParam As Long) As String
Dim HiW As Integer
Dim LoW As Integer

Dim HiBHiW As Byte
Dim HiBLoW As Byte

Dim LoBHiW As Byte
Dim LoBLoW As Byte

HiW = HiWord(pParam)
LoW = LoWord(pParam)

HiBHiW = HiByte(HiW)
HiBLoW = HiByte(LoW)

LoBHiW = LoByte(HiW)
LoBLoW = LoByte(LoW)

AlignDWORD8 = IIf(Len(hex(LoBLoW)) = 1, "0" & hex(LoBLoW), hex(LoBLoW)) & _
         IIf(Len(hex(HiBLoW)) = 1, "0" & hex(HiBLoW), hex(HiBLoW)) & _
         IIf(Len(hex(LoBHiW)) = 1, "0" & hex(LoBHiW), hex(LoBHiW)) & _
         IIf(Len(hex(HiBHiW)) = 1, "0" & hex(HiBHiW), hex(HiBHiW)) & _
         IIf(Len(hex(LoBLoW)) = 1, "0" & hex(LoBLoW), hex(LoBLoW)) & _
         IIf(Len(hex(HiBLoW)) = 1, "0" & hex(HiBLoW), hex(HiBLoW)) & _
         IIf(Len(hex(LoBHiW)) = 1, "0" & hex(LoBHiW), hex(LoBHiW)) & _
         IIf(Len(hex(HiBHiW)) = 1, "0" & hex(HiBHiW), hex(HiBHiW))
End Function

Public Function HiByte(ByVal wParam As Integer) As Byte

    HiByte = (wParam And &HFF00&) \ (&H100)

End Function

Public Function LoByte(ByVal wParam As Integer) As Byte

LoByte = wParam And &HFF&

End Function

Function LoWord(DWord As Long) As Integer
   If DWord And &H8000& Then '
      LoWord = DWord Or &HFFFF0000
   Else
      LoWord = DWord And &HFFFF&
   End If
End Function

Function HiWord(DWord As Long) As Integer
   HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function LogFile(sStr As String)
Open CurDir & "\Log.txt" For Append As #1
Print #1, ";==> " & Time & " " & sStr & " ;<=="
Close #1
End Function

Public Function AttachKO() As Boolean
If FindWindow(vbNullString, KO_TITLE) Then
GetWindowThreadProcessId FindWindow(vbNullString, KO_TITLE), KO_PID
KO_HANDLE = OpenProcess(PROCESS_ALL_ACCESS, False, KO_PID)
If KO_HANDLE = 0 Then
MsgBox ("Cannot get handle from KO(" & KO_PID & ").")
AttachKO = False
End If
If ENB_HOOK = True Then AutoHook
'If ENB_DI8 = True Then HookDI8
AttachKO = True
Else
If Form1.optTr.Value = True Then MsgBox ("Pencere bulunamadý."), vbCritical, "Hata" Else MsgBox ("Could not find the window."), vbCritical, "Error"
End If
End Function

Public Function EstablishMailSlot(ByVal MailSlotName As String, Optional MaxMessageSize As Long = 0, Optional ReadTimeOut As Long = 50) As Long
EstablishMailSlot = CreateMailslot(MailSlotName, MaxMessageSize, ReadTimeOut, ByVal 0&)
End Function

Function LoadOffsets()
KO_TITLE = Form1.txtTitle

'Pointers/
KO_PTR_CHR = &HC1DAA0
KO_PTR_DLG = &HC1DD94
KO_PTR_PKT = &HC1DD60
KO_SND_FNC = &H475060
KO_RECVHK = &HC1D914
KO_RCVHKB = &HC1D918
KO_KEYBPTR = &HC1DD5C
KO_SENDPTR = &HC131F8
KO_PARTY = &HC1DD80
KO_NODC = &HC16138
KO_FLDB = &HC1DA9C
KO_FPOZ = &H4A7A30
KO_FPOX = &H5C93F0
KO_FPOB = &H4A6B80
KO_STMB = &H81F540
KO_FNCZ = &H81EA70
KO_FNCX = &H81EB80
KO_FNCB = &H81EC80
'\

'AutoHook/
KO_RCVHKB1 = &H7F4130
KO_RCVHKB2 = &H7F6940
KO_RCVHKB3 = &H7F9080
KO_RECVHK1 = &H9BDEDC
KO_RECVHK2 = &H9BDEE0
KO_RECVHK3 = &H9BDEE4
'\
KO_OFF_MX = &HD44
KO_OFF_MY = &HD4C
KO_OFF_MZ = &HD48
KO_OFF_MOVE = &HD38

KO_OFF_NICK = &H5BC
KO_OFF_LVL = &H5DC
KO_OFF_CLASS = &H5D8
KO_OFF_ID = &H5B4
KO_OFF_SWIFT = &H680
KO_OFF_NT = &H584
KO_OFF_HP = &H5E4
KO_OFF_MAXHP = &H5E0
KO_OFF_MP = &H9A8
KO_OFF_MAXMP = &H9A4
KO_OFF_SIT = &HAE6
KO_OFF_WH = &H5E8
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
End Function

Function FindDLLFunc(pDLLName As String, pFuncName As String) As Long
Dim LoadAddr As Long
Dim ProcAddr As Long
Dim offset As Long
Dim RemoteAddr As Long

LoadAddr = LoadLibrary(pDLLName)
If LoadAddr = 0 Then End
ProcAddr = GetProcAddress(LoadAddr, pFuncName)
offset = ProcAddr - LoadAddr
FreeLibrary LoadAddr

RemoteAddr = FindModuleHandle(pDLLName)
Do While RemoteAddr = 0
    RemoteAddr = FindModuleHandle(pDLLName)
    DoEvents
Loop
FindDLLFunc = RemoteAddr + offset
End Function
Sub HookRecvYeni()
        'kojd//
        Dim CreateFileAADDR As Long, WriteFileADDR As Long, CloseHandleADDR As Long
        Dim pBytesMSName() As Byte, pBytes() As Byte
        Dim pStr As String, pStrKO_RECVFNC As String

        CreateFileAADDR = FindDLLFunc("kernel32.dll", "CreateFileA") 'Kernel: dosya oluþturma adresi.
        WriteFileADDR = FindDLLFunc("kernel32.dll", "WriteFile") 'Kernel: dosyaya yazma adresi.
        CloseHandleADDR = FindDLLFunc("kernel32.dll", "CloseHandle") 'Kernel: dosyayý kapatma adresi.

        KO_RCVFNC = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
        '<-- Hafýzada yer aç (yer açýlan adres: KO_RCVFNC)
        
        ConvHEX2ByteArray HexString2(MSName), pBytesMSName 'Önceki 1 byte kaçýrýyordu bakalým böyle ne olacak.
        WriteByteArray KO_RCVFNC + &H400, pBytesMSName, UBound(pBytesMSName) - LBound(pBytesMSName) + 1
        '<-- Mailslot dizinini yaz.
        
        pStr = AlignDWORD(CreateFileAADDR)
        ConvHEX2ByteArray pStr, pBytes
        WriteByteArray KO_RCVFNC + &H32A, pBytes, UBound(pBytes) - LBound(pBytes) + 1
        '<-- Mailslot'a dosya oluþturma rutini.

        pStr = AlignDWORD(WriteFileADDR)
        ConvHEX2ByteArray pStr, pBytes
        WriteByteArray KO_RCVFNC + &H334, pBytes, UBound(pBytes) - LBound(pBytes) + 1
        '<-- Mailslot'taki dosyaya yazma rutini.

        pStr = AlignDWORD(CloseHandleADDR)
        ConvHEX2ByteArray pStr, pBytes
        WriteByteArray KO_RCVFNC + &H33E, pBytes, UBound(pBytes) - LBound(pBytes) + 1
        '<-- Mailslot'taki dosyayý kapatma rutini.
        
        pStr = AlignDWORD(KO_RCVHKB)
        ConvHEX2ByteArray pStr, pBytes
        WriteByteArray KO_RCVFNC + &H208, pBytes, UBound(pBytes) - LBound(pBytes) + 1
        
        pStr = AlignDWORD(KO_RECVHK)
        ConvHEX2ByteArray pStr, pBytes
        WriteByteArray KO_RCVFNC + &H212, pBytes, UBound(pBytes) - LBound(pBytes) + 1

        pStr = AlignDWORD(KO_RCVFNC)
        ConvHEX2ByteArray pStr, pBytes
        WriteByteArray KO_RCVFNC + &H21C, pBytes, UBound(pBytes) - LBound(pBytes) + 1
        
        pStr = "52" + "890D" + AlignDWORD(KO_RCVFNC + &H320) + "8905" + AlignDWORD(KO_RCVFNC + &H3B6) + "8B4E04890d" + AlignDWORD(KO_RCVFNC + &H1F4) + "8B56088915" + AlignDWORD(KO_RCVFNC + &H1FE) + "81F9001000007D3E5068800000006A036A006A01680000004068" + AlignDWORD(KO_RCVFNC + &H400) + "FF15" + AlignDWORD(KO_RCVFNC + &H32A) + "83F8FF741D506A0054FF35" + AlignDWORD(KO_RCVFNC + &H1F4) + "ff35" + AlignDWORD(KO_RCVFNC + &H1FE) + "50ff15" + AlignDWORD(KO_RCVFNC + &H334) + "ff15" + AlignDWORD(KO_RCVFNC + &H33E) + "8b0d" + AlignDWORD(KO_RCVFNC + &H320) + "8b05" + AlignDWORD(KO_RCVFNC + &H3B6) + "5aff25" + AlignDWORD(KO_RCVFNC + &H208)
        ConvHEX2ByteArray pStr, pBytes
        WriteByteArray KO_RCVFNC, pBytes, UBound(pBytes) - LBound(pBytes) + 1
        '<-- Gelen paketleri mailslot'a yollama rutini.
        
        pStrKO_RECVFNC = AlignDWORD(KO_RCVFNC)
        ConvHEX2ByteArray pStrKO_RECVFNC, pBytes
        WriteByteArray KO_RECVHK, pBytes, UBound(pBytes) - LBound(pBytes) + 1
        '<-- Paket gelince KO_RCVFNC (Hook) adresine gitme rutini.
        MsgBox hex(KO_RCVFNC)
End Sub
Public Sub AutoHook()
       Dim HookFix As Long
       
       '/Update 30/03/10 - kojd - multi rpr fix
       Dim a, b, c, d As Integer
        Randomize 'her seferinde zar atmamýz gerekiyor unutmuþuz :P
        a = CInt(Rnd * 9)
        Randomize
        b = CInt(Rnd * 9)
        Randomize
        c = CInt(Rnd * 9)
        Randomize
        d = CInt(Rnd * 9)
        Randomize
        
        'mailslot: KoJD_"PID Son 2 hanesi"_4 random sayý
        MSName = "\\.\mailslot\KoJD_" & Right(App.ThreadID, 2) & "_" & a & b & c & d & CInt(Rnd * 9999)
        Debug.Print MSName
        MSHandle = EstablishMailSlot(MSName)
        
        HookFix = KO_PTR_DLG + &H84
        'MsgBox ReadByte(HookFix)
        
        Select Case ReadByte(HookFix)
            Case 8
                KO_RCVHKB = KO_RCVHKB1
                KO_RECVHK = KO_RECVHK1
                HookRecvYeni
            Case 9
                KO_RCVHKB = KO_RCVHKB2
                KO_RECVHK = KO_RECVHK2
                HookRecvYeni
            Case 10
                KO_RCVHKB = KO_RCVHKB3
                KO_RECVHK = KO_RECVHK3
                HookRecvYeni
            Case Else
                KO_RCVHKB = KO_RCVHKB1
                KO_RECVHK = KO_RECVHK1
                HookRecvYeni
        End Select
        
    End Sub
Public Function HexVal(pStrHex As String) As Long
Dim TmpStr As String
Dim TmpHex As String
Dim hexcode As String
Dim i As Long
TmpStr = ""
For i = Len(pStrHex) To 1 Step -1
    TmpHex = hex(Asc(Mid(pStrHex, i, 1)))
    If Len(TmpHex) = 1 Then TmpHex = "0" & TmpHex
    TmpStr = TmpStr & TmpHex
Next
hexcode = TmpStr

End Function

Private Function ReadMessage(MailMessage As String, MessagesLeft As Long)
Dim lBytesRead As Long
Dim lNextMsgSize As Long
Dim lpBuffer     As String
ReadMessage = False
Call GetMailslotInfo(MSHandle, ByVal 0&, lNextMsgSize, MessagesLeft, ByVal 0&)
If MessagesLeft > 0 And lNextMsgSize <> MAILSLOT_NO_MESSAGE Then
    lBytesRead = 0
    lpBuffer = String$(lNextMsgSize, Chr$(0))
    Call ReadFile(MSHandle, ByVal lpBuffer, Len(lpBuffer), lBytesRead, ByVal 0&)
    If lBytesRead <> 0 Then
        MailMessage = Left(lpBuffer, lBytesRead)
        ReadMessage = True
        Call GetMailslotInfo(MSHandle, ByVal 0&, lNextMsgSize, MessagesLeft, ByVal 0&)
    End If
End If
End Function

Private Function CheckForMessages(MessageCount As Long)
Dim lMsgCount    As Long
Dim lNextMsgSize As Long
CheckForMessages = False
GetMailslotInfo MSHandle, ByVal 0&, lNextMsgSize, lMsgCount, ByVal 0&
MessageCount = lMsgCount
CheckForMessages = True
End Function
Public Function HexString(EvalString As String) As String
Dim intStrLen As Integer
Dim intLoop As Integer
Dim strHex As String

EvalString = Trim(EvalString)
intStrLen = Len(EvalString)
For intLoop = 1 To intStrLen
strHex = strHex & hex(Asc(Mid(EvalString, intLoop, 1)))
Next
hexword = strHex
End Function


Public Function Hex2Val(pStrHex As String) As Long
Dim TmpStr As String
Dim TmpHex As String
Dim i As Long
TmpStr = ""
For i = Len(pStrHex) To 1 Step -1
    TmpHex = hex(Asc(Mid(pStrHex, i, 1)))
    If Len(TmpHex) = 1 Then TmpHex = "0" & TmpHex
    TmpStr = TmpStr & TmpHex
Next
Hex2Val = CLng("&H" & TmpStr)
End Function
Public Function ReadLong1(Addr As Long) As Long '1 byte lýk deðer okur
    Dim Value As Long
    ReadProcessMem KO_HANDLE, Addr, Value, 1, 0&
    ReadLong1 = Value
End Function

'Public Sub MaxMana() 'The manasave
'Dim pPtr As Long
'pPtr = ReadLong(KO_CHRDMA)
'WriteLong pPtr + 2360, totalMP
'End Sub
'
Public Sub SetupKeyboard()
Keyboard = ReadLong(KO_KEYBPTR + 0)
Keyboard = Keyboard + 620
End Sub

Public Sub SendK(ByVal Key As Long)
WriteLong Keyboard + Key * 4, 1
Sleep (50)
WriteLong Keyboard + Key * 4, 0
End Sub

'Burda aslýnda KnightOnline ustemi deðilmi onu kontrol etmek lazým
'Eðer deðilse SendK yý kullansýn, Ama aktifse Knight o zaman Yolla Tuþ("Z") olabilir.
'ben size yolun nasýl gideceðini söylerim sizlerde nasýl yapýlýr onu araþtýrýp uygularsýnýz :D
Function PM(ID As String)
Dim pStr As String
Dim pBytes() As Byte
pStr = "35010" & "7" & "00" & hexword
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Function
Function Sendkeys(Keys As Long)
Dim a, b, c As Long
Dim kb As Long
Dim kb2 As Long
kb = ReadLong(KO_PTR_PKT - 4)
a = kb + KO_DIKKEY
WriteLong a + Keys * 4, 1
Sleep (50)
WriteLong a + Keys * 4, 0
End Function
Sub DInputSendKeys(pKey As Long, Optional pTimeMS As Long = 50)
WriteByte pKey, 128
f_Sleep pTimeMS, True
WriteByte pKey, 0
End Sub

Public Function TersCevir(KCode As String)
Dim msgBuffer As String
msgBuffer = KCode
Dim i As Integer
Dim Sonuc As String
Dim TekCift As String
TekCift = Len(msgBuffer) / 2
Dim a As Integer
a = InStr(1, TekCift, ",")

If a = 1 Then
    Dim TekEkle As String
    TekEkle = Mid(msgBuffer, 1, 1)
    msgBuffer = Mid(msgBuffer, 2)
End If
For i = (Len(msgBuffer) - 1) To 1 Step -2
    Sonuc = Sonuc & Mid(msgBuffer, i, 2)
Next i
TersCevir = Sonuc & TekEkle
End Function

Sub DispatchMailSlot()
Dim MsgCount As Long
Dim rc As Long
Dim MessageBuffer As String
Dim pVal As Long
Dim fullcode
Dim code
Dim sKey
MsgCount = 1
Do While MsgCount <> 0
rc = CheckForMessages(MsgCount)
If CBool(rc) And MsgCount > 0 Then
If ReadMessage(MessageBuffer, MsgCount) Then
Call HexVal(MessageBuffer)
code = MessageBuffer
On Error Resume Next
'Debug.Print Asc(Left(MessageBuffer, 1))
Select Case Asc(Left(MessageBuffer, 1))

Case &H23 '//Kutu düþtü
If Form2.chAutoLoot.Value = 1 Then
If Form2.chRuntochest.Value = 1 Then

Dim z As Integer
For z = 0 To UBound(Chest) - LBound(Chest) + 1
If Chest(z) = 0 Then
Chest(z) = Hex2Val(Mid(MessageBuffer, 4, 4))
ChestMob(z) = Hex2Val(Mid(MessageBuffer, 2, 2))
Chest_Exec ChestMob(z), Chest(z)
End If
Next

Else

Box_Target = Hex2Val(Mid(MessageBuffer, 2, 2))
Box_Target_X = ReadFloat(GetMobBase(Box_Target) + &HB4)
Box_Target_Y = ReadFloat(GetMobBase(Box_Target) + &HBC)
Box_Target_Dis = GetMyDistance(Box_Target_X, Box_Target_Y, Box_Target)

BoxIDNew = Hex2Val(Mid(MessageBuffer, 4, 4))
BoxIDNew = hex(BoxIDNew)
BoxIDNew = FormatHex(BoxIDNew, 8)
OpenBox
End If
End If

Case &H24 '//Kutu açýldý
If Form2.chAutoLoot.Value = 1 Then
If Box_Target_Dis <= 3 Then
DropItem1 = hex(Hex2Val(Mid(MessageBuffer, 2, 4)))
DropItem1 = FormatHex(DropItem1, 8)
DropItem2 = hex(Hex2Val(Mid(MessageBuffer, 8, 4)))
DropItem2 = FormatHex(DropItem2, 8)
DropItem3 = hex(Hex2Val(Mid(MessageBuffer, 14, 4)))
DropItem3 = FormatHex(DropItem3, 8)
DropItem4 = hex(Hex2Val(Mid(MessageBuffer, 20, 4)))
DropItem4 = FormatHex(DropItem4, 8)
LootBox
End If
End If

Case &H2F '//Party isteði
If Form2.chAutoParty.Value = 1 Then
If Hex2Val(Mid(MessageBuffer, 2, 1)) = 2 Then
SendPacket "2F0201"
End If
Debug.Print "Got Party Request"
End If

Case &H16 '//Char bilgilerini al (sað,sol silah)
If Form2.chRepair.Value = 1 Or TestRpr = True Or Form2.tmRpr.Enabled = True Then
Dim NameLen As Integer, ClanLen As Integer, ID As Long
ID = Hex2Val(Mid(MessageBuffer, 5, 2))
If FormatHex(hex(ID), 4) = GetCharID Then
NameLen = Hex2Val(Mid(MessageBuffer, 7, 1))
ClanLen = Hex2Val(Mid(MessageBuffer, 15 + NameLen, 1))
RightItem = Hex2Val(Mid(MessageBuffer, 100 + NameLen + ClanLen, 4))
durRightItem = Hex2Val(Mid(MessageBuffer, 104 + NameLen + ClanLen, 2))
LeftItem = Hex2Val(Mid(MessageBuffer, 108 + NameLen + ClanLen, 4))
durLeftItem = Hex2Val(Mid(MessageBuffer, 112 + NameLen + ClanLen, 2))
Debug.Print "Right Item: " & RightItem & ", Dur: " & durRightItem & ", Left Item: " & LeftItem & ", Dur: " & durLeftItem
Form2.CheckRPR
'TestRpr = False
End If
End If

Case &H22 '//Hedef HP
Form2.lbTarHP.Caption = Hex2Val(Mid(MessageBuffer, 9, 3)) & "/" & Hex2Val(Mid(MessageBuffer, 5, 3)) & " (%" & GetMobHP & ")"

Case &H1D '//Slot Al
If GetSlot = 1 Then
If tmpMob = Hex2Val(Mid(MessageBuffer, 4, 2)) Then
sSid = Hex2Val(Mid(MessageBuffer, 6, 2))
Lvl = Hex2Val(Mid(MessageBuffer, 27, 1))
Form2.txtsSid = sSid
Form2.txtSlotLVL = Lvl
End If
If sSid = Hex2Val(Mid(MessageBuffer, 6, 2)) And sSid > 0 Then Form2.lstSlot.AddItem FormatHex(hex(Hex2Val(Mid(MessageBuffer, 4, 2))), 4)
Form2.txtMobCount = Form2.lstSlot.ListCount
End If

Case &H7 '//USERINOUT - Slot giriþ çýkýþ kontrol
If Form2.chGMTown = 1 Or Form2.chGMDC.Value = 1 Or Form2.chGMOffKO.Value = 1 Or Form2.chGMOffPC.Value = 1 Or Form2.lbNotifyPT.Value = 1 Then
If Hex2Val(Mid(MessageBuffer, 2, 1)) <> 2 Then
Dim NameLen2 As Integer, ClanLen2 As Integer
NameLen2 = Hex2Val(Mid(MessageBuffer, 6, 1))
ClanLen2 = Hex2Val(Mid(MessageBuffer, 14 + NameLen2, 1))
UserInAuth = Hex2Val(Mid(MessageBuffer, 47 + NameLen2 + ClanLen2, 2))
GMID = Mid(MessageBuffer, 7, NameLen2)
Debug.Print "GMPROTECTION:USERIN:AUTH:" & UserInAuth & ":ID:" & GMID
If UserInAuth = 0 Then Form2.OnGmDetect
End If
End If


If Form2.chSlotLog.Value = 1 Then '//Slot log
Dim InOutType As Integer, strTypeTR As String, strTypeEN As String, NameLen4 As Integer, ClanLen4 As Integer, UserID As Long, UserNick As String, UserJob As Integer, UserLevel As Integer, Clan As String, UserNation As Integer, strUserJob As String, strUserNation As String
InOutType = Hex2Val(Mid(MessageBuffer, 2, 1))
UserID = Hex2Val(Mid(MessageBuffer, 4, 2))
If InOutType <> 2 Then
NameLen4 = Hex2Val(Mid(MessageBuffer, 6, 1))
ClanLen4 = Hex2Val(Mid(MessageBuffer, 14 + NameLen4, 1))
UserNick = Mid(MessageBuffer, 7, NameLen4)
UserNation = Hex2Val(Mid(MessageBuffer, 7 + NameLen4, 1))
Clan = Mid(MessageBuffer, 15 + NameLen4, ClanLen4)
UserLevel = Hex2Val(Mid(MessageBuffer, 26 + NameLen4 + ClanLen4, 1))
UserJob = Hex2Val(Mid(MessageBuffer, 28 + NameLen4 + ClanLen4, 1))
End If

Select Case InOutType
Case 1
strTypeTR = "Slota Girdi"
strTypeEN = "Entered Slot"
Case 2
strTypeTR = "Slottan Ayrýldý"
strTypeEN = "Left Slot"
Case Else
strTypeTR = "Slotta Doðdu"
strTypeEN = "Spawned in Slot"
End Select

If InOutType <> 2 Then
Select Case UserNation
Case 1: strUserNation = "Karus"
Case 2: strUserNation = "Human"
End Select
End If

If InOutType <> 2 Then
Select Case UserJob
Case 101, 105, 106, 201, 205, 206: strUserJob = "Warrior"
Case 102, 107, 109, 202, 207, 209: strUserJob = "Rogue"
Case 103, 109, 110, 203, 209, 210: strUserJob = "Mage"
Case 104, 111, 112, 204, 211, 212: strUserJob = "Priest"
End Select
End If

If Clan = "" Then Clan = "-"

If Form1.optTr.Value = True Then
If InOutType <> 2 Then
Form2.LogOut "user.log", "Kullanýcý " & strTypeTR & ": " & UserID & " - " & UserNick & " (" & strUserNation & ") (" & strUserJob & ") (Lvl: " & UserLevel & ") (Clan: " & Clan & ")"
Debug.Print "Kullanýcý " & strTypeTR & ": " & UserID & " - " & UserNick & " (" & strUserNation & ") (" & strUserJob & ") (Lvl: " & UserLevel & ") (Clan: " & Clan & ")"
Else
Form2.LogOut "user.log", "Kullanýcý " & strTypeTR & ": " & UserID
Debug.Print "Kullanýcý " & strTypeTR & ": " & UserID
End If
Else
If InOutType <> 2 Then
Form2.LogOut "user.log", "User " & strTypeEN & ": " & UserID & " - " & UserNick & " (" & strUserNation & ") (" & strUserJob & ") (Lvl: " & UserLevel & ") (Clan: " & Clan & ")"
Debug.Print "User " & strTypeEN & ": " & UserID & " - " & UserNick & " (" & strUserNation & ") (" & strUserJob & ") (Lvl: " & UserLevel & ") (Clan: " & Clan & ")"
Else
Form2.LogOut "user.log", "User " & strTypeEN & ": " & UserID
Debug.Print "User " & strTypeEN & ": " & UserID
End If
End If
End If

If Form2.chConRace2.Value = 1 Then '//Irk kontrol
If Hex2Val(Mid(MessageBuffer, 2, 1)) <> 2 Then
Dim Opposite As Integer, UserNation2 As Integer, NameLen5 As Integer, InOutType2 As Integer, UserID2 As Long
InOutType2 = Hex2Val(Mid(MessageBuffer, 2, 1))
UserID2 = Hex2Val(Mid(MessageBuffer, 4, 2))
NameLen5 = Hex2Val(Mid(MessageBuffer, 6, 1))
UserNation2 = Hex2Val(Mid(MessageBuffer, 7 + NameLen5, 1))
Select Case ReadLong(KO_ADR_CHR + nation)
Case 1: Opposite = 2
Case 2: Opposite = 1
End Select
Debug.Print "me=" & ReadLong(KO_ADR_CHR + nation)
Debug.Print "opposite=" & Opposite
Debug.Print "usernation=" & UserNation2
If Opposite = UserNation2 Then
If Form2.opRaceTown.Value = True Then
If conNtnTowned = False Then
SendPacket "4800"
conNtnTowned = True
End If
Else
WaitingUser = UserID2
SendPacket "290103"
WaitingUserLeave = True
Debug.Print "Started Pretending to die..."
End If
End If
End If
End If

If Hex2Val(Mid(MessageBuffer, 2, 1)) = 2 Then
If WaitingUserLeave = True Then
Dim UserID3 As Long
UserID3 = Hex2Val(Mid(MessageBuffer, 4, 2))
If UserID3 = WaitingUser Then
SendPacket "290106"
SendPacket "290101"
WaitingUserLeave = False
Debug.Print "End Pretending to die..."
End If
End If
End If


Case &H1A '//Exp deðiþimi
LastExp = ReadLong(KO_ADR_CHR + KO_OFF_EXP)
If Form2.chExpLog.Value = 1 Then
If Form1.optTr.Value = True Then Form2.LogOut "exp.log", "EXP Kazanýldý: " & LastExp & "/" & ReadLong(KO_ADR_CHR + KO_OFF_MAXEXP) Else Form2.LogOut "exp.log", "Earned EXP: " & LastExp & "/" & ReadLong(KO_ADR_CHR + KO_OFF_MAXEXP)
End If

Case &H1B '//Level deðiþimi
If Form2.chExpLog.Value = 1 Then
Dim RecvID As Long
RecvID = Hex2Val(Mid(MessageBuffer, 2, 2))
If RecvID = ReadLong(KO_ADR_CHR + KO_OFF_ID) Then
Dim NewLevel As Integer
NewLevel = Hex2Val(Mid(MessageBuffer, 4, 1))
Form2.LogOut "exp.log", "Level UP (" & NewLevel & ")"
Debug.Print "Level UP (" & NewLevel & ")"
End If
End If

Case &H10 '//Chat,pm
If Form2.chChatLog.Value = 1 Then
Dim RecvType As Integer, NameLen3 As Integer, UserName As String, ChatLen As Integer, ChatString As String
RecvType = Hex2Val(Mid(MessageBuffer, 2, 1))
NameLen3 = Hex2Val(Mid(MessageBuffer, 6, 1))
UserName = Mid(MessageBuffer, 7, NameLen3)
ChatLen = Hex2Val(Mid(MessageBuffer, 7 + NameLen3, 1))
ChatString = Mid(MessageBuffer, 9 + NameLen3, ChatLen)
Select Case RecvType
Case 1 'ALLCHAT
Form2.LogOut "chat.log", "(Normal) " & UserName & ": " & ChatString
Debug.Print "(Normal) " & UserName & ": " & ChatString
Case 2 'PM
Form2.LogOut "chat.log", "(PM) " & UserName & ": " & ChatString
Debug.Print "(PM) " & UserName & ": " & ChatString
Case 3 'PARTY
Form2.LogOut "chat.log", "(Party) " & UserName & ": " & ChatString
Debug.Print "(Party) " & UserName & ": " & ChatString
Case 4 'ALLY
Form2.LogOut "chat.log", "(Alliance) " & UserName & ": " & ChatString
Debug.Print "(Alliance) " & UserName & ": " & ChatString
Case 5 'SHOUT
Form2.LogOut "chat.log", "(Shout) " & UserName & ": " & ChatString
Debug.Print "(Shout) " & UserName & ": " & ChatString
Case 6 'CLAN
Form2.LogOut "chat.log", "(Clan) " & UserName & ": " & ChatString
Debug.Print "(Clan) " & UserName & ": " & ChatString
Case 7 'NOTICE
Form2.LogOut "chat.log", "(Notice) " & UserName & ": " & ChatString
Debug.Print "(Notice) " & UserName & ": " & ChatString
Case 13 'COMMANDER
Form2.LogOut "chat.log", "(Commander) " & UserName & ": " & ChatString
Debug.Print "(Commander) " & UserName & ": " & ChatString
Case 14 'ALLCHAT
Form2.LogOut "chat.log", "(Merchant) " & UserName & ": " & ChatString
Debug.Print "(Merchant) " & UserName & ": " & ChatString
End Select
End If

If Form2.CmdEnabled = True Then
Dim RecvType2 As Integer, NameLen3_2 As Integer, UserName2 As String, ChatLen2 As Integer, ChatString2 As String
RecvType2 = Hex2Val(Mid(MessageBuffer, 2, 1))
NameLen3_2 = Hex2Val(Mid(MessageBuffer, 6, 1))
UserName2 = Mid(MessageBuffer, 7, NameLen3_2)
ChatLen2 = Hex2Val(Mid(MessageBuffer, 7 + NameLen3_2, 1))
ChatString2 = Mid(MessageBuffer, 9 + NameLen3_2, ChatLen2)

If RecvType2 = 2 Then 'PM
Select Case ChatString2
Case Form2.CmdAdd
InviteParty (UserName2)
Case Form2.CmdTown
GoTown
Case Else
End Select
ElseIf RecvType2 = 3 Then 'PARTY
If ChatString2 = Form2.CmdTP Then
TeleportUser (UserName2)
End If
End If
End If

If Form2.CmdPriEnabled = True Then
PTCount = ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H360)
If PTCount > 0 Then
Dim RecvType3 As Integer, NameLen3_3 As Integer, UserName3 As String, ChatLen3 As Integer, ChatString3 As String
RecvType3 = Hex2Val(Mid(MessageBuffer, 2, 1))
NameLen3_3 = Hex2Val(Mid(MessageBuffer, 6, 1))
UserName3 = Mid(MessageBuffer, 7, NameLen3_3)
ChatLen3 = Hex2Val(Mid(MessageBuffer, 7 + NameLen3_3, 1))
ChatString3 = Mid(MessageBuffer, 9 + NameLen3_3, ChatLen3)
If RecvType3 = 3 Then 'PARTY
Select Case ChatString3
Case Form2.CmdBuff
BuffByChat (UserName3)
Case Form2.CmdAc
ACByChat (UserName3)
Case Form2.CmdCure
CureByChat (UserName3)
Case Form2.CmdMind
ResistByChat (UserName3)
Case Form2.CmdStr
STRByChat (UserName3)
Case Else
End Select
End If
End If
End If

'Case &H6A 'Recv Inventory
'If onGoing = False Then
'If Form2.chAutoPot.Value = 1 Then
'InvSlot1 = Hex2Val(Mid(MessageBuffer, 3, 4))
'InvSlot2 = Hex2Val(Mid(MessageBuffer, 14, 4))
'InvQuantity1 = Hex2Val(Mid(MessageBuffer, 9, 1))
'InvQuantity2 = Hex2Val(Mid(MessageBuffer, 20, 1))
'Debug.Print "InvSlot1=" & InvSlot1 & ", InvSlot2=" & InvSlot2 & " - Quantity1=" & InvQuantity1 & ", Quantity2=" & InvQuantity2
'Form2.CheckPots
'Else
'InvSlot1 = 0
'InvSlot2 = 0
'InvQuantity1 = 0
'InvQuantity2 = 0
'End If
'End If

End Select
End If
End If
Loop

Exit Sub

chest_far:
Debug.Print "Could not open chest (Far away)"
End Sub


Function Notice(NoticeYazi As String)

Dim pStr As String

Dim pBytes() As Byte

HexString NoticeYazi

 

pStr = "1013FF00" & hexword

ConvHEX2ByteArray pStr, pBytes

SendPackets pBytes

End Function

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

Public Function GetActiveWindow()
Static lHWnd As Long
    Dim lCurHwnd As Long
    Dim sText As String * 255
'
    lCurHwnd = GetForegroundWindow
    If lCurHwnd = lHWnd Then Exit Function
    lHWnd = lCurHwnd
        GetActiveWindow = "ActiveWidow: " & Left$(sText, _
            GetWindowText(lHWnd, ByVal sText, 255))
End Function

Function OpenBox()
SendPacket "24" + BoxIDNew
End Function
Function LootBox()
If Form2.chCoins.Value = False Then
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
Dim Mob As String
Mob = hex(ReadLong(KO_ADR_CHR + KO_OFF_MOB))
If Mob = "FFFFFFFF" Then
Mob = "FFFF"
Else
Mob = FormatHex(Mob, 4)
End If
GetMobID = Mob
End Function

Function GetCharID()
Dim Char As String
Char = hex(ReadLong(KO_ADR_CHR + KO_OFF_ID))
Char = FormatHex(Char, 4)
'char = Right(char, 4)
GetCharID = Char
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

Function GetMobZ()
Dim tmpPtr As Long
Dim MobBase As Long
Dim MobZ As Single
tmpPtr = ReadLong(KO_PTR_DLG)
MobBase = ReadLong(tmpPtr + &H3D4)
MobZ = ReadFloat(MobBase + &H4C)
GetMobZ = MobZ
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

Function GetMobDistance2()
Dim tmpPtr As Long
Dim cx As Single, cy As Single
Dim frkx As Single, frky As Single, uz As Single
tmpPtr = ReadLong(KO_PTR_CHR)
cx = ReadFloat(tmpPtr + &HB4)
cy = ReadFloat(tmpPtr + &HBC)
frkx = (GetMobX - cx) * (GetMobX - cx)
frky = (GetMobY - cy) * (GetMobY - cy)
uz = ((frkx + frky) ^ 0.5) / 4
GetMobDistance2 = uz
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
'KoJD...
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
Form4.lstInventory.clear
      Dim tmpBase As Long, tmpLng1 As Long, tmpLng2 As Long, tmpLng3 As Long, tmpLng4 As Long
      Dim lngItemID As Long, lngItemID_Ext As Long, lngItemNameLen As Long, AdrItemName As Long, ItemQuantity As Long
      Dim ItemNameB() As Byte
      Dim ItemName As String
      Dim i As Integer
      
      tmpBase = ReadLong(KO_PTR_DLG) 'read KO_DLGBMA adress
      tmpLng1 = ReadLong(tmpBase + &H1A0) 'first pointer
      
For i = 26 To 53 'read 0 to 41 inventory slots (0=earring, 1=helmet, 2=earring, 3=necklace, 4=pauldron ....14=first inventory slot)
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
          
        ItemNum(i) = lngItemID
        ItemQua(i) = ItemQuantity
        If ItemNum(i) = 370004000 Then WolfSlot = i
          
Form4.lstInventory.AddItem i - 25 & " - " & ItemName
      Next
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
'// function "auto heal calculation" ~ kojd' tarihi boþver
'update 27.03.10 !
If Form2.chAutoSit.Value = 1 Then
StandUp
End If

Dim Fark As Long, UserID As Long, SkillID As String
Fark = 0
UserID = 0
Select Case UserNo
Case 0 'kendi charým
Fark = CurMaxHP - CurHP
UserID = ReadLong(KO_ADR_CHR + KO_OFF_ID)
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

Function ReadString(Addr As Long, Size As Long)
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

Function GetMobName()
Dim tmp1 As Long, tmp2 As Long
tmp1 = ReadLong(ReadLong(KO_PTR_DLG) + &H1B8)
tmp2 = ReadLong(tmp1 + &HD4)
GetMobName = ReadString(tmp2 + &H150, 25)
End Function

Function ReadString2(ByVal pAddy As Long, ByVal OtoSize As Boolean, Optional ByVal LSize As Long = 1) As String
        Dim Value As Byte
        Dim tex() As Byte
        On Error Resume Next
If OtoSize = True Then
             ReadProcessMem KO_HANDLE, pAddy, Value, 1, 0&
            LSize = Value
ReDim tex(1 To LSize)
            ReadProcessMem KO_HANDLE, pAddy, tex(1), LSize, 0&
          ReadString2 = StrConv(tex, vbUnicode)
          Else
            If LSize = 0 Then
                MsgBox "Fazla Karakter içeriyor..", vbCritical, "Error"
                Exit Function
            Else
                ReDim tex(1 To LSize)
                 ReadProcessMem KO_HANDLE, pAddy, tex(1), LSize, 0&
                ReadString2 = StrConv(tex, vbUnicode)
            End If
        End If
 End Function

Public Function MobName2()
Dim pPtr As Long, pStr As String
pPtr = ReadLong(KO_PTR_DLG)
Dim FuckThisGame As Long
FuckThisGame = ReadLong(pPtr + &H1B8)
FuckThisGame = ReadLong(FuckThisGame + &HD4)
FuckThisGame = ReadLong(FuckThisGame + &H150)
pStr = ReadString2(FuckThisGame, False, 25)
MobName2 = pStr
End Function
Public Function MobName3()
Dim pPtr As Long, pStr As String
pPtr = ReadLong(KO_PTR_DLG)
Dim FuckThisGame As Long
FuckThisGame = ReadLong(pPtr + &H1B8)
FuckThisGame = ReadLong(FuckThisGame + &HD4)
FuckThisGame = ReadLong(FuckThisGame + &H150)
pStr = ReadString2(FuckThisGame, False, 25)
Form6.lbmob.Caption = pStr
MobName3 = Form6.lbmob.Caption
End Function

Function ClassOku()
Dim a As Long
a = ReadLong(KO_ADR_CHR + KO_OFF_CLASS)
ClassOku = a
End Function

Function CharName()
Dim a As Long, pStr As String
a = ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_NICK)
pStr = ReadString2(a, False, 20)
CharName = pStr
End Function

Function TeleportUser(UserName As String)
On Error Resume Next
Dim tmpUser As String
Dim UserID
Dim NameLen As Long
Dim pBytes() As Byte
tmpUser = ""
Dim i As Integer

PTCount = ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H360)

If PTCount > 0 Then
NameLen = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H34)
Debug.Print "Teleport-NameLen:" & NameLen
ReadByteArray ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H30) + &H0, pBytes, NameLen
Debug.Print "Teleport:" & hex(pBytes(1)) & "," & hex(pBytes(2)) & "," & hex(pBytes(3)) & "," & hex(pBytes(4)) & "," & hex(pBytes(5)) & "," & hex(pBytes(6)) & "," & hex(pBytes(7)) & "," & hex(pBytes(8))
For i = 1 To NameLen
tmpUser = tmpUser & hex(pBytes(i))
Next
Debug.Print "Teleport-Name:" & HexToStr(tmpUser)
If HexToStr(tmpUser) = UserName Then
UserID = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H8)
CastTP (UserID)
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
CastTP (UserID)
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
CastTP (UserID)
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
CastTP (UserID)
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
CastTP (UserID)
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
CastTP (UserID)
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
CastTP (UserID)
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
CastTP (UserID)
Exit Function
End If
End If
End Function

Function CastTP(UserID As Long)
SendPacket "3101" & GetSkillNum("004") & "00" & GetCharID & FormatHex(hex(UserID), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("004") & "00" & GetCharID & FormatHex(hex(UserID), 4) & "000000000000000000000000"
End Function

Sub MemoryAllocation()
FuncPtr = VirtualAllocEx(KO_PID, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
BytesAddr = VirtualAllocEx(KO_PID, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End Sub

Public Function HexString2(EvalString As String) As String
Dim intStrLen As Integer
Dim intLoop As Integer
Dim strHex As String

EvalString = Trim(EvalString)
intStrLen = Len(EvalString)
For intLoop = 1 To intStrLen
strHex = strHex & hex(Asc(Mid(EvalString, intLoop, 1)))
Next
HexString2 = strHex
End Function

Function SelectMob3()
AutoMobZ
End Function

Function SelectMob()
'Tuþ yolla("Z")
'Sendkeys (DIK_Z)
Select Case Form2.attacktype(0).Value
Case True
AutoMobZ
Case False
WriteLong KO_ADR_CHR + KO_OFF_MOB, GetZMob
WriteByte ReadLong(ReadLong(KO_PTR_DLG) + &H1B8) + &HB0, 0
End Select
End Function
Function SelectMob2()
Tuþ yolla("Z")
Sendkeys (DIK_Z)
End Function

Function CharMoving() As Boolean
If ReadByte(KO_ADR_CHR + KO_OFF_MOVE) = 1 Then CharMoving = True Else CharMoving = False
End Function

Function GetCharX()
GetCharX = ReadFloat(KO_ADR_CHR + KO_OFF_X)
End Function

Function GetCharY()
GetCharY = ReadFloat(KO_ADR_CHR + KO_OFF_Y)
End Function

Function GetCharZ()
GetCharZ = ReadFloat(KO_ADR_CHR + KO_OFF_Z)
End Function

Function GetMX()
GetMX = ReadFloat(KO_ADR_CHR + KO_OFF_MX)
End Function

Function GetMY()
GetMY = ReadFloat(KO_ADR_CHR + KO_OFF_MY)
End Function

Function DeleteSkill(ID As String)
SendPacket "3106" & ID & "00" & GetCharID & GetCharID & "000000000000000000000000"
End Function

Function GetCurrentSkill(SkillNo As Integer)
'//kojd// get current skills (sað üstteki skill,buff,sc ler :P)
Dim i As Integer
Dim Ptr As Long, tmpBase As Long
Ptr = ReadLong(KO_PTR_DLG)
tmpBase = ReadLong(Ptr + &H1B4)
tmpBase = ReadLong(tmpBase + &H4)
tmpBase = ReadLong(tmpBase + &HCC)
For i = 1 To SkillNo
tmpBase = ReadLong(tmpBase + &H0) 'soldan kaçýncý skill?
Next
tmpBase = ReadLong(tmpBase + &H8)
If tmpBase > 0 Then
tmpBase = ReadLong(tmpBase + &H0) 've nihayet burasý skill kodunu tutan adres
GetCurrentSkill = tmpBase
Else
GetCurrentSkill = 0
End If
End Function

Function GetSkillCount()
'//kojd// get current skill count (sað üstte kaç skill basýlý ??)
Dim Ptr As Long, tmpBase As Long
Ptr = ReadLong(KO_PTR_DLG)
tmpBase = ReadLong(Ptr + &H1B4)
tmpBase = ReadLong(tmpBase + &H4)
tmpBase = ReadLong(tmpBase + &HD0)
GetSkillCount = tmpBase
End Function

Function GetLastSkill()
'//kojd// get last skill (en son kullanýlan skill..)
Dim Ptr As Long, tmpBase As Long
Ptr = ReadLong(KO_PTR_DLG)
tmpBase = ReadLong(Ptr + &H3D8)
tmpBase = ReadLong(tmpBase + &H6C)
tmpBase = ReadLong(tmpBase + &H10)
GetLastSkill = tmpBase
End Function

Function GetMobBase(TargetMob As Long)
'melih/kojd
Dim Ptr As Long, tmpMobBase As Long, tmpBase As Long, IDArray As Long, BaseAddr As Long, Mob As Long
Mob = TargetMob
Ptr = ReadLong(KO_FLDB)
tmpMobBase = ReadLong(Ptr + &H2C) 'mob=0x2C
tmpBase = ReadLong(tmpMobBase + &H4) '0x1DD8B1B8
While tmpBase <> 0
IDArray = ReadLong(tmpBase + &HC)
If IDArray >= Mob Then
If IDArray = Mob Then
BaseAddr = ReadLong(tmpBase + &H10) 'BASE
End If
tmpBase = ReadLong(tmpBase + &H0) 'Aþaðý
Else
tmpBase = ReadLong(tmpBase + &H8) 'Yukarý
End If
Wend
GetMobBase = BaseAddr
End Function

Function GetCharBase(TargetChar As Long)
'melih**kojd
'player için tek fark 1 offset farklý: Ptr + &H3C
Dim Ptr As Long, tmpCharBase As Long, tmpBase As Long, IDArray As Long, BaseAddr As Long, Mob As Long
Mob = TargetChar
Ptr = ReadLong(KO_FLDB)
tmpCharBase = ReadLong(Ptr + &H3C) 'char=0x3C
tmpBase = ReadLong(tmpCharBase + &H4) '0x1DD8B1B8
While tmpBase <> 0
IDArray = ReadLong(tmpBase + &HC)
If IDArray >= Mob Then
If IDArray = Mob Then
BaseAddr = ReadLong(tmpBase + &H10) 'BASE
End If
tmpBase = ReadLong(tmpBase + &H0)
Else
tmpBase = ReadLong(tmpBase + &H8)
End If
Wend
GetCharBase = BaseAddr
End Function

Function AutoMobZ() 'Z (Enemy)
'29/04
'// melih (kojd)

Dim xCode() As Byte, xStr As String

xStr = "60" & _
        "8B0D" & _
        AlignDWORD(KO_PTR_DLG) & _
        "BF" & _
        AlignDWORD(KO_FNCZ) & _
        "FFD7" & _
        "61C3"

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True

'// kojd - Auto select Z mob
'PUSHAD
'MOV ECX,DWORD PTR DS:[KO_PTR_DLG]
'MOV EDI,KO_FNCZ
'CALL EDI
'POPAD
'RETN
'//

End Function

Function AutoMobX() 'X (Party Member)
'29/04
'// melih (kojd)

Dim xCode() As Byte, xStr As String

xStr = "60" & _
        "8B0D" & _
        AlignDWORD(KO_PTR_DLG) & _
        "BF" & _
        AlignDWORD(KO_FNCX) & _
        "FFD761C3"

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True

End Function

Function AutoMobB() 'B (NPC)
'29/04
'// melih (kojd)

Dim xCode() As Byte, xStr As String

xStr = "60" & _
        "8B0D" & _
        AlignDWORD(KO_PTR_DLG) & _
        "BF" & _
        AlignDWORD(KO_FNCB) & _
        "FFD761C3"

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True

End Function

Function GetZMob() 'Enemy
'26/04
'// melih (kojd)
If BytesAddr2 = 0 Then
BytesAddr2 = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If

'hafýza açalým
If BytesAddr2 <> 0 Then
Dim Mob_Adr As Long, LocX As Long, xCode() As Byte, xStr As String

LocX = ReadLong(KO_PTR_CHR) + KO_OFF_X

xStr = "6068" & _
        AlignDWORD(LocX) & _
        "8B0D" & _
        AlignDWORD(KO_FLDB) & _
        "6A00BF" & _
        AlignDWORD(KO_FPOZ) & _
        "FFD7A3" & _
        AlignDWORD(BytesAddr2) & _
        "61C3"

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True
Mob_Adr = ReadLong(BytesAddr2) 'mob base oku
GetZMob = ReadLong(Mob_Adr + &H5B4)
'hafýza boþaltalým
End If

VirtualFreeEx KO_HANDLE, BytesAddr2, 0, MEM_RELEASE&
End Function

Function GetXMob() 'Party Member
'kojd
If BytesAddr3 = 0 Then
BytesAddr3 = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If

'hafýza açalým
If BytesAddr3 <> 0 Then
Dim Mob_Adr As Long, LocX As Long, PartyBase As Long, xCode() As Byte, xStr As String

LocX = ReadLong(KO_PTR_CHR) + KO_OFF_X
PartyBase = ReadLong(KO_PTR_DLG) + &H1C8

xStr = "6068" & _
        AlignDWORD(LocX) & _
        "8B0D" & _
        AlignDWORD(PartyBase) & _
        "BF" & _
        AlignDWORD(KO_FPOX) & _
        "FFD7A3" & _
        AlignDWORD(BytesAddr3) & _
        "61C3"

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True
Mob_Adr = ReadLong(BytesAddr3) 'mob base oku
GetXMob = ReadLong(Mob_Adr + &H5B4)
'hafýza boþaltalým
End If

VirtualFreeEx KO_HANDLE, BytesAddr3, 0, MEM_RELEASE&
End Function

Function GetBMob() 'NPC
'kojd
If BytesAddr4 = 0 Then
BytesAddr4 = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If

'hafýza açalým
If BytesAddr4 <> 0 Then
Dim Mob_Adr As Long, LocX As Long, xCode() As Byte, xStr As String

LocX = ReadLong(KO_PTR_CHR) + KO_OFF_X

xStr = "6068" & _
        AlignDWORD(LocX) & _
        "8B0D" & _
        AlignDWORD(KO_FLDB) & _
        "BF" & _
        AlignDWORD(KO_FPOB) & _
        "FFD7A3" & _
        AlignDWORD(BytesAddr4) & _
        "61C3"

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True
Mob_Adr = ReadLong(BytesAddr4) 'mob base oku
GetBMob = ReadLong(Mob_Adr + &H5B4)
'hafýza boþaltalým
End If

VirtualFreeEx KO_HANDLE, BytesAddr4, 0, MEM_RELEASE&
End Function

Function SetMob(MobID As Long)
'yaratýðý da client'e seçtirelim öteki türlü bozukluk oluyor :D
'//kojd
If MobID <> 0 Then
Dim MobBase As Long, xCode() As Byte, xStr As String

MobBase = GetMobBase(MobID)

xStr = "6068" & _
        AlignDWORD(MobBase) & _
        "8B0D" & _
        AlignDWORD(KO_PTR_DLG) & _
        "BF" & _
        AlignDWORD(KO_STMB) & _
        "FFD761C3"
        
'// ASM:
        'PUSHAD
        'PUSH MobBase
        'MOV ECX,DWORD PTR DS:[KO_PTR_DLG]
        'MOV EDI,KO_STMB
        'CALL EDI
        'POPAD
        'RETN
'// Kojd

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True
Else: Exit Function
End If
End Function

Function OpenBox2(Box As Long)
SendPacket "24" & FormatHex(hex(Box), 8)
End Function
Function Chest_Exec(Mob As Long, Box As Long)
'tamamlanacak..
End Function

Function GetMyDistance(Target_X As Single, Target_Y As Single, Optional Target_ID As Long)
'/kojd
If Target_ID = ReadLong(KO_ADR_CHR + KO_OFF_ID) Then GetMyDistance = 0: Exit Function
Dim tmpPtr, a, b, mX, mY, cx, cy As Long
Dim frkx, frky, uz As Single
tmpPtr = ReadLong(KO_PTR_CHR)
cx = ReadFloat(tmpPtr + &HB4)
cy = ReadFloat(tmpPtr + &HBC)
frkx = (Target_X - cx) * (Target_X - cx)
frky = (Target_Y - cy) * (Target_Y - cy)
uz = Fix(((frkx + frky) ^ 0.5) / 4)
GetMyDistance = uz
End Function

Function GetRandom(U_Bound As Long) As Long
Randomize
GetRandom = Rnd * U_Bound
End Function

Function GetDistance(TargetID As Long, Mob As Boolean)
'/kojd
If TargetID = ReadLong(KO_ADR_CHR + KO_OFF_ID) Then GetDistance = 0: Exit Function
Dim Target_X As Single, Target_Y As Single
If Mob = True Then Target_X = ReadFloat(GetMobBase(TargetID) + &HB4) Else Target_X = ReadFloat(GetCharBase(TargetID) + &HB4)
If Mob = True Then Target_Y = ReadFloat(GetMobBase(TargetID) + &HBC) Else Target_Y = ReadFloat(GetCharBase(TargetID) + &HBC)
Dim tmpPtr, a, b, mX, mY, cx, cy As Long
Dim frkx, frky, uz As Single
tmpPtr = ReadLong(KO_PTR_CHR)
cx = ReadFloat(tmpPtr + &HB4)
cy = ReadFloat(tmpPtr + &HBC)
frkx = (Target_X - cx) * (Target_X - cx)
frky = (Target_Y - cy) * (Target_Y - cy)
uz = Fix(((frkx + frky) ^ 0.5) / 4)
GetDistance = uz
End Function

Function GetHP(TargetID As Long, Mob As Boolean)
'/kojd
Dim hp As Long
If TargetID = ReadLong(KO_ADR_CHR + KO_OFF_ID) Then GetHP = 0: Exit Function
If Mob = True Then hp = ReadLong(GetMobBase(TargetID) + &H5E4) Else hp = ReadLong(GetCharBase(TargetID) + &H5E4)
GetHP = hp
End Function

Function SetMobName(name As String)
'/melihh
Dim NameAdr As Long
NameAdr = ReadLong(KO_PTR_DLG)
NameAdr = ReadLong(NameAdr + &H1B8)
NameAdr = ReadLong(NameAdr + &HD4)
NameAdr = ReadLong(NameAdr + &H150)
'NameAdr = ReadLong(NameAdr)

Dim str_isim As String, hex_isim As String, byte_isim() As Byte
str_isim = name
hex_isim = HexString2(str_isim)
ConvHEX2ByteArray hex_isim, byte_isim

WriteByteArray NameAdr, byte_isim, UBound(byte_isim) - LBound(byte_isim) + 1

End Function

Function GetName(TargetID As Long, Mob As Boolean)
'/kojd
Dim name As String
If TargetID = ReadLong(KO_ADR_CHR + KO_OFF_ID) Then GetName = "": Exit Function
If Mob = True Then name = ReadString2(ReadLong(GetMobBase(TargetID) + &H5BC), False, 25) Else name = ReadString2(ReadLong(GetCharBase(TargetID) + &H5BC), False, 25)
GetName = name
End Function

Function asdasd1()

Dim xCode() As Byte, xStr As String


xStr = "60BF" & _
        AlignDWORD(&H86EFA0) & _
        "FFD761C3"

ConvHEX2ByteArray xStr, xCode
ExecuteRemoteCode xCode, True
End Function

Function CurHP()
CurHP = ReadLong(KO_ADR_CHR + KO_OFF_HP)
End Function

Function CurMaxHP()
CurMaxHP = ReadLong(KO_ADR_CHR + KO_OFF_MAXHP)
End Function

Function CurMP()
CurMP = ReadLong(KO_ADR_CHR + KO_OFF_MP)
End Function

Function CurMaxMP()
CurMaxMP = ReadLong(KO_ADR_CHR + KO_OFF_MAXMP)
End Function

Function MyLvl()
MyLvl = ReadLong(KO_ADR_CHR + KO_OFF_LVL)
End Function

Function pNewPacket()
packet.Value = ""
End Function

Function pPutByte(pByte As Byte)
packet.Value = packet.Value & FormatHex(hex(pByte), 2)
End Function

Function pLoopByte(pByte As Byte, pCount As Integer)
If pCount > 0 Then
Dim i As Integer, tmp As String
For i = 1 To pCount
tmp = tmp & FormatHex(hex(pByte), 2)
Next
packet.Value = packet.Value & tmp
End If
End Function

Function pPutWord(pWord As Long)
packet.Value = packet.Value & FormatHex(hex(pWord), 4)
End Function

Function pPutDWORD(pDWORD As Long)
packet.Value = packet.Value & FormatHex(hex(pDWORD), 8)
End Function

Function pPutString(pString As String)
packet.Value = packet.Value & HexString2(pString)
End Function

Function pSend()
If packet.Value = vbNullString Then Exit Function
SendPacket packet.Value
End Function

Function mSockID() As Long
mSockID = ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_ID)
End Function
Function tSockID() As Long
tSockID = ReadLong(ReadLong(KO_PTR_CHR) + KO_OFF_MOB)
End Function
