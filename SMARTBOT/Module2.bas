Attribute VB_Name = "Dinput"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Keyboard As Long
Public Const DIK_0 As Long = 1
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
Public Const DIK_X As Long = &H2D
Public Const KeybPtr As Long = &HB86AEC
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
Public DINPUT_K_R As Long
Public DINPUT_K_S As Long
Public Const KO_DIKKEY As Long = &H26C
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
End If
End Sub
Public Sub SetupKeyboard()
Keyboard = ReadLong(KO_KEYBPTR + 0)
Keyboard = Keyboard + 620
End Sub

Public Sub SendK(ByVal Key As Long)
WriteLong Keyboard + Key * 4, 1
Sleep (50)
WriteLong Keyboard + Key * 4, 0
End Sub
Public Sub SendKeys(ByVal Key As Long)
WriteLong Keyboard + Key * 4, 1
Sleep (50)
WriteLong Keyboard + Key * 4, 0
End Sub
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
End Select
End Function


