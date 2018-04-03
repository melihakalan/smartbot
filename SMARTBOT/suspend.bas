Attribute VB_Name = "suspend"
Option Explicit

Private Type THREADENTRY32
        dwSize As Long
        cntUsage As Long
        th32ThreadID As Long
        th32OwnerProcessID As Long
        tpBasePri As Long
        tpDeltaPri As Long
        dwFlags As Long
End Type

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const THREAD_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF

Private Const TH32CS_SNAPTHREAD = &H4

Private Declare Function SuspendThread Lib "kernel32" (ByVal hthread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hthread As Long) As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwProcessId As Long) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Sub GetThreads(pID As Long)
Form1.lstThreads.ListItems.clear
Dim hsnapshot As Long
Dim htthread As Long
Dim pthread As Boolean
Dim pt As THREADENTRY32
Dim mList As ListItem

hsnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0)
pt.dwSize = Len(pt)
pthread = Thread32First(hsnapshot, pt)

While pthread
        htthread = OpenThread(THREAD_ALL_ACCESS, 0, pt.th32ThreadID)
        If htthread <> 0 And pt.th32OwnerProcessID = pID Then
        Set mList = Form1.lstThreads.ListItems.Add(, , htthread)
        mList.SubItems(1) = pt.th32ThreadID
        mList.SubItems(2) = pt.cntUsage
        mList.SubItems(3) = pt.dwFlags
        mList.SubItems(4) = pt.dwSize
        mList.SubItems(5) = pt.tpBasePri
        mList.SubItems(6) = pt.tpDeltaPri
        mList.SubItems(7) = "Active"
        End If
        pthread = Thread32Next(hsnapshot, pt)
Wend
CloseHandle hsnapshot
End Sub

Public Function GetKOPID() As Long
Dim pID As Long, wHandle As Long

wHandle = FindWindow(vbNullString, Form1.txtTitle)
If wHandle > 0 Then
        GetWindowThreadProcessId wHandle, pID
        If pID > 0 Then GetKOPID = pID
End If
End Function

Public Function SuspendMe(ThreadID As Long) As Boolean
Dim ret As Long
ret = SuspendThread(ThreadID)
If ret <> -1 Then SuspendMe = True
End Function

Public Function ResumeMe(ThreadID As Long) As Boolean
Dim ret As Long
ret = ResumeThread(ThreadID)
If ret <> -1 Then ResumeMe = True
End Function

Function EnableTPT()
'//KoJD
If Form1.lstThreads.ListItems.Count > 0 Then
        Dim tptID
        tptID = Form1.lstThreads.ListItems.Item(Form1.lstThreads.ListItems.Count).Text
        Debug.Print tptID
        If CLng(tptID) > 0 Then
        Dim ret As Boolean
        ret = ResumeMe(CLng(tptID))
        If ret Then
        'MsgBox "Thread Activated.", vbInformation, "Succeed"
        Else
        MsgBox "Thread Resuming Failed (" & tptID & ")"
        End If
        End If
End If
'//KoJD
End Function

Function DisableTPT()
'//KoJD
If Form1.lstThreads.ListItems.Count > 0 Then
        Dim tptID
        tptID = Form1.lstThreads.ListItems.Item(Form1.lstThreads.ListItems.Count).Text
        Debug.Print tptID
        If CLng(tptID) > 0 Then
        Dim ret As Boolean
        ret = SuspendMe(CLng(tptID))
        If ret Then
        'MsgBox "Thread Suspended.", vbInformation, "Succeed"
        Else
        MsgBox "Thread Suspending Failed (" & tptID & ")"
        End If
        End If
End If
'//KoJD
End Function
