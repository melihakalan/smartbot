Attribute VB_Name = "openlocation"
Option Explicit

Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWDEFAULT = 10
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
    
Public Function openlocation1(URL As String, _
WindowState As Long) As Long

    Dim lHWnd As Long
    Dim lAns As Long

    lAns = ShellExecute(lHWnd, "open", URL, vbNullString, _
    vbNullString, WindowState)
   
    openlocation1 = lAns

End Function







