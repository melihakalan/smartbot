Attribute VB_Name = "submain"
Public KOPATH As String

Sub Main()

If Dir(App.Path & "/config.ini") <> "" Then
KOPATH = ReadIni(App.Path & "/config.ini", "Main", "KOPATH")
If KOPATH <> "" And Dir(KOPATH & "/") <> "" Then
If Dir(KOPATH & "\log.klg") <> "" Then Kill KOPATH & "\log.klg"
If Dir(KOPATH & "\Scheduler.ini") <> "" Then Kill KOPATH & "\Scheduler.ini"
Load Form1
Form1.Show
Else
MsgBox "Could not find KoPath !"
Dim a As String
a = InputBox("KOPATH", "KOPATH", KOPATH)
If a <> "" Then WriteIni App.Path & "/config.ini", "MAIN", "KOPATH", a
End
End If
Else
MsgBox "Could not find config.ini !"
End
End If

End Sub

Sub CloseBot()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
End Sub
