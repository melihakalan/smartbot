VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SB0T INI CRYPTION v1"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function HexToStr(ByVal Data As String) As String
On Error Resume Next
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

Public Function HexString(EvalString As String) As String
On Error Resume Next
Dim intStrLen As Integer
Dim intLoop As Integer
Dim strHex As String

EvalString = Trim(EvalString)
intStrLen = Len(EvalString)
For intLoop = 1 To intStrLen
strHex = strHex & Hex(Asc(Mid(EvalString, intLoop, 1)))
Next
HexString = strHex
End Function

Private Sub Command1_Click()
On Error Resume Next
Text1 = Hex(3 * CDec("&H" & HexString(Text1)))
End Sub

Private Sub Command2_Click()
On Error Resume Next
Text1 = HexToStr(Hex(CDec("&H" & Text1) / 3))
End Sub

Private Sub Command3_Click()
Text1.Text = vbNullString
End Sub
