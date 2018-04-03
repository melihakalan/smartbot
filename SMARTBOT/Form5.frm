VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chaos Menu"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Chaos Skills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1150
         Left            =   3000
         Top             =   840
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Manuel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         ItemData        =   "Form5.frx":0000
         Left            =   120
         List            =   "Form5.frx":0016
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If List1.SelCount = 0 Then Exit Sub
If Check1.Value = 1 Then
Timer1.Enabled = True
Check1.Caption = "Atak Durdur"
Else
Timer1.Enabled = False
Check1.Caption = "Atak Baþlat"
End If
End Sub

Private Sub Command1_Click()
asdasd1
End Sub

Private Sub Timer1_Timer()
If List1.SelCount = 0 Then Exit Sub
If Check2.Value = 0 Then SelectMob2
CaosAttack
End Sub
