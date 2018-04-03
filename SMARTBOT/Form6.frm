VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slot"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2295
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   2295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClear 
      Caption         =   "Temizle"
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton btnok 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton addoto 
      Caption         =   "Seçili Mob Ekle"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton addmanuel 
      Caption         =   "Manuel ekle"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox txtmob 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Mob Name"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ListBox lstmob 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lbmob 
      AutoSize        =   -1  'True
      Caption         =   "mob"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lbdesc 
      Caption         =   "Atak yapýlacak yaratýk listesi:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addmanuel_Click()
If txtmob <> vbNullString Then lstmob.AddItem txtmob
End Sub

Private Sub addoto_Click()
If GetMobID <> "FFFF" Then lstmob.AddItem MobName2
End Sub

Private Sub btnClear_Click()
lstmob.clear
End Sub

Private Sub btnok_Click()
Me.Hide
End Sub
