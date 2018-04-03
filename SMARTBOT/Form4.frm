VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oto Buy Sell"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "refresh"
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TAMAM"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtWolf 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Text            =   "100"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CheckBox chSell 
      Caption         =   "Oto sell aktif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox minWolf 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "5"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CheckBox buywolf 
      Caption         =   "Wolf al"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton clear 
      Caption         =   "temizle"
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   ">>"
      Height          =   2600
      Left            =   1980
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.ListBox lstSell 
      Height          =   2595
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox lstInventory 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lbselldesc 
      Caption         =   "Seçtiðiniz itemler'ýn bulunduðu kutular her RPR'de boþaltýlýr."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   4230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "kaç tane alsýn:"
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "not: Satým/Alým iþlemi Rpr yaparken gerçekleþtirilir.."
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "minimum wolf miktarý:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Satýlacak listesi:"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Inventory:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAdd_Click()
If lstInventory.ListCount > 0 And lstInventory.ListIndex > -1 Then
lstSell.AddItem val(Left(lstInventory.Text, 2)) + 25
End If
End Sub

Private Sub clear_Click()
lstSell.clear
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
GetInventory
End Sub
