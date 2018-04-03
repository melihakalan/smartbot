VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SB0T 1831"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame segman 
      Caption         =   "zaman segmant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8520
      TabIndex        =   16
      Top             =   840
      Width           =   375
      Begin VB.Label lblEvade 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   216
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label labelz 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Labellf 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label labelswift 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label labelcat 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label labellup 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label labeldef3 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label labelmaster 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label labeldef 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.Label labelwolf 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Genel"
      TabPicture(0)   =   "SmartBot.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frPot"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frGenel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frRepair"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frTarget"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frTS"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tmMain1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chTop"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tmBeep"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btnreskill"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tmLogKill"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Korunma"
      TabPicture(1)   =   "SmartBot.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frChat"
      Tab(1).Control(1)=   "frTimed"
      Tab(1).Control(2)=   "frDetails"
      Tab(1).Control(3)=   "frGM"
      Tab(1).Control(4)=   "frSlot"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Atak"
      TabPicture(2)   =   "SmartBot.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Diðer"
      TabPicture(3)   =   "SmartBot.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Priest"
      TabPicture(4)   =   "SmartBot.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frPTInfo"
      Tab(4).Control(1)=   "frPriest"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Save Bot"
      TabPicture(5)   =   "SmartBot.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "btnSaveall"
      Tab(5).Control(1)=   "Label1"
      Tab(5).ControlCount=   2
      Begin VB.CommandButton btnSaveall 
         Caption         =   "Ayarlarý Kaydet"
         Height          =   375
         Left            =   -74040
         TabIndex        =   229
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete Chaos TS"
         Height          =   285
         Left            =   120
         TabIndex        =   224
         Top             =   7200
         Width           =   1575
      End
      Begin VB.Timer tmLogKill 
         Enabled         =   0   'False
         Interval        =   20000
         Left            =   1320
         Top             =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Chaos TS (Illegal)"
         Height          =   285
         Left            =   120
         TabIndex        =   223
         Top             =   6840
         Width           =   1575
      End
      Begin VB.CommandButton btnreskill 
         Caption         =   "Free Reskill"
         Height          =   285
         Left            =   3000
         TabIndex        =   221
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Timer tmBeep 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2760
         Top             =   480
      End
      Begin VB.CheckBox chTop 
         Caption         =   "Top"
         Height          =   195
         Left            =   3480
         TabIndex        =   218
         Top             =   360
         Width           =   615
      End
      Begin VB.Timer tmMain1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   3840
         Top             =   480
      End
      Begin VB.Frame frPTInfo 
         Caption         =   "      ID                   HP                    MP                  LVL"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   178
         Top             =   480
         Width           =   3975
         Begin VB.Timer tmPriest 
            Enabled         =   0   'False
            Interval        =   800
            Left            =   2880
            Top             =   1320
         End
         Begin VB.Label lbPTIndex 
            Caption         =   "1  2 3 4 5 6 7 8"
            ForeColor       =   &H00000000&
            Height          =   1575
            Left            =   50
            TabIndex        =   211
            Top             =   240
            Width           =   135
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   0
            Left            =   240
            TabIndex        =   210
            Top             =   260
            Width           =   495
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   1
            Left            =   240
            TabIndex        =   209
            Top             =   450
            Width           =   495
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   2
            Left            =   240
            TabIndex        =   208
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   3
            Left            =   240
            TabIndex        =   207
            Top             =   860
            Width           =   495
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   4
            Left            =   240
            TabIndex        =   206
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   5
            Left            =   240
            TabIndex        =   205
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   6
            Left            =   240
            TabIndex        =   204
            Top             =   1420
            Width           =   495
         End
         Begin VB.Label lbUserID 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   135
            Index           =   7
            Left            =   240
            TabIndex        =   203
            Top             =   1620
            Width           =   495
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   0
            Left            =   960
            TabIndex        =   202
            Top             =   260
            Width           =   1095
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   1
            Left            =   960
            TabIndex        =   201
            Top             =   450
            Width           =   1095
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   2
            Left            =   960
            TabIndex        =   200
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   3
            Left            =   960
            TabIndex        =   199
            Top             =   860
            Width           =   1095
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   4
            Left            =   960
            TabIndex        =   198
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   5
            Left            =   960
            TabIndex        =   197
            Top             =   1230
            Width           =   1095
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   6
            Left            =   960
            TabIndex        =   196
            Top             =   1425
            Width           =   1095
         End
         Begin VB.Label lbUserHP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   135
            Index           =   7
            Left            =   960
            TabIndex        =   195
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   0
            Left            =   2060
            TabIndex        =   194
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   1
            Left            =   2060
            TabIndex        =   193
            Top             =   435
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   2
            Left            =   2060
            TabIndex        =   192
            Top             =   645
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   3
            Left            =   2060
            TabIndex        =   191
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   4
            Left            =   2060
            TabIndex        =   190
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   5
            Left            =   2060
            TabIndex        =   189
            Top             =   1215
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   6
            Left            =   2060
            TabIndex        =   188
            Top             =   1410
            Width           =   1095
         End
         Begin VB.Label lbUserMP 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   135
            Index           =   7
            Left            =   2060
            TabIndex        =   187
            Top             =   1605
            Width           =   1095
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   0
            Left            =   3380
            TabIndex        =   186
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   1
            Left            =   3380
            TabIndex        =   185
            Top             =   435
            Width           =   495
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   2
            Left            =   3380
            TabIndex        =   184
            Top             =   645
            Width           =   495
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   3
            Left            =   3380
            TabIndex        =   183
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   4
            Left            =   3380
            TabIndex        =   182
            Top             =   1035
            Width           =   495
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   5
            Left            =   3380
            TabIndex        =   181
            Top             =   1215
            Width           =   495
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   6
            Left            =   3380
            TabIndex        =   180
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label lbUserLVL 
            Alignment       =   2  'Center
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   135
            Index           =   7
            Left            =   3380
            TabIndex        =   179
            Top             =   1605
            Width           =   495
         End
      End
      Begin VB.Frame frPriest 
         Caption         =   "Priest"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   148
         Top             =   2400
         Width           =   3975
         Begin VB.CheckBox chParty 
            Caption         =   "Party Aktif"
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
            TabIndex        =   225
            Top             =   3120
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.Timer tmGroup 
            Enabled         =   0   'False
            Interval        =   5000
            Left            =   3480
            Top             =   2880
         End
         Begin VB.Timer tmRestore 
            Enabled         =   0   'False
            Interval        =   5000
            Left            =   3480
            Top             =   3240
         End
         Begin VB.ComboBox comRestore 
            ForeColor       =   &H00808000&
            Height          =   315
            ItemData        =   "SmartBot.frx":00A8
            Left            =   960
            List            =   "SmartBot.frx":00C4
            TabIndex        =   154
            Text            =   "Combo1"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.ComboBox comGroupHeal 
            ForeColor       =   &H00808000&
            Height          =   315
            ItemData        =   "SmartBot.frx":0175
            Left            =   960
            List            =   "SmartBot.frx":017F
            TabIndex        =   157
            Text            =   "Combo1"
            Top             =   2325
            Width           =   2175
         End
         Begin VB.ComboBox comAutoHeal 
            ForeColor       =   &H000000FF&
            Height          =   315
            ItemData        =   "SmartBot.frx":01B8
            Left            =   960
            List            =   "SmartBot.frx":01D7
            TabIndex        =   169
            Text            =   "Combo1"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtLmtHeal 
            ForeColor       =   &H000000FF&
            Height          =   310
            Left            =   3480
            TabIndex        =   168
            Text            =   "50"
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox comAutoBuff 
            ForeColor       =   &H00800080&
            Height          =   315
            ItemData        =   "SmartBot.frx":0285
            Left            =   960
            List            =   "SmartBot.frx":02AA
            TabIndex        =   167
            Text            =   "Combo1"
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox comAutoAc 
            ForeColor       =   &H00FF0000&
            Height          =   315
            ItemData        =   "SmartBot.frx":0358
            Left            =   960
            List            =   "SmartBot.frx":0377
            TabIndex        =   166
            Text            =   "Combo1"
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CheckBox chAutoMalice 
            Caption         =   "Auto Z + Malice"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   120
            TabIndex        =   165
            Top             =   3720
            Width           =   1455
         End
         Begin VB.CheckBox chAutoCure 
            Caption         =   "Auto Cure"
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   120
            TabIndex        =   164
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CheckBox chAutoHeal 
            Caption         =   "Heal"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   120
            TabIndex        =   163
            Top             =   320
            Width           =   615
         End
         Begin VB.CheckBox chAutoBuff 
            Caption         =   "Buff"
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   120
            TabIndex        =   162
            Top             =   1395
            Width           =   615
         End
         Begin VB.Timer tmBuffAC 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   3480
            Top             =   1440
         End
         Begin VB.Timer tmAutoMalice 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   3240
            Top             =   1920
         End
         Begin VB.CheckBox chAutoSit 
            Caption         =   "Boþtayken Otur"
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   1800
            TabIndex        =   161
            Top             =   3480
            Width           =   1575
         End
         Begin VB.CheckBox chAutoSTR 
            Caption         =   "Auto STR"
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   1800
            TabIndex        =   160
            Top             =   3720
            Width           =   1335
         End
         Begin VB.ComboBox comAutoResist 
            ForeColor       =   &H00008000&
            Height          =   315
            ItemData        =   "SmartBot.frx":045B
            Left            =   960
            List            =   "SmartBot.frx":046E
            TabIndex        =   159
            Text            =   "Combo1"
            Top             =   1995
            Width           =   2175
         End
         Begin VB.CheckBox chGroupHeal 
            Caption         =   "Group"
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   120
            TabIndex        =   158
            Top             =   2355
            Width           =   1215
         End
         Begin VB.TextBox txtIntGroupHeal 
            ForeColor       =   &H00000000&
            Height          =   310
            Left            =   3480
            TabIndex        =   156
            Text            =   "10"
            Top             =   2325
            Width           =   375
         End
         Begin VB.CheckBox chRestore 
            Caption         =   "Restore"
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   120
            TabIndex        =   155
            Top             =   2655
            Width           =   975
         End
         Begin VB.TextBox txtIntRestore 
            ForeColor       =   &H00000000&
            Height          =   310
            Left            =   3480
            TabIndex        =   153
            Text            =   "25"
            Top             =   2640
            Width           =   375
         End
         Begin VB.Timer tmCure 
            Enabled         =   0   'False
            Interval        =   800
            Left            =   3120
            Top             =   1560
         End
         Begin VB.CheckBox chSaveParty 
            Caption         =   "Party Kurtar"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   480
            TabIndex        =   152
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton opGroupMassive 
            Caption         =   "Group Massive Healing"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   480
            TabIndex        =   151
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton opGroupComplete 
            Caption         =   "Group Complete Healing"
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   480
            TabIndex        =   150
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtSavePTLimit 
            ForeColor       =   &H000000FF&
            Height          =   310
            Left            =   3120
            TabIndex        =   149
            Text            =   "75"
            Top             =   800
            Width           =   375
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3240
            TabIndex        =   177
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "sn."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3180
            TabIndex        =   176
            Top             =   2355
            Width           =   345
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "sn."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3180
            TabIndex        =   175
            Top             =   2670
            Width           =   225
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "AC    ->"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   360
            TabIndex        =   174
            Top             =   1725
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Mind ->"
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   360
            TabIndex        =   173
            Top             =   2040
            Width           =   555
         End
         Begin VB.Label Label9 
            Caption         =   "|_ | |_"
            ForeColor       =   &H00800080&
            Height          =   615
            Left            =   195
            TabIndex        =   172
            Top             =   1635
            Width           =   135
         End
         Begin VB.Line LineTemp 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Index           =   4
            X1              =   220
            X2              =   220
            Y1              =   520
            Y2              =   700
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2880
            TabIndex        =   171
            Top             =   860
            Width           =   165
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "}"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   36
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   870
            Left            =   2520
            TabIndex        =   170
            Top             =   440
            Width           =   345
         End
         Begin VB.Line LineTemp 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Index           =   5
            X1              =   210
            X2              =   400
            Y1              =   720
            Y2              =   720
         End
      End
      Begin VB.Frame frTS 
         Caption         =   "Legal TS (60dk)"
         Height          =   1095
         Left            =   120
         TabIndex        =   143
         Top             =   4560
         Width           =   3975
         Begin VB.Timer tmTS 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   3480
            Top             =   600
         End
         Begin VB.CheckBox chTS 
            Caption         =   "Oto TS Baþlat"
            Height          =   195
            Left            =   120
            TabIndex        =   145
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox comTS 
            Height          =   315
            ItemData        =   "SmartBot.frx":04BC
            Left            =   120
            List            =   "SmartBot.frx":04F6
            TabIndex        =   144
            Text            =   "Combo1"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lbTSRem2 
            AutoSize        =   -1  'True
            Caption         =   "?"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3120
            TabIndex        =   147
            Top             =   360
            Width           =   75
         End
         Begin VB.Label lbtemp 
            AutoSize        =   -1  'True
            Caption         =   "Kalan dakika:"
            Height          =   195
            Index           =   15
            Left            =   1920
            TabIndex        =   146
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Frame frChat 
         Caption         =   "Oto Chat"
         Height          =   735
         Left            =   -74880
         TabIndex        =   85
         Top             =   6360
         Width           =   3975
         Begin VB.Timer tmChat 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   120
            Top             =   240
         End
         Begin VB.CheckBox chChat 
            Caption         =   "Check24"
            Height          =   195
            Left            =   3680
            TabIndex        =   89
            Top             =   280
            Width           =   200
         End
         Begin VB.TextBox delChat 
            Height          =   285
            Left            =   3000
            TabIndex        =   88
            Text            =   "3000"
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox comChat 
            Height          =   315
            ItemData        =   "SmartBot.frx":05BE
            Left            =   2040
            List            =   "SmartBot.frx":05D4
            TabIndex        =   87
            Text            =   "Combo3"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtChat 
            Height          =   285
            Left            =   120
            TabIndex        =   86
            Text            =   ":D"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame frTimed 
         Caption         =   "Zamanlý Ýþlemler"
         Height          =   855
         Left            =   -74880
         TabIndex        =   77
         Top             =   5520
         Width           =   3975
         Begin VB.CheckBox chTimedOffPC 
            Caption         =   "PC Kapat"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2040
            TabIndex        =   81
            Top             =   520
            Width           =   975
         End
         Begin VB.CheckBox chTimedOffKO 
            Caption         =   "KO Kapat"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   960
            TabIndex        =   80
            Top             =   520
            Width           =   975
         End
         Begin VB.TextBox txtMin 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   480
            MaxLength       =   2
            TabIndex        =   79
            Text            =   "00"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   78
            Text            =   "12"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lbMin 
            AutoSize        =   -1  'True
            Caption         =   "Dk"
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   600
            TabIndex        =   83
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lbHr 
            AutoSize        =   -1  'True
            Caption         =   "Sa"
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   240
            TabIndex        =   82
            Top             =   240
            Width           =   180
         End
      End
      Begin VB.Frame frDetails 
         Caption         =   "Detay Kayýtlarý"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   73
         Top             =   4080
         Width           =   3975
         Begin VB.CheckBox chExpLog 
            Caption         =   "Alýnan Levelleri ve gelen tüm Expleri kaydet        ( Log/Exp.log)"
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   3735
         End
         Begin VB.CheckBox chSlotLog 
            Caption         =   "Slota gelip giden herkesi ID,Nick,Level,Job,Clan þeklinde kaydet (Log/user.log)"
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   120
            TabIndex        =   75
            Top             =   480
            Width           =   3735
         End
         Begin VB.CheckBox chChatLog 
            Caption         =   "Tüm Chat konuþmalarýný kaydet (Log/chat.log)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame frGM 
         Caption         =   "GM Korunma"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   65
         Top             =   2520
         Width           =   3975
         Begin VB.CheckBox chGMDC 
            Caption         =   "DC->"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   960
            TabIndex        =   84
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox chGMLog 
            Caption         =   "Yapýlan iþlemi Log olarak kaydet (Log/gm.log)"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1200
            Value           =   2  'Grayed
            Width           =   3735
         End
         Begin VB.CheckBox chGMTown 
            Caption         =   "Town->"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   840
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtNotifyPT 
            Height          =   285
            Left            =   1440
            TabIndex        =   69
            Text            =   "{GMID} Online!"
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox lbNotifyPT 
            Caption         =   "Partiye bildir:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chGMOffKO 
            Caption         =   "KO Kapat->"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1680
            TabIndex        =   67
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chGMOffPC 
            Caption         =   "PC Kapat"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2880
            TabIndex        =   66
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lbGM 
            Caption         =   "Slota GM Gelirse:"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame frTarget 
         Caption         =   "Hedef Bilgileri"
         Height          =   975
         Left            =   120
         TabIndex        =   44
         Top             =   5760
         Width           =   3975
         Begin VB.Label lbTarName 
            AutoSize        =   -1  'True
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   720
            TabIndex        =   54
            Top             =   240
            Width           =   300
         End
         Begin VB.Label lbtemp 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   465
         End
         Begin VB.Label lbTarID 
            AutoSize        =   -1  'True
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   720
            TabIndex        =   52
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lbtemp 
            AutoSize        =   -1  'True
            Caption         =   "ID:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   225
         End
         Begin VB.Label lbTarDis 
            AutoSize        =   -1  'True
            Caption         =   "Null"
            ForeColor       =   &H00404000&
            Height          =   195
            Left            =   3120
            TabIndex        =   50
            Top             =   720
            Width           =   255
         End
         Begin VB.Label lbTarY 
            AutoSize        =   -1  'True
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3120
            TabIndex        =   49
            Top             =   480
            Width           =   255
         End
         Begin VB.Label lbTarX 
            AutoSize        =   -1  'True
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   2520
            TabIndex        =   48
            Top             =   480
            Width           =   255
         End
         Begin VB.Label lbTarHP 
            AutoSize        =   -1  'True
            Caption         =   "Null"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   720
            TabIndex        =   47
            Top             =   480
            Width           =   300
         End
         Begin VB.Label lbtemp 
            AutoSize        =   -1  'True
            Caption         =   "Uzaklýk:"
            Height          =   195
            Index           =   3
            Left            =   2520
            TabIndex        =   46
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lbtemp 
            AutoSize        =   -1  'True
            Caption         =   "HP:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   255
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   7095
         Left            =   -75000
         TabIndex        =   28
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   12515
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Ýþlemler"
         TabPicture(0)   =   "SmartBot.frx":0604
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frConditions"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frOther"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Oto Buy - Sell"
         TabPicture(1)   =   "SmartBot.frx":0620
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Chaos Atak"
         TabPicture(2)   =   "SmartBot.frx":063C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.Frame frOther 
            Caption         =   "Karýþýk"
            Height          =   3975
            Left            =   50
            TabIndex        =   232
            Top             =   3000
            Width           =   3975
            Begin VB.Timer tm_other 
               Enabled         =   0   'False
               Index           =   1
               Interval        =   1000
               Left            =   3000
               Top             =   960
            End
            Begin VB.TextBox txt_other 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   5
               Left            =   2280
               TabIndex        =   241
               Text            =   "Y2"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txt_other 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   1560
               TabIndex        =   240
               Text            =   "X2"
               Top             =   600
               Width           =   615
            End
            Begin VB.CommandButton btn_other 
               Caption         =   "Git-Gel"
               Height          =   285
               Index           =   2
               Left            =   3000
               TabIndex        =   239
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txt_other 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   840
               TabIndex        =   238
               Text            =   "Y1"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txt_other 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   237
               Text            =   "X1"
               Top             =   600
               Width           =   615
            End
            Begin VB.Timer tm_other 
               Enabled         =   0   'False
               Index           =   0
               Interval        =   1000
               Left            =   3480
               Top             =   960
            End
            Begin VB.CommandButton btn_other 
               Caption         =   "Teleport"
               Height          =   285
               Index           =   1
               Left            =   2760
               TabIndex        =   236
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton btn_other 
               Caption         =   "Koþ"
               Height          =   285
               Index           =   0
               Left            =   1560
               TabIndex        =   235
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txt_other 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   234
               Text            =   "Y"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txt_other 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   233
               Text            =   "X"
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame frConditions 
            Caption         =   "Koþullar"
            Height          =   2655
            Left            =   50
            TabIndex        =   116
            Top             =   360
            Width           =   3975
            Begin VB.Frame frConRace 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1095
               Left            =   2160
               TabIndex        =   140
               Top             =   600
               Visible         =   0   'False
               Width           =   1695
               Begin VB.OptionButton opRaceTown 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Town"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   142
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1335
               End
               Begin VB.OptionButton opRaceSeemDie 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Slottan ayrýlana kadar ölü numarasý yap"
                  Height          =   675
                  Left            =   120
                  TabIndex        =   141
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Shape shConRace 
                  Height          =   1095
                  Left            =   0
                  Top             =   0
                  Width           =   1695
               End
            End
            Begin VB.OptionButton opExpAlarm 
               Caption         =   "Alarm"
               ForeColor       =   &H00808000&
               Height          =   195
               Left            =   2280
               TabIndex        =   119
               Top             =   1800
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.Timer tmReturnSlot 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   3480
               Top             =   1800
            End
            Begin VB.CommandButton btnConRace 
               Caption         =   "VVV Ýþem VVV"
               Height          =   255
               Left            =   2160
               TabIndex        =   117
               Top             =   360
               Width           =   1695
            End
            Begin VB.Timer tm3MinNoExp 
               Enabled         =   0   'False
               Interval        =   60000
               Left            =   3480
               Top             =   2160
            End
            Begin VB.CheckBox chConReturn 
               Caption         =   "Öldüðün Yere Dön"
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   120
               TabIndex        =   137
               Top             =   2160
               Width           =   2895
            End
            Begin VB.CheckBox chConNoExp 
               Caption         =   "3 Dk. Exp Gelmezse:"
               ForeColor       =   &H00008000&
               Height          =   435
               Left            =   120
               TabIndex        =   136
               Top             =   1680
               Width           =   2655
            End
            Begin VB.CheckBox chConRace2 
               Caption         =   "Slota Karþý Irk Gelirse:"
               ForeColor       =   &H00800080&
               Height          =   495
               Left            =   120
               TabIndex        =   134
               Top             =   240
               Width           =   3735
            End
            Begin VB.Frame frConDc 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   2280
               TabIndex        =   129
               Top             =   600
               Width           =   1575
               Begin VB.OptionButton opDcAlarm 
                  Caption         =   "Alarm"
                  ForeColor       =   &H00808000&
                  Height          =   195
                  Left            =   0
                  TabIndex        =   130
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton opDcOff 
                  Caption         =   "Off PC"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Left            =   720
                  TabIndex        =   131
                  Top             =   120
                  Width           =   855
               End
            End
            Begin VB.Frame frConDead 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2280
               TabIndex        =   126
               Top             =   960
               Width           =   1575
               Begin VB.OptionButton opDieAlarm 
                  Caption         =   "Alarm"
                  ForeColor       =   &H00808000&
                  Height          =   195
                  Left            =   0
                  TabIndex        =   128
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton opDieOff 
                  Caption         =   "Off PC"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Left            =   720
                  TabIndex        =   127
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.Frame frConNoPt 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2280
               TabIndex        =   123
               Top             =   1200
               Width           =   1575
               Begin VB.OptionButton opNoPtTown 
                  Caption         =   "Town"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   720
                  TabIndex        =   125
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton opNoPtAlarm 
                  Caption         =   "Alarm"
                  ForeColor       =   &H00808000&
                  Height          =   195
                  Left            =   0
                  TabIndex        =   124
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.Frame frConNoMana 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2280
               TabIndex        =   120
               Top             =   1440
               Width           =   1575
               Begin VB.OptionButton opNoManaTown 
                  Caption         =   "Town"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   720
                  TabIndex        =   122
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton opNoManaAlarm 
                  Caption         =   "Alarm"
                  ForeColor       =   &H00808000&
                  Height          =   195
                  Left            =   0
                  TabIndex        =   121
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.OptionButton opExpOff 
               Caption         =   "Off PC"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   3000
               TabIndex        =   118
               Top             =   1800
               Width           =   855
            End
            Begin VB.CheckBox chConDC 
               Caption         =   "DC Olursa:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   132
               Top             =   720
               Width           =   3735
            End
            Begin VB.CheckBox chConDead 
               Caption         =   "Yatarsa:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   138
               Top             =   960
               Width           =   3735
            End
            Begin VB.CheckBox chConNoParty 
               Caption         =   "Party Bozulursa:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   133
               Top             =   1200
               Width           =   3735
            End
            Begin VB.CheckBox chConNoMana 
               Caption         =   "Mana Biterse:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   135
               Top             =   1440
               Width           =   3735
            End
            Begin VB.Label lbTmpMin 
               AutoSize        =   -1  'True
               Caption         =   "3"
               Height          =   195
               Left            =   3360
               TabIndex        =   139
               Top             =   2400
               Visible         =   0   'False
               Width           =   90
            End
         End
      End
      Begin VB.Frame frSlot 
         Caption         =   "Slot Kayýt"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Width           =   3975
         Begin VB.OptionButton opmobname 
            Alignment       =   1  'Right Justify
            Caption         =   "Mob adýna göre"
            Height          =   255
            Left            =   1320
            TabIndex        =   228
            Top             =   1680
            Width           =   1455
         End
         Begin VB.OptionButton opmobid 
            Caption         =   "Mob ID göre"
            Height          =   255
            Left            =   120
            TabIndex        =   227
            Top             =   1680
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton btnmobname 
            Caption         =   "Adlar>>"
            Height          =   285
            Left            =   2880
            TabIndex        =   226
            Top             =   1680
            Width           =   975
         End
         Begin VB.Timer tmSlot 
            Enabled         =   0   'False
            Interval        =   2500
            Left            =   3480
            Top             =   120
         End
         Begin VB.CommandButton btnClearSlot 
            Caption         =   "Temizle"
            Height          =   285
            Left            =   2880
            TabIndex        =   58
            Top             =   1380
            Width           =   975
         End
         Begin VB.CommandButton btnSlot 
            Caption         =   "Slot Ekle"
            Height          =   285
            Left            =   1680
            TabIndex        =   57
            Top             =   1380
            Width           =   1215
         End
         Begin VB.CheckBox chSlot 
            Caption         =   "Aktifleþtir"
            Height          =   255
            Left            =   1680
            TabIndex        =   56
            Top             =   240
            Width           =   2175
         End
         Begin VB.ListBox lstSlot 
            Height          =   1425
            ItemData        =   "SmartBot.frx":0658
            Left            =   120
            List            =   "SmartBot.frx":065A
            TabIndex        =   55
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label txtSlotLVL 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2880
            TabIndex        =   64
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label txtMobCount 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2880
            TabIndex        =   63
            Top             =   840
            Width           =   945
         End
         Begin VB.Label txtsSid 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2880
            TabIndex        =   62
            Top             =   600
            Width           =   945
         End
         Begin VB.Label lbMobCount 
            AutoSize        =   -1  'True
            Caption         =   "Yaratýk Sayýsý:"
            Height          =   195
            Left            =   1680
            TabIndex        =   61
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label lbSlotLVL 
            AutoSize        =   -1  'True
            Caption         =   "Slot Level:"
            Height          =   195
            Left            =   1680
            TabIndex        =   60
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label lbsSid 
            AutoSize        =   -1  'True
            Caption         =   "Slot sSid:"
            Height          =   195
            Left            =   1680
            TabIndex        =   59
            Top             =   600
            Width           =   660
         End
      End
      Begin VB.Frame frRepair 
         Caption         =   "Oto RPR"
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   3855
         Begin VB.Timer tmTPGecikme 
            Enabled         =   0   'False
            Interval        =   2500
            Left            =   2040
            Top             =   480
         End
         Begin VB.Timer tmTownGecikme 
            Enabled         =   0   'False
            Interval        =   2500
            Left            =   2400
            Top             =   480
         End
         Begin VB.CheckBox chRprTown 
            Caption         =   "Town"
            Height          =   195
            Left            =   1440
            TabIndex        =   222
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtMinDur 
            Height          =   285
            Left            =   2880
            TabIndex        =   220
            Text            =   "3000"
            Top             =   960
            Width           =   855
         End
         Begin VB.Timer tmSlotMove 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2760
            Top             =   480
         End
         Begin VB.Timer tmRprMove 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   3120
            Top             =   480
         End
         Begin VB.Timer tmRpr 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   3480
            Top             =   480
         End
         Begin VB.CheckBox chRepair 
            Caption         =   "Oto Repair"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton btnRecNpc 
            Caption         =   "Npc Kayýt"
            Height          =   320
            Left            =   120
            TabIndex        =   38
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lbtemp 
            AutoSize        =   -1  'True
            Caption         =   "Minimum Durability:"
            Height          =   195
            Index           =   17
            Left            =   1440
            TabIndex        =   219
            Top             =   1000
            Width           =   1380
         End
         Begin VB.Label lbNpcID 
            AutoSize        =   -1  'True
            Caption         =   "Npc ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   540
         End
         Begin VB.Label txtNpcID 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   720
            TabIndex        =   42
            Top             =   600
            Width           =   585
         End
         Begin VB.Label txtNpcX 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1440
            TabIndex        =   41
            Top             =   600
            Width           =   585
         End
         Begin VB.Label txtNpcY 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Null"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2160
            TabIndex        =   40
            Top             =   600
            Width           =   585
         End
      End
      Begin VB.Frame frGenel 
         Caption         =   "Genel"
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   3855
         Begin VB.CheckBox chRuntochest 
            Caption         =   "Kutuya Git"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   231
            Top             =   720
            Width           =   1695
         End
         Begin VB.Timer tmWH 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   3360
            Top             =   120
         End
         Begin VB.CheckBox chCoins 
            Caption         =   "Sadece Para"
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chNoDC 
            Caption         =   "No Dc"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chWH 
            Caption         =   "Wall Hack"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chAutoParty 
            Caption         =   "Oto Party"
            Height          =   255
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chAutoLoot 
            Caption         =   "Oto Kutu"
            Height          =   195
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.Line LineTemp 
            BorderWidth     =   2
            Index           =   3
            X1              =   430
            X2              =   550
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line LineTemp 
            BorderWidth     =   2
            Index           =   2
            X1              =   435
            X2              =   435
            Y1              =   720
            Y2              =   840
         End
      End
      Begin VB.Frame frPot 
         Caption         =   "Oto Pot"
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3855
         Begin VB.TextBox txtMinor 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   600
            TabIndex        =   34
            Text            =   "90"
            Top             =   840
            Width           =   375
         End
         Begin VB.Timer tmMinor 
            Enabled         =   0   'False
            Interval        =   150
            Left            =   3000
            Top             =   840
         End
         Begin VB.CheckBox Chmin 
            Caption         =   "Auto Minor"
            Height          =   195
            Left            =   120
            TabIndex        =   214
            Top             =   600
            Width           =   1095
         End
         Begin VB.Timer tmCheckHPMP 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   3480
            Top             =   360
         End
         Begin VB.Timer tmMP 
            Enabled         =   0   'False
            Interval        =   1700
            Left            =   3360
            Top             =   480
         End
         Begin VB.Timer tmHP 
            Enabled         =   0   'False
            Interval        =   1700
            Left            =   3240
            Top             =   240
         End
         Begin VB.CheckBox chAutoHPMP 
            Caption         =   "Oto Pot"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtHPLimit 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1920
            TabIndex        =   32
            Text            =   "50"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtMPLimit 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3000
            TabIndex        =   33
            Text            =   "50"
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox comHPPot 
            Appearance      =   0  'Flat
            ForeColor       =   &H000000FF&
            Height          =   315
            ItemData        =   "SmartBot.frx":065C
            Left            =   1320
            List            =   "SmartBot.frx":0672
            TabIndex        =   31
            Text            =   "Combo1"
            Top             =   600
            Width           =   1010
         End
         Begin VB.ComboBox comMPPot 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            ItemData        =   "SmartBot.frx":0693
            Left            =   2400
            List            =   "SmartBot.frx":06A9
            TabIndex        =   30
            Text            =   "Combo1"
            Top             =   600
            Width           =   1010
         End
         Begin VB.Line LineTemp 
            Index           =   1
            X1              =   330
            X2              =   200
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line LineTemp 
            Index           =   0
            X1              =   210
            X2              =   210
            Y1              =   840
            Y2              =   960
         End
         Begin VB.Label lbtemp 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   16
            Left            =   360
            TabIndex        =   217
            Top             =   860
            Width           =   165
         End
         Begin VB.Label lbHPLimit 
            Caption         =   "HP    %"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1320
            TabIndex        =   37
            Top             =   255
            Width           =   735
         End
         Begin VB.Label lbMPLimit 
            Caption         =   "MP    %"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   36
            Top             =   255
            Width           =   735
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   7095
         Left            =   -75000
         TabIndex        =   1
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   12515
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Rogue"
         TabPicture(0)   =   "SmartBot.frx":06CD
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frRogue"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frRogueTimed"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Warrior"
         TabPicture(1)   =   "SmartBot.frx":06E9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frWrr"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Priest"
         TabPicture(2)   =   "SmartBot.frx":0705
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frPri"
         Tab(2).ControlCount=   1
         Begin VB.Frame frPri 
            Caption         =   "Priest"
            Height          =   3255
            Left            =   -74880
            TabIndex        =   104
            Top             =   360
            Width           =   3855
            Begin VB.CommandButton btnAtkLoad 
               Caption         =   "Yükle"
               Height          =   285
               Left            =   120
               TabIndex        =   114
               Top             =   480
               Width           =   975
            End
            Begin VB.OptionButton opAtkDebuff 
               Caption         =   "Debuff Atak"
               Height          =   255
               Left            =   2280
               TabIndex        =   113
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton opAtkBuff 
               Caption         =   "Buff Atak"
               Height          =   255
               Left            =   1200
               TabIndex        =   112
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton opAtkHeal 
               Caption         =   "Heal Atak"
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.Timer tmPriBot 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   3360
               Top             =   960
            End
            Begin VB.Timer tmPriRAttack 
               Enabled         =   0   'False
               Interval        =   1325
               Left            =   2880
               Top             =   480
            End
            Begin VB.CheckBox chPriAutoMob 
               Caption         =   "Oto Yaratýk Seç"
               Height          =   255
               Left            =   2160
               TabIndex        =   109
               Top             =   840
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox chPriStartBot 
               Caption         =   "Start"
               Height          =   375
               Left            =   2160
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   1920
               Width           =   1575
            End
            Begin VB.OptionButton opPriRunMob 
               Caption         =   "Yaratýða Koþ"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   2160
               TabIndex        =   107
               Top             =   1095
               Width           =   1215
            End
            Begin VB.OptionButton opPriWaitMob 
               Caption         =   "Yaratýðý Bekle"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   2160
               TabIndex        =   106
               Top             =   1320
               Width           =   1335
            End
            Begin VB.CheckBox chMaliceMob 
               Caption         =   "Yaratýða Malice"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2160
               TabIndex        =   105
               Top             =   1560
               Width           =   1455
            End
            Begin MSFlexGridLib.MSFlexGrid PriflexSkill 
               Height          =   2175
               Left            =   120
               TabIndex        =   110
               Top             =   840
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   3836
               _Version        =   393216
               Rows            =   1
               FixedCols       =   0
               BackColor       =   16777215
               BackColorSel    =   8421504
               ForeColorSel    =   16777215
               BackColorBkg    =   16777215
               GridColor       =   0
               AllowBigSelection=   0   'False
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
               AllowUserResizing=   1
               FormatString    =   "Num |Skill                   "
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame frWrr 
            Caption         =   "Warrior"
            Height          =   3255
            Left            =   -74880
            TabIndex        =   92
            Top             =   360
            Width           =   3855
            Begin VB.Timer tmDefense 
               Enabled         =   0   'False
               Interval        =   16000
               Left            =   1920
               Top             =   2640
            End
            Begin VB.CheckBox chAutoSprint 
               Caption         =   "Sprint"
               Height          =   255
               Left            =   120
               TabIndex        =   103
               Top             =   2640
               Width           =   735
            End
            Begin VB.CheckBox chAutoDefense 
               Caption         =   "Defense"
               Height          =   255
               Left            =   960
               TabIndex        =   102
               Top             =   2640
               Width           =   900
            End
            Begin VB.Timer tmBindingInt 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   3360
               Top             =   960
            End
            Begin VB.Timer tmRockInterval 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   3360
               Top             =   600
            End
            Begin VB.Timer tmWrrBot 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   3360
               Top             =   240
            End
            Begin VB.Timer tmWrrRAttack 
               Enabled         =   0   'False
               Interval        =   1325
               Left            =   3000
               Top             =   240
            End
            Begin VB.OptionButton opWrrWaitMob 
               Caption         =   "Yaratýðý Bekle"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   2160
               TabIndex        =   100
               Top             =   840
               Width           =   1335
            End
            Begin VB.OptionButton opWrrRunMob 
               Caption         =   "Yaratýða Koþ"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   2160
               TabIndex        =   99
               Top             =   615
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.Frame frBindMob 
               Caption         =   "Yaratýðý Çek"
               Height          =   975
               Left            =   2160
               TabIndex        =   95
               Top             =   1080
               Width           =   1575
               Begin VB.CheckBox chBindMob 
                  Caption         =   "Yaratýðý Çek"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   98
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.OptionButton opBinding 
                  Caption         =   "Binding"
                  Enabled         =   0   'False
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   360
                  TabIndex        =   97
                  Top             =   480
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton opRock 
                  Caption         =   "Taþ"
                  Enabled         =   0   'False
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   360
                  TabIndex        =   96
                  Top             =   720
                  Width           =   855
               End
            End
            Begin VB.CheckBox chWrrStartBot 
               Caption         =   "Start"
               Height          =   375
               Left            =   2160
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   2160
               Width           =   1575
            End
            Begin VB.CheckBox chWrrAutoMob 
               Caption         =   "Oto Yaratýk Seç"
               Height          =   255
               Left            =   2160
               TabIndex        =   93
               Top             =   360
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin MSFlexGridLib.MSFlexGrid WrrflexSkill 
               Height          =   2175
               Left            =   120
               TabIndex        =   101
               Top             =   360
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   3836
               _Version        =   393216
               Rows            =   1
               FixedCols       =   0
               BackColor       =   16777215
               BackColorSel    =   8421504
               ForeColorSel    =   16777215
               BackColorBkg    =   16777215
               GridColor       =   0
               AllowBigSelection=   0   'False
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
               AllowUserResizing=   1
               FormatString    =   "Num |Skill                   "
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame frRogueTimed 
            Caption         =   "Rogue Zamanlý Skiller"
            Height          =   1575
            Left            =   120
            TabIndex        =   13
            Top             =   3240
            Width           =   3855
            Begin VB.Timer rogueZAMAN 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   120
               Top             =   480
            End
            Begin VB.Timer zhidetmr 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   120
               Top             =   960
            End
            Begin VB.CheckBox chckzhde 
               Caption         =   "Z Hide"
               Height          =   255
               Left            =   600
               TabIndex        =   26
               Top             =   720
               Width           =   975
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Start"
               Height          =   375
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Durdur"
               Height          =   375
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ListBox roglist2 
               Height          =   1230
               ItemData        =   "SmartBot.frx":0721
               Left            =   1680
               List            =   "SmartBot.frx":0740
               MultiSelect     =   1  'Simple
               TabIndex        =   14
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Frame frRogue 
            Caption         =   "Rogue Atak"
            Height          =   2775
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   3855
            Begin VB.Frame Frame1 
               Caption         =   "Attack"
               Height          =   735
               Left            =   1920
               TabIndex        =   242
               Top             =   1200
               Width           =   1815
               Begin VB.OptionButton attacktype 
                  Caption         =   "HF Style"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   162
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   244
                  Top             =   480
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton attacktype 
                  Caption         =   "Normal"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   162
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   243
                  Top             =   240
                  Width           =   1575
               End
            End
            Begin VB.Timer Rogcs 
               Enabled         =   0   'False
               Interval        =   1300
               Left            =   3360
               Top             =   720
            End
            Begin VB.Timer rogatak 
               Enabled         =   0   'False
               Interval        =   1300
               Left            =   3480
               Top             =   240
            End
            Begin VB.CheckBox chmobdead 
               Caption         =   "Mob ölmeden Z yapma"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1920
               TabIndex        =   215
               Top             =   960
               Width           =   1815
            End
            Begin VB.OptionButton cs 
               Caption         =   "Counter Strike"
               Height          =   195
               Left            =   1920
               TabIndex        =   213
               Top             =   720
               Width           =   1335
            End
            Begin VB.OptionButton listeden 
               Caption         =   "Listeden seç"
               Height          =   195
               Left            =   1920
               TabIndex        =   212
               Top             =   480
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Atak hýzý..."
               Height          =   285
               Left            =   2040
               TabIndex        =   115
               Top             =   2400
               Width           =   1695
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Start"
               Height          =   375
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   91
               Top             =   2040
               Width           =   1695
            End
            Begin VB.CheckBox chAutoMob 
               Caption         =   "Oto Yaratýk Seç"
               Height          =   255
               Left            =   1920
               TabIndex        =   90
               Top             =   240
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.ListBox List2 
               Height          =   2010
               ItemData        =   "SmartBot.frx":0795
               Left            =   70
               List            =   "SmartBot.frx":07D8
               TabIndex        =   12
               Top             =   240
               Width           =   1695
            End
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "KoJD && Orion && GamerTurk www.Onlinehile.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74160
         TabIndex        =   230
         Top             =   1920
         Width           =   2535
      End
   End
   Begin VB.Label Label14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   15
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pStr As String
Dim pBytes() As Byte

Private Const VK_F11 As Long = &H7A
Private Const VK_LCONTROL As Long = &HA2
Private Const VK_NUMPAD0 As Long = &H60

Public BeepIndex As Integer
Public LastMalicedID As Long, LastMalicedID2 As Long
Public Maliced As Boolean
Public LastRow As Integer, LastRow2 As Integer
Public CmdEnabled As Boolean, tmpCmdEnabled As String, CmdPriEnabled As Boolean, tmpCmdPriEnabled As String
Public CmdBuff As String, CmdAc As String, CmdMind As String, CmdCure As String, CmdStr As String
Public CmdAdd As String, CmdKick As String, CmdTown As String, CmdTP As String
Public LastAP As Long
Public LastX As Long, LastY As Long, LastZ As Long
Public DeadX As Long, DeadY As Long, DeadZ As Long, LastBotStatus2 As Integer, LastRprStatus2 As Integer
Public LastAutoMob As Integer, LastAutoMob2 As Integer
Public LastBotStatus As Integer, LastRprStatus As Integer
Public LastWrrBotStatus As Integer, LastPriBotStatus As Integer, LastArcBotStatus1 As Integer, LastArcBotStatus2 As Integer
Public LastWrrBotStatus2 As Integer, LastPriBotStatus2 As Integer, LastArcBotStatus12 As Integer, LastArcBotStatus22 As Integer


Private Sub btn_other_Click(Index As Integer)
Select Case Index
Case 0

If IsNumeric(txt_other(0)) = True And IsNumeric(txt_other(1)) = True And Len(txt_other(0)) > 0 And Len(txt_other(1)) > 0 Then
tm_other(0).Enabled = True
End If

Case 1

If IsNumeric(txt_other(0)) = True And IsNumeric(txt_other(1)) = True And Len(txt_other(0)) > 0 And Len(txt_other(1)) > 0 Then
WriteFloat KO_ADR_CHR + KO_OFF_X, CSng(txt_other(0))
WriteFloat KO_ADR_CHR + KO_OFF_Y, CSng(txt_other(1))
End If

Case 2

If tm_other(1).Enabled = False Then
If IsNumeric(txt_other(2)) = True And IsNumeric(txt_other(3)) = True And IsNumeric(txt_other(4)) = True And IsNumeric(txt_other(5)) = True And Len(txt_other(2)) > 0 And Len(txt_other(3)) > 0 And Len(txt_other(4)) > 0 And Len(txt_other(5)) > 0 Then
onPatrol = 1
tm_other(1).Enabled = True
If Form1.optTr.Value = True Then btn_other(2).Caption = "Dur" Else btn_other(2).Caption = "Stop"
End If
Else
tm_other(1).Enabled = False
If Form1.optTr.Value = True Then btn_other(2).Caption = "Git-Gel" Else btn_other(2).Caption = "Patrol"
End If

End Select
End Sub

Private Sub btnAtkLoad_Click()
PriflexSkill.Clear

LoadBasicSkillNum
LoadBasicSkillName

LoadMainSkillNum
LoadMainSkillName

LoadMasterSkillNum
LoadMasterSkillName
End Sub

Private Sub btnClearSlot_Click()
lstSlot.Clear
txtsSid = "Null"
txtMobCount = "Null"
txtSlotLVL = "Null"
End Sub

Private Sub btnConRace_Click()
If frConRace.Visible = False Then
frConRace.Visible = True
btnConRace.Caption = "^^^"
Else
frConRace.Visible = False
If Form1.optTr.Value = True Then btnConRace.Caption = "VVV Ýþlem VVV" Else btnConRace.Caption = "VVV Do VVV"
End If
End Sub

Private Sub btnmobname_Click()
Form6.Show
End Sub

Private Sub btnRecNpc_Click()
If GetMobID <> "FFFF" Then
txtNpcID = GetMobID
txtNpcX = GetMobX
txtNpcY = GetMobY
Else
If Form1.optTr.Value = True Then MsgBox "NPC Seçilmemiþ.", vbExclamation, "Sorun"
If Form1.optEn.Value = True Then MsgBox "No NPC Chosen.", vbExclamation, "Error"
End If
End Sub

Private Sub btnreskill_Click()
SendPacket "3201"
SendPacket "3404"
If Form1.optTr.Value = True Then MsgBox "Reskill uygulandý. Üzerinizde 1 skill point varsa baþarýyla olmuþtur. Relog atýn." Else MsgBox "Executed reskill. It was successfully If you have 1 skill point. Please relog."
End Sub

Private Sub btnSaveall_Click()
On Error Resume Next
SaveSetting "SB0T", "AllSettings", "HPLimit", txtHPLimit
SaveSetting "SB0T", "AllSettings", "MPLimit", txtMPLimit
SaveSetting "SB0T", "AllSettings", "HPPot", comHPPot.ListIndex
SaveSetting "SB0T", "AllSettings", "MPPot", comMPPot.ListIndex
SaveSetting "SB0T", "AllSettings", "MinorLimit", txtMinor
SaveSetting "SB0T", "AllSettings", "RPR_NpcID", txtNpcID.Caption
SaveSetting "SB0T", "AllSettings", "RPR_NpcLocX", txtNpcX.Caption
SaveSetting "SB0T", "AllSettings", "RPR_NpcLocY", txtNpcY.Caption
SaveSetting "SB0T", "AllSettings", "RPR_Durability", txtMinDur
SaveSetting "SB0T", "AllSettings", "Pri_HealType", comAutoHeal.ListIndex
SaveSetting "SB0T", "AllSettings", "Pri_HealLimit", txtLmtHeal
SaveSetting "SB0T", "AllSettings", "Pri_BuffType", comAutoBuff.ListIndex
SaveSetting "SB0T", "AllSettings", "Pri_ACType", comAutoAc.ListIndex
SaveSetting "SB0T", "AllSettings", "Pri_MindType", comAutoResist.ListIndex
SaveSetting "SB0T", "AllSettings", "Pri_GroupHeal", comGroupHeal.ListIndex
SaveSetting "SB0T", "AllSettings", "Pri_Restore", comRestore.ListIndex
SaveSetting "SB0T", "AllSettings", "Pri_SaveParty", txtSavePTLimit
SaveSetting "SB0T", "AllSettings", "AttackSpeed", Form3.Text1

SaveSetting "SB0T", "AllSettings", "SellCount", Form4.lstSell.ListCount
For i = 0 To Form4.lstSell.ListCount
Form4.lstSell.ListIndex = i
SaveSetting "SB0T", "AllSettings", "Sell" & i, Form4.lstSell.Text
Next

SaveSetting "SB0T", "AllSettings", "SlotCount_ID", lstSlot.ListCount
For i = 0 To lstSlot.ListCount
lstSlot.ListIndex = i
SaveSetting "SB0T", "AllSettings", "Slot_ID" & i, lstSlot.Text
Next

SaveSetting "SB0T", "AllSettings", "SlotCount_Name", Form6.lstmob.ListCount
For i = 0 To Form6.lstmob.ListCount
Form6.lstmob.ListIndex = i
SaveSetting "SB0T", "AllSettings", "Slot_Name" & i, Form6.lstmob.Text
Next

SaveSetting "SB0T", "AllSettings", "MinWolf", Form4.minWolf
SaveSetting "SB0T", "AllSettings", "BuyWolf", Form4.txtWolf
End Sub


Private Sub btnSlot_Click()
tmpMob = ReadLong(KO_ADR_CHR + KO_OFF_MOB)
If GetMobID <> "FFFF" Then
btnSlot.Caption = "..."
btnSlot.Enabled = False
SendPacket ("1D0100" & FormatHex(hex(tmpMob), 4))
For i = 1 To 7
SendPacket ("1D0100" & FormatHex(hex(tmpMob + i), 4))
Next
For i = 1 To 7
SendPacket ("1D0100" & FormatHex(hex(tmpMob - i), 4))
Next
GetSlot = 1
tmSlot.Enabled = True
End If
End Sub

Private Sub chAutoCure_Click()
If chAutoCure.Value = 1 Then
tmCure.Enabled = True
Else
tmCure.Enabled = False
End If
End Sub

Private Sub chAutoDefense_Click()
If chAutoDefense.Value = 1 Then
tmDefense.Enabled = True
SendPacket "3103" & GetSkillNum("007") & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Else
tmDefense.Enabled = False
End If
End Sub

Private Sub chAutoHPMP_Click()
If chAutoHPMP.Value = 1 Then
If IsNumeric(txtHPLimit) = True And IsNumeric(txtMPLimit) = True Then
tmCheckHPMP.Enabled = True
Else
chAutoHPMP.Value = 0
If Form1.optTr.Value = True Then MsgBox "Limitler hatalý.", vbExclamation, "Sorun" Else MsgBox "Invalid Limits.", vbExclamation, "Error"
End If
Else
tmCheckHPMP.Enabled = False
End If
End Sub

Private Sub chAutoMalice_Click()
If chAutoMalice.Value = 1 Then tmAutoMalice.Enabled = True Else tmAutoMalice.Enabled = False
End Sub

Private Sub chAutoMob_Click()
'If chAutoMob.Value = 1 Then tmRogueZ.Enabled = True Else tmRogueZ.Enabled = False
End Sub

Private Sub chAutoSit_Click()
If chAutoSit.Value = 1 Then
Sit
Else
StandUp
End If
End Sub

Private Sub chBindMob_Click()
If chBindMob.Value = 0 Then
opBinding.Enabled = False
opRock.Enabled = False
Else
opBinding.Enabled = True
opRock.Enabled = True
End If
End Sub

Private Sub chChat_Click()
If chChat.Value = 1 Then
tmChat.Interval = Int(delChat.Text)
tmChat.Enabled = True
Else
tmChat.Enabled = False
End If
End Sub

Private Sub chckzhde_Click()
If chckzhde.Value = 1 Then
GizleBeni
zhidetmr.Enabled = True
Else
zhidetmr.Enabled = False
End If
End Sub

Private Sub chConDC_Click()
conDCOff = False
End Sub

Private Sub chConDead_Click()
conDeadOff = False
End Sub

Private Sub chConNoExp_Click()
If chConNoExp.Value = 1 Then
tm3MinNoExp.Enabled = True
Else
tm3MinNoExp.Enabled = False
End If
lbTmpMin.Caption = "3"
End Sub

Private Sub chConNoMana_Click()
conNoManaTowned = False
End Sub

Private Sub chConNoParty_Click()
conNoPtTowned = False
End Sub

Private Sub chConRace2_Click()
conNtnTowned = False
End Sub

Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
rogatak.Interval = val(Form3.Text1)
Rogcs.Interval = val(Form3.Text1)
If listeden.Value = True Then rogatak.Enabled = True
If cs.Value = True Then Rogcs.Enabled = True
'tmRogueZ.Enabled = True
Else
rogatak.Enabled = False
Rogcs.Enabled = False
'tmRogueZ.Enabled = False
End If
End Sub

Private Sub Checkcs_Click()
If Checkcs.Value = 1 Then
Rogcs.Enabled = True
Else
Rogcs.Enabled = False
End If
End Sub

Private Sub chGroupHeal_Click()
If IsNumeric(txtIntGroupHeal) = True And val(txtIntGroupHeal) <= 60 Then
If chGroupHeal.Value = 1 Then
chSaveParty.Enabled = False
chSaveParty.Value = 0
tmGroup.Interval = val(txtIntGroupHeal) * 1000
tmGroup.Enabled = True
'GroupHealNow
Else
chSaveParty.Enabled = True
tmGroup.Enabled = False
End If
Else
If Form1.optTr.Value = True Then MsgBox "Saniye hatalý. Yeinden giriniz. (En fazla 60 saniye.)", vbExclamation, "Sorun" Else MsgBox "Group Heal interval is invalid. Please re-enter. (Max. 60 seconds.)", vbExclamation, "Error"
chGroupHeal.Value = 0
End If
End Sub

Private Sub Chmin_Click()
If Chmin.Value = 1 Then tmMinor.Enabled = True Else tmMinor.Enabled = False
End Sub

Private Sub chNoDC_Click()
If chNoDC.Value = 1 Then WriteLong KO_NODC, 1 Else WriteLong KO_NODC, 0
End Sub

Private Sub chPriStartBot_Click()
If val(PriflexSkill.Text) > 0 Then
If chPriStartBot.Value = 1 Then
tmPriBot.Enabled = True
tmPriRAttack.Enabled = True
Else
tmPriBot.Enabled = False
tmPriRAttack.Enabled = False
End If
Else
If Form1.optTr.Value = True Then MsgBox "Skill seçmediniz.", vbExclamation, "Sorun"
If Form1.optEn.Value = True Then MsgBox "You didn't choose any skill.", vbExclamation, "Error"
chPriStartBot.Value = 0
End If
End Sub

Private Sub chRepair_Click()
If txtNpcID = "Null" Or txtNpcX = "Null" Or txtNpcY = "Null" Then
chRepair.Value = 0
If Form1.optTr.Value = True Then MsgBox "Önce NPC Kaydedin!", vbExclamation, "Sorun"
If Form1.optEn.Value = True Then MsgBox "Save NPC First!", vbExclamation, "Error"
Else
If chRepair.Value = 1 Then
tmRpr.Enabled = True
Else
tmRpr.Enabled = False
End If
End If
End Sub

Private Sub chRestore_Click()
If IsNumeric(txtIntRestore) = True And val(txtIntRestore) <= 60 Then
If chRestore.Value = 1 Then
tmRestore.Interval = val(txtIntRestore) * 1000
tmRestore.Enabled = True
'RestoreNow
Else
tmRestore.Enabled = False
End If
Else
If Form1.optTr.Value = True Then MsgBox "Saniye hatalý. Yeinden giriniz. (En fazla 60 saniye.)", vbExclamation, "Sorun" Else MsgBox "Group Heal interval is invalid. Please re-enter. (Max. 60 seconds.)", vbExclamation, "Error"
chRestore.Value = 0
End If
End Sub

Private Sub chSaveParty_Click()
If chSaveParty.Value = 1 Then
opGroupMassive.Enabled = True
opGroupComplete.Enabled = True
chGroupHeal.Enabled = False
chGroupHeal.Value = 0
Else
opGroupMassive.Enabled = True
opGroupComplete.Enabled = True
chGroupHeal.Enabled = True
End If
End Sub

Private Sub chSlot_Click()
'If lstSlot.ListCount = 0 Then
'If Form1.optTr.Value = True Then MsgBox "Slota Yaratýk Eklenmemiþ.", vbExclamation, "Sorun" Else MsgBox "You didn't add any mob.", vbExclamation, "Error"
'chSlot.Value = 0
'End If
End Sub

Private Sub chTop_Click()
If chTop.Value = 1 Then
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Else
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub

Private Sub chWH_Click()
If chWH.Value = 1 Then
tmWH.Enabled = True
WriteLong KO_ADR_CHR + KO_OFF_WH, 0
Else
tmWH.Enabled = False
WriteLong KO_ADR_CHR + KO_OFF_WH, 1
End If
End Sub

Private Sub chWrrStartBot_Click()
If val(WrrflexSkill.Text) > 0 Then
If chWrrStartBot.Value = 1 Then
tmWrrBot.Enabled = True
tmWrrRAttack.Enabled = True
Else
tmWrrBot.Enabled = False
tmWrrRAttack.Enabled = False
End If
Else
If Form1.optTr.Value = True Then MsgBox "Skill seçmediniz.", vbExclamation, "Sorun"
If Form1.optEn.Value = True Then MsgBox "You didn't choose any skill.", vbExclamation, "Error"
chWrrStartBot.Value = 0
End If
End Sub

Private Sub chTimedOffKO_Click()
If IsNumeric(txtHr) = False Or val(txtHr) > 24 Or IsNumeric(txtMin) = False Or val(txtMin) > 60 Then
chTimedOffKO.Value = 0
If Form1.optTr.Value = True Then MsgBox "Saat ve Dakika hatalý girilmiþ.", vbExclamation, "Sorun" Else MsgBox "Invalid Hour and Minute values.", vbExclamation, "Error"
End If
GotTimedAction = False
End Sub

Private Sub chTimedOffPC_Click()
If IsNumeric(txtHr) = False Or val(txtHr) > 24 Or IsNumeric(txtMin) = False Or val(txtMin) > 60 Then
chTimedOffPC.Value = 0
If Form1.optTr.Value = True Then MsgBox "Saat ve Dakika hatalý girilmiþ.", vbExclamation, "Sorun" Else MsgBox "Invalid Hour and Minute values.", vbExclamation, "Error"
End If
GotTimedAction = False
End Sub

Private Sub chTS_Click()
If chTS.Value = 1 Then
UseTS
lbTSRem2.Caption = "61"
tmTS.Enabled = True
Else
lbTSRem2.Caption = "?"
tmTS.Enabled = False
End If
End Sub

Private Sub comAutoHeal_Change()
If comAutoHeal.ListIndex = 8 Then
If MyLvl < 45 Then
MsgBox "You must be level 45 or greater."
comAutoHeal.ListIndex = 0
End If
End If
End Sub

Private Sub Command1_Click()
KrowazTS
End Sub

Private Sub Command12_Click()
Command12.Visible = False
Command5.Visible = True
Command5.Left = 120
Command5.Top = 240

rogueZAMAN.Enabled = False
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
DeleteKrowazTS
End Sub

Private Sub Command4_Click()
asdasd1
End Sub

Private Sub Command5_Click()
Command5.Visible = False
Command12.Visible = True
Command12.Left = 120
Command12.Top = 240

rogueZAMAN.Enabled = True
Wolf
Def
Def2
Def3
CatEyes
Lupin
mastersc
sw
NewLF

End Sub

Private Sub cs_Click()
List2.Enabled = False
End Sub

Private Sub Form_Initialize()
Me.Caption = Form1.txtWnd & " - SB0T"
End Sub

Public Sub LoadBasicSkillNum()
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
Open App.Path & "/priest/basic.txt" For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   PriflexSkill.AddItem Left(sNextLine, 3)
   Loop
Close nFileNum
End Sub

Public Sub LoadBasicSkillName()
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
Open App.Path & "/priest/basic.txt" For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   PriflexSkill.Row = lLineCount
   PriflexSkill.Col = 1
   PriflexSkill.Text = Mid(sNextLine, 4, Len(sNextLine) - 3)
   lLineCount = lLineCount + 1
   Loop
Close nFileNum
LastRow2 = PriflexSkill.Rows - 1
End Sub

Public Sub LoadMainSkillNum()
PriflexSkill.Row = LastRow2
Dim tmpFile As String 'attacklist
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
If opAtkHeal.Value = True Then tmpFile = "heal.txt"
If opAtkBuff.Value = True Then tmpFile = "buff.txt"
If opAtkDebuff.Value = True Then tmpFile = "debuff.txt"
Open App.Path & "/priest/" & tmpFile For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   PriflexSkill.AddItem Left(sNextLine, 3)
   Loop
Close nFileNum
End Sub

Public Sub LoadMainSkillName()
PriflexSkill.Row = LastRow2
Dim tmpFile As String 'attacklist
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
If opAtkHeal.Value = True Then tmpFile = "heal.txt"
If opAtkBuff.Value = True Then tmpFile = "buff.txt"
If opAtkDebuff.Value = True Then tmpFile = "debuff.txt"
Open App.Path & "/priest/" & tmpFile For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   PriflexSkill.Row = LastRow2 + lLineCount
   PriflexSkill.Col = 1
   PriflexSkill.Text = Mid(sNextLine, 4, Len(sNextLine) - 3)
   lLineCount = lLineCount + 1
   Loop
Close nFileNum
LastRow2 = PriflexSkill.Rows - 1
End Sub

Public Sub LoadMasterSkillNum()
PriflexSkill.Row = LastRow2
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
Open App.Path & "/priest/master.txt" For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   PriflexSkill.AddItem Left(sNextLine, 3)
   Loop
Close nFileNum
End Sub

Public Sub LoadMasterSkillName()
PriflexSkill.Row = LastRow2
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
Open App.Path & "/priest/master.txt" For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   PriflexSkill.Row = LastRow2 + lLineCount
   PriflexSkill.Col = 1
   PriflexSkill.Text = Mid(sNextLine, 4, Len(sNextLine) - 3)
   lLineCount = lLineCount + 1
   Loop
Close nFileNum
End Sub

Private Sub Form_Load()
On Error Resume Next

Kill App.Path & "/Log/exp.log"
Kill App.Path & "/Log/gm.log"
Kill App.Path & "/Log/user.log"

LoadSkillNum
LoadSkillName

LoadCmd
For i = 1 To WrrflexSkill.Rows - 1
WrrflexSkill.Row = i
WrrflexSkill.Col = 0
WrrflexSkill.CellAlignment = 0
WrrflexSkill.Col = 1
WrrflexSkill.CellAlignment = 0
Next
For i = 1 To PriflexSkill.Rows - 1
PriflexSkill.Row = i
PriflexSkill.Col = 0
PriflexSkill.CellAlignment = 0
PriflexSkill.Col = 1
PriflexSkill.CellAlignment = 0
Next

comHPPot.ListIndex = 0
comMPPot.ListIndex = 0
comChat.ListIndex = 0
comTS.ListIndex = 0

If MyLvl >= 45 Then comAutoHeal.ListIndex = 8 Else comAutoHeal.ListIndex = 0
comAutoBuff.ListIndex = 0
comAutoResist.ListIndex = 0
comRestore.ListIndex = 0
comGroupHeal.ListIndex = 0
comAutoAc.ListIndex = 0

GetLastAP

LoadBotSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SaveSetting "SmartBot", "Settings", "Run", Form1.Run + 1
'openlocation1 "http://www.onlinehile.com/", 1
CloseBot
End Sub

Private Sub listeden_Click()
List2.Enabled = True
End Sub

Private Sub rogatak_Timer()
Dim a As Integer, b As Integer
On Error Resume Next

If chAutoMob.Value = 1 Then
If chmobdead.Value = 1 Then
If GetMobID = "FFFF" Or GetMobHP = 0 Or GetMobDistance >= 13 Or FormatDec(GetMobID, 4) <= 9999 Then SelectMob
Else
SelectMob
End If
Else
End If

If GetMobID <> "FFFF" And GetMobDistance <= 12 Then
If chSlot.Value = 1 Then
If opmobid.Value = True And lstSlot.ListCount > 0 Then
lstSlot.ListIndex = 0
For a = 0 To lstSlot.ListCount
If lstSlot.Text = GetMobID Then
Skill1to50
Exit For
Else
lstSlot.ListIndex = a
End If
Next
ElseIf opmobname.Value = True And Form6.lstmob.ListCount > 0 Then
Form6.lstmob.ListIndex = 0
For b = 0 To Form6.lstmob.ListCount
If Form6.lstmob.Text = MobName3 Then
Skill1to50
Exit For
Else
Form6.lstmob.ListIndex = b
End If
Next
Else
Exit Sub
End If
Else
'SendPacket "1F01F1B18B090008"
Skill1to50
'SendPacket "1F01312590060006"
'SendPacket "3103F9A10100" & GetCharID & GetMobID & "0100010000000000000000000000"
'SendPacket "080101" & GetMobID & "FF00000000"
End If
End If
End Sub

Private Sub Rogcs_Timer()
Dim a As Integer, b As Integer
On Error Resume Next

If chAutoMob.Value = 1 Then
If chmobdead.Value = 1 Then
If GetMobID = "FFFF" Or GetMobHP = 0 Or GetMobDistance >= 13 Or FormatDec(GetMobID, 4) <= 9999 Then SelectMob
Else
SelectMob
End If
Else
End If

If GetMobID <> "FFFF" And GetMobDistance <= 12 Then
If chSlot.Value = 1 Then
If opmobid.Value = True And lstSlot.ListCount > 0 Then
lstSlot.ListIndex = 0
For a = 0 To lstSlot.ListCount
If lstSlot.Text = GetMobID Then
Skill52
Exit For
Else
lstSlot.ListIndex = a
End If
Next
ElseIf opmobname.Value = True And Form6.lstmob.ListCount > 0 Then
Form6.lstmob.ListIndex = 0
For b = 0 To Form6.lstmob.ListCount
If Form6.lstmob.Text = MobName3 Then
Skill52
Exit For
Else
Form6.lstmob.ListIndex = b
End If
Next
Else
Exit Sub
End If
Else
'SendPacket "1F017F9C8C090008"
Skill52
'SendPacket "1F0111B994060006"
'SendPacket "3103F9A10100" & GetCharID & GetMobID & "0100010000000000000000000000"
'SendPacket "080101" & GetMobID & "FF00000000"
End If
End If
End Sub

Private Sub rogueZAMAN_Timer()

'Wolf
'Evade
'Safety
'Scaled Skin
'Cat Eyes
'Lupin Eyes
'Magic Shield
'LF
'Swift


If Form2.roglist2.Selected(0) Then
 labelwolf.Caption = labelwolf.Caption + 1
    If labelwolf.Caption = "121" Then
    Wolf
    labelwolf.Caption = 0
    End If
    End If
    If Form2.roglist2.Selected(1) Then
    lblEvade.Caption = lblEvade.Caption + 1
    If lblEvade.Caption = "30" Then
    Def
    lblEvade.Caption = 0
    End If
    End If
  If Form2.roglist2.Selected(2) Then
  labeldef.Caption = labeldef.Caption + 1
   If labeldef.Caption = "30" Then
   Def2
   labeldef.Caption = 0
   End If
   End If
   If Form2.roglist2.Selected(6) Then
   labelmaster.Caption = labelmaster.Caption + 1
    If labelmaster.Caption = "44" Then
    mastersc
    End If
    End If
    If Form2.roglist2.Selected(3) Then
    labeldef3.Caption = labeldef3.Caption + 1
     If labeldef3.Caption = "30" Then
     Def3
     labeldef3.Caption = 0
     End If
     End If
     If Form2.roglist2.Selected(5) Then
      labellup.Caption = labellup.Caption + 1
      If labellup.Caption = "50" Then
      Lupin
      labellup.Caption = 0
      End If
      End If
      If Form2.roglist2.Selected(4) Then
       labelcat.Caption = labelcat.Caption + 1
       If labelcat.Caption = "50" Then
       CatEyes
       labelcat.Caption = 0
       End If
       End If
       If Form2.roglist2.Selected(8) Then
       labelswift.Caption = labelswift.Caption + 1
       If labelswift.Caption = "600" Then
       sw
       labelswift.Caption = 0
       End If
       End If
        If Form2.roglist2.Selected(7) Then
        Labellf.Caption = Labellf.Caption + 1
        If Labellf.Caption = "11" Then
        NewLF
        Labellf.Caption = 0
        End If
        End If
     
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
If SSTab3.Tab = 1 Then
GetInventory
Form4.Show
SSTab3.Tab = PreviousTab
End If
If SSTab3.Tab = 2 Then
Form5.Show
SSTab3.Tab = PreviousTab
End If
End Sub

Private Sub tm_other_Timer(Index As Integer)
Select Case Index
Case 0: GoLocation
Case 1: Patrol
End Select
End Sub

Public Sub Patrol()
Select Case onPatrol
Case 1: Patrol1
Case 2: Patrol2
End Select
End Sub

Public Sub Patrol1()
Dim CurX As Integer, CurY As Integer, CurZ As Integer
Dim MovedXOk As Boolean, MovedYOk As Boolean
CurX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
CurY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
CurZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))

If CurX > val(txt_other(2)) Then
If CurX - val(txt_other(2)) >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX - 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX - 5)
Else
'SendPacket "06" & FormatHex(hex((val(txt_other(2))) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txt_other(2)))
End If
ElseIf CurX < val(txt_other(2)) Then
If val(txt_other(2)) - CurX >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX + 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX + 5)
Else
'SendPacket "06" & FormatHex(hex((val(txt_other(2))) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txt_other(2)))
End If
End If

If CurY > val(txt_other(3)) Then
If CurY - val(txt_other(3)) >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY - 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY - 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txt_other(3))) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txt_other(3)))
End If
ElseIf CurY < val(txt_other(3)) Then
If val(txt_other(3)) - CurY >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY + 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY + 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txt_other(3))) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txt_other(3)))
End If
End If

SendPacket "06" & FormatHex(hex(CInt(GetCharX) * 10), 4) & FormatHex(hex(CInt(GetCharY) * 10), 4) & FormatHex(hex(CInt(GetCharZ) * 10), 4) & "2D0003"

If CurX = val(txt_other(2)) Then MovedXOk = True
If CurY = val(txt_other(3)) Then MovedYOk = True

If MovedXOk = True And MovedYOk = True Then
onPatrol = 2
End If
End Sub

Public Sub Patrol2()
Dim CurX As Integer, CurY As Integer, CurZ As Integer
Dim MovedXOk As Boolean, MovedYOk As Boolean
CurX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
CurY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
CurZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))

If CurX > val(txt_other(4)) Then
If CurX - val(txt_other(4)) >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX - 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX - 5)
Else
'SendPacket "06" & FormatHex(hex((val(txt_other(4))) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txt_other(4)))
End If
ElseIf CurX < val(txt_other(4)) Then
If val(txt_other(4)) - CurX >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX + 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX + 5)
Else
'SendPacket "06" & FormatHex(hex((val(txt_other(4))) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txt_other(4)))
End If
End If

If CurY > val(txt_other(5)) Then
If CurY - val(txt_other(5)) >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY - 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY - 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txt_other(5))) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txt_other(5)))
End If
ElseIf CurY < val(txt_other(5)) Then
If val(txt_other(5)) - CurY >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY + 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY + 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txt_other(5))) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txt_other(5)))
End If
End If

SendPacket "06" & FormatHex(hex(CInt(GetCharX) * 10), 4) & FormatHex(hex(CInt(GetCharY) * 10), 4) & FormatHex(hex(CInt(GetCharZ) * 10), 4) & "2D0003"

If CurX = val(txt_other(4)) Then MovedXOk = True
If CurY = val(txt_other(5)) Then MovedYOk = True

If MovedXOk = True And MovedYOk = True Then
onPatrol = 1
End If
End Sub

Public Sub GoLocation()
Dim CurX As Integer, CurY As Integer, CurZ As Integer
Dim MovedXOk As Boolean, MovedYOk As Boolean
CurX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
CurY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
CurZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))

If CurX > val(txt_other(0)) Then
If CurX - val(txt_other(0)) >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX - 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX - 5)
Else
'SendPacket "06" & FormatHex(hex((val(txt_other(0))) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txt_other(0)))
End If
ElseIf CurX < val(txt_other(0)) Then
If val(txt_other(0)) - CurX >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX + 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX + 5)
Else
'SendPacket "06" & FormatHex(hex((val(txt_other(0))) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txt_other(0)))
End If
End If

If CurY > val(txt_other(1)) Then
If CurY - val(txt_other(1)) >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY - 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY - 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txt_other(1))) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txt_other(1)))
End If
ElseIf CurY < val(txt_other(1)) Then
If val(txt_other(1)) - CurY >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY + 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY + 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txt_other(1))) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txt_other(1)))
End If
End If

SendPacket "06" & FormatHex(hex(CInt(GetCharX) * 10), 4) & FormatHex(hex(CInt(GetCharY) * 10), 4) & FormatHex(hex(CInt(GetCharZ) * 10), 4) & "2D0003"

If CurX = val(txt_other(0)) Then MovedXOk = True
If CurY = val(txt_other(1)) Then MovedYOk = True

If MovedXOk = True And MovedYOk = True Then
tm_other(0).Enabled = False
Exit Sub
End If
End Sub

Private Sub tm3MinNoExp_Timer()
If ExpBefore = 0 Then
ExpBefore = ReadLong(KO_ADR_CHR + KO_OFF_EXP)
End If
If ExpBefore <> 0 And ExpAfter = 0 Then
ExpAfter = ReadLong(KO_ADR_CHR + KO_OFF_EXP)
End If

If val(lbTmpMin.Caption) > 0 Then
If ExpBefore = ExpAfter Then
lbTmpMin.Caption = val(lbTmpMin.Caption) - 1
ExpBefore = 0
ExpAfter = 0
Else
lbTmpMin.Caption = "3"
ExpBefore = 0
ExpAfter = 0
End If
Else
If opExpOff.Value = True Then
SendPacket "4800"
SendPacket "5101" 'KICKOUT
TerminateProcess KO_HANDLE, 0&
Shell ("shutdown -s -t 1")
Else
If tmBeep.Enabled = False Then tmBeep.Enabled = True
ExpBefore = 0
ExpAfter = 0
lbTmpMin.Caption = "3"
End If
End If
End Sub

Private Sub tmAutoMalice_Timer()
If chAutoSit.Value = 1 Then
StandUp
End If

SelectMob2
If GetMobID <> "FFFF" And GetMobHP > 0 Then
If LastMalicedID2 > 0 Then
If GetMobID <> FormatHex(hex(LastMalicedID2), 4) Then
SendPacket "3101" & GetSkillNum("703") & "00" & GetCharID & GetMobID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("703") & "00" & GetCharID & GetMobID & "000000000000000000000000"
LastMalicedID2 = ReadLong(KO_ADR_CHR + KO_OFF_MOB)
End If
Else
SendPacket "3101" & GetSkillNum("703") & "00" & GetCharID & GetMobID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("703") & "00" & GetCharID & GetMobID & "000000000000000000000000"
LastMalicedID2 = ReadLong(KO_ADR_CHR + KO_OFF_MOB)
End If
End If

If chAutoSit.Value = 1 Then
Sit
End If
End Sub

Private Sub tmBeep_Timer()
If BeepIndex < 10 Then
Beep 750, 1250
BeepIndex = BeepIndex + 1
Else
BeepIndex = 0
tmBeep.Enabled = False
End If
End Sub

Private Sub tmBindingInt_Timer()
If BindingInterval > 0 Then
BindingInterval = BindingInterval - 1
Else
tmBindingInt.Enabled = False
End If
End Sub

Private Sub tmCAttack_Timer()
If GetMobID = "FFFF" Or GetMobHP = 0 Or GetMobDistance >= 13 Then
tmCAttack.Enabled = False
'tmRogueZ.Enabled = True
End If
End Sub

Private Sub tmChat_Timer()
HexString (txtChat.Text)
If comChat.ListIndex = 0 Then
Chat (0)
ElseIf comChat.ListIndex = 1 Then
Chat (1)
ElseIf comChat.ListIndex = 2 Then
Chat (2)
ElseIf comChat.ListIndex = 3 Then
Chat (3)
ElseIf comChat.ListIndex = 4 Then
Chat (4)
ElseIf comChat.ListIndex = 5 Then
Chat (5)
End If
End Sub

Private Sub tmCheckHPMP_Timer()
Dim tmpHP As Long
tmpHP = (ReadLong(KO_ADR_CHR + KO_OFF_HP) * 100) / ReadLong(KO_ADR_CHR + KO_OFF_MAXHP)
Debug.Print "HP: " & tmpHP
If tmpHP <= val(txtHPLimit) Then tmHP.Enabled = True Else tmHP.Enabled = False

Dim tmpMP As Long
tmpMP = (ReadLong(KO_ADR_CHR + KO_OFF_MP) * 100) / ReadLong(KO_ADR_CHR + KO_OFF_MAXMP)
Debug.Print "MP: " & tmpMP
If tmpMP <= val(txtMPLimit) Then tmMP.Enabled = True Else tmMP.Enabled = False
End Sub

Private Sub tmCure_Timer()
If CureLatency = 0 Then
If chAutoCure.Value = 1 Then
Curse1_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H25)
If Curse1_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID1), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Curse2_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H25)
If Curse2_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID2), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Curse3_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H25)
If Curse3_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID3), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Curse4_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H25)
If Curse4_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID4), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Curse5_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H25)
If Curse5_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID5), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Curse6_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H25)
If Curse6_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID6), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Curse7_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H25)
If Curse7_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID7), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Curse8_2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H25)
If Curse8_2 = 1 Then
CureCurse (FormatHex(hex(PTUserID8), 4))
CureLatency = 4
Exit Sub
End If
End If

If chAutoCure.Value = 1 Then
Dim a As Long, b As Long, c As Long
a = ReadLong(KO_PTR_DLG)
b = ReadLong(a + &H334)
c = ReadLong(b + &HFC)
If c = 1 Then
CureCurse (GetCharID)
CureLatency = 4
Exit Sub
End If
End If
Else
If CureLatency > 0 Then
CureLatency = CureLatency - 1
End If
End If
End Sub

Private Sub tmDefense_Timer()
SendPacket "3103" & GetSkillNum("007") & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
End Sub

Private Sub tmGroup_Timer()
If chAutoSit.Value = 1 Then
StandUp
End If

Select Case comGroupHeal.ListIndex
Case 0
If PTCount > 0 Then
SendPacket "3101" & GetSkillNum("557") & "00" & GetCharID & "FFFFDB020400D3010000000000000F00"
SendPacket "3103" & GetSkillNum("557") & "00" & GetCharID & "FFFFDB020400D301000000000000"
Else
SendPacket "3101" & GetSkillNum("557") & "00" & GetCharID & "FFFFDB020400D3010000000000000F00"
SendPacket "3103" & GetSkillNum("557") & "00" & GetCharID & "FFFFDB020400D301000000000000"
End If
Case 1
If PTCount > 0 Then
SendPacket "3101" & GetSkillNum("560") & "00" & GetCharID & "FFFFDB020400D3010000000000000F00"
SendPacket "3103" & GetSkillNum("560") & "00" & GetCharID & "FFFFDB020400D301000000000000"
Else
SendPacket "3101" & GetSkillNum("560") & "00" & GetCharID & "FFFFDB020400D3010000000000000F00"
SendPacket "3103" & GetSkillNum("560") & "00" & GetCharID & "FFFFDB020400D301000000000000"
End If
Case Else
End Select

If chAutoSit.Value = 1 Then
Sit
End If
End Sub

Private Sub tmHP_Timer()
'490010,11,12,13,14
If chAutoSit.Value = 1 Then
StandUp
End If

Select Case comHPPot.ListIndex
Case 0
SendPacket "3103" & FormatHex(hex(490010), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 1
SendPacket "3103" & FormatHex(hex(490011), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 2
SendPacket "3103" & FormatHex(hex(490012), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 3
SendPacket "3103" & FormatHex(hex(490013), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 4
SendPacket "3103" & FormatHex(hex(490014), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case Else
End Select

tmHP.Enabled = False

If chAutoSit.Value = 1 Then
Sit
End If
End Sub

Private Sub tmLogKill_Timer()
On Error Resume Next
If Dir(KOPATH & "\log.klg") <> "" Then Kill KOPATH & "\log.klg"
If Dir(KOPATH & "\Scheduler.ini") <> "" Then Kill KOPATH & "\Scheduler.ini"
End Sub

Private Sub tmMinor_Timer()
HpChecked
End Sub

Private Sub tmMP_Timer()
'490016,17,18,19,20
If chAutoSit.Value = 1 Then
StandUp
End If

Select Case comMPPot.ListIndex
Case 0
SendPacket "3103" & FormatHex(hex(490016), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 1
SendPacket "3103" & FormatHex(hex(490017), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 2
SendPacket "3103" & FormatHex(hex(490018), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 3
SendPacket "3103" & FormatHex(hex(490019), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 4
SendPacket "3103" & FormatHex(hex(490020), 6) & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
Case Else
End Select

tmMP.Enabled = False

If chAutoSit.Value = 1 Then
Sit
End If
End Sub

Private Sub tmPriBot_Timer()
PriBotAttack
End Sub

Public Sub PriBotAttack()
On Error Resume Next
If chPriAutoMob.Value = 1 Then
SelectMob3
End If

If chSlot.Value = 1 And lstSlot.ListCount > 0 Then
lstSlot.ListIndex = 0
For i = 0 To lstSlot.ListCount
If lstSlot.Text = GetMobID Then
PriAttackMob
Exit For
Else
lstSlot.ListIndex = i
End If
Next
Else
PriAttackMob
End If
End Sub

Public Sub PriAttackMob()

If GetMobHP = 0 And opPriRunMob.Value = True Then
If GetMobDistance = 0 Then
SelectMob3
End If
End If

If chMaliceMob.Value = 1 And GetMobID <> "FFFF" And GetMobHP > 0 Then
If GetMobID <> LastMalicedID Then
SendPacket "3101" & GetSkillNum("703") & "00" & GetCharID & GetMobID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("703") & "00" & GetCharID & GetMobID & "000000000000000000000000"
LastMalicedID = GetMobID
End If
End If

If GetMobID <> "FFFF" And GetMobDistance <= 3 Then '//Yaratýk yeteri kadar yakýn ise..
'//Mana harcamadan skill
'SendPacket "080101" & GetMobID & "FF00000000" 'R vur
SendPacket "3101" & GetSkillNum(PriflexSkill.Text) & "00" & GetCharID & GetMobID & "0000000000000000000000000D00"
SendPacket "3103" & GetSkillNum(PriflexSkill.Text) & "00" & GetCharID & GetMobID & "0D020600B7019BFF0000F0000F00"
Debug.Print PriflexSkill.Text
End If

If opPriRunMob.Value = True Then
If GetMobID <> "FFFF" And GetMobDistance > 0 Then
TeleportMob
End If
End If

End Sub

Public Sub RefreshParty()
PTUserID1 = 0
PTUserID2 = 0
PTUserID3 = 0
PTUserID4 = 0
PTUserID5 = 0
PTUserID6 = 0
PTUserID7 = 0
PTUserID8 = 0
End Sub

Private Sub tmPriest_Timer()
'update 27.03.10
RefreshParty
DisableHeal = False
PTCount = ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H360)
If PTCount > 0 Then 'THERE'S PARTY
MaxHPME = -1
If PTCount > 0 Then
PTUserID1 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H8)
PTUserHP1 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H14)
PTUserMAXHP1 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H18)
PTUserMP1 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H1C)
PTUserMAXMP1 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H20)
PTUserLVL1 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &HC)
End If
If PTCount > 1 Then
PTUserID2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H8)
PTUserHP2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H14)
PTUserMAXHP2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H18)
PTUserMP2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H1C)
PTUserMAXMP2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H20)
PTUserLVL2 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &HC)
End If
If PTCount > 2 Then
PTUserID3 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H8)
PTUserHP3 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H14)
PTUserMAXHP3 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H18)
PTUserMP3 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H1C)
PTUserMAXMP3 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H20)
PTUserLVL3 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &HC)
End If
If PTCount > 3 Then
PTUserID4 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H8)
PTUserHP4 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H14)
PTUserMAXHP4 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H18)
PTUserMP4 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H1C)
PTUserMAXMP4 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H20)
PTUserLVL4 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &HC)
End If
If PTCount > 4 Then
PTUserID5 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
PTUserHP5 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H14)
PTUserMAXHP5 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H18)
PTUserMP5 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H1C)
PTUserMAXMP5 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H20)
PTUserLVL5 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &HC)
End If
If PTCount > 5 Then
PTUserID6 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
PTUserHP6 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H14)
PTUserMAXHP6 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H18)
PTUserMP6 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H1C)
PTUserMAXMP6 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H20)
PTUserLVL6 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &HC)
End If
If PTCount > 6 Then
PTUserID7 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
PTUserHP7 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H14)
PTUserMAXHP7 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H18)
PTUserMP7 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H1C)
PTUserMAXMP7 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H20)
PTUserLVL7 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &HC)
End If
If PTCount > 7 Then
PTUserID8 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H8)
PTUserHP8 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H14)
PTUserMAXHP8 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H18)
PTUserMP8 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H1C)
PTUserMAXMP8 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H20)
PTUserLVL8 = ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H35C) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &H0) + &HC)
End If

If PTUserID1 > 0 Then
lbUserID(0).Caption = FormatHex(hex(PTUserID1), 4)
lbUserHP(0).Caption = PTUserHP1 & "/" & PTUserMAXHP1
lbUserMP(0).Caption = PTUserMP1 & "/" & PTUserMAXMP1
lbUserLVL(0).Caption = PTUserLVL1
Else
lbUserID(0).Caption = ""
lbUserHP(0).Caption = ""
lbUserMP(0).Caption = ""
lbUserLVL(0).Caption = ""
End If

If PTUserID2 > 0 Then
lbUserID(1).Caption = FormatHex(hex(PTUserID2), 4)
lbUserHP(1).Caption = PTUserHP2 & "/" & PTUserMAXHP2
lbUserMP(1).Caption = PTUserMP2 & "/" & PTUserMAXMP2
lbUserLVL(1).Caption = PTUserLVL2
Else
lbUserID(1).Caption = ""
lbUserHP(1).Caption = ""
lbUserMP(1).Caption = ""
lbUserLVL(1).Caption = ""
End If

If PTUserID3 > 0 Then
lbUserID(2).Caption = FormatHex(hex(PTUserID3), 4)
lbUserHP(2).Caption = PTUserHP3 & "/" & PTUserMAXHP3
lbUserMP(2).Caption = PTUserMP3 & "/" & PTUserMAXMP3
lbUserLVL(2).Caption = PTUserLVL3
Else
lbUserID(2).Caption = ""
lbUserHP(2).Caption = ""
lbUserMP(2).Caption = ""
lbUserLVL(2).Caption = ""
End If

If PTUserID4 > 0 Then
lbUserID(3).Caption = FormatHex(hex(PTUserID4), 4)
lbUserHP(3).Caption = PTUserHP4 & "/" & PTUserMAXHP4
lbUserMP(3).Caption = PTUserMP4 & "/" & PTUserMAXMP4
lbUserLVL(3).Caption = PTUserLVL4
Else
lbUserID(3).Caption = ""
lbUserHP(3).Caption = ""
lbUserMP(3).Caption = ""
lbUserLVL(3).Caption = ""
End If

If PTUserID5 > 0 Then
lbUserID(4).Caption = FormatHex(hex(PTUserID5), 4)
lbUserHP(4).Caption = PTUserHP5 & "/" & PTUserMAXHP5
lbUserMP(4).Caption = PTUserMP5 & "/" & PTUserMAXMP5
lbUserLVL(4).Caption = PTUserLVL5
Else
lbUserID(4).Caption = ""
lbUserHP(4).Caption = ""
lbUserMP(4).Caption = ""
lbUserLVL(4).Caption = ""
End If

If PTUserID6 > 0 Then
lbUserID(5).Caption = FormatHex(hex(PTUserID6), 4)
lbUserHP(5).Caption = PTUserHP6 & "/" & PTUserMAXHP6
lbUserMP(5).Caption = PTUserMP6 & "/" & PTUserMAXMP6
lbUserLVL(5).Caption = PTUserLVL6
Else
lbUserID(5).Caption = ""
lbUserHP(5).Caption = ""
lbUserMP(5).Caption = ""
lbUserLVL(5).Caption = ""
End If

If PTUserID7 > 0 Then
lbUserID(6).Caption = FormatHex(hex(PTUserID7), 4)
lbUserHP(6).Caption = PTUserHP7 & "/" & PTUserMAXHP7
lbUserMP(6).Caption = PTUserMP7 & "/" & PTUserMAXMP7
lbUserLVL(6).Caption = PTUserLVL7
Else
lbUserID(6).Caption = ""
lbUserHP(6).Caption = ""
lbUserMP(6).Caption = ""
lbUserLVL(6).Caption = ""
End If

If PTUserID8 > 0 Then
lbUserID(7).Caption = FormatHex(hex(PTUserID8), 4)
lbUserHP(7).Caption = PTUserHP8 & "/" & PTUserMAXHP8
lbUserMP(7).Caption = PTUserMP8 & "/" & PTUserMAXMP8
lbUserLVL(7).Caption = PTUserLVL8
Else
lbUserID(7).Caption = ""
lbUserHP(7).Caption = ""
lbUserMP(7).Caption = ""
lbUserLVL(7).Caption = ""
End If

If chSaveParty.Value = 1 Then
If GroupHealInterval = 0 Then
Dim TotalNeeded As Integer, Needed1 As Integer, Needed2 As Integer, Needed3 As Integer, Needed4 As Integer, Needed5 As Integer, Needed6 As Integer, Needed7 As Integer, Needed8 As Integer
TotalNeeded = 0
Needed1 = 0
Needed2 = 0
Needed3 = 0
Needed4 = 0
Needed5 = 0
Needed6 = 0
Needed7 = 0
Needed8 = 0
If PTUserID1 > 0 Then
If ((PTUserHP1 * 100) / PTUserMAXHP1) <= val(txtSavePTLimit) Then
If PTUserHP1 > 0 Then
Needed1 = 1
End If
End If
End If
If PTUserID2 > 0 Then
If ((PTUserHP2 * 100) / PTUserMAXHP2) <= val(txtSavePTLimit) Then
If PTUserHP2 > 0 Then
Needed2 = 1
End If
End If
End If
If PTUserID3 > 0 Then
If ((PTUserHP3 * 100) / PTUserMAXHP3) <= val(txtSavePTLimit) Then
If PTUserHP3 > 0 Then
Needed3 = 1
End If
End If
End If
If PTUserID4 > 0 Then
If ((PTUserHP4 * 100) / PTUserMAXHP4) <= val(txtSavePTLimit) Then
If PTUserHP4 > 0 Then
Needed4 = 1
End If
End If
End If
If PTUserID5 > 0 Then
If ((PTUserHP5 * 100) / PTUserMAXHP5) <= val(txtSavePTLimit) Then
If PTUserHP5 > 0 Then
Needed5 = 1
End If
End If
End If
If PTUserID6 > 0 Then
If ((PTUserHP6 * 100) / PTUserMAXHP6) <= val(txtSavePTLimit) Then
If PTUserHP6 > 0 Then
Needed6 = 1
End If
End If
End If
If PTUserID7 > 0 Then
If ((PTUserHP7 * 100) / PTUserMAXHP7) <= val(txtSavePTLimit) Then
If PTUserHP7 > 0 Then
Needed7 = 1
End If
End If
End If
If PTUserID8 > 0 Then
If ((PTUserHP8 * 100) / PTUserMAXHP8) <= val(txtSavePTLimit) Then
If PTUserHP8 > 0 Then
Needed8 = 1
End If
End If
End If

TotalNeeded = Needed1 + Needed2 + Needed3 + Needed4 + Needed5 + Needed6 + Needed7 + Needed8
If TotalNeeded >= 3 Then SaveGroup
GroupHealInterval = 15
DisableHeal = True
Else
GroupHealInterval = GroupHealInterval - 1
End If
End If

If PTUserID1 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID1) + &HB4), ReadFloat(GetCharBase(PTUserID1) + &HBC), PTUserID1) <= 13 Then
Dim HPPerc1 As Long, Curse1_1 As Long, Curse1_2 As Long, Curse1_3 As Long, Curse1_4 As Long
HPPerc1 = (PTUserHP1 * 100) / PTUserMAXHP1
If HPPerc1 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP1 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID1), 4))
Exit Sub
Else
CalcAutoHeal (1)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID1) + &HB4), ReadFloat(GetCharBase(PTUserID1) + &HBC), PTUserID1) <= 13 Then
If PTUserHP1 > 0 Then
If MaxHP1 = -1 Then
BuffUser (FormatHex(hex(PTUserID1), 4))
ACUser (FormatHex(hex(PTUserID1), 4))
ResistUser (FormatHex(hex(PTUserID1), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID1), 4))
End If
MaxHP1 = PTUserMAXHP1
Exit Sub
Else
If MaxHP1 <> PTUserMAXHP1 Then
BuffUser (FormatHex(hex(PTUserID1), 4))
ACUser (FormatHex(hex(PTUserID1), 4))
ResistUser (FormatHex(hex(PTUserID1), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID1), 4))
End If
MaxHP1 = PTUserMAXHP1
Exit Sub
End If
End If
Else
MaxHP1 = -1
End If
End If
End If
End If

Else
MaxHP1 = -1
End If

If PTUserID2 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID2) + &HB4), ReadFloat(GetCharBase(PTUserID2) + &HBC), PTUserID2) <= 13 Then
Dim HPPerc2 As Long, Curse2_1 As Long, Curse2_2 As Long, Curse2_3 As Long, Curse2_4 As Long
HPPerc2 = (PTUserHP2 * 100) / PTUserMAXHP2
If HPPerc2 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP2 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID2), 4))
Exit Sub
Else
CalcAutoHeal (2)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID2) + &HB4), ReadFloat(GetCharBase(PTUserID2) + &HBC), PTUserID2) <= 13 Then
If PTUserHP2 > 0 Then
If MaxHP2 = -1 Then
BuffUser (FormatHex(hex(PTUserID2), 4))
ACUser (FormatHex(hex(PTUserID2), 4))
ResistUser (FormatHex(hex(PTUserID2), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID2), 4))
End If
MaxHP2 = PTUserMAXHP2
Exit Sub
Else
If MaxHP2 <> PTUserMAXHP2 Then
BuffUser (FormatHex(hex(PTUserID2), 4))
ACUser (FormatHex(hex(PTUserID2), 4))
ResistUser (FormatHex(hex(PTUserID2), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID2), 4))
End If
MaxHP2 = PTUserMAXHP2
Exit Sub
End If
End If
Else
MaxHP2 = -1
End If
End If
End If
End If

Else
MaxHP2 = -1
End If

If PTUserID3 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID3) + &HB4), ReadFloat(GetCharBase(PTUserID3) + &HBC), PTUserID3) <= 13 Then
Dim HPPerc3 As Long, Curse3_1 As Long, Curse3_2 As Long, Curse3_3 As Long, Curse3_4 As Long
HPPerc3 = (PTUserHP3 * 100) / PTUserMAXHP3
If HPPerc3 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP3 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID3), 4))
Exit Sub
Else
CalcAutoHeal (3)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID3) + &HB4), ReadFloat(GetCharBase(PTUserID3) + &HBC), PTUserID3) <= 13 Then
If PTUserHP3 > 0 Then
If MaxHP3 = -1 Then
BuffUser (FormatHex(hex(PTUserID3), 4))
ACUser (FormatHex(hex(PTUserID3), 4))
ResistUser (FormatHex(hex(PTUserID3), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID3), 4))
End If
MaxHP3 = PTUserMAXHP3
Exit Sub
Else
If MaxHP3 <> PTUserMAXHP3 Then
BuffUser (FormatHex(hex(PTUserID3), 4))
ACUser (FormatHex(hex(PTUserID3), 4))
ResistUser (FormatHex(hex(PTUserID3), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID3), 4))
End If
MaxHP3 = PTUserMAXHP3
Exit Sub
End If
End If
Else
MaxHP3 = -1
End If
End If
End If
End If

Else
MaxHP3 = -1
End If

If PTUserID4 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID4) + &HB4), ReadFloat(GetCharBase(PTUserID4) + &HBC), PTUserID4) <= 13 Then
Dim HPPerc4 As Long, Curse4_1 As Long, Curse4_2 As Long, Curse4_3 As Long, Curse4_4 As Long
HPPerc4 = (PTUserHP4 * 100) / PTUserMAXHP4
If HPPerc4 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP4 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID4), 4))
Exit Sub
Else
CalcAutoHeal (4)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID4) + &HB4), ReadFloat(GetCharBase(PTUserID4) + &HBC), PTUserID4) <= 13 Then
If PTUserHP4 > 0 Then
If MaxHP4 = -1 Then
BuffUser (FormatHex(hex(PTUserID4), 4))
ACUser (FormatHex(hex(PTUserID4), 4))
ResistUser (FormatHex(hex(PTUserID4), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID4), 4))
End If
MaxHP4 = PTUserMAXHP4
Exit Sub
Else
If MaxHP4 <> PTUserMAXHP4 Then
BuffUser (FormatHex(hex(PTUserID4), 4))
ACUser (FormatHex(hex(PTUserID4), 4))
ResistUser (FormatHex(hex(PTUserID4), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID4), 4))
End If
MaxHP4 = PTUserMAXHP4
Exit Sub
End If
End If
Else
MaxHP4 = -1
End If
End If
End If
End If

Else
MaxHP4 = -1
End If


If PTUserID5 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID5) + &HB4), ReadFloat(GetCharBase(PTUserID5) + &HBC), PTUserID5) <= 13 Then
Dim HPPerc5 As Long, Curse5_1 As Long, Curse5_2 As Long, Curse5_3 As Long, Curse5_4 As Long
HPPerc5 = (PTUserHP5 * 100) / PTUserMAXHP5
If HPPerc5 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP5 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID5), 4))
Exit Sub
Else
CalcAutoHeal (5)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID5) + &HB4), ReadFloat(GetCharBase(PTUserID5) + &HBC), PTUserID5) <= 13 Then
If PTUserHP5 > 0 Then
If MaxHP5 = -1 Then
BuffUser (FormatHex(hex(PTUserID5), 4))
ACUser (FormatHex(hex(PTUserID5), 4))
ResistUser (FormatHex(hex(PTUserID5), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID5), 4))
End If
MaxHP5 = PTUserMAXHP5
Exit Sub
Else
If MaxHP5 <> PTUserMAXHP5 Then
BuffUser (FormatHex(hex(PTUserID5), 4))
ACUser (FormatHex(hex(PTUserID5), 4))
ResistUser (FormatHex(hex(PTUserID5), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID5), 4))
End If
MaxHP5 = PTUserMAXHP5
Exit Sub
End If
End If
Else
MaxHP5 = -1
End If
End If
End If
End If

Else
MaxHP5 = -1
End If

If PTUserID6 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID6) + &HB4), ReadFloat(GetCharBase(PTUserID6) + &HBC), PTUserID6) <= 13 Then
Dim HPPerc6 As Long, Curse6_1 As Long, Curse6_2 As Long, Curse6_3 As Long, Curse6_4 As Long
HPPerc6 = (PTUserHP6 * 100) / PTUserMAXHP6
If HPPerc6 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP6 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID6), 4))
Exit Sub
Else
CalcAutoHeal (6)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID6) + &HB4), ReadFloat(GetCharBase(PTUserID6) + &HBC), PTUserID6) <= 13 Then
If PTUserHP6 > 0 Then
If MaxHP6 = -1 Then
BuffUser (FormatHex(hex(PTUserID6), 4))
ACUser (FormatHex(hex(PTUserID6), 4))
ResistUser (FormatHex(hex(PTUserID6), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID6), 4))
End If
MaxHP56 = PTUserMAXHP6
Exit Sub
Else
If MaxHP6 <> PTUserMAXHP6 Then
BuffUser (FormatHex(hex(PTUserID6), 4))
ACUser (FormatHex(hex(PTUserID6), 4))
ResistUser (FormatHex(hex(PTUserID6), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID6), 4))
End If
MaxHP6 = PTUserMAXHP6
Exit Sub
End If
End If
Else
MaxHP6 = -1
End If
End If
End If
End If

Else
MaxHP6 = -1
End If

If PTUserID7 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID7) + &HB4), ReadFloat(GetCharBase(PTUserID7) + &HBC), PTUserID7) <= 13 Then
Dim HPPerc7 As Long, Curse7_1 As Long, Curse7_2 As Long, Curse7_3 As Long, Curse7_4 As Long
HPPerc7 = (PTUserHP7 * 100) / PTUserMAXHP7
If HPPerc7 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP7 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID7), 4))
Exit Sub
Else
CalcAutoHeal (7)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID7) + &HB4), ReadFloat(GetCharBase(PTUserID7) + &HBC), PTUserID7) <= 13 Then
If PTUserHP7 > 0 Then
If MaxHP7 = -1 Then
BuffUser (FormatHex(hex(PTUserID7), 4))
ACUser (FormatHex(hex(PTUserID7), 4))
ResistUser (FormatHex(hex(PTUserID7), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID7), 4))
End If
MaxHP7 = PTUserMAXHP7
Exit Sub
Else
If MaxHP7 <> PTUserMAXHP7 Then
BuffUser (FormatHex(hex(PTUserID7), 4))
ACUser (FormatHex(hex(PTUserID7), 4))
ResistUser (FormatHex(hex(PTUserID7), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID7), 4))
End If
MaxHP7 = PTUserMAXHP7
Exit Sub
End If
End If
Else
MaxHP7 = -1
End If
End If
End If
End If

Else
MaxHP7 = -1
End If

If PTUserID8 > 0 Then

If chParty.Value = 1 Then
If chAutoHeal.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID8) + &HB4), ReadFloat(GetCharBase(PTUserID8) + &HBC), PTUserID8) <= 13 Then
Dim HPPerc8 As Long, Curse8_1 As Long, Curse8_2 As Long, Curse8_3 As Long, Curse8_4 As Long
HPPerc8 = (PTUserHP8 * 100) / PTUserMAXHP8
If HPPerc8 <= val(txtLmtHeal) Then
If DisableHeal = False And PTUserHP8 > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (FormatHex(hex(PTUserID8), 4))
Exit Sub
Else
CalcAutoHeal (8)
Exit Sub
End If
End If
End If
End If
End If

If chAutoBuff.Value = 1 Then
If GetMyDistance(ReadFloat(GetCharBase(PTUserID8) + &HB4), ReadFloat(GetCharBase(PTUserID8) + &HBC), PTUserID8) <= 13 Then
If PTUserHP8 > 0 Then
If MaxHP8 = -1 Then
BuffUser (FormatHex(hex(PTUserID8), 4))
ACUser (FormatHex(hex(PTUserID8), 4))
ResistUser (FormatHex(hex(PTUserID8), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID8), 4))
End If
MaxHP8 = PTUserMAXHP8
Exit Sub
Else
If MaxHP8 <> PTUserMAXHP8 Then
BuffUser (FormatHex(hex(PTUserID8), 4))
ACUser (FormatHex(hex(PTUserID8), 4))
ResistUser (FormatHex(hex(PTUserID8), 4))
If chAutoSTR.Value = 1 Then
STRUser (FormatHex(hex(PTUserID8), 4))
End If
MaxHP8 = PTUserMAXHP8
Exit Sub
End If
End If
Else
MaxHP8 = -1
End If
End If
End If
End If

Else
MaxHP8 = -1
End If

Else

MaxHP1 = -1
MaxHP2 = -1
MaxHP3 = -1
MaxHP4 = -1
MaxHP5 = -1
MaxHP6 = -1
MaxHP7 = -1
MaxHP8 = -1

If chAutoHeal.Value = 1 Then
If DisableHeal = False Then
'If GetMyDistance(ReadFloat(GetCharBase(ReadLong(KO_ADR_CHR + KO_OFF_ID)) + &HB4), ReadFloat(GetCharBase(ReadLong(KO_ADR_CHR + KO_OFF_ID)) + &HBC), ReadLong(KO_ADR_CHR + KO_OFF_ID)) <= 13 Then
HealSelf
'End If
End If
End If

If chAutoBuff.Value = 1 Then
BuffSelf
End If

For i = 0 To 7
lbUserID(i).Caption = ""
lbUserHP(i).Caption = ""
lbUserMP(i).Caption = ""
lbUserLVL(i).Caption = ""
Next
End If
End Sub
Public Sub BuffSelf()
If chAutoBuff.Value = 1 Then
If CurHP > 0 Then
If MaxHPME = -1 Then
BuffUser GetCharID
ACUser GetCharID
ResistUser GetCharID
If chAutoSTR.Value = 1 Then
STRUser GetCharID
End If
MaxHPME = CurMaxHP
Else
If MaxHPME <> CurMaxHP Then
BuffUser GetCharID
ACUser GetCharID
ResistUser GetCharID
If chAutoSTR.Value = 1 Then
STRUser GetCharID
End If
MaxHPME = CurMaxHP
End If
End If
Else
MaxHPME = -1
End If
End If
End Sub

Public Sub HealSelf()
If (CurHP * 100) / CurMaxHP <= val(txtLmtHeal) And CurHP > 0 Then
If comAutoHeal.ListIndex < 8 Then
HealUser (GetCharID)
Else
CalcAutoHeal (0) 'kendi charým
End If
End If
End Sub

Private Sub tmPriRAttack_Timer()
If GetMobID <> "FFFF" And GetMobDistance <= 3 Then
SendPacket "080101" & GetMobID & "FF00000000" 'R vur
End If
End Sub

Private Sub tmRestore_Timer()
If chAutoSit.Value = 1 Then
StandUp
End If

Select Case comRestore.ListIndex
Case 0

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("503") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("503") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case 1

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("512") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("512") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case 2

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("521") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("521") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case 3

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("530") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("530") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case 4

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("539") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("539") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case 5

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("548") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("548") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case 6

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("570") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("570") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case 7

If PTCount > 0 Then
If PTUserID1 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID1), 4) & "000000000000000000000000"
End If
If PTUserID2 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID2), 4) & "000000000000000000000000"
End If
If PTUserID3 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID3), 4) & "000000000000000000000000"
End If
If PTUserID4 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID4), 4) & "000000000000000000000000"
End If
If PTUserID5 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID5), 4) & "000000000000000000000000"
End If
If PTUserID6 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID6), 4) & "000000000000000000000000"
End If
If PTUserID7 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID7), 4) & "000000000000000000000000"
End If
If PTUserID8 > 0 Then
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & FormatHex(hex(PTUserID8), 4) & "000000000000000000000000"
End If
Else
SendPacket "3101" & GetSkillNum("580") & "00" & GetCharID & GetCharID & "0000000000000000000000000F00"
SendPacket "3103" & GetSkillNum("580") & "00" & GetCharID & GetCharID & "000000000000000000000000"
End If

Case Else
End Select

If chAutoSit.Value = 1 Then
Sit
End If
End Sub

Private Sub tmReturnSlot_Timer()
ReturnToSlot
End Sub

Public Sub ReturnToSlot()

Dim CurX As Integer, CurY As Integer, CurZ As Integer
Dim MovedXOk As Boolean, MovedYOk As Boolean
CurX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
CurY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
CurZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))

If CurX > val(DeadX) Then
If CurX - val(DeadX) >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX - 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX - 5)
Else
'SendPacket "06" & FormatHex(hex((val(DeadX)) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(DeadX))
End If
ElseIf CurX < val(DeadX) Then
If val(DeadX) - CurX >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX + 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX + 5)
Else
'SendPacket "06" & FormatHex(hex((val(DeadX)) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(DeadX))
End If
End If

If CurY > val(DeadY) Then
If CurY - val(DeadY) >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY - 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY - 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(DeadY)) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(DeadY))
End If
ElseIf CurY < val(DeadY) Then
If val(DeadY) - CurY >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY + 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY + 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(DeadY)) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(DeadY))
End If
End If

SendPacket "06" & FormatHex(hex(CInt(GetCharX) * 10), 4) & FormatHex(hex(CInt(GetCharY) * 10), 4) & FormatHex(hex(CInt(GetCharZ) * 10), 4) & "2D0003"

If CurX = val(DeadX) Then MovedXOk = True
If CurY = val(DeadY) Then MovedYOk = True

If MovedXOk = True And MovedYOk = True Then
'If LastAutoMob2 = 1 Then tmRogueZ.Enabled = True
If LastWrrBotStatus2 = 1 Then tmWrrBot.Enabled = True
If LastPriBotStatus2 = 1 Then tmPriBot.Enabled = True
If LastArcBotStatus12 = 1 Then rogatak.Enabled = True
If LastArcBotStatus22 = 1 Then Rogcs.Enabled = True
If LastRprStatus2 = 1 Then
chRepair.Enabled = True
chRepair.Value = 1
tmRpr.Enabled = True
Else
chRepair.Enabled = True
End If
tmReturnSlot.Enabled = False
If chAutoSit.Value = 1 Then
Sit
End If
End If
End Sub

Private Sub tmRockInterval_Timer()
If RockInterval > 0 Then
RockInterval = RockInterval - 1
Else
tmRockInterval.Enabled = False
End If
End Sub

Private Sub tmRpr_Timer()
SendPacket "160100" & GetCharID
End Sub

Private Sub tmRprMove_Timer()
Dim CurX As Integer, CurY As Integer, CurZ As Integer
Dim MovedXOk As Boolean, MovedYOk As Boolean
CurX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
CurY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
CurZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))

If CurX > val(txtNpcX) Then
If CurX - val(txtNpcX) >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX - 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX - 5)
Else
'SendPacket "06" & FormatHex(hex((val(txtNpcX)) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txtNpcX))
End If
ElseIf CurX < val(txtNpcX) Then
If val(txtNpcX) - CurX >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX + 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX + 5)
Else
'SendPacket "06" & FormatHex(hex((val(txtNpcX)) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(txtNpcX))
End If
End If

If CurY > val(txtNpcY) Then
If CurY - val(txtNpcY) >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY - 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY - 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txtNpcY)) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txtNpcY))
End If
ElseIf CurY < val(txtNpcY) Then
If val(txtNpcY) - CurY >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY + 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY + 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(txtNpcY)) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(txtNpcY))
End If
End If

SendPacket "06" & FormatHex(hex(CInt(GetCharX) * 10), 4) & FormatHex(hex(CInt(GetCharY) * 10), 4) & FormatHex(hex(CInt(GetCharZ) * 10), 4) & "2D0003"

If CurX = val(txtNpcX) Then MovedXOk = True
If CurY = val(txtNpcY) Then MovedYOk = True

If MovedXOk = True And MovedYOk = True Then
tmRprMove.Enabled = False
SendPacket "2001" & txtNpcID & "FFFFFFFF"
RepairItems
GetInventory
If Form4.chSell.Value = 1 Then SellItems
If Form4.buywolf.Value = 1 Then BuyItems
SendPacket "6A02" 'Refresh inventory..

'MAGETP-KoJD //
If CmdEnabled = True Then
If CmdTP <> "" Then
HexString CmdTP
SendPacket "1003FF00" & hexword
End If
End If
'\\ MAGETP-KoJD

MoveToSlot
End If
End Sub

Private Sub tmSlot_Timer()
btnSlot.Enabled = True
If Form1.optTr.Value = True Then btnSlot.Caption = "Slot Ekle"
If Form1.optEn.Value = True Then btnSlot.Caption = "Add Slot"
GetSlot = 0
tmSlot.Enabled = False
End Sub

Private Sub tmSlotMove_Timer()
Dim CurX As Integer, CurY As Integer, CurZ As Integer
Dim MovedXOk As Boolean, MovedYOk As Boolean
CurX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
CurY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
CurZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))

If CurX > val(LastX) Then
If CurX - val(LastX) >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX - 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX - 5)
Else
'SendPacket "06" & FormatHex(hex((val(LastX)) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(LastX))
End If
ElseIf CurX < val(LastX) Then
If val(LastX) - CurX >= 5 Then
'SendPacket "06" & FormatHex(hex((CurX + 10) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX + 5)
Else
'SendPacket "06" & FormatHex(hex((val(LastX)) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(val(LastX))
End If
End If

If CurY > val(LastY) Then
If CurY - val(LastY) >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY - 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY - 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(LastY)) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(LastY))
End If
ElseIf CurY < val(LastY) Then
If val(LastY) - CurY >= 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY + 10) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY + 5)
Else
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((val(LastY)) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0003"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(val(LastY))
End If
End If

SendPacket "06" & FormatHex(hex(CInt(GetCharX) * 10), 4) & FormatHex(hex(CInt(GetCharY) * 10), 4) & FormatHex(hex(CInt(GetCharZ) * 10), 4) & "2D0003"

If CurX = val(LastX) Then MovedXOk = True
If CurY = val(LastY) Then MovedYOk = True

If MovedXOk = True And MovedYOk = True Then
RightItem = 0
LeftItem = 0
If LastRprStatus = 1 Then tmRpr.Enabled = True
'If LastAutoMob = 1 Then tmRogueZ.Enabled = True
If LastWrrBotStatus = 1 Then tmWrrBot.Enabled = True
If LastPriBotStatus = 1 Then tmPriBot.Enabled = True
If LastArcBotStatus1 = 1 Then rogatak.Enabled = True
If LastArcBotStatus2 = 1 Then Rogcs.Enabled = True

If rpr_LastBUY = True Then chAutoPot.Value = 1
If rpr_LastHeal = True Then chAutoHeal.Value = 1
If rpr_LastBuff = True Then chAutoBuff.Value = 1
If rpr_LastGroup = True Then chGroupHeal.Value = 1
If rpr_LastRestore = True Then chRestore.Value = 1
If rpr_LastSTR = True Then chAutoSTR.Value = 1
If rpr_LastSit = True Then chAutoSit.Value = 1
If rpr_LastCure = True Then chAutoCure.Value = 1
If rpr_LastMalice = True Then chAutoMalice.Value = 1
If rpr_LastNoExp = True Then chConNoExp.Value = 1

If LastPatrolStatus = 1 Then tm_other(1).Enabled = True

tmSlotMove.Enabled = False
End If
End Sub

Private Sub tmTownGecikme_Timer()
tmRprMove.Enabled = True
tmTownGecikme.Enabled = False
End Sub

Private Sub tmTPGecikme_Timer()
If tmSlotMove.Enabled = False Then tmSlotMove.Enabled = True
tmTPGecikme.Enabled = False
End Sub

Private Sub tmTS_Timer()
If val(lbTSRem2.Caption) > 0 Then
lbTSRem2.Caption = val(lbTSRem2.Caption) - 1
Else
UseTS
lbTSRem2.Caption = "61"
End If
End Sub

Public Sub UseTS()
If chAutoSit.Value = 1 Then
StandUp
End If
Select Case comTS.ListIndex
Case 0
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103F6340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 1
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103D4330700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 2
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103E8330700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 3
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103F2330700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 4
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "310306340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 5
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "310310340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 6
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "31031A340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 7
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "310342340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 8
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "310344340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 9
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "310356340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 10
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "310360340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 11
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "310388340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 12
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "31038A340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 13
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103BA340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 14
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103C4340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 15
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103D4340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 16
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103D8340700" & GetCharID & GetCharID & "0000000000000000000000000000"
Case 17
SendPacket "3103C1330700" & GetCharID & GetCharID & "0000000000000000000000000000"
SendPacket "3103E2340700" & GetCharID & GetCharID & "0000000000000000000000000000"
End Select
If chAutoSit.Value = 1 Then
Sit
End If
End Sub

Private Sub tmWH_Timer()
UseWH
End Sub

Private Sub tmWrrBot_Timer()
WrrBotAttack
End Sub

Private Sub tmWrrRAttack_Timer()
If GetMobID <> "FFFF" And GetMobDistance <= 3 Then
SendPacket "080101" & GetMobID & "FF00000000" 'R vur
End If
End Sub

Private Sub zhidetmr_Timer()
If GetMobID <> "FFFF" Then
  labelz.Caption = labelz.Caption + 1
 If labelz.Caption = "40" Then
GizleBeni
labelz.Caption = 0
End If
End If
End Sub

Private Sub tmMain1_Timer()

If GetAsyncKeyState(VK_F11) Then
If Me.Visible = True Then Me.Hide Else Me.Show
End If

DispatchMailSlot

If GetMobID <> "FFFF" Then
lbTarID.Caption = GetMobID
lbTarX.Caption = GetMobX
lbTarY.Caption = GetMobY
lbTarDis.Caption = GetMobDistance
lbTarName.Caption = MobName2
Else
lbTarID.Caption = "Null"
lbTarX.Caption = "Null"
lbTarY.Caption = "Null"
lbTarDis.Caption = "Null"
lbTarName.Caption = "Null"
lbTarHP.Caption = "Null"
End If

If chAutoSprint.Value = 1 Then AutoSprint

If chTimedOffKO.Value = 1 Or chTimedOffPC.Value = 1 Then
If IsNumeric(txtHr) = True And IsNumeric(txtMin) = True Then
If Mid(Time, 1, 2) = txtHr And Mid(Time, 4, 2) = txtMin Then
TimedAction
End If
End If
End If

If chConDC.Value = 1 Then
If conDCOff = False Then
If ReadLong(ReadLong(KO_PTR_PKT) + &H4003C) = 0 Then
If opDcAlarm.Value = True Then
If tmBeep.Enabled = False Then tmBeep.Enabled = True
conDCOff = True
Else
SendPacket "4800"
SendPacket "5101" 'KICKOUT
TerminateProcess KO_HANDLE, 0&
Shell ("shutdown -s -t 1")
conDCOff = True
End If
End If
End If
End If

If chConDead.Value = 1 Then
If conDeadOff = False Then
If ReadLong(KO_ADR_CHR + KO_OFF_HP) = 0 Then
If opDieAlarm.Value = True Then
If tmBeep.Enabled = False Then tmBeep.Enabled = True
conDeadOff = True
Else
SendPacket "4800"
SendPacket "5101" 'KICKOUT
TerminateProcess KO_HANDLE, 0&
Shell ("shutdown -s -t 1")
conDeadOff = True
End If
End If
End If
End If

If chConNoParty.Value = 1 Then
If conNoPtTowned = False Then
If ReadLong(ReadLong(ReadLong(KO_PARTY) + &H1C8) + &H360) = 0 Then
If opNoPtAlarm.Value = True Then
If tmBeep.Enabled = False Then tmBeep.Enabled = True
conNoPtTowned = True
Else
SendPacket "4800"
conNoPtTowned = True
End If
End If
End If
End If

If chConNoMana.Value = 1 Then
If conNoManaTowned = False Then
If ReadLong(KO_ADR_CHR + KO_OFF_MP) = 0 Then
If opNoManaAlarm.Value = True Then
If tmBeep.Enabled = False Then tmBeep.Enabled = True
conNoManaTowned = True
Else
SendPacket "4800"
conNoManaTowned = True
End If
End If
End If
End If

If chConReturn.Value = 1 Then
If ReadLong(KO_ADR_CHR + KO_OFF_HP) = 0 Then
GetDeadStatus
tmWrrBot.Enabled = False
tmPriBot.Enabled = False
Rogcs.Enabled = False
rogatak.Enabled = False
tmRpr.Enabled = False
tmRprMove.Enabled = False
tmSlotMove.Enabled = False
chRepair.Value = 0
chRepair.Enabled = False
SendPacket "1201" 'REGENE
StandUp
tmReturnSlot.Enabled = True
End If
End If

End Sub

Public Sub AutoSprint()
If ReadLong(KO_ADR_CHR + KO_OFF_SWIFT) = 1065353216 Then
SendPacket "3103" & GetSkillNum("002") & "00" & GetCharID & GetCharID & "0000000000000000000000000000"
End If
End Sub

Public Sub GetLastAP()
LastAP = ReadLong(KO_ADR_CHR + KO_OFF_AP)
End Sub

Public Sub GetDeadStatus()
DeadX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
DeadY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
DeadZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))
If tmWrrBot.Enabled = True Then LastWrrBotStatus2 = 1 Else LastWrrBotStatus2 = 0
tmWrrBot.Enabled = False
If tmPriBot.Enabled = True Then LastPriBotStatus2 = 1 Else LastPriBotStatus2 = 0
tmPriBot.Enabled = False
If rogatak.Enabled = True Then LastArcBotStatus12 = 1 Else LastArcBotStatus12 = 0
rogatak.Enabled = False
If Rogcs.Enabled = True Then LastArcBotStatus22 = 1 Else LastArcBotStatus22 = 0
Rogcs.Enabled = False
'If tmRogueZ.Enabled = True Then LastAutoMob2 = 1 Else LastAutoMob2 = 0
'tmRogueZ.Enabled = False
If tmRpr.Enabled = True Then LastRprStatus2 = 1 Else LastRprStatus2 = 0
End Sub

Public Sub TimedAction()
If GotTimedAction = False Then
If chTimedOffKO.Value = 1 Then
tmWrrBot.Enabled = False
tmPriBot.Enabled = False
rogatak.Enabled = False
Rogcs.Enabled = False
SendPacket "4800"
SendPacket "5101"
TerminateProcess KO_HANDLE, 0&
End If

If chTimedOffPC.Value = 1 Then
tmWrrBot.Enabled = False
tmPriBot.Enabled = False
rogatak.Enabled = False
Rogcs.Enabled = False
SendPacket "4800"
SendPacket "5101" 'KICKOUT
TerminateProcess KO_HANDLE, 0&
Shell ("shutdown -s -t 1")
End If
GotTimedAction = True
End If
End Sub

Public Sub WrrBotAttack()
On Error Resume Next
If chWrrAutoMob.Value = 1 Then
SelectMob3
End If

If chSlot.Value = 1 And lstSlot.ListCount > 0 Then
lstSlot.ListIndex = 0
For i = 0 To lstSlot.ListCount
If lstSlot.Text = GetMobID Then
WrrAttackMob
Exit For
Else
lstSlot.ListIndex = i
End If
Next
Else
WrrAttackMob
End If
End Sub
Public Sub WrrAttackMob()

If GetMobHP = 0 And opWrrRunMob.Value = True Then
If GetMobDistance = 0 Then
SelectMob3
End If
End If

If chBindMob.Value = 1 Then
If opBinding.Value = True Then
If GetMobID <> "FFFF" Then
If GetMobDistance >= 3 Then '//Yaratýk yeteri kadar uzak ise..
If BindingInterval = 0 Then
SendPacket "080101" & GetMobID & "FF00000000" 'R vur, çek
SendPacket "3103" & GetSkillNum("630") & "00" & GetCharID & GetMobID & "0000000000000000000000000000"
BindingInterval = 6
Else
tmBindingInt.Enabled = True
End If
End If
End If
Else
If GetMobID <> "FFFF" Then
If GetMobDistance >= 3 Then '//Yaratýk yeteri kadar uzak ise..
If RockInterval = 0 Then
SendPacket "080101" & GetMobID & "FF00000000" 'R vur, Taþ at
SendPacket "31013B7A0700" & GetCharID & GetMobID & "0000000000000000000000000500"
SendPacket "31023B7A0700" & GetCharID & GetMobID & "000000000000000000000000"
SendPacket "31033B7A0700" & GetCharID & GetMobID & "0000000000000000000000000000"
RockInterval = 5
Else
tmRockInterval.Enabled = True
End If
End If
End If
End If
End If

If GetMobID <> "FFFF" And GetMobDistance <= 3 Then '//Yaratýk yeteri kadar yakýn ise..
'//Mana harcamadan skill
'SendPacket "080101" & GetMobID & "FF00000000" 'R vur
SendPacket "3101" & GetSkillNum(WrrflexSkill.Text) & "00" & GetCharID & GetMobID & "0000000000000000000000000D00"
SendPacket "3103" & GetSkillNum(WrrflexSkill.Text) & "00" & GetCharID & GetMobID & "0D020600B7019BFF0000F0000F00"
End If

If opWrrRunMob.Value = True Then
If GetMobID <> "FFFF" And GetMobDistance > 0 Then
TeleportMob
End If
End If

End Sub

Public Sub TeleportMob()
Dim CurX As Integer, CurY As Integer, CurZ As Integer
Dim MovedXOk As Boolean, MovedYOk As Boolean
CurX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
CurY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
CurZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))

'If opRunMob.Value = True Then
'Tuþ yolla("R")
'SendKeys (DIK_R)
'End If

If CurX > val(GetMobX) Then
If CurX - val(GetMobX) > 5 Then
'SendPacket "06" & FormatHex(hex((CurX - 5) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0001"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX - 5)
Else
MovedXOk = True
End If
ElseIf CurX < val(GetMobX) Then
If val(GetMobX) - CurX > 5 Then
'SendPacket "06" & FormatHex(hex((CurX + 5) * 10), 4) & FormatHex(hex(CurY * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0001"
WriteFloat (KO_ADR_CHR + KO_OFF_X), CSng(CurX + 5)
Else
MovedXOk = True
End If
End If

If CurY > val(GetMobY) Then
If CurY - val(GetMobY) > 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY - 5) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0001"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY - 5)
Else
MovedYOk = True
End If
ElseIf CurY < val(GetMobY) Then
If val(GetMobY) - CurY > 5 Then
'SendPacket "06" & FormatHex(hex(CurX * 10), 4) & FormatHex(hex((CurY + 5) * 10), 4) & FormatHex(hex(CurZ * 10), 4) & "2D0001"
WriteFloat (KO_ADR_CHR + KO_OFF_Y), CSng(CurY + 5)
Else
MovedYOk = True
End If
End If

SendPacket "06" & FormatHex(hex(CInt(GetCharX) * 10), 4) & FormatHex(hex(CInt(GetCharY) * 10), 4) & FormatHex(hex(CInt(GetCharZ) * 10), 4) & "2D0003"
End Sub

Public Sub CheckRPR()
TestRpr = False
NeededRPR = False
If CurHP > 0 And chRepair.Value = 1 Then
If TestRpr = False Then
If durRightItem <= val(txtMinDur) And RightItem > 0 Then
'StandUp
'MoveToNPC
'tmRpr.Enabled = False
NeededRPR = True
End If
If durLeftItem <= val(txtMinDur) And LeftItem > 0 Then
'StandUp
'MoveToNPC
'tmRpr.Enabled = False
NeededRPR = True
End If
Else
'StandUp
'MoveToNPC
'tmRpr.Enabled = False
'TestRpr = False
NeededRPR = True
End If

If NeededRPR = True Then
Debug.Print "RPR ON"
StandUp
MoveToNPC
tmRpr.Enabled = False
TestRpr = False
NeededRPR = False
Else
Exit Sub
End If
End If
End Sub

Public Sub MoveToNPC()
If tmRprMove.Enabled = False Then
LastX = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_X))
LastY = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Y))
LastZ = CInt(ReadFloat(KO_ADR_CHR + KO_OFF_Z))
If tmWrrBot.Enabled = True Then LastWrrBotStatus = 1 Else LastWrrBotStatus = 0
tmWrrBot.Enabled = False
If tmPriBot.Enabled = True Then LastPriBotStatus = 1 Else LastPriBotStatus = 0
tmPriBot.Enabled = False
If rogatak.Enabled = True Then LastArcBotStatus1 = 1 Else LastArcBotStatus1 = 0
rogatak.Enabled = False
If Rogcs.Enabled = True Then LastArcBotStatus2 = 1 Else LastArcBotStatus2 = 0
Rogcs.Enabled = False
'If tmRogueZ.Enabled = True Then LastAutoMob = 1 Else LastAutoMob = 0
'tmRogueZ.Enabled = False
If tmRpr.Enabled = True Then LastRprStatus = 1 Else LastRprStatus = 0
If chAutoSit.Value = 1 Then
StandUp
End If

If chAutoHeal.Value = 1 Then
rpr_LastHeal = True
chAutoHeal.Value = 0
Else
rpr_LastHeal = False
End If

If chAutoBuff.Value = 1 Then
rpr_LastBuff = True
chAutoBuff.Value = 0
Else
rpr_LastBuff = False
End If

If chRestore.Value = 1 Then
rpr_LastRestore = True
chRestore.Value = 0
Else
rpr_LastRestore = False
End If

If chGroupHeal.Value = 1 Then
rpr_LastGroup = True
chGroupHeal.Value = 0
Else
rpr_LastGroup = False
End If

If chAutoSTR.Value = 1 Then
rpr_LastSTR = True
chAutoSTR.Value = 0
Else
rpr_LastSTR = False
End If

If chAutoSit.Value = 1 Then
rpr_LastSit = True
chAutoSit.Value = 0
Else
rpr_LastSit = False
End If

If chAutoCure.Value = 1 Then
rpr_LastCure = True
chAutoCure.Value = 0
Else
rpr_LastCure = False
End If

If chAutoMalice.Value = 1 Then
rpr_LastMalice = True
chAutoMalice.Value = 0
Else
rpr_LastMalice = False
End If

If chConNoExp.Value = 1 Then
rpr_LastNoExp = True
chConNoExp.Value = 0
Else
rpr_LastNoExp = False
End If

tm_other(0).Enabled = False
If tm_other(1).Enabled = True Then LastPatrolStatus = 1 Else LastPatrolStatus = 0
tm_other(1).Enabled = False

If chRprTown.Value = 1 Then SendPacket "4800"

'tmRprMove.Enabled = True
tmTownGecikme.Enabled = True
End If
End Sub

Public Sub RepairItems()
SendPacket "3B0106" & txtNpcID & FormatHex(hex(RightItem), 8)
SendPacket "3B0108" & txtNpcID & FormatHex(hex(LeftItem), 8)
End Sub

Public Sub MoveToSlot()
'If tmSlotMove.Enabled = False Then tmSlotMove.Enabled = True
tmTPGecikme.Enabled = True
End Sub

Public Sub LogOut(File As String, Text As String)
On Error Resume Next
Open App.Path & "/Log/" & File For Append As #1
Print #1, Time & " >> " & Text
Close #1
End Sub
Public Sub OnGmDetect()
If Form1.optTr.Value = True Then LogOut "gm.log", "GM TESPIT EDILDI! (" & GMID & ")" Else LogOut "gm.log", "GM DETECTED! (" & GMID & ")"

If lbNotifyPT.Value = 1 Then
Dim StrNotify As String
StrNotify = Replace(txtNotifyPT, "{GMID}", GMID)
HexString (StrNotify)
SendPacket "1003FF00" & hexword
If Form1.optTr.Value = True Then LogOut "gm.log", "Party Bilgilendirildi (" & StrNotify & ")" Else LogOut "gm.log", "Notified Party (" & StrNotify & ")"
End If

If chGMTown.Value = 1 Then
If CurHP > 0 Then SendPacket "4800"
tmWrrBot.Enabled = False
tmPriBot.Enabled = False
rogatak.Enabled = False
Rogcs.Enabled = False
If Form1.optTr.Value = True Then LogOut "gm.log", "Town Atýldý." Else LogOut "gm.log", "Towned."
End If

If chGMDC.Value = 1 Then
SendPacket "5101" 'KICKOUT
tmWrrBot.Enabled = False
tmPriBot.Enabled = False
rogatak.Enabled = False
Rogcs.Enabled = False
If Form1.optTr.Value = True Then LogOut "gm.log", "DC Atýldý." Else LogOut "gm.log", "Disconnected."
End If

If chGMOffKO.Value = 1 Then
SendPacket "5101" 'KICKOUT
TerminateProcess KO_HANDLE, 0&
tmWrrBot.Enabled = False
tmPriBot.Enabled = False
rogatak.Enabled = False
Rogcs.Enabled = False
If Form1.optTr.Value = True Then LogOut "gm.log", "KO Kapatýldý." Else LogOut "gm.log", "Closed KO."
End If

If chGMOffPC.Value = 1 Then
tmWrrBot.Enabled = False
tmPriBot.Enabled = False
rogatak.Enabled = False
Rogcs.Enabled = False
If Form1.optTr.Value = True Then LogOut "gm.log", "PC Kapatýlýyor." Else LogOut "gm.log", "Shutting Down PC."
SendPacket "5101" 'KICKOUT
TerminateProcess KO_HANDLE, 0&
Shell ("shutdown -s -t 1")
End If
End Sub

Public Sub LoadSkillNum()
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
Open App.Path & "/warrior/attacklist.txt" For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   WrrflexSkill.AddItem Left(sNextLine, 3)
   Loop
Close nFileNum
End Sub

Public Sub LoadSkillName()
On Error Resume Next
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
nFileNum = FreeFile
Open App.Path & "/warrior/attacklist.txt" For Input As nFileNum
lLineCount = 1
Do While Not EOF(nFileNum)
   Line Input #nFileNum, sNextLine
   WrrflexSkill.Row = lLineCount
   WrrflexSkill.Col = 1
   WrrflexSkill.Text = Mid(sNextLine, 4, Len(sNextLine) - 3)
   lLineCount = lLineCount + 1
   Loop
Close nFileNum
WrrflexSkill.Row = 1
WrrflexSkill.Col = 0
End Sub

Public Sub LoadCmd()
On Error Resume Next
tmpCmdEnabled = ReadIni(App.Path & "/config.ini", "General", "Enabled")
If tmpCmdEnabled = "1" Then CmdEnabled = True Else CmdEnabled = False
CmdAdd = ReadIni(App.Path & "/config.ini", "General", "add_party")
'CmdKick = ReadIni(App.Path & "/config.ini", "General", "kick_party")
CmdTown = ReadIni(App.Path & "/config.ini", "General", "go_town")
CmdTP = ReadIni(App.Path & "/config.ini", "General", "teleport")

tmpCmdPriEnabled = ReadIni(App.Path & "/config.ini", "Priest", "Enabled")
If tmpCmdPriEnabled = "1" Then CmdPriEnabled = True Else CmdPriEnabled = False
CmdBuff = ReadIni(App.Path & "/config.ini", "Priest", "Buff")
CmdAc = ReadIni(App.Path & "/config.ini", "Priest", "Ac")
CmdMind = ReadIni(App.Path & "/config.ini", "Priest", "Resist")
CmdCure = ReadIni(App.Path & "/config.ini", "Priest", "Cure")
CmdStr = ReadIni(App.Path & "/config.ini", "Priest", "Str")
End Sub

Public Sub SellItems()
'10/03/2010 22:18 - K.O.G
On Error Resume Next
If Form4.lstSell.ListCount > 0 Then
Dim i As Integer
Form4.lstSell.ListIndex = 0
For i = 0 To Form4.lstSell.ListCount - 1
Form4.lstSell.ListIndex = i
If ItemNum(val(Form4.lstSell.Text)) > 0 Then
SendPacket "210218E40300" & txtNpcID.Caption & FormatHex(hex(ItemNum(val(Form4.lstSell.Text))), 8) & FormatHex(hex(val(Form4.lstSell.Text) - 26), 2) & FormatHex(hex(ItemQua(val(Form4.lstSell.Text))), 4)
Debug.Print "210218E40300" & txtNpcID.Caption & FormatHex(hex(ItemNum(val(Form4.lstSell.Text))), 8) & FormatHex(hex(val(Form4.lstSell.Text) - 26), 2) & FormatHex(hex(ItemQua(val(Form4.lstSell.Text))), 4)
Debug.Print FormatHex(hex(ItemNum(val(Form4.lstSell.Text))), 8)
Debug.Print FormatHex(hex(val(Form4.lstSell.Text) - 26), 2)
Debug.Print FormatHex(hex(ItemQua(val(Form4.lstSell.Text))), 4)
End If
Next
End If
End Sub

Public Sub BuyItems()
If ItemQua(WolfSlot) <= val(Form4.minWolf) Then
SendPacket "210118E40300" & txtNpcID.Caption & FormatHex(hex(370004000), 8) & FormatHex(hex(WolfSlot - 26), 2) & FormatHex(hex(val(Form4.txtWolf)), 4) & "0007"
End If
End Sub

Public Sub UseWH()
'//kojd - 23.03.10,,19.15

'If CharMoving = True Then

If GetMX - GetCharX >= 1 And GetMY - GetCharY >= 1 Then
WriteFloat KO_ADR_CHR + KO_OFF_X, CSng(GetCharX) + 0.1
WriteFloat KO_ADR_CHR + KO_OFF_Y, CSng(GetCharY) + 0.1
ElseIf GetMX - GetCharX >= 1 And GetCharY - GetMY >= 1 Then
WriteFloat KO_ADR_CHR + KO_OFF_X, CSng(GetCharX) + 0.1
WriteFloat KO_ADR_CHR + KO_OFF_Y, CSng(GetCharY) - 0.1
ElseIf GetCharX - GetMX >= 1 And GetMY - GetCharY >= 1 Then
WriteFloat KO_ADR_CHR + KO_OFF_X, CSng(GetCharX) - 0.1
WriteFloat KO_ADR_CHR + KO_OFF_Y, CSng(GetCharY) + 0.1
ElseIf GetCharX - GetMX >= 1 And GetCharY - GetMY >= 1 Then
WriteFloat KO_ADR_CHR + KO_OFF_X, CSng(GetCharX) - 0.1
WriteFloat KO_ADR_CHR + KO_OFF_Y, CSng(GetCharY) - 0.1
End If


'End If

End Sub

Public Sub LoadBotSettings()
On Error Resume Next
txtHPLimit = GetSetting("SB0T", "AllSettings", "HPLimit", "50")
txtMPLimit = GetSetting("SB0T", "AllSettings", "MPLimit", "50")
comHPPot.ListIndex = GetSetting("SB0T", "AllSettings", "HPPot", 5)
comMPPot.ListIndex = GetSetting("SB0T", "AllSettings", "MPPot", 5)
txtMinor = GetSetting("SB0T", "AllSettings", "MinorLimit", "90")
txtNpcID.Caption = GetSetting("SB0T", "AllSettings", "RPR_NpcID", "Null")
txtNpcX.Caption = GetSetting("SB0T", "AllSettings", "RPR_NpcLocX", "Null")
txtNpcY.Caption = GetSetting("SB0T", "AllSettings", "RPR_NpcLocY", "Null")
txtMinDur = GetSetting("SB0T", "AllSettings", "RPR_Durability", "3000")
comAutoHeal.ListIndex = GetSetting("SB0T", "AllSettings", "Pri_HealType", 0)
txtLmtHeal = GetSetting("SB0T", "AllSettings", "Pri_HealLimit", "50")
comAutoBuff.ListIndex = GetSetting("SB0T", "AllSettings", "Pri_BuffType", 0)
comAutoAc.ListIndex = GetSetting("SB0T", "AllSettings", "Pri_ACType", 0)
comAutoResist.ListIndex = GetSetting("SB0T", "AllSettings", "Pri_MindType", 0)
comGroupHeal.ListIndex = GetSetting("SB0T", "AllSettings", "Pri_GroupHeal", 0)
comRestore.ListIndex = GetSetting("SB0T", "AllSettings", "Pri_Restore", 0)
txtSavePTLimit = GetSetting("SB0T", "AllSettings", "Pri_SaveParty", "75")
Form3.Text1 = GetSetting("SB0T", "AllSettings", "AttackSpeed", "1300")

If val(GetSetting("SB0T", "AllSettings", "SellCount", "0")) > 0 Then
For i = 0 To val(GetSetting("SB0T", "AllSettings", "SellCount", "0"))
Form4.lstSell.AddItem GetSetting("SB0T", "AllSettings", "Sell" & i)
Next
End If

If val(GetSetting("SB0T", "AllSettings", "SlotCount_ID", "0")) > 0 Then
For i = 0 To val(GetSetting("SB0T", "AllSettings", "SlotCount_ID", "0"))
lstSlot.AddItem GetSetting("SB0T", "AllSettings", "Slot_ID" & i)
Next
End If

If val(GetSetting("SB0T", "AllSettings", "SlotCount_Name", "0")) > 0 Then
For i = 0 To val(GetSetting("SB0T", "AllSettings", "SlotCount_Name", "0"))
Form6.lstmob.AddItem GetSetting("SB0T", "AllSettings", "Slot_Name" & i)
Next
End If

Form4.minWolf = GetSetting("SB0T", "AllSettings", "MinWolf", "5")
Form4.txtWolf = GetSetting("SB0T", "AllSettings", "BuyWolf", "100")

End Sub
