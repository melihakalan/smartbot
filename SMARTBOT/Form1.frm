VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SB0T 1831"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   3420
      TabIndex        =   7
      Top             =   120
      Width           =   3420
      Begin SHDocVwCtl.WebBrowser WebMain 
         Height          =   1830
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3735
         ExtentX         =   6588
         ExtentY         =   3219
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Start"
      Height          =   2175
      Left            =   275
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
      Begin VB.CheckBox chNoDC 
         Caption         =   "No DC Bugu Kapat"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtWnd 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Text            =   "SB0T"
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optEn 
         Caption         =   "English"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optTr 
         Caption         =   "Türkçe"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtTitle 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   0
         Text            =   "Knight OnLine Client"
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton btnStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ko:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   625
         Width           =   240
      End
   End
   Begin MSComctlLib.ListView lstThreads 
      Height          =   1215
      Left            =   3720
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2143
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   12582912
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "a"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "b"
         Object.Width           =   1641
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "c"
         Object.Width           =   1288
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "d"
         Object.Width           =   1288
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "e"
         Object.Width           =   1288
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "f"
         Object.Width           =   935
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "g"
         Object.Width           =   354
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "h"
         Object.Width           =   1853
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "www.OnlineHile.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   4560
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "KoJD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   4320
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Run As Integer
Public ReklamIndex As Integer
Public DisableWeb As Boolean


Private Sub btnStart_Click()
LoadBot
End Sub

Private Sub Form_Initialize()
'LoadPrivate
ENB_HOOK = True
ENB_DI8 = True
DisableWeb = False

ReklamIndex = 0
If DisableWeb = False Then WebMain.Navigate "http://www.onlinehile.com/kojd/smart/smart_tr.htm"
UserInAuth = 1

Run = GetSetting("SmartBot", "Settings", "Run", 0)
If Run > 2 Then
'btnStart.Enabled = False
Else
btnStart.Enabled = True
End If
Debug.Print Run
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "SmartBot", "Settings", "Run", Run + 1
CloseBot
End Sub

Private Sub optEn_Click()
If DisableWeb = False Then WebMain.Navigate "http://www.onlinehile.com/kojd/smart/smart_en.htm"

chNoDC.Caption = "Fix No DC Bug"

Form3.frSpeed.Caption = "Atak speed"

Form2.SSTab1.TabCaption(0) = "General"
Form2.SSTab1.TabCaption(1) = "Protection"
Form2.SSTab1.TabCaption(2) = "Attack"
Form2.SSTab1.TabCaption(3) = "Other"

Form2.frGenel.Caption = "General"
Form2.frTS.Caption = "Legal TS (60min)"
Form2.frTarget.Caption = "Target info"
Form2.frSlot.Caption = "Save Slot"
Form2.frGM.Caption = "GM Protection"
Form2.frDetails.Caption = "Detail log"
Form2.frTimed.Caption = "Timed Functions"
Form2.frRogue.Caption = "Rogue Attack"
Form2.frRogueTimed.Caption = "Rogue Timed Skills"
Form2.frConditions.Caption = "Conditions"

Form2.chRuntochest.Caption = "Go to chest"
Form2.chAutoLoot.Caption = "AutoLoot"
Form2.chCoins.Caption = "Only coins"
Form2.btnRecNpc.Caption = "Save Npc"
Form2.lbtemp(15).Caption = "Remaining min:"
Form2.chTS.Caption = "Start Auto TS"
Form2.lbtemp(3).Caption = "Dist."
Form2.chSlot.Caption = "Enable Slot Limitting"
Form2.lbMobCount.Caption = "Mob Count:"
Form2.lbGM.Caption = "If a GM enters the slot:"
Form2.lbNotifyPT.Caption = "Notify party:"
Form2.chGMLog.Caption = "Save actions as log (Log/gm.log)"
Form2.chChatLog.Caption = "Save all chat messages as log (Log/chat.log)"
Form2.chSlotLog.Caption = "Save everyone who enters the slot as ID,Nick,Lvl,Job,Clan (Log/user.log)"
Form2.chExpLog.Caption = "Save exping/leveling        ( Log/Exp.log)"
Form2.lbHr.Caption = "Hr"
Form2.lbMin.Caption = "Min"
Form2.chTimedOffKO.Caption = "Shut KO"
Form2.chTimedOffPC.Caption = "Shut PC"
Form2.chAutoMob.Caption = "Auto Mob"
Form2.listeden.Caption = "From the list"
Form2.chmobdead.Caption = "Don't Z before kill"
Form2.btnSlot.Caption = "Add Slot"
Form2.btnClearSlot.Caption = "Clear"
Form2.Command2.Caption = "Atak speed..."
Form2.chWrrAutoMob.Caption = "Auto Mob"
Form2.opWrrRunMob.Caption = "Run Mob"
Form2.opWrrWaitMob.Caption = "Wait Mob"
Form2.frBindMob.Caption = "Bring Mob"
Form2.opRock.Caption = "Rock"
Form2.btnAtkLoad.Caption = "Load"
Form2.chPriAutoMob.Caption = "Auto Mob"
Form2.opPriRunMob.Caption = "Run Mob"
Form2.opPriWaitMob.Caption = "Wait Mob"
Form2.chMaliceMob.Caption = "Malice Mob"
Form2.chConRace2.Caption = "If detects other nation:"
Form2.btnConRace.Caption = "VVV Action VVV"
Form2.chConDC.Caption = "If DC:"
Form2.chConDead.Caption = "If Dies:"
Form2.chConNoParty.Caption = "If No Party:"
Form2.chConNoMana.Caption = "If No Mana:"
Form2.chConNoExp.Caption = "If No Exp for 3 mins:"
Form2.chConReturn.Caption = "Return back when die"
Form2.opRaceSeemDie.Caption = "Pretend to die till leaves the slot"
Form2.chSaveParty.Caption = "Save Party"
Form2.chParty.Caption = "Party Enabled"
Form2.chAutoSit.Caption = "Sit when idle"

Form3.Label1.Caption = "speed (MS)"
Form3.Label2.Caption = "Should be between 1250-1350"

Form4.Label2.Caption = "Sell list:"
Form4.clear.Caption = "clear"
Form4.chSell.Caption = "Auto sell enabled"
Form4.buywolf.Caption = "Buy wolf"
Form4.Label3.Caption = "minimum wolf quantity:"
Form4.Label5.Caption = "buy quantity:"
Form4.Label4.Caption = "note: selling/buying actions will happen while repairing."
Form4.Command1.Caption = "O K"

Form4.lbselldesc.Caption = "The items where in selected slots will be sold on every repairing."

Form2.opmobid.Caption = "Mob ID"
Form2.opmobname.Caption = "Mob Name"
Form2.btnmobname.Caption = "Mobs>>"
Form6.lbdesc.Caption = "Mob names to attack:"
Form6.btnClear.Caption = "Clear"
Form6.addoto.Caption = "Add selected mob"
Form6.addmanuel.Caption = "Add Manual"
End Sub

Private Sub optTr_Click()
If DisableWeb = False Then WebMain.Navigate "http://www.onlinehile.com/kojd/smart/smart_tr.htm"

chNoDC.Caption = "No DC Bugu Kapat"

Form3.frSpeed.Caption = "Atak hýzý"

Form2.SSTab1.TabCaption(0) = "Genel"
Form2.SSTab1.TabCaption(1) = "Korunma"
Form2.SSTab1.TabCaption(2) = "Atak"
Form2.SSTab1.TabCaption(3) = "Diðer"

Form2.chRuntochest.Caption = "Kutuya Git"
Form2.frGenel.Caption = "Genel"
Form2.frTS.Caption = "Legal TS (60dk)"
Form2.frTarget.Caption = "Hedef Bilgileri"
Form2.frSlot.Caption = "Slot Kayýt"
Form2.frGM.Caption = "GM Korunma"
Form2.frDetails.Caption = "Detay Kayýtlarý"
Form2.frTimed.Caption = "Zamanlý Ýþlemler"
Form2.frRogue.Caption = "Rogue Atak"
Form2.frRogueTimed.Caption = "Rogue Zamanlý Skiller"
Form2.frConditions.Caption = "Koþullar"

Form2.chAutoLoot.Caption = "Oto Kutu"
Form2.chCoins.Caption = "Sadece Para"
Form2.btnRecNpc.Caption = "Npc Kayýt"
Form2.lbtemp(15).Caption = "Kalan dakika"
Form2.chTS.Caption = "Oto TS Baþlat"
Form2.lbtemp(3).Caption = "Uzaklýk"
Form2.chSlot.Caption = "Sadece Bu Slota Saldýr"
Form2.lbMobCount.Caption = "Yaratýk Sayýsý:"
Form2.lbGM.Caption = "Slota GM Gelirse:"
Form2.lbNotifyPT.Caption = "Partiye bildir:"
Form2.chGMLog.Caption = "Yapýlan iþlemi Log olarak kaydet (Log/gm.log)"
Form2.chChatLog.Caption = "Tüm Chat konuþmalarýný kaydet (Log/chat.log)"
Form2.chSlotLog.Caption = "Slota gelip giden herkesi ID,Nick,Level,Job,Clan þeklinde kaydet (Log/user.log)"
Form2.chExpLog.Caption = "Alýnan Levelleri ve gelen tüm Expleri kaydet        ( Log/Exp.log)"
Form2.lbHr.Caption = "Sa"
Form2.lbMin.Caption = "Dk"
Form2.chTimedOffKO.Caption = "KO Kapat"
Form2.chTimedOffPC.Caption = "PC Kapat"
Form2.chAutoMob.Caption = "Oto Yaratýk Seç"
Form2.listeden.Caption = "Listeden seç"
Form2.chmobdead.Caption = "Mob ölmeden Z yapma"
Form2.btnSlot.Caption = "Slot Ekle"
Form2.btnClearSlot.Caption = "Temizle"
Form2.Command2.Caption = "Atak hýzý..."
Form2.chWrrAutoMob.Caption = "Oto Yaratýk Seç"
Form2.opWrrRunMob.Caption = "Yaratýða Koþ"
Form2.opWrrWaitMob.Caption = "Yaratýðý Bekle"
Form2.frBindMob.Caption = "Yaratýðý Çek"
Form2.opRock.Caption = "Taþ"
Form2.btnAtkLoad.Caption = "Yükle"
Form2.chPriAutoMob.Caption = "Oto Yaratýk Seç"
Form2.opPriRunMob.Caption = "Yaratýða Koþ"
Form2.opPriWaitMob.Caption = "Yaratýðý Bekle"
Form2.chMaliceMob.Caption = "Yaratýða Malice"
Form2.chConRace2.Caption = "Slota Karþý Irk Gelirse:"
Form2.btnConRace.Caption = "VVV Ýþem VVV"
Form2.chConDC.Caption = "DC Olursa:"
Form2.chConDead.Caption = "Yatarsa:"
Form2.chConNoParty.Caption = "Party Bozulursa:"
Form2.chConNoMana.Caption = "Mana Biterse:"
Form2.chConNoExp.Caption = "3 Dk. Exp Gelmezse:"
Form2.chConReturn.Caption = "Öldüðün Yere Dön"
Form2.opRaceSeemDie.Caption = "Slottan ayrýlana kadar ölü numarasý yap"
Form2.chSaveParty.Caption = "Party Kurtar"
Form2.chParty.Caption = "Party Aktif"
Form2.chAutoSit.Caption = "Boþtayken Otur"

Form3.Label1.Caption = "Hýz (MS)"
Form3.Label2.Caption = "Ýdeal Atak Hýzý 1250 <> 1350 Arasýdýr"

Form4.Label2.Caption = "Satýlacak listesi:"""
Form4.clear.Caption = "temizle"
Form4.chSell.Caption = "Oto sell aktif"
Form4.buywolf.Caption = "Wolf al"
Form4.Label3.Caption = "minimum wolf miktarý:"
Form4.Label5.Caption = "kaç tane alsýn:"
Form4.Label4.Caption = "not: Satým/Alým iþlemi Rpr yaparken gerçekleþtirilir.."
Form4.Command1.Caption = "TAMAM"

Form4.lbselldesc.Caption = "Seçtiðiniz itemler'ýn bulunduðu kutular her RPR'de boþaltýlýr."

Form2.opmobid.Caption = "Mob ID göre"
Form2.opmobname.Caption = "Mob adýna göre"
Form2.btnmobname.Caption = "Adlar>>"
Form6.lbdesc.Caption = "Atak yapýlacak yaratýk listesi:"
Form6.btnClear.Caption = "Temizle"
Form6.addoto.Caption = "Seçili Mob Ekle"
Form6.addmanuel.Caption = "Manuel ekle"
End Sub

Private Sub webReklam_DownloadBegin()
If ReklamIndex <= 1 Then
ReklamIndex = ReklamIndex + 1
ElseIf ReklamIndex = 2 Then
btnStart.Enabled = True
Run = 0
SaveSetting "SmartBot", "Settings", "Run", 0
End If
End Sub

Public Sub LoadBot()
'Shell "cmd /c ipconfig/flushdns", vbHide
'Shell "cmd /c ipconfig/renew", vbHide

On Error Resume Next

If Dir(KOPATH & "\log.klg") <> "" Then Kill KOPATH & "\log.klg"
If Dir(KOPATH & "\Scheduler.ini") <> "" Then Kill KOPATH & "\Scheduler.ini"

LoadOffsets
If AttachKO = False Then
Exit Sub
Else
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
KO_ADR_DLG = ReadLong(KO_PTR_DLG)

MemoryAllocation

If Form1.chNoDC.Value = 1 Then WriteLong KO_NODC, 0
SetupKeyboard
SNDFIX

'Dim pID As Long
'pID = GetKOPID
'If pID > 0 Then
'GetThreads pID
'DisableTPT
'Else
'MsgBox "Couldn't find Knight Online Process ID."
'End If

Me.Hide
Form2.Show
Form2.tmMain1.Enabled = True
Form2.tmLogKill.Enabled = True
If ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 104 Or ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 111 Or ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 112 Or ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 204 Or ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 211 Or ReadLong(KO_ADR_CHR + KO_OFF_CLASS) = 212 Then Form2.tmPriest.Enabled = True
End If
End Sub

Public Sub SNDFIX()
WriteLong KO_SENDPTR, &H73233F17
End Sub

Public Sub LoadPrivate()
On Error Resume Next

If p_Dec(ReadIni(App.Path & "/private.ini", "OPT", "ENB_WEB")) = "no" Then
DisableWeb = True
End If

If p_Dec(ReadIni(App.Path & "/private.ini", "OPT", "ENB_HOOK")) = "yes" Then
ENB_HOOK = True
End If

If p_Dec(ReadIni(App.Path & "/private.ini", "OPT", "ENB_DI8")) = "yes" Then
ENB_DI8 = True
End If

If p_Dec(ReadIni(App.Path & "/private.ini", "OPT", "ENB_WH")) = "yes" Then
Form2.chWH.Enabled = True
End If

End Sub

Function p_Dec(EncString As String)
On Error Resume Next
p_Dec = HexToStr(hex(CDec("&H" & EncString) / 3))
End Function
