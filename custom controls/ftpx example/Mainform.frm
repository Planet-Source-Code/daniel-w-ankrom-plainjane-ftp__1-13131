VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{2971AEE0-B6E1-11D4-81CA-00C0F010F316}#12.0#0"; "FTPX.OCX"
Begin VB.Form Mainform 
   Caption         =   "PlainJane FTP"
   ClientHeight    =   5145
   ClientLeft      =   -420
   ClientTop       =   1830
   ClientWidth     =   8910
   FillColor       =   &H80000012&
   Icon            =   "Mainform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin FTPx.engine engine1 
      Left            =   8490
      Top             =   1290
      _ExtentX        =   1349
      _ExtentY        =   873
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   90
      TabIndex        =   28
      Top             =   1935
      Width           =   9405
      Begin RichTextLib.RichTextBox SVresponse 
         Height          =   720
         Left            =   105
         TabIndex        =   36
         Top             =   270
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   1270
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"Mainform.frx":0442
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1815
         TabIndex        =   29
         Top             =   15
         Visible         =   0   'False
         Width           =   2310
      End
   End
   Begin VB.ComboBox SVname 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   5340
   End
   Begin VB.Timer KeepAlive 
      Interval        =   10000
      Left            =   8580
      Top             =   2220
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7095
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":04F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":0604
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":0718
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   4770
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "10:06 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/26/00"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Port 
      Height          =   285
      Left            =   4620
      TabIndex        =   3
      Text            =   "21"
      Top             =   1200
      Width           =   750
   End
   Begin VB.TextBox Pword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2865
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox UrName 
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   1200
      Width           =   2550
   End
   Begin VB.CommandButton CMDdisconnect 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   1275
      TabIndex        =   5
      Top             =   1605
      Width           =   1035
   End
   Begin VB.CommandButton CMDconnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1605
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   75
      Width           =   8145
      Begin VB.OptionButton OptAnon 
         Caption         =   "Login Anonymous"
         Height          =   195
         Left            =   6255
         TabIndex        =   30
         Top             =   1470
         Width           =   1770
      End
      Begin VB.Label Label9 
         Caption         =   "Server Name [ftp.servername.com]"
         Height          =   255
         Left            =   75
         TabIndex        =   35
         Top             =   300
         Width           =   4185
      End
      Begin VB.Label Label8 
         Caption         =   "Port [21]"
         Height          =   195
         Left            =   4515
         TabIndex        =   34
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "Password"
         Height          =   210
         Left            =   2790
         TabIndex        =   33
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Username"
         Height          =   270
         Left            =   105
         TabIndex        =   31
         Top             =   900
         Width           =   1920
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remote"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Index           =   1
      Left            =   4260
      TabIndex        =   14
      Top             =   3045
      Width           =   5250
      Begin VB.OptionButton Option6 
         Caption         =   "Auto"
         Height          =   285
         Left            =   3060
         TabIndex        =   38
         Top             =   405
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.TextBox RemPath 
         Height          =   285
         Left            =   135
         TabIndex        =   20
         Top             =   870
         Width           =   4470
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1950
         Left            =   120
         TabIndex        =   19
         Top             =   1155
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   3440
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Attributes"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Binary"
         Height          =   285
         Left            =   2100
         TabIndex        =   18
         Top             =   405
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ascii"
         Height          =   195
         Left            =   1305
         TabIndex        =   17
         Top             =   450
         Width           =   1245
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "DownLoad"
         Height          =   330
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1065
      End
      Begin VB.CommandButton cmdUP 
         Height          =   285
         Left            =   4650
         Picture         =   "Mainform.frx":0B6C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   870
         Width           =   480
      End
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   4650
      TabIndex        =   7
      Top             =   3825
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   5685
      TabIndex        =   6
      Top             =   4845
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   90
      TabIndex        =   21
      Top             =   3075
      Width           =   4065
      Begin VB.OptionButton Option5 
         Caption         =   "Auto"
         Height          =   225
         Left            =   2865
         TabIndex        =   37
         Top             =   420
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Binary"
         Height          =   375
         Left            =   2025
         TabIndex        =   27
         Top             =   330
         Width           =   1275
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ascii"
         Height          =   210
         Left            =   1275
         TabIndex        =   26
         Top             =   420
         Width           =   1365
      End
      Begin VB.CommandButton cmdUpLoad 
         Caption         =   "Upload"
         Height          =   315
         Left            =   165
         TabIndex        =   25
         Top             =   360
         Width           =   1005
      End
      Begin VB.CommandButton cmdUP1Lvl 
         Height          =   300
         Left            =   3390
         Picture         =   "Mainform.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox LocalPath 
         Height          =   300
         Left            =   150
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   825
         Width           =   3165
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1965
         Left            =   135
         TabIndex        =   22
         Top             =   1125
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   3466
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   3228
         EndProperty
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   270
      Left            =   2850
      TabIndex        =   32
      Top             =   945
      Width           =   1290
   End
   Begin VB.Label Label4 
      Caption         =   "Server Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      Caption         =   "Port"
      Height          =   210
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   240
      Left            =   2880
      TabIndex        =   9
      Top             =   720
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1080
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnulistview1 
      Caption         =   "Directory Commands"
      Visible         =   0   'False
      Begin VB.Menu mnulistview1Rename 
         Caption         =   "Rename File"
      End
      Begin VB.Menu mnulistview1Delete 
         Caption         =   "Delete File"
      End
      Begin VB.Menu mnuViewFile 
         Caption         =   "View File"
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ServerSelected As Boolean
Public Connected As Boolean
Public Transfering As Boolean

Private Sub CMDconnect_Click()
Dim Result As Boolean
Dim x As Integer
CMDconnect.Enabled = False
engine1.RemoteHost = SVname.Text
engine1.UserName = UrName.Text
engine1.Password = Pword.Text
engine1.RemotePort = Val(Port.Text)
StatusBar1.Panels(1).Text = "Connecting..."
Connected = engine1.Connect
If Connected Then
 StatusBar1.Panels(1).Text = "Connected"
 CMDconnect.Enabled = False
 CMDdisconnect.Enabled = True
 cmdUpLoad.Enabled = True
 cmdDownload.Enabled = True
 If SVname.ListIndex = -1 Then
  SVname.AddItem SVname.Text
  x = SVname.ListCount - 1
 Else
  x = SVname.ListIndex
 End If
 
 SaveSetting "PlainJane", "Server" & Str(x), "Server Name", SVname.Text
 SaveSetting "PlainJane", "Server" & Str(x), "UserName", UrName.Text
 SaveSetting "PlainJane", "Server" & Str(x), "Password", Pword.Text
 SaveSetting "PlainJane", "Server" & Str(x), "Port", Port.Text
 SaveSetting "PlainJane", vbNullChar, "Servers", SVname.ListCount - 1
 SaveSetting "PlainJane", "Last Server", "index", Str(x)
 LoadCurrentDirectory
 RemPath.Text = "Root/"
 Frame1(1).Enabled = True
 OptAnon.Value = False
Else
 StatusBar1.Panels(1).Text = "Connection Failed"
 CMDconnect.Enabled = True
End If

End Sub

Private Sub CMDdisconnect_Click()
engine1.Disconnect
CMDdisconnect.Enabled = False
CMDconnect.Enabled = True
cmdUpLoad.Enabled = False
cmdDownload.Enabled = False
StatusBar1.Panels(1).Text = "Disconnected"
Connected = False
Frame1(1).Enabled = False
End Sub

Private Sub CMDdownload_Click()
Dim BinXfer As Boolean
Dim strFileToGet As String
Dim FileToGetSize As Long
Dim PathToStore As String
Dim FileNameToStore As String

If ListView1.SelectedItem.SmallIcon = 1 Then
 Exit Sub
End If

If Option1.Value Then
 BinXfer = False
Else
 BinXfer = True
End If

strFileToGet = ListView1.SelectedItem.Text
FileToGetSize = Val(ListView1.SelectedItem.ListSubItems(1).Text)
PathToStore = PathToStore & UCase(Dir1.List(Dir1.ListIndex)) & "\"
FileNameToStore = InputBox("Store this file as:" & vbCrLf & "<enter> to use the same filename.", "Store as")
If FileNameToStore = "" Then
 FileNameToStore = strFileToGet
End If
Screen.MousePointer = vbHourglass
ProgDisp.Show
ProgDisp.ProgressDisplay1.BytesToGet = FileToGetSize
ProgDisp.ProgressDisplay1.FileNameToGet = strFileToGet
ProgDisp.ProgressDisplay1.SavingAs = PathToStore & FileNameToStore
ProgDisp.ProgressDisplay1.Initialize
ProgDisp.ProgressDisplay1.BeginCalcs
Transfering = True
engine1.GetFile Replace(RemPath, "Root/", "", 1, -1), strFileToGet, PathToStore & FileNameToStore, Not BinXfer
Unload ProgDisp
Transfering = False
LoadLocalDirectory
Screen.MousePointer = vbNormal
End Sub

Private Sub cmdUP_Click()
Dim Marker As Integer
Dim strData As String
Dim strResponse As String
cmdUP.Enabled = False
strData = RemPath.Text
If strData = "Root/" Then GoTo skip
Marker = Len(strData) - 1
While Mid(strData, Marker, 1) <> "/"
 Marker = Marker - 1
Wend
strData = Mid(strData, 1, Marker)
RemPath.Text = strData
StatusBar1.Panels(1).Text = "Getting Directory Info."
engine1.Execute "CDUP", strResponse
skip:
LoadCurrentDirectory
StatusBar1.Panels(1).Text = "Connected"
cmdUP.Enabled = True
End Sub

Private Sub cmdUP1Lvl_Click()
Dim Marker1 As Integer
Dim strData As String

strData = LocalPath.Text
Marker1 = Len(strData)
While Mid(strData, Marker1, 1) <> "\"
 Marker1 = Marker1 - 1
Wend
strData = Mid(strData, 1, Len(strData) - (Len(strData) - Marker1 + 1))
If InStr(1, strData, "\") = 0 Then
 strData = strData & "\"
End If
LocalPath.Text = strData
Dir1.Path = strData
LoadLocalDirectory
End Sub

Private Sub cmdUpLoad_Click()
If ListView2.SelectedItem.SmallIcon <> 2 Then
 Exit Sub
End If
If MsgBox("Upload " & ListView2.SelectedItem.Text & "?", vbYesNo, "Confirm Upload") = vbNo Then
 Exit Sub
End If
UpLoadFile
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Path
End Sub

Private Sub engine1_CommandSentToServer(CommandSent As String)
Dim txtStart As Long
txtStart = Len(SVresponse.Text)
SVresponse.Text = SVresponse.Text & CommandSent & vbCrLf
SVresponse.SelColor = &HC0&
SVresponse.SelStart = txtStart
SVresponse.SelLength = Len(SVresponse.Text) - txtStart
SVresponse.SelBold = True
SVresponse.SelStart = Len(SVresponse.Text)

End Sub

Private Sub engine1_ReceiveProgress(ByVal BytesToGet As Long, BytesGot As Long)
ProgDisp.ProgressDisplay1.BytesReceived = BytesGot
End Sub

Private Sub engine1_SendProgress(BytesSent As Long, BytesTotal As Long)
StatusBar1.Panels(4).Text = Str(BytesSent)
End Sub

Private Sub engine1_ServerResponse(Response As String)
SVresponse.Text = SVresponse.Text & Response
SVresponse.SelStart = Len(SVresponse.Text)
End Sub

Private Sub Form_Load()
Dim x As Integer
Dim LastServer As String
StatusBar1.Panels(1).Text = "Not Connected"
CMDconnect.Enabled = False
CMDdisconnect.Enabled = False
cmdUpLoad.Enabled = False
cmdDownload.Enabled = False
ServerSelected = False
Frame1(1).Enabled = False
For x = 0 To Val(GetSetting("PlainJane", vbNullChar, "Servers"))
 SVname.AddItem GetSetting("PlainJane", "Server" & Str(x), "Server Name")
Next

SVname.ListIndex = 0
UrName.Text = GetSetting("PlainJane", "Server 0", "Username")
Pword.Text = GetSetting("PlainJane", "Server 0", "Password")
Port.Text = GetSetting("PlainJane", "Server 0", "Port")
Dir1.Path = GetSetting("PlainJane", vbNullChar, "LocalPath")
LastServer = GetSetting("PlainJane", "Last Server", "index")
SVname.ListIndex = Val(LastServer)
LoadLocalDirectory
End Sub



Private Sub Form_Unload(Cancel As Integer)
SaveSetting "PlainJane", vbNullChar, "Servers", SVname.ListCount - 1
engine1.Cancel
Unload Viewer
End Sub

Private Sub KeepAlive_Timer()
Dim strResponse As String
If Not Transfering And Connected Then
 engine1.Execute "NOOP", strResponse
End If
End Sub



Private Sub ListView1_DblClick()
'
Dim a As Integer
Dim NewDir As String: Dim strResponse As String
a = ListView1.SelectedItem.SmallIcon
If a <> 1 Then Exit Sub
NewDir = ListView1.SelectedItem.Text
If NewDir = "." Or NewDir = ".." Then Exit Sub
RemPath.Text = RemPath.Text & NewDir & "/"
engine1.Execute "CWD " & NewDir, strResponse
LoadCurrentDirectory
End Sub



Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then
 If ListView1.SelectedItem.SmallIcon = 2 Then
  If MsgBox("Delete " & ListView1.SelectedItem.Text & vbCrLf & "Are you sure?", vbYesNo, "Confirm Delete") = vbYes Then
   engine1.KillFile "", ListView1.SelectedItem.Text
   LoadCurrentDirectory
  End If
 End If
End If
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
 Me.PopupMenu mnulistview1
End If
End Sub



Private Sub ListView2_DblClick()
If ListView2.SelectedItem.Icon = 3 Then
 Drive1.Drive = ListView2.SelectedItem.Text
 GoTo LoadIt
End If
If ListView2.SelectedItem.Icon = 2 Then
 'doubleclicked a file
 Exit Sub
End If
On Error GoTo LoadIt 'in case of floppy or CD not ready
Dir1.Path = ListView2.SelectedItem.Text
LoadIt:
LoadLocalDirectory
End Sub

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strData As String
If KeyCode = 46 Then ' delete key pressed
 If ListView2.SelectedItem.SmallIcon = 3 Then
  Exit Sub ' cannot delete a drive
 End If
 If ListView2.SelectedItem.SmallIcon = 2 Then
  'delete a file
  strData = ListView2.SelectedItem.Text
  If MsgBox("Delete " & strData & vbCrLf & "Are you sure?", vbYesNo, "Confirm Delete") = vbYes Then
   If Mid(LocalPath.Text, Len(LocalPath.Text), 1) = "\" Then
    strData = LocalPath.Text & ListView2.SelectedItem.Text
   Else
    strData = LocalPath.Text & "\" & ListView2.SelectedItem.Text
   End If
   Kill strData
   LoadLocalDirectory
  End If
 End If
End If
End Sub


Private Sub ListView2_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim strDropData As String
If ListView1.SelectedItem.SmallIcon <> 2 Then Exit Sub
strDropData = Data.GetData(vbCFText)
If MsgBox("Download " & strDropData & "?", vbYesNo, "Confirm Download") = vbNo Then
 Exit Sub
End If

CMDdownload_Click
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim strDropData As String
If ListView2.SelectedItem.SmallIcon <> 2 Then Exit Sub
strDropData = Data.GetData(vbCFText)
If MsgBox("Upload " & strDropData & "?", vbYesNo, "Confirm Upload") = vbNo Then
 Exit Sub
End If

UpLoadFile
End Sub

Private Sub mnuExit_Click()
SaveSetting "PlainJane", vbNullChar, "Servers", SVname.ListCount - 1
Unload Me
End Sub

Private Sub mnulistview1Rename_Click()
Dim strData As String
strData = InputBox("Enter new filename [Blank to cancel]:", "Rename File")
If strData = "" Then Exit Sub
engine1.Rename "", ListView1.SelectedItem.Text, strData
LoadCurrentDirectory
End Sub

Private Sub mnulistview1Delete_Click()
 If ListView1.SelectedItem.SmallIcon = 2 Then
  If MsgBox("Delete " & ListView1.SelectedItem.Text & vbCrLf & "Are you sure?", vbYesNo, "Confirm Delete") = vbYes Then
   engine1.KillFile "", ListView1.SelectedItem.Text
   LoadCurrentDirectory
  End If
 End If
End Sub

Private Sub mnuViewFile_Click()
 If ListView1.SelectedItem.SmallIcon = 2 Then
  StatusBar1.Panels(1).Text = "Getting File"
  engine1.GetFile "", ListView1.SelectedItem.Text, App.Path & "\qrzxd.zzz", True
  Viewer.Show
  Viewer.RichTextBox1.LoadFile App.Path & "\qrzxd.zzz"
  StatusBar1.Panels(1).Text = "Connected"
  Kill App.Path & "\qrzxd.zzz"
 End If
End Sub

Private Sub OptAnon_Click()
If OptAnon.Value Then
 UrName.Text = "anonymous"
 Pword.Text = "Me@me.com"
 Port.Text = "21"
End If
End Sub



Private Sub Port_Change()
CheckServerFilledIn
End Sub

Private Sub Pword_Change()
CheckServerFilledIn
End Sub



Private Sub CheckServerFilledIn()
If SVname.Text <> "" And UrName.Text <> "" And Pword.Text <> "" And Port.Text <> "" Then
 ServerSelected = True
 CMDconnect.Enabled = True
Else
 ServerSelected = False
 CMDconnect.Enabled = False
End If
End Sub



Private Sub SVname_Change()

a = SVname.ListIndex
CheckServerFilledIn
End Sub



Private Sub SVname_GotFocus()
UrName.Text = GetSetting("PlainJane", "Server" & Str(SVname.ListIndex), "UserName")
Pword.Text = GetSetting("PlainJane", "Server" & Str(SVname.ListIndex), "Password")
Port.Text = GetSetting("PlainJane", "Server" & Str(SVname.ListIndex), "Port")
End Sub

Private Sub SVname_LostFocus()
UrName.Text = GetSetting("PlainJane", "Server" & Str(SVname.ListIndex), "UserName")
Pword.Text = GetSetting("PlainJane", "Server" & Str(SVname.ListIndex), "Password")
Port.Text = GetSetting("PlainJane", "Server" & Str(SVname.ListIndex), "Port")
End Sub

Private Sub UrName_Change()
CheckServerFilledIn
End Sub

Private Sub LoadCurrentDirectory()
Dim FileNum As Integer: FileNum = FreeFile()
Dim strData As String
Dim Marker1 As Integer: Dim Marker2 As Integer
Dim Filename As String: Dim FileType As String
Dim FileSize As Long: Dim Ftype As Integer
StatusBar1.Panels(1).Text = "Getting directory..."
Transfering = True
ListView1.ListItems.Clear
engine1.DirListing App.Path & "\tmpdir.dat"
Open App.Path & "\tmpdir.dat" For Input As #FileNum
While Not EOF(FileNum)
 Input #FileNum, strData
 If Len(strData) < 30 Then GoTo SkipLine
 Marker1 = Len(strData): Marker2 = Marker1
 While InStr(Marker1, strData, " ") = 0
  Marker1 = Marker1 - 1
 Wend
 Filename = Mid(strData, Marker1 + 1)
 Marker1 = 31
 FileType = Mid(strData, 1, 1)
 If FileType = "d" Then Ftype = 1 Else Ftype = 2
 Marker1 = 1: Marker2 = 0
 While Marker2 <> 4
  While Mid(strData, Marker1, 1) <> " "
   Marker1 = Marker1 + 1
  Wend
  While Mid(strData, Marker1, 1) = " "
   Marker1 = Marker1 + 1
  Wend
  Marker2 = Marker2 + 1
 Wend
 FileSize = Val(Mid(strData, Marker1))
 With ListView1
  .ListItems.Add .ListItems.Count + 1, , Filename, , Ftype
  .ListItems.Item(.ListItems.Count).ListSubItems.Add 1, , Str(FileSize)
  .ListItems.Item(.ListItems.Count).ListSubItems.Add 2, , Mid(strData, 1, 10)
 End With
SkipLine:
Wend
Close #FileNum
StatusBar1.Panels(1).Text = "Connected"
Transfering = False
End Sub

Private Sub LoadLocalDirectory()
Dim x As Integer
Dim old As Integer
Dim strPath As String

LocalPath.Text = Dir1.Path
strPath = LocalPath.Text
If Right(strPath, 1) = "\" Then
 strPath = Mid(strPath, 1, Len(strPath) - 1)
End If
File1.Refresh
With ListView2
 .ListItems.Clear
 For x = 0 To Drive1.ListCount - 1
  .ListItems.Add x + 1, , Drive1.List(x), , 3
 Next x
 old = x
 If Dir1.ListCount Then
  For x = 0 To Dir1.ListCount - 1
   .ListItems.Add x + old, , Dir1.List(x), , 1
  Next x
  old = x
 End If
 For x = 0 To File1.ListCount - 1
  .ListItems.Add x + old, , File1.List(x), , 2
  .ListItems(x + old).ListSubItems.Add 1, , Str(FileLen(strPath & "\" & File1.List(x)))
  .ListItems(x + old).ListSubItems.Add 2, , FileDateTime(strPath & "\" & File1.List(x))
 Next
End With
SaveSetting "PlainJane", vbNullChar, "LocalPath", strPath
End Sub

Private Sub UpLoadFile()
Dim UpFileName As String
Dim Extension As String
Dim XferMode As String: XferMode = "0"
Dim XferFlag As Boolean

Screen.MousePointer = vbHourglass
UpFileName = LocalPath.Text & "\" & ListView2.SelectedItem.Text
UpFileName = Replace(UpFileName, "\\", "\", 1, -1)
Extension = Mid(UpFileName, InStr(1, UpFileName, ".") + 1)

If Option5.Value Then
 XferMode = GetSetting("PlainJane", "Extensions", Extension, 1)
Else
 If Option4.Value Then
  XferMode = 1
 End If
End If
SVresponse.Text = SVresponse.Text & "UpLoading " & ListView2.SelectedItem.Text
If XferMode = "1" Then
 SVresponse.Text = SVresponse.Text & " Binary Mode."
 XferFlag = False
Else
 SVresponse.Text = SVresponse.Text & " Ascii Mode."
 XferFlag = True
End If
SVresponse.Text = SVresponse.Text & vbCrLf
SaveSetting "PlainJane", "Extensions", Extension, XferMode

Transfering = True
engine1.PutFile "", ListView2.SelectedItem.Text, UpFileName, XferFlag
Transfering = False
StatusBar1.Panels(4).Text = ""
LoadCurrentDirectory
Screen.MousePointer = vbNormal
End Sub
