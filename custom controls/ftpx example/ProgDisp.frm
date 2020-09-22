VERSION 5.00
Object = "{27AC8959-B6DE-11D4-81CA-00C0F010F316}#2.0#0"; "FTPPROGRESS.OCX"
Begin VB.Form ProgDisp 
   BackColor       =   &H80000008&
   Caption         =   "Transfer Progress"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "ProgDisp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   240
      Left            =   2190
      TabIndex        =   1
      Top             =   3180
      Width           =   1125
   End
   Begin FTPprogress.ProgressDisplay ProgressDisplay1 
      Height          =   2925
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   5159
      PercentDone     =   ""
      PercentDone     =   ""
   End
End
Attribute VB_Name = "ProgDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Mainform.engine1.Cancel
Unload Me
End Sub
