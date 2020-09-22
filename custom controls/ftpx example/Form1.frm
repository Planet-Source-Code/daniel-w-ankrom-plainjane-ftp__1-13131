VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Viewer 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2505
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   4419
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   270
      Left            =   75
      TabIndex        =   1
      Top             =   2730
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   285
      Left            =   3765
      TabIndex        =   0
      Top             =   2760
      Width           =   870
   End
End
Attribute VB_Name = "Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

