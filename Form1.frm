VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Cool writer -By [NL]The_Magic"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "C:\Program\Microsoft Office\Office\WinWord.exe"
      Top             =   3120
      Width           =   3615
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   99999
      SelStart        =   20000
      TickStyle       =   3
      Value           =   20000
   End
   Begin VB.TextBox Text1 
      Height          =   2445
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Type it!!!"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fast                                                   Slow"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
wrd = Shell(Text2, 1)
AppActivate wrd
Dim aass
For aa = 1 To Len(Text1)
aass = Mid(Text1, aa, 1)
If aass = Chr(10) Then GoTo enterkey

SendKeys aass, True

For gg = 1 To Slider1.Value
DoEvents
Next gg
enterkey:
Next aa
End Sub

Private Sub Command2_Click()
CD1.Filter = "Microsoft Word or Word Pad|winword.exe;wordpad.exe"
CD1.Action = 1
CD1.DialogTitle = "Select location"
Text2 = CD1.FileName
End Sub
