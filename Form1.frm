VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "fLASH eXTRACTOR"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3960
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "Open Flash-Projector"
      Height          =   495
      Left            =   150
      TabIndex        =   3
      Top             =   840
      Width           =   1950
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Unten ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   1530
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   615
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.exe"
      DialogTitle     =   "Flash-Projektor öffnen..."
      Filter          =   "Projektor|*.exe"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001 SilverMoon Software e.K."
      Height          =   240
      Left            =   675
      TabIndex        =   2
      Top             =   345
      Width           =   6195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "fLASH eXTRACOR"
      Height          =   210
      Left            =   675
      TabIndex        =   1
      Top             =   120
      Width           =   6045
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":0442
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  On Error Resume Next
  With CommonDialog1
    .ShowOpen
    If Err.Number > 0 Then Exit Sub
      Dim TeMp As String
      TeMp = String((FileLen(.FileName) - 286744), Chr(0))
      If .FileName <> "" Then
      ProgressBar1.Value = 25
      Open .FileName For Binary As #1
      Open .FileName & ".swf" For Binary As #2
        Seek #1, 286721
        ProgressBar1.Value = 50
        Get #1, , TeMp
        ProgressBar1.Value = 75
        Put #2, , TeMp
      Close #1
      Close #2
      ProgressBar1.Value = 100
      ret = MsgBox("Test FLASH-Animation?", vbYesNo + vbQuestion)
      If ret = vbYes Then Form2.Tag = .FileName & ".swf": Form2.play
    End If
  End With
End Sub

Private Sub Form_Load()

End Sub
