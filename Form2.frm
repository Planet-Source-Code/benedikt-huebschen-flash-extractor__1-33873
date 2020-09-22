VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "fLASH pLAYER"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   6930
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9960
      _cx             =   17568
      _cy             =   12224
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub play()
  On Error Resume Next
  ShockwaveFlash1.Movie = Me.Tag
  ShockwaveFlash1.play
  Me.Show vbModal, Form1
End Sub
