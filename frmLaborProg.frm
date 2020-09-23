VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaborProg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Please Wait Updating Labor..."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmLaborProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'that form unloaded correctly
'if progress bar in 'test' mode
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Form_Activate()
Dim Count As Long
  Dim Progress As Long
  Dim lTimer As Long
  
  For Progress = 1 To 100
    PB.Value = Progress
    lTimer = timeGetTime
    Do: Loop Until timeGetTime - lTimer > 3
    DoEvents    'Allows user to change styles etc whilst in progress
  Next
Unload Me
End Sub

Private Sub Form_Load()
Left = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmWorkorderLabor.Show
End Sub


