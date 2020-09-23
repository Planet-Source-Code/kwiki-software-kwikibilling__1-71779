VERSION 5.00
Begin VB.Form frmDB 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1320
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin KwikiBilling.XP_ProgressBar PB 
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         Picture         =   "frmDB.frx":0000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Public m_strType As String
Dim m_intTimer As Integer

Private Sub Form_Activate()
On Error Resume Next
'----------------------------------------------------------
Dim count As Long
  Dim Progress As Long
  Dim lTimer As Long

  For Progress = 1 To 100
    PB.Value = Progress
    lTimer = timeGetTime
    Do: Loop Until timeGetTime - lTimer > 60
        If PB.Value <= 5 Then
lblProgress.Caption = "Please Wait...."
lblStatus.Caption = "Status in progress.."
ElseIf PB.Value <= 35 Then
lblProgress.Caption = "Verifying Database...."
ElseIf PB.Value <= 65 Then
lblProgress.Caption = "Validating Integrity...."
ElseIf PB.Value <= 90 Then
lblProgress.Caption = "Checking Database..."
ElseIf PB.Value <= 100 Then
lblStatus.Caption = "Completed Successfully.."
End If
DoEvents
Next
Call Timer1_Timer
'-----------------------------------------------------------
Select Case m_strType
Case "Backup"
'lblProgress.Caption = "Backing Up Current Data ..."
Call BackupDatabase
'Timer1.Enabled = True
Case "Restore"
'lblProgress.Caption = "Restoring Program Data ..."
Call RestoreDatabase
'Timer1.Enabled = True
End Select
End Sub

Private Sub Timer1_Timer()
m_intTimer = m_intTimer + 1
   
If m_intTimer = PB.Value <= 100 Then
Select Case m_strType
Case "Backup"
Unload Me

Case "Restore"
Unload Me

End Select
End If
End Sub
