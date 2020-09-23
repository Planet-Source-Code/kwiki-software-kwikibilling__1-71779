VERSION 5.00
Begin VB.Form frmSysUpdate 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   825
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin KwikiBilling.XP_ProgressBar PB 
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   344
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
         Color           =   49152
         Scrolling       =   2
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Updating System Data ..."
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSysUpdate"
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
    Do: Loop Until timeGetTime - lTimer > 30
        If PB.Value < 5 Then
ElseIf PB.Value > 80 Then
lblProgress.Caption = "Update Completed .."
End If
DoEvents
Next
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

