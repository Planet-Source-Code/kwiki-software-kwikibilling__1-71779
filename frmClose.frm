VERSION 5.00
Begin VB.Form frmClose 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1065
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin KwikiBilling.XP_ProgressBar PB 
         Height          =   135
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   238
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
         Scrolling       =   2
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblProgress 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Closing Database Connections"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmClose"
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
Screen.MousePointer = vbHourglass
'----------------------------------------------------------
  Dim Progress As Long
  Dim lTimer As Long

  For Progress = 1 To 100
    PB.Value = Progress
    lTimer = timeGetTime
    Do: Loop Until timeGetTime - lTimer > 30
If PB.Value < 10 Then
lblStatus.Caption = "Terminating Connections ..."
End If

DoEvents
Next
  For Progress = 1 To 100
    PB.Value = Progress
    lTimer = timeGetTime
    Do: Loop Until timeGetTime - lTimer > 30
If PB.Value < 20 Then
lblStatus.Caption = "Ending Program.."
End If

DoEvents
Next
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub
