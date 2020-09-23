VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMsg 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   5430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin MSComctlLib.ProgressBar PB 
         Height          =   135
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Form_Activate()
On Error Resume Next
Dim count As Long
  Dim Progress As Long
  Dim lTimer As Long

  For Progress = 1 To 100
    PB.Value = Progress
    lTimer = timeGetTime
    Do: Loop Until timeGetTime - lTimer > 5
If PB.Value <= 100 Then
Label1.Caption = frmParts.Label8.Caption
End If

DoEvents

Next
Unload Me

End Sub


