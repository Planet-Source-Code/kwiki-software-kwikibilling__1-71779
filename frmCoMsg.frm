VERSION 5.00
Begin VB.Form frmCoMsg 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   840
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   3495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin KwikiBilling.XP_ProgressBar PB 
         Height          =   135
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
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
         Color           =   12937777
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmCoMsg"
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
    Do: Loop Until timeGetTime - lTimer > 15
        If PB.Value <= 5 Then
Label1.Caption = "Please Wait...."
ElseIf PB.Value <= 65 Then
Label1.Caption = "Updating Company Records.."
End If
DoEvents
Next
With frmCompanySetup.rsSetup
.Delete
.MoveNext
If .EOF Then .MoveLast
End With
Unload Me
frmCompanySetup.Show
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub


