VERSION 5.00
Begin VB.Form frmCustProg 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   840
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   3600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin KwikiBilling.XP_ProgressBar PB 
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
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
         Color           =   32768
         Scrolling       =   5
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmCustProg"
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
    Do: Loop Until timeGetTime - lTimer > 30
        If PB.Value <= 5 Then
Label1.Caption = "Please Wait...."
ElseIf PB.Value <= 70 Then
Label1.Caption = "Please Wait..Updating Records...."
ElseIf PB.Value <= 100 Then
Label1.Caption = "ReLoading Customers Records..."
End If
DoEvents
    
    frmTree.FillListView
    frmTree.UpdateTree
    frmCustomers.LoadCustomerListView
    frmWorkorders.rsOrder.Requery
    frmWorkorders.dcbCustName.Refresh
    
Next
Unload Me
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub
