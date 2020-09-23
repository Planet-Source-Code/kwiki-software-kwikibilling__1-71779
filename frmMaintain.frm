VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMaintain 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaintain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Close"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMaintain.frx":000C
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maintaince"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5655
      Begin VB.CommandButton cmdBckupData 
         Caption         =   "&Backup Database ( Advised )"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   5415
      End
      Begin VB.CommandButton cmdRepairDB 
         Caption         =   "&Repair Database"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdRestoreDB 
         Caption         =   "&Restore Database"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "frmMaintain.frx":0A06
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "      Database Maintaince"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBckupData_Click()
   On Error GoTo Error_Handler
   
   frmDB.m_strType = "Backup"
   frmDB.Show vbModal, MDIForm1
   
   Exit Sub
Error_Handler:
   MsgBox "An un-known error occurred while Backing Up the Database!" & vbCrLf & _
         "Sorry for the inconvenience"

End Sub

Private Sub cmdCancel_Click()
SndClick
Unload Me
End Sub

'--------------------------------------
'Private Sub cmdCompactDB_Click()
'Call CompactDB
'End Sub
'--------------------------------------


Private Sub cmdRepairDB_Click()
On Error GoTo Error_Handler:
Dim sNewName As String
 Screen.MousePointer = vbHourglass
   'the file name to repair
   sNewName = App.Path & "\KwikiDat\db2.mdb"
   
   Screen.MousePointer = vbHourglass
   MsgBar "Repairing " & sNewName, True
   
   'unload all forms & close the database
   UnloadAllForms
   DB.Close
   Set DB = Nothing
   
   'DBEngine.RepairDatabase sNewName
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   're-open the compacted database
   Call OpenDatabase
   Load MDIForm1
   MDIForm1.UpdateTree
   
   MsgBox "Num: M3 0" & vbCrLf & _
   "No errors was detected to be repaired."
   
   Exit Sub
Error_Handler:
MsgBox "An un-known error occurred while Repairing the Database!" & vbCrLf & _
"Sorry for the inconvenience"
Screen.MousePointer = vbDefault
MsgBar vbNullString, False
End Sub

Private Sub cmdRestoreDB_Click()
   'Restore data
   On Error GoTo Error_Handler
   
   If FileExists(App.Path & "\KwikiDat\db2BACKUP.mdb") = False Then
      MsgBox "There is no Backup Data to Restore from"
      Exit Sub
   End If
   
   If MsgBox("Restoring the Database from backup files will replace the existing database." & vbCrLf & _
         "Are you sure you want to Continue?", vbYesNo, "Restore From Backup Data") = vbYes Then
      'close the running database
      UnloadAllForms
      DB.Close
      Set DB = Nothing
    
      frmDB.m_strType = "Restore"
      frmDB.Show vbModal, MDIForm1
      
   Else
      MsgBox "Database Restore Canceled"
   End If
   
   Exit Sub
Error_Handler:
   MsgBox "An un-known error occurred while Restoring the Database!" & vbCrLf & _
         "Sorry for the inconvenience"
End Sub

Private Sub Form_Load()
frmTree.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmTree.Show
frmTree.TvwCustomer.SetFocus
End Sub
