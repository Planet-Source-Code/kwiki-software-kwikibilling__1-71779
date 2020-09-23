VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2880
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin lvButton.lvButtons_H Command2 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "&Register"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   -2147483634
         cFHover         =   -2147483634
         cBhover         =   49152
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   32768
         mPointer        =   99
         mIcon           =   "frmSplash.frx":000C
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   8000
         Left            =   720
         Top             =   1800
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   8000
         Left            =   240
         Top             =   1800
      End
      Begin VB.PictureBox Picture 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         Picture         =   "frmSplash.frx":016E
         ScaleHeight     =   735
         ScaleWidth      =   3375
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         Begin VB.Label lblValid 
            BackColor       =   &H80000009&
            Caption         =   "6502-X03M-D139-L4N1"
            Height          =   255
            Left            =   3360
            TabIndex        =   2
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtValid 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   14
         Text            =   "6502-X03M-D139-L4N1"
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin KwikiBilling.XP_ProgressBar PB 
         Height          =   270
         Left            =   3480
         TabIndex        =   16
         Top             =   2400
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   476
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
      Begin VB.Label Label3 
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   11
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Kwiki Billing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Author: Kwiki Billing - Developer Team"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Kwiki Billing Business Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   8
         Top             =   1080
         Width           =   3000
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright (2008)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "LicenseTo : Kwiki Billing - Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         Height          =   135
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "   Single Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private mobjConn As ADODB.Connection
Private mobjCmd As ADODB.Command
Private mobjRst As ADODB.Recordset

Dim rsKey As Recordset
Dim sql As String

Public Sub Command1_Click()
On Error Resume Next
Dim str1 As String
Dim str2 As String
Dim dlen
Dim regi
Dim pass As String

Close
'OPEN THE FILE AND GET BOTH STRINGS OF INFO
Open "C:\WINDOWS\system\Reg.ini" For Input As #1
Do Until EOF(1)
Line Input #1, str1
Line Input #1, str2
Loop
'CHECK TO SEE IF IT HAS BEEN RIGISTERED
'IF IT HAS THERE IS NO NEED TO GO ON JUST EXIT SUB
If str1 = "Registered" And str2 = " " Then
Label6.Caption = "Registered"
Call Timer1_Timer
Exit Sub
Else
End If
'SO IT HASNT BEEN REGISTERED CHECK TO SEE IF IT HAS RAN
'ITS TRIAL PERIOD
dlen = Mid(str1, 1, 2)
'DLEN IS THE NUMBER OF USES
regi = Mid(str2, 1, 12)
'REGI IS IF IT HAS BEEN REGISTERED OR NOT
Close #1

'IF THE NUMBER OF USES EXCEEDS NINE ASK FOR THE REGISTRATION CODE
If dlen > 9 Then
DisableSplash
    pass = InputBox("Please enter Registration Key, To obtain activation key go to http://invoice.x10hosting.com", "Enter Serial Key")
'CHECK TO SEE IF THE CODE IS CORRECT

        If pass = frmSplash.txtValid.Text Then
'IT IS? OH GOOD THEN OPEN THE REG.INI FILE AND MARK IT SO
            Open "C:\WINDOWS\system\Reg.ini" For Output As #1
            Print #1, "Registered" & vbCrLf & " "
            Close #1
      MsgBox ("Thank You, you have successfully registered this program")
      txtKey.Text = txtValid.Text
      With rsKey
      .AddNew
      !LicenseKey = txtKey.Text
      .Update
      End With
      rsKey.Close
      Command2.Visible = False
      Label6.Caption = "Registered"
      Call Timer1_Timer
      Exit Sub
        Else
'IT ISN'T? OH YOU ARE NAUGHTY WELL HAVE TO STOP THE PROGRAM

        MsgBox "Invalid Registration Key", , "Invalid Key"
        End
        End If
Else
'SO IT HASN'T RAN ITS TRAIL PEROID? WELL HE HAD BETTER
'TELL THE REG.INI FILE THAT THE TRIAL PERIOD IS GETTING
'CLOSER TO EXPIRING

    Open "C:\WINDOWS\system\Reg.ini" For Output As #1
    Print #1, dlen + 1 & vbCrLf & "Unregistered"
    Close #1
End If
MsgBox ("You are in demo mode giving you access to 10 times of use, After you need to register this software via PayPal for 29.95 to receive your activation key")
Command2.Visible = True
Label6.Caption = "Unregistered"
Call Timer1_Timer
End Sub


Private Sub Form_Activate()
Screen.MousePointer = vbHourglass
  Dim Progress As Long
  Dim lTimer As Long

  For Progress = 1 To 100
    PB.Value = Progress
    lTimer = timeGetTime
    Do: Loop Until timeGetTime - lTimer > 30
    If PB.Value <= 5 Then
Screen.MousePointer = vbDefault
Label3.Caption = "Initializing....."
ElseIf PB.Value <= 30 Then
Screen.MousePointer = vbHourglass
Label3.Caption = "Verifying Database....."
ElseIf PB.Value <= 60 Then
Screen.MousePointer = vbDefault
Label3.Caption = "Integrating Database...."
ElseIf PB.Value <= 80 Then
Screen.MousePointer = vbDefault
Label3.Caption = "Updating Database..."
ElseIf PB.Value <= 100 Then
Screen.MousePointer = vbHourglass
Label3.Caption = "Loading Data...."

End If
DoEvents
Next
LoadPB1
'Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblProductName.Caption = App.Title

'Check to see if app is already running
If App.PrevInstance Then
MsgBox "This program is already running"
End
End If

CreateKeyDat
OpenKeyDat
ConnectKey
CreateKeyTable

On Error Resume Next
sql = "SELECT * FROM License"
Set rsKey = KD.OpenRecordset(sql)
If (IsNull(rsKey!LicenseKey)) Then txtKey = rsKey!LicenseKey


'CHECK TO SEE IF REG.INI EXISTS IF NOT
'IF NOT THEN CREATE IT WITH 1 TRY USED
'IF IT IS THERE CHECK TO SEE IF
'A) IT NEEDS REGISTERING
'B) IT HAS BEEN REGISTERED
'-------------------------------------
'CHECK THAT THIS IS THE FIRST TIME THIS FORM HAS BEEN OPENED
'SO IF I IS GREATER THAN ONE IT HAS BEEN LOADED B4 SO IGNORE
'ALL THE CHECKING

Dim i As Long

i = i + 1
If i > 1 Then
MsgBox "Application already loaded"
Exit Sub
End If

If Dir$("C:\WINDOWS\system\Reg.ini") <> "" Then
GoTo bext
Else
Open "C:\WINDOWS\system\Reg.ini" For Output As #1
Print #1, "1  " & vbCrLf & "Unregistered"
Close #1
Label6.Caption = "Unregistered"
'MDIForm1.Text1.Text = "Unregistered"
End If
bext:
Call Command1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
Set rsKey = Nothing
Set mobjCmd = Nothing
Set mobjConn = Nothing
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'Unload Me
'MDIForm1.Show
'End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 8000
'MDIForm1.Show
Me.Show
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = 8000
MDIForm1.Show
Unload Me
End Sub

Private Sub DisableSplash()
Timer1.Interval = 0
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Command2_Click()
ShellExecute 3, "open", "http://invoice.x10hosting.com", vbNullString, vbNullString, 1
End Sub

Private Sub CreateKeyTable()
On Error GoTo EH:
mobjCmd.CommandText = "CREATE TABLE License ([LicenseKey] text(50))"
mobjCmd.Execute
EH:
End Sub

Private Sub ConnectKey()
Set mobjConn = New ADODB.Connection
mobjConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\KwikiDat\key.mdb" & ";Persist Security Info=False"
mobjConn.Open

Set mobjCmd = New ADODB.Command
Set mobjCmd.ActiveConnection = mobjConn
mobjCmd.CommandType = adCmdText
End Sub

Private Sub LoadPB1()
  Dim Progress As Long
  Dim lTimer As Long

  For Progress = 1 To 100
    PB.Value = Progress
    PB.ShowText = True
    lTimer = timeGetTime
    Do: Loop Until timeGetTime - lTimer > 30
    If PB.Value <= 5 Then
Screen.MousePointer = vbHourglass
Label3.Caption = "Populating Data.."
End If
DoEvents
Next
End Sub
