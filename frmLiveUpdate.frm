VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLiveUpdate 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1530
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   4020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLiveUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LiveUpdate.XP_ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
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
      Color           =   49152
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   600
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmLiveUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
On Error GoTo Err:
    Dim Reply As Variant
    Dim TransferSuccess As Boolean
    UpdateTime = 0
    Timer2.Interval = 1000
    'Command1.Enabled = False
    ProgressBar1.Value = 3
    status$ = "Checking For Updated Version ..."
    TransferSuccess = GetInternetFile(Inet1, "http://invoice.x10hosting.com/liveupdate/update.html", App.Path & "\Update\")

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Timer2.Interval = 0
        Exit Sub
    End If
       
    ProgressBar1.Value = 25
    status$ = "Connected.."
    
    Open App.Path & "\Update\update.html" For Input As #1
        Input #1, updatever$
    Close #1
      
    If updatever$ > myVer Then
        Label1.Caption = "Please Wait .."
    Else
        Label1.Caption = "There are no updates available at this time"
        ProgressBar1.Value = 3
        MsgBox "There are no updates available at this time"
        Timer2.Interval = 0
        Me.Hide
        Reply = MsgBox("Would you like to restart the application now? " & "", vbYesNo, "Restart Kwiki Billing")
        Select Case Reply
        Case vbYes:
        Shell "KwikiBilling.exe"

        Case vbNo:
        End
        End Select
        'Shell "Bias(Admin).exe"
        End
    End If

    status$ = "Downloading Updates.."

    'MsgBox ("")
    TransferSuccess = GetInternetFile(Inet1, "http://invoice.x10hosting.com/liveupdate/Update.exe", App.Path)

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Command1.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If
    Label1.Caption = "Completed Downloading Updated Version " + updatever
    ProgressBar1.Value = 100
    Timer2.Interval = 0
    
    'X = MsgBox("Live Update has completed downloading necessary files")
    'Command1.Enabled = True
    
    Shell "Update.exe"
    End
Err:
End Sub

Private Sub Form_Load()
On Local Error GoTo 200
WindowState = 0
' myVer = App.Major & "." & App.Minor & "." & App.Revision


' this is where the updated program needs to write it's current version
' number to.  The above commented out line puts the version number in
' the correct format.

status$ = "Idle"
UpdateTime = 0


Open App.Path & "\Update\ver.dat" For Input As #1
    Input #1, myVer
Close #1

Exit Sub

200 myVer = App.Major & "." & App.Minor & "." & App.Revision
'X = MsgBox("Current Version " & myVer)

Resume 205

205 End Sub

Private Sub lblClose_Click()
End
End Sub

Private Sub Timer1_Timer()
If Inet1.StillExecuting = False Then
    Label2.Caption = "Status: Server down.."
    lblClose.Visible = True
Else
    Label2.Caption = "Status: " & status$
    lblClose.Visible = False
End If

End Sub

Private Sub Timer2_Timer()
    UpdateTime = UpdateTime + 1
    Label3.Caption = "Connection Time:" & Str$(UpdateTime) & " Seconds"
End Sub

