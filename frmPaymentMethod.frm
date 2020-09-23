VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPaymentMethod 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Payment Methods"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6285
   Icon            =   "frmPaymentMethod.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmPaymentMethod.frx":000C
      DataSource      =   "Adodc1"
      Height          =   1425
      Left            =   4080
      TabIndex        =   11
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2514
      _Version        =   393216
      BackColor       =   12648384
      ForeColor       =   -2147483640
      ListField       =   "PaymentMethod"
      BoundColumn     =   "PaymentMethodID"
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "If the New Payment Method is a credit card be sure to validate the credit card box by unchecking and rechecking after clicking me"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      DataField       =   "PaymentMethod"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CheckBox chkCreditCard 
      Alignment       =   1  'Right Justify
      Caption         =   "Check For Credit Card"
      DataField       =   "CreditCard"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   2250
   End
   Begin VB.TextBox txtPaymentMethodID 
      BackColor       =   &H80000016&
      DataField       =   "PaymentMethodID"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   300
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   0
      Top             =   2400
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   12648384
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT `Payment Methods`.* FROM `Payment Methods`"
      Caption         =   "Payment Method Navigation"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdRequery 
      Caption         =   "&Requery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Current Pay Methods"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PaymentMethod:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1590
   End
End
Attribute VB_Name = "frmPaymentMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo EH
Adodc1.Recordset.CancelUpdate
EnableCont
EH:
End Sub
Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew
chkCreditCard.Value = 0
DisableCont
End Sub

Private Sub cmdRequery_Click()
Adodc1.Refresh
End Sub

Private Sub cmdUpdate_Click()
Adodc1.Recordset.UpdateBatch
Adodc1.Refresh
EnableCont
Label2.Caption = Adodc1.Recordset.RecordCount & " Existing Categories"
End Sub

Private Sub cmdClose_Click()
SndClick
frmPayments.Adodc1.Refresh
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
If MsgBox("Are you sure you want to Remove ? " & Text1.Text, vbYesNo, "Confirm") = vbYes Then

Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext

If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveLast
End If

Label2.Caption = Adodc1.Recordset.RecordCount & " Existing Categories"
End If

frmTree.TvwCustomer.SetFocus

Exit Sub
DeleteErr:
End Sub

'Private Sub DataList1_Click()
'txtPaymentMethodID.Text = DataList1.BoundText

'Dim rsPMeth As Recordset
'Dim sql As String
'sql = "Select * FROM [Payment Methods] Where PaymentMethodID = " & DataList1.BoundText
'Set rsPMeth = DB.OpenRecordset(sql)
'chkCreditCard.DataSource = rsPMeth
'chkCreditCard.DataField = rsPMeth!CreditCard

'Text1.Text = rsPMeth!PaymentMethod
'chkCreditCard.Value = rsPMeth!CreditCard
'End Sub

Private Sub Form_Load()
Label2.Caption = Adodc1.Recordset.RecordCount & " Existing Categories"
chkCreditCard.Value = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub DisableCont()
Adodc1.Enabled = False
cmdAdd.Enabled = False
cmdUpdate.Enabled = True
cmdCancel.Enabled = True
cmdDelete.Enabled = False
cmdRequery.Enabled = False
cmdClose.Enabled = False
End Sub

Private Sub EnableCont()
Adodc1.Enabled = True
cmdAdd.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
cmdDelete.Enabled = True
cmdRequery.Enabled = True
cmdClose.Enabled = True
End Sub
