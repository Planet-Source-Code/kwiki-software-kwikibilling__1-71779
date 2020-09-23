VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmVendors 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Vendor Accounts"
   ClientHeight    =   3705
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   9120
   Icon            =   "frmVendors.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo dcbVendor 
      Height          =   330
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   7680
      TabIndex        =   25
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVendors.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   6120
      TabIndex        =   24
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVendors.frx":0A06
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVendors.frx":12E0
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVendors.frx":1BBA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Edit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVendors.frx":2494
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Add"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVendors.frx":2D6E
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "Supplier Accounts"
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   8895
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpAddDate 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483634
         CheckBox        =   -1  'True
         Format          =   45416449
         CurrentDate     =   39687
      End
      Begin VB.Label Label2 
         Caption         =   "Account # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Supplier Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Street Address :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "City, State, ZipCode :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Phone Number :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Fax Number :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Special Notes :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
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
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Date Added :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Update"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVendors.frx":3648
      cBack           =   -2147483633
   End
   Begin VB.Label Label11 
      Caption         =   "Select Vendor :"
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
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private CNN As ADODB.Connection
Dim Clear As String

Private Sub cmdEdit_Click()
DisableCont
EnableFields
cmdSave.Visible = False
cmdUpdate.Visible = True
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo EH:
Dim rsEdit As Recordset
Dim sql As String
sql = "Select * FROM Vendors Where VenderID = " & Text1
Set rsEdit = DB.OpenRecordset(sql)

With rsEdit
.Edit
!AddDate = dtpAddDate.Value
!AccountNum = Text2.Text
!SupplierName = Text3.Text
!Address = Text4.Text
!Phone = Text5.Text
!Fax = Text6.Text
!Notes = Text7.Text
.Update
End With

    EnableCont
    DisableFields
    'FillFields
    
cmdSave.Enabled = False
cmdUpdate.Visible = False
EH:
End Sub

Private Sub dcbVendor_Change()
On Error GoTo EH:
'SndPlayEx App.Path & "\Sounds\OpenMenu.wav"

If Not rs.BOF Then rs.MoveFirst
rs.Find "VenderID = " & dcbVendor.BoundText, 0, adSearchForward, 0

FillFields
Exit Sub
EH:
End Sub

Private Sub Form_Load()
'----------------------------------------------
Set CNN = New ADODB.Connection
CNN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\KwikiDat\db2.mdb"
CNN.Open
  
Set rs = New ADODB.Recordset
rs.Open "Select * from Vendors", CNN, adOpenStatic, adLockOptimistic

Set dcbVendor.RowSource = rs
dcbVendor.ListField = "SupplierName"
dcbVendor.BoundColumn = "VenderID"
'----------------------------------------------

'FillFields
DisableFields
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddErr
DisableCont
cmdSave.Enabled = True
EnableFields
ClearFields
dtpAddDate = Date
Text2.SetFocus

Exit Sub
AddErr:

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
cmdEdit.Visible = True
cmdSave.Visible = True
cmdSave.Enabled = False
cmdUpdate.Visible = False
cmdAdd.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdDelete.Enabled = True
'cmdRefresh.Enabled = True
cmdClose.Enabled = True
'cmdUpdate.Visible = False
dcbVendor.Enabled = True

FillFields
DisableFields
End Sub

Private Sub cmdClose_Click()
SndClick
Set rs = Nothing
Unload Me
End Sub


Private Sub cmdDelete_Click()
On Error GoTo DeleteErr:
If Text2.Text = "" Then
MsgBox "No vendor selected to delete"
Else

Dim sql As String
sql = "FROM Vendors WHERE VendorID = " & Text1.Text

If MsgBox("Are you sure you want to Remove ? " & Text3.Text, vbYesNo, "Confirm") = vbYes Then
rs.Delete
ClearFields
ElseIf vbNo Then
Exit Sub
End If
End If
Exit Sub
DeleteErr:
End Sub

Private Sub cmdSave_Click()
On Error GoTo EH:
If ValidateFields = False Then Exit Sub

With rs
.AddNew
!AddDate = dtpAddDate.Value
!AccountNum = Text2.Text
!SupplierName = Text3.Text
!Address = Text4.Text
!Phone = Text5.Text
!Fax = Text6.Text
!Notes = Text7.Text
.Update
.MoveLast
End With

EnableCont
DisableFields
FillFields
'Unload Me
Exit Sub
EH:
End Sub

Private Sub DisableCont()
cmdAdd.Enabled = False
cmdCancel.Enabled = True
cmdEdit.Visible = False
cmdSave.Visible = True
cmdDelete.Enabled = False
'cmdRefresh.Enabled = False
cmdClose.Enabled = False
'cmdUpdate.Visible = False
dcbVendor.Enabled = False
End Sub

Private Sub EnableCont()
cmdAdd.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Visible = True
cmdEdit.Enabled = True
cmdSave.Visible = True
cmdDelete.Enabled = True
'cmdRefresh.Enabled = True
cmdClose.Enabled = True
dcbVendor.Enabled = True

End Sub

Private Sub DisableFields()
Text1.Locked = True ' Auto inc
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
dtpAddDate.Enabled = False
End Sub

Private Sub EnableFields()
Text1.Locked = True ' Auto inc
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
dtpAddDate.Enabled = True
End Sub

Private Sub FillFields()
On Error GoTo EH
Text1.Text = rs!VenderID
Text2.Text = rs!AccountNum
Text3.Text = rs!SupplierName
Text4.Text = rs!Address
Text5.Text = rs!Phone
Text6.Text = rs!Fax
Text7.Text = rs!Notes
dtpAddDate.Value = rs!AddDate

Exit Sub
EH:
End Sub

Private Sub ClearFields()
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
Text6.Text = Clear
Text7.Text = Clear
dtpAddDate.Value = Date
dcbVendor.Text = Clear
End Sub

Private Function ValidateFields() As Boolean

If Text3.Text = "" Then
MsgBox "Supplier Name must not be blank"
ValidateFields = False
Text3.SetFocus
Exit Function
End If

If Text2.Text = "" Then
MsgBox "Account Number must not be blank"
ValidateFields = False
Text2.SetFocus
Exit Function
End If

If Text4.Text = "" Then
MsgBox "Address must not be blank"
ValidateFields = False
Text4.SetFocus
Exit Function
End If

If Text5.Text = "" Then
MsgBox "Phone Number must not be blank"
ValidateFields = False
Text5.SetFocus
Exit Function
End If

ValidateFields = True
End Function

Private Sub Form_Unload(Cancel As Integer)
frmTree.TvwCustomer.SetFocus
End Sub
