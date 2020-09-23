VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmRecieveStock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Vendor Stock Receiving Window"
   ClientHeight    =   3450
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   7920
   Icon            =   "frmRecieveStock.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Received Stock"
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7695
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   375
         Left            =   6120
         TabIndex        =   27
         Top             =   2400
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
         Image           =   "frmRecieveStock.frx":000C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdAddStock 
         Height          =   375
         Left            =   4560
         TabIndex        =   26
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Add Stock"
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
         Image           =   "frmRecieveStock.frx":0A06
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdView 
         Height          =   375
         Left            =   2160
         TabIndex        =   25
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "&View Stock Entries"
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
         Image           =   "frmRecieveStock.frx":12E0
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Delete Entry"
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
         Image           =   "frmRecieveStock.frx":1732
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   480
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text7 
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
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
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
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text3 
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
         Height          =   360
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   2175
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
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dcbPart 
         Height          =   360
         Left            =   5280
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcbVendor 
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Current Stock :"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Date Added :"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Supplier Name :"
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
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "SKU Number:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Part Name :"
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
         Left            =   4200
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Part Description :"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Stock Received :"
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
         Left            =   3720
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Select Part :"
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
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   135
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "   Stocked Received That is available to add to part index                                    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmRecieveStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private WithEvents rsPart As ADODB.Recordset
Attribute rsPart.VB_VarHelpID = -1
Private CNN As ADODB.Connection
'Dim Clear As String

Private Sub cmdAddStock_Click()
On Error GoTo EH:

If dcbVendor.Text = "" Then
MsgBox "You must select a vendor first"
Exit Sub
End If

If Text5.Text = "" Then
MsgBox "Received stock cannot contain a null value"
Text5.SetFocus
Exit Sub
End If

If Text5.Text = "0" Then
MsgBox "Received stock cannot contain a null value"
Text5.SetFocus
Exit Sub
End If

OpenDatabase
Dim rsTrack As Recordset
Dim sqlTrack As String
sqlTrack = "SELECT * FROM StockTrack"
Set rsTrack = DB.OpenRecordset(sqlTrack)

With rsTrack 'ADD TO STOCKTRACK
.AddNew
!PartID = Text1.Text
!PartCode = Text2.Text
!PartName = Text3.Text
!PartDescription = Text4.Text
!StockDate = Text6.Text
!ReceivedAmount = Text5.Text
!Vendor = dcbVendor.Text
!VendorAcctNum = Text9.Text
.Update
.MoveLast
End With

UpdatePartStock
Unload Me
frmStockMsg.Show
Exit Sub
EH:
MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
SndClick
Set rs = Nothing
Set rsPart = Nothing
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DelErr:
'GetReceivedAmount

'Dim rsChk As Recordset
'Dim sqlChk As String

'sqlChk = "SELECT UnitsStock, PartID FROM Parts WHERE PartID = " & Text1.Text
'Set rsChk = DB.OpenRecordset(sqlChk)

'Text11.Text = Format$(rsChk!UnitsStock, "0;0")


'If Not Text11.Text = Text12.Text Then
'MsgBox ("This stock cannot be deleted because it is assigned to an order")
'Exit Sub
'Else

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

If MsgBox("Are you sure you want to Remove This Stock Entry", vbYesNo, "Confirm") = vbYes Then

OpenDatabase
Dim rsTrack As Recordset
Dim sqlTrack As String

sqlTrack = "SELECT * FROM StockTrack WHERE ID = " & Text10.Text
Set rsTrack = DB.OpenRecordset(sqlTrack)
rsTrack.Delete
RemovePartStock

frmStockMsg.Show
Unload Me

ElseIf vbNo Then
Exit Sub

End If
'End If
Exit Sub
DelErr: MsgBox Err.Description
End Sub

Private Sub cmdView_Click()
frmViewStock.Show
End Sub

Private Sub dcbPart_Change()
On Error GoTo EH:
'SndPlayEx App.Path & "\Sounds\OpenMenu.wav"

If Not rsPart.BOF Then rsPart.MoveFirst
rsPart.Find "PartID = " & dcbPart.BoundText, 0, adSearchForward, 0

FillFields
cmdAddStock.Enabled = True
Exit Sub
EH:
End Sub

Private Sub dcbVendor_Change()
On Error GoTo EH:
'SndPlayEx App.Path & "\Sounds\OpenMenu.wav"

If Not rs.BOF Then rs.MoveFirst
rs.Find "VenderID = " & dcbVendor.BoundText, 0, adSearchForward, 0

GetVenNum
Exit Sub
EH:
End Sub

Private Sub Form_Load()
Set CNN = New ADODB.Connection
CNN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\KwikiDat\db2.mdb"
CNN.Open
  
Set rs = New ADODB.Recordset
rs.Open "Select * from Vendors", CNN, adOpenStatic, adLockOptimistic

Set dcbVendor.RowSource = rs
dcbVendor.ListField = "SupplierName"
dcbVendor.BoundColumn = "VenderID"

Text6 = Date
'------------------------------------------------------------------------

Set rsPart = New ADODB.Recordset
rsPart.Open "Select * from Parts", CNN, adOpenStatic, adLockOptimistic

Set dcbPart.RowSource = rsPart
dcbPart.ListField = "PartName"
dcbPart.BoundColumn = "PartID"

End Sub

Private Sub FillFields()
On Error GoTo EH
Text1.Text = rsPart!PartID
Text2.Text = rsPart!PartCode
Text3.Text = rsPart!PartName
Text4.Text = rsPart!PartDescription
Text7.Text = rsPart!UnitsStock

Exit Sub
EH:
End Sub

Private Sub UpdatePartStock()
On Err GoTo EH:
Dim rs As Recordset
Dim sql As String
sql = "SELECT PartID, UnitsStock FROM Parts WHERE PartID = " & Text1.Text
Set rs = DB.OpenRecordset(sql)

Text8.Text = Format$(rs!UnitsStock + Text5, "0;0")

With rs
.Edit
!UnitsStock = Text8.Text
.Update
End With
Exit Sub
EH:
End Sub

Private Sub GetVenNum()
On Error GoTo EH:
Text9.Text = rs!AccountNum
Exit Sub
EH:
End Sub

Private Sub RemovePartStock()
On Err GoTo EH:
Dim rs As Recordset
Dim sql As String
sql = "SELECT PartID, UnitsStock FROM Parts WHERE PartID = " & Text1.Text
Set rs = DB.OpenRecordset(sql)

Text8.Text = Format$(rs!UnitsStock - Text5, "0;0")

With rs
.Edit
!UnitsStock = Text8.Text
.Update
End With
Exit Sub
EH:
End Sub

Public Sub GetReceivedAmount()
OpenDatabase
Dim rs As Recordset
Dim sqlList As String

sqlList = "SELECT * From StockTrack WHERE ID = " & Text10.Text
Set rs = DB.OpenRecordset(sqlList)

Text12.Text = Format$(rs!ReceivedAmount, "0;0")
End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub
