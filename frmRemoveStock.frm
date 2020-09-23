VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Load Stock Entry.."
   ClientHeight    =   3915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8895
   Icon            =   "frmRemoveStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLoadStock 
      Caption         =   "&Load Entry"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwStock 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      Left            =   7320
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset
Dim sqlList As String


Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "PartID", 0, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "Vendor", 0, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "Vendor Acct #", 0, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "PartCode", 1400, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "PartName", 2200, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "Description", 2200, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "Received #", 1200, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "Stock Date", 1200, lvwColumnLeft)

lvwStock.View = lvwReport
End Sub

Public Sub LoadStockList()
Dim sqlItem As ListItem


lvwStock.ListItems.Clear

sqlList = "SELECT * From StockTrack "
Set rs = DB.OpenRecordset(sqlList)

If (rs.RecordCount > 0) Then
rs.MoveFirst
End If

While Not rs.EOF
Set sqlItem = lvwStock.ListItems.Add(, , _
rs!PartID)
sqlItem.SubItems(1) = rs!Vendor
sqlItem.SubItems(2) = rs!VendorAcctNum
sqlItem.SubItems(3) = rs!PartCode
sqlItem.SubItems(4) = rs!PartName
sqlItem.SubItems(5) = rs!PartDescription
sqlItem.SubItems(6) = rs!ReceivedAmount
sqlItem.SubItems(7) = rs!StockDate

rs.MoveNext
Wend
End Sub

Private Sub cmdClose_Click()
Set rs = Nothing
Unload Me
End Sub

Private Sub Form_Load()
OpenDatabase

If (Not OpenDatabase()) Then
  MsgBox "Database could not be openend !"
End If

setUpListView
LoadStockList
End Sub

Private Sub cmdLoadStock_Click()
On Error Resume Next

If lvwStock.SelectedItem.SubItems(1) = "" Then
MsgBox "No stocked parts to load"
Exit Sub
End If

frmRecieveStock.Text1.Text = lvwStock.SelectedItem.Text
frmRecieveStock.dcbVendor.Text = lvwStock.SelectedItem.SubItems(1)
frmRecieveStock.Text8.Text = lvwStock.SelectedItem.SubItems(2)
frmRecieveStock.Text2.Text = lvwStock.SelectedItem.SubItems(3)
frmRecieveStock.Text3.Text = lvwStock.SelectedItem.SubItems(4)
frmRecieveStock.Text4.Text = lvwStock.SelectedItem.SubItems(5)
frmRecieveStock.Text5.Text = lvwStock.SelectedItem.SubItems(6)
frmRecieveStock.Text6.Text = lvwStock.SelectedItem.SubItems(7)

frmRecieveStock.Text10.Text = Format$(frmRecieveStock.Text7 - frmRecieveStock.Text5, "0;0")

frmRecieveStock.cmdDelEntry.Visible = True
frmRecieveStock.cmdDelete.Visible = False
Unload Me
End Sub

