VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmViewStock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Viewing Stock"
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   11100
   Icon            =   "frmViewStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   8160
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwStock 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8916
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16056314
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   9720
      TabIndex        =   0
      Top             =   5280
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
Add(, , "ID", 0, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "PartID", 0, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "Vendor", 1200, lvwColumnLeft)
Set clmHdr = lvwStock.ColumnHeaders. _
Add(, , "Vendor Acct #", 1200, lvwColumnLeft)
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
rs!ID)
sqlItem.SubItems(1) = rs!PartID
sqlItem.SubItems(2) = rs!Vendor
sqlItem.SubItems(3) = rs!VendorAcctNum
sqlItem.SubItems(4) = rs!PartCode
sqlItem.SubItems(5) = rs!PartName
sqlItem.SubItems(6) = rs!PartDescription
sqlItem.SubItems(7) = rs!ReceivedAmount
sqlItem.SubItems(8) = rs!StockDate

rs.MoveNext
Wend
End Sub

Private Sub cmdClose_Click()
SndClick
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

frmRecieveStock.Text1.Text = lvwStock.SelectedItem.SubItems(1)
frmRecieveStock.dcbVendor.Text = lvwStock.SelectedItem.SubItems(2)
frmRecieveStock.Text8.Text = lvwStock.SelectedItem.SubItems(3)
frmRecieveStock.Text2.Text = lvwStock.SelectedItem.SubItems(4)
frmRecieveStock.Text3.Text = lvwStock.SelectedItem.SubItems(5)
frmRecieveStock.Text4.Text = lvwStock.SelectedItem.SubItems(6)
frmRecieveStock.Text5.Text = lvwStock.SelectedItem.SubItems(7)
frmRecieveStock.Text6.Text = lvwStock.SelectedItem.SubItems(8)
frmRecieveStock.Text10.Text = lvwStock.SelectedItem.Text

frmRecieveStock.cmdDelete.Enabled = True

Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmRecieveStock.Visible = True Then
frmRecieveStock.Show
Else
frmTree.TvwCustomer.SetFocus
End If
End Sub

Private Sub lvwStock_DblClick()
Call cmdLoadStock_Click
End Sub

Private Sub lvwStock_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'-------------------------------------------------------------------------
    
    ' sort the listview on the column clicked
    With lvwStock
        If (.Sorted) And (ColumnHeader.SubItemIndex = .SortKey) Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .Sorted = True
            .SortKey = ColumnHeader.SubItemIndex
            .SortOrder = lvwAscending
        End If
        .Refresh
    End With
        
    ' If an item was selected prior to the sort,
    ' make sure it is still visible now that the sort is done.
    If Not lvwStock.SelectedItem Is Nothing Then
        lvwStock.SelectedItem.EnsureVisible
    End If

End Sub
