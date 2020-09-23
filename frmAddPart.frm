VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddPart 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Viewing Products"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8190
   Icon            =   "frmAddPart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin lvButton.lvButtons_H Close 
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
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
      Image           =   "frmAddPart.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H AddPart 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Add Part"
      CapAlign        =   2
      BackStyle       =   2
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
      Image           =   "frmAddPart.frx":0A06
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView LvwAddPart 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16056314
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "frmAddPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAddParts As Recordset
Dim sqlList As String

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwAddPart.ColumnHeaders. _
Add(, , "ID", 0, lvwColumnLeft)
Set clmHdr = LvwAddPart.ColumnHeaders. _
Add(, , "PartName", 2200, lvwColumnLeft)
Set clmHdr = LvwAddPart.ColumnHeaders. _
Add(, , "Description", 2800, lvwColumnLeft)
Set clmHdr = LvwAddPart.ColumnHeaders. _
Add(, , "Unit Price", 1400, lvwColumnLeft)
Set clmHdr = LvwAddPart.ColumnHeaders. _
Add(, , "Units Stock", 1400, lvwColumnLeft)

LvwAddPart.View = lvwReport
End Sub

Public Sub LoadPartList()
Dim sqlItem As ListItem

LvwAddPart.ListItems.Clear
sqlList = "SELECT PartID, PartName, PartDescription, UnitPrice, UnitsStock "
sqlList = sqlList & "FROM Parts"
Set rsAddParts = DB.OpenRecordset(sqlList)

If (rsAddParts.RecordCount > 0) Then
rsAddParts.MoveFirst
End If

While Not rsAddParts.EOF
Set sqlItem = LvwAddPart.ListItems.Add(, , _
rsAddParts!PartID)
sqlItem.SubItems(1) = rsAddParts!PartName
sqlItem.SubItems(2) = rsAddParts!PartDescription
sqlItem.SubItems(3) = Format$(rsAddParts!UnitPrice, "$#,##0.00;(#,##0.00)")
sqlItem.SubItems(4) = rsAddParts!UnitsStock
rsAddParts.MoveNext
Wend
End Sub

Private Sub AddPart_Click()
On Error Resume Next
If LvwAddPart.SelectedItem.SubItems(4) = 0 Then
MsgBox ("Part is not available, has no stock")
Exit Sub
End If
frmWorkorderParts.Text2.Text = LvwAddPart.SelectedItem.SubItems(1)
frmWorkorderParts.Text3.Text = LvwAddPart.SelectedItem.SubItems(2)
frmWorkorderParts.Text12.Text = LvwAddPart.SelectedItem.SubItems(3)
frmWorkorderParts.Text10.Text = LvwAddPart.SelectedItem.Text
frmWorkorderParts.Text11.Text = ""
frmWorkorderParts.Text11.Locked = False
frmWorkorderParts.txtTotal.Text = ""
frmWorkorderParts.Text11.SetFocus
frmWorkorderParts.CStock
frmWorkorderParts.cmdDelete.Enabled = False
frmWorkorderParts.cmdClose.Visible = False
frmWorkorderParts.cmdCancel.Visible = True
Unload Me
End Sub

Private Sub SearchList()
On Error Resume Next
Dim itm As ListItem
With LvwAddPart
Set itm = .FindItem(Text1.Text, lvwSubItem, lvwPartial)
Label2.BackColor = vbRed
Label2.ForeColor = vbWhite
Label2.Caption = "Searched Record Not Found"
If Not itm Is Nothing Then
Label2.BackColor = vbRed
Label2.ForeColor = vbWhite
Label2.Caption = "Searched Record Found"
.ListItems(itm.Index).Selected = True
.SetFocus
End If
End With
Set itm = Nothing
End Sub

Private Sub Close_Click()
SndClick
Set rsAddParts = Nothing
Unload Me
End Sub

Private Sub Form_Activate()
LoadPartList
End Sub

Private Sub Form_Load()
setUpListView
End Sub

Private Sub LvwAddPart_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the listview on the column clicked
    With LvwAddPart
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
    If Not LvwAddPart.SelectedItem Is Nothing Then
        LvwAddPart.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub LvwAddPart_DblClick()
On Error Resume Next
AddPart_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SearchList
End If
End Sub
