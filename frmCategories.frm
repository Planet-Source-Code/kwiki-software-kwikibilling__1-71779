VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCategories 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "     Product Categories"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7650
   Icon            =   "frmCategories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   6120
      TabIndex        =   29
      Top             =   3840
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
      Image           =   "frmCategories.frx":000C
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   3
      Left            =   3960
      Picture         =   "frmCategories.frx":0A06
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   2
      Left            =   3360
      Picture         =   "frmCategories.frx":0D48
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   1
      Left            =   2760
      Picture         =   "frmCategories.frx":108A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   0
      Left            =   2160
      Picture         =   "frmCategories.frx":13CC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   644
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Part Categories"
      TabPicture(0)   =   "frmCategories.frx":170E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "View Categories"
      TabPicture(1)   =   "frmCategories.frx":172A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   6975
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
            Height          =   315
            Left            =   4800
            TabIndex        =   15
            Top             =   2520
            Width           =   1695
         End
         Begin MSComctlLib.ListView LvwCat 
            Height          =   2175
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   3836
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label4 
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label3 
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   2640
            Width           =   2655
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
            Left            =   3960
            TabIndex        =   18
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label2 
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   2640
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   6975
         Begin VB.TextBox txtFields 
            BackColor       =   &H80000018&
            DataField       =   "CategoryDecsription"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Index           =   2
            Left            =   1920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H80000018&
            DataField       =   "CategoryName"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            DataField       =   "CategoryID"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   285
            Index           =   0
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame Frame2 
            Height          =   735
            Left            =   120
            TabIndex        =   7
            Top             =   1800
            Width           =   6735
            Begin lvButton.lvButtons_H cmdAdd 
               Height          =   375
               Left            =   5520
               TabIndex        =   28
               Top             =   240
               Width           =   1095
               _ExtentX        =   1931
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
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdUpdate 
               Height          =   375
               Left            =   4440
               TabIndex        =   27
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
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
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdCancel 
               Height          =   375
               Left            =   3360
               TabIndex        =   26
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
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
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdDelete 
               Height          =   375
               Left            =   2280
               TabIndex        =   25
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
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
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdEdit 
               Height          =   375
               Left            =   1200
               TabIndex        =   23
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
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
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdRefresh 
               Height          =   375
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Caption         =   "&Requery"
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
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdSave 
               Height          =   375
               Left            =   1200
               TabIndex        =   24
               Top             =   240
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
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
               cBack           =   -2147483633
            End
         End
         Begin VB.Label Label5 
            Caption         =   "(Not Required)"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblStatus 
            Height          =   615
            Left            =   1680
            TabIndex        =   8
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label lblLabels 
            Caption         =   "Decsription:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Category Name:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCats As Recordset
Dim sqlCats As String
Dim Clear As String

Private Sub cmdCancel_Click()
On Error GoTo CancelErr
cmdEdit.Visible = True
cmdSave.Visible = False
cmdAdd.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdRefresh.Enabled = True
cmdClose.Enabled = True
cmdBrowse(0).Enabled = True
cmdBrowse(1).Enabled = True
cmdBrowse(2).Enabled = True
cmdBrowse(3).Enabled = True

FillCatFields
DisableFields

If txtFields(0).Text = "" Then
cmdEdit.Enabled = False
cmdDelete.Enabled = False
Else
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End If
'LvwCat.Enabled = True
Exit Sub
CancelErr:
MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
On Error GoTo EditErr
cmdEdit.Visible = False
cmdSave.Visible = True
cmdAdd.Enabled = False
cmdUpdate.Enabled = False
cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdRefresh.Enabled = False
cmdClose.Enabled = False
cmdBrowse(0).Enabled = False
cmdBrowse(1).Enabled = False
cmdBrowse(2).Enabled = False
cmdBrowse(3).Enabled = False
'LvwCat.Enabled = False
EnableFields

If rsCats.RecordCount > 0 Then
rsCats.MoveFirst
End If

Exit Sub
EditErr:
End Sub

Private Sub cmdSave_Click()
On Error GoTo EH:
With rsCats
.Edit
rsCats!CategoryName = txtFields(1).Text
rsCats!CategoryDecsription = txtFields(2).Text
.Update
.MoveLast
End With

cmdEdit.Visible = True
cmdSave.Visible = False
cmdAdd.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdRefresh.Enabled = True
cmdClose.Enabled = True
cmdBrowse(0).Enabled = True
cmdBrowse(1).Enabled = True
cmdBrowse(2).Enabled = True
cmdBrowse(3).Enabled = True
DisableFields
rsCats.Requery
LoadCatList
'LvwCat.Enabled = True
Label7.Caption = Clear
Exit Sub
EH:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
OpenDatabase
If (Not OpenDatabase()) Then
  MsgBox "Database could not be openend !"
End If

sqlCats = "Select * From Categories"
Set rsCats = DB.OpenRecordset(sqlCats)

If rsCats.RecordCount > 0 Then
rsCats.MoveFirst
End If

FillCatFields
DisableFields
setUpListView
LoadCatList

If txtFields(0).Text = "" Then
cmdEdit.Enabled = False
cmdDelete.Enabled = False
Else
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End If

End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        rsCats.MoveFirst
        Beep
        FillCatFields        'call subprocedure
        
    Case 1
        rsCats.MovePrevious
        If rsCats.BOF Then
        rsCats.MoveFirst
        End If
        FillCatFields
        
    Case 2
        rsCats.MoveNext
        If rsCats.EOF Then
        rsCats.MoveLast
        End If
        FillCatFields
        
    Case 3
        rsCats.MoveLast
        Beep
        FillCatFields
        
End Select
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddErr
DisableCont
EnableFields
ClearCatFields
Exit Sub
AddErr:
MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
If txtFields(0) = "" Then
MsgBox ("No records to remove")
Exit Sub
Else

If MsgBox("Are you sure you want to Remove ? " & txtFields(1).Text, vbYesNo, "Confirm") = vbYes Then
ClearCatFields
With rsCats
If .EOF Then .MoveLast
.Delete
.MoveNext
End With
rsCats.Requery
LoadCatList
End If

End If
Exit Sub
DeleteErr:
MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo EH
rsCats.Requery
FillCatFields
LoadCatList
Label7.Caption = Clear
EnableCont
Exit Sub
EH:
MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo UpdateErr
EnableCont
With rsCats
.AddNew
rsCats!CategoryName = txtFields(1).Text
rsCats!CategoryDecsription = txtFields(2).Text
.Update
.MoveLast
End With
FillCatFields
DisableFields

rsCats.Requery
LoadCatList
Label7.Caption = Clear
Exit Sub

UpdateErr:
MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
SndClick
frmParts.Adodc1.Refresh

If frmParts.Visible = True Then
Unload Me
frmParts.Show
End If

Set rsCats = Nothing
Unload Me

End Sub

Private Sub DisableCont()
cmdAdd.Enabled = False
cmdUpdate.Enabled = True
cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdRefresh.Enabled = False
cmdClose.Enabled = False
cmdBrowse(0).Enabled = False
cmdBrowse(1).Enabled = False
cmdBrowse(2).Enabled = False
cmdBrowse(3).Enabled = False
End Sub

Private Sub EnableCont()
cmdAdd.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdRefresh.Enabled = True
cmdClose.Enabled = True
cmdBrowse(0).Enabled = True
cmdBrowse(1).Enabled = True
cmdBrowse(2).Enabled = True
cmdBrowse(3).Enabled = True
End Sub

Private Sub DisableFields()
txtFields(0).Locked = True
txtFields(1).Locked = True
txtFields(2).Locked = True
End Sub

Private Sub EnableFields()
txtFields(0).Locked = True
txtFields(1).Locked = False
txtFields(2).Locked = False
End Sub

Private Sub FillCatFields()
On Error GoTo EH
txtFields(0).Text = rsCats!CategoryID
txtFields(1).Text = rsCats!CategoryName
txtFields(2).Text = rsCats!CategoryDecsription
Exit Sub
EH:
End Sub

Private Sub ClearCatFields()
txtFields(0).Text = Clear
txtFields(1).Text = Clear
txtFields(2).Text = Clear
End Sub

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwCat.ColumnHeaders. _
Add(, , "CategoryID", 0, lvwColumnLeft)
Set clmHdr = LvwCat.ColumnHeaders. _
Add(, , "Categories", 2400, lvwColumnLeft)
Set clmHdr = LvwCat.ColumnHeaders. _
Add(, , "Decsription", 3800, lvwColumnLeft)
LvwCat.View = lvwReport
End Sub

Public Sub LoadCatList()
On Error Resume Next
Dim sqlItem As ListItem

LvwCat.ListItems.Clear

While Not rsCats.EOF
Set sqlItem = LvwCat.ListItems.Add(, , _
rsCats!CategoryID)
sqlItem.SubItems(1) = rsCats!CategoryName
sqlItem.SubItems(2) = rsCats!CategoryDecsription
rsCats.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsCats = Nothing
frmTree.TvwCustomer.SetFocus
End Sub

Private Sub LvwCat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the listview on the column clicked
    With LvwCat
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
    If Not LvwCat.SelectedItem Is Nothing Then
        LvwCat.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SearchList
End If
End Sub

Private Sub SearchList()
On Error Resume Next
Dim itm As ListItem
With LvwCat
Set itm = .FindItem(Text1.Text, lvwSubItem, lvwPartial)
Label4.Caption = "Searched Record Not Found"
If Not itm Is Nothing Then
Label4.Caption = "Searched Record Found"
.ListItems(itm.Index).Selected = True
.SetFocus
End If
End With
Set itm = Nothing
End Sub
