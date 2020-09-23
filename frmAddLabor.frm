VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddLabor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Viewing Employees"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8055
   Icon            =   "frmAddLabor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8055
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
      Left            =   1080
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   6720
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
      Image           =   "frmAddLabor.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H AddLabor 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Add Labor"
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
      Image           =   "frmAddLabor.frx":0A06
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView LvwAddEmployees 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
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
   Begin VB.Label Label3 
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
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
      Width           =   975
   End
End
Attribute VB_Name = "frmAddLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmployees As Recordset
Dim sqlEmployee As String

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "EmpID", 0, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "Employee Name", 2000, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "Contact Phone", 1800, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "Salary Rate", 1500, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "SSN Number", 1500, lvwColumnLeft)

LvwAddEmployees.View = lvwReport
End Sub

Private Sub cmdClose_Click()
SndClick
Set rsEmployees = Nothing
Unload Me
End Sub

Private Sub Form_Activate()
LoadEmpList
End Sub

Private Sub Form_Load()
If (Not OpenDatabase()) Then
MsgBox "Could not load database"
End If
setUpListView
End Sub

Public Sub LoadEmpList()
Dim sqlEmp As ListItem
LvwAddEmployees.ListItems.Clear

sqlEmployee = "Select Employees.EmployeeID, Employees.FullName, Employees.ContactPhone, Employees.BillingRate, Employees.SSNNumber "
sqlEmployee = sqlEmployee & " From Employees"
Set rsEmployees = DB.OpenRecordset(sqlEmployee)

If (rsEmployees.RecordCount > 0) Then
rsEmployees.MoveFirst
End If
On Error Resume Next
While Not rsEmployees.EOF
Set sqlEmp = LvwAddEmployees.ListItems.Add(, , _
rsEmployees!EmployeeID)
sqlEmp.SubItems(1) = rsEmployees!FullName
sqlEmp.SubItems(2) = rsEmployees!ContactPhone
sqlEmp.SubItems(3) = Format$(rsEmployees!BillingRate, "$#,##0.00;(#,##0.00)")
sqlEmp.SubItems(4) = Format$(rsEmployees!SSNNumber, "###-##-####;(###-##-####)")

rsEmployees.MoveNext
Wend
Set rsEmployees = Nothing
End Sub

Private Sub LvwAddEmployees_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the listview on the column clicked
    With LvwAddEmployees
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
    If Not LvwAddEmployees.SelectedItem Is Nothing Then
        LvwAddEmployees.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub LvwAddEmployees_DblClick()
On Error Resume Next
AddLabor_Click
End Sub

Private Sub AddLabor_Click()
On Error Resume Next
frmWorkorderLabor.EmployeeID.Text = LvwAddEmployees.SelectedItem.Text
frmWorkorderLabor.Text2.Text = LvwAddEmployees.SelectedItem.SubItems(1)
frmWorkorderLabor.BillingRate.Text = Format$(LvwAddEmployees.SelectedItem.SubItems(3), "$#,##0.00;(#,##0.00)")
frmWorkorderLabor.BillableHours.Text = ""
frmWorkorderLabor.BillableHours.SetFocus
Unload Me
End Sub

Private Sub SearchList()
On Error Resume Next
Dim itm As ListItem

With LvwAddEmployees
Set itm = .FindItem(Text1.Text, lvwSubItem, lvwPartial)
Label3.BackColor = vbRed
Label3.ForeColor = vbWhite
Label3.Caption = "Searched Record Not Found"
If Not itm Is Nothing Then
Label3.BackColor = vbRed
Label3.ForeColor = vbWhite
Label3.Caption = "Searched Record Found"
.ListItems(itm.Index).Selected = True
.SetFocus
End If
End With

Set itm = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SearchList
End If
End Sub
