VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmWorkorderLabor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Labor For Order "
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWorkorderLabor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H Close 
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
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
      Image           =   "frmWorkorderLabor.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H AddLabor 
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   4320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   " &Add Labor To Order"
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
      Image           =   "frmWorkorderLabor.frx":0A06
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H RemoveLabor 
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "&Remove Labor"
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
      Image           =   "frmWorkorderLabor.frx":12E0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H ShowEmployees 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "&Load Employees"
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
      Image           =   "frmWorkorderLabor.frx":1BBA
      cBack           =   -2147483633
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.TextBox WorkorderID 
         DataField       =   "WorkorderID"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox WorkorderLaborID 
         DataField       =   "WorkorderLaborID"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox EmployeeID 
         DataField       =   "EmployeeID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox BillableHours 
         BackColor       =   &H80000018&
         DataField       =   "BillableHours"
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
         Left            =   1440
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "You can load the employee list by double clicking this box"
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox BillingRate 
         BackColor       =   &H80000018&
         DataField       =   "BillingRate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox Text2 
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
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox Comment 
         BackColor       =   &H80000018&
         DataField       =   "Comment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   960
         Width           =   4095
      End
      Begin MSMask.MaskEdBox Text7 
         Height          =   405
         Left            =   6960
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "BillableHours:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SalaryRate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3000
         TabIndex        =   12
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label12 
         Caption         =   "Amount To Be Added To Order"
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
         Left            =   5760
         TabIndex        =   11
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Employee:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   930
      End
   End
   Begin MSComctlLib.ListView LvwLabor 
      Height          =   1455
      Left            =   720
      TabIndex        =   14
      Top             =   2160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   8535
      Begin lvButton.lvButtons_H cmdUpdate 
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Update Hours"
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
         Image           =   "frmWorkorderLabor.frx":2494
         cBack           =   -2147483633
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000FF&
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
         Left            =   2160
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Viewing Order >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   18
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label10 
         Caption         =   "Total Labor"
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
         Left            =   4800
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmWorkorderLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLabor As Recordset
Dim sqlLabor As String
Dim rsViewLabor As Recordset
Dim sqlViewLabor As String

Private Sub cmdUpdate_Click()
On Error Resume Next
DB.Execute "UPDATE [Workorder Labor] SET BillableHours=" & BillableHours.Text & " WHERE WorkorderLaborID=" & WorkorderLaborID.Text & " "
ViewLabor

frmTree.FillListView
frmTree.GetTax
frmTree.FillInfo

AddTotal
cmdUpdate.Visible = False
Exit Sub
End Sub

Private Sub Form_Activate()
BillableHours.SetFocus
Label1.Caption = " Customer > " & frmTree.lblName.Caption
End Sub

Private Sub Form_Load()
sqlLabor = "Select * From [Workorder Labor]"
Set rsLabor = DB.OpenRecordset(sqlLabor)

GetLaborWorkorderID
setUpLaborListView
ClearLaborFields
ViewLabor
AddTotal
End Sub

'--------------------------Labor Section-------------'

Public Sub setUpLaborListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwLabor.ColumnHeaders. _
Add(, , "WOLID", 0, lvwColumnLeft)
Set clmHdr = LvwLabor.ColumnHeaders. _
Add(, , "WOID", 0, lvwColumnLeft)
Set clmHdr = LvwLabor.ColumnHeaders. _
Add(, , "EmpID", 0, lvwColumnLeft)
Set clmHdr = LvwLabor.ColumnHeaders. _
Add(, , "Employee Name", 2200, lvwColumnLeft)
Set clmHdr = LvwLabor.ColumnHeaders. _
Add(, , "Hours", 1600, lvwColumnLeft)
Set clmHdr = LvwLabor.ColumnHeaders. _
Add(, , "SalaryRate", 1600, lvwColumnLeft)
Set clmHdr = LvwLabor.ColumnHeaders. _
Add(, , "Total", 1400, lvwColumnLeft)

LvwLabor.View = lvwReport
End Sub

Public Sub AddLabor_Click()
On Error Resume Next

GetLaborWorkorderID

If WorkorderID.Text = "(Null)" Then
MsgBox ("You must select an order for the customer")
Exit Sub
End If

With rsLabor
    .AddNew
    
    ![WorkorderID] = WorkorderID.Text
    ![EmployeeID] = EmployeeID.Text
    ![BillingRate] = BillingRate.Text
    ![BillableHours] = BillableHours.Text
    ![Comment] = Comment.Text
    .Update
    .MoveLast
     End With
    
    ViewLabor
    AddTotal
    ClearLaborFields
    
    frmTree.FillListView
    frmTree.FillInfo
    frmTree.GetTax

End Sub

Private Sub LvwLabor_Click()
On Error Resume Next
If LvwLabor.ListItems.count = 0 Then
ClearLaborFields
Else
GetLaborWorkorderID
cmdUpdate.Visible = True

sqlViewLabor = "SELECT [Workorder Labor].WorkorderLaborID, [Workorder Labor].WorkorderID, [Workorder Labor].EmployeeID, Employees.FullName, [Workorder Labor].Comment, [Workorder Labor].BillableHours, [Workorder Labor].BillingRate, [Workorder Labor].BillableHours*[Workorder Labor].BillingRate AS Total "
sqlViewLabor = sqlViewLabor & "FROM Employees INNER JOIN [Workorder Labor] ON Employees.EmployeeID = [Workorder Labor].EmployeeID "
sqlViewLabor = sqlViewLabor & "WHERE [Workorder Labor].WorkorderLaborID =" & LvwLabor.SelectedItem.Text & " "
Set rsViewLabor = DB.OpenRecordset(sqlViewLabor)

WorkorderLaborID.Text = rsViewLabor!WorkorderLaborID
WorkorderID.Text = rsViewLabor!WorkorderID
EmployeeID.Text = rsViewLabor!EmployeeID
Text2.Text = rsViewLabor!FullName
Comment.Text = rsViewLabor!Comment
BillableHours.Text = rsViewLabor!BillableHours
BillingRate.Text = Format$(rsViewLabor!BillingRate, "$#,##0.00;(#,##0.00)")
Text7.Text = Format$(rsViewLabor!Total, "$#,##0.00;(#,##0.00)")
End If
End Sub

Private Sub LvwLabor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the listview on the column clicked
    With LvwLabor
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
    If Not LvwLabor.SelectedItem Is Nothing Then
        LvwLabor.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub RemoveLabor_Click()
On Error Resume Next
If LvwLabor.ListItems.count = 0 Then
MsgBox "There's no Labor to remove"
Exit Sub
ElseIf MsgBox("Are you sure you want to Remove " & LvwLabor.SelectedItem.SubItems(3) & " from labor ?", vbYesNo, "Confirm") = vbYes Then
DB.Execute "DELETE FROM [Workorder Labor] WHERE WorkorderLaborID=" & LvwLabor.SelectedItem.Text & " "
LvwLabor.ListItems.Remove LvwLabor.SelectedItem.Index
MinusTotal
ClearLaborFields

frmTree.FillListView
frmTree.FillInfo
frmTree.GetTax

ElseIf vbNo Then
MsgBox "No Record Deleted"
Exit Sub
End If
End Sub

Private Sub ShowEmployees_Click()
ClearLaborFields
GetLaborWorkorderID
frmAddLabor.Show
cmdUpdate.Visible = False
End Sub

Private Sub BillableHours_Change()
On Error Resume Next
Text7 = BillableHours.Text * BillingRate.Text
End Sub

Private Sub BillableHours_DblClick()
ShowEmployees_Click
End Sub

Private Sub BillableHours_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
AddLabor_Click
End If
End Sub

Private Sub AddTotal()
On Error Resume Next
Dim i As Integer
Dim cTotal As Currency
With LvwLabor
For i = 1 To .ListItems.count
cTotal = cTotal + CCur(.ListItems(i).SubItems(6))
Next
End With
Text4.Text = Format$(cTotal, "$#,##0.00;($#,##0.00)")
End Sub

Private Sub MinusTotal()
On Error Resume Next
Dim i As Integer
Dim cTotal As Currency
With LvwLabor
For i = 1 To .ListItems.count
cTotal = cTotal - CCur(.ListItems(i).SubItems(6))
Next
End With
Text4.Text = Format$(cTotal, "($#,##0.00);$#,##0.00")
End Sub

Private Sub ClearLaborFields()
BillableHours = ""
BillingRate.Text = ""
EmployeeID.Text = ""
WorkorderID.Text = frmTree.txtWorkorderID.Text
WorkorderLaborID.Text = ""
Comment.Text = " "
Text2.Text = ""
Text7.Text = ""
End Sub

Private Sub GetLaborWorkorderID()
If frmWorkorders.Visible = True Then
WorkorderID.Text = " " & frmWorkorders.Text2.Text
Label7.Caption = " " & frmWorkorders.Text2.Text
Else
WorkorderID.Text = frmTree.txtWorkorderID.Text
Label7.Caption = " " & WorkorderID.Text
End If
End Sub

Private Sub ViewLabor()
On Error GoTo EH:
Dim itmx As ListItem
LvwLabor.ListItems.Clear

sqlViewLabor = "SELECT [Workorder Labor].WorkorderLaborID, [Workorder Labor].WorkorderID, [Workorder Labor].EmployeeID, Employees.FullName, [Workorder Labor].BillableHours, [Workorder Labor].BillingRate, [Workorder Labor].BillableHours*[Workorder Labor].BillingRate AS Total "
sqlViewLabor = sqlViewLabor & "FROM Employees INNER JOIN [Workorder Labor] ON Employees.EmployeeID = [Workorder Labor].EmployeeID "
sqlViewLabor = sqlViewLabor & "WHERE [Workorder Labor].WorkorderID =" & Label7.Caption & " "
Set rsViewLabor = DB.OpenRecordset(sqlViewLabor)

While Not rsViewLabor.EOF
Set itmx = LvwLabor.ListItems.Add(, , _
rsViewLabor!WorkorderLaborID)
itmx.SubItems(1) = rsViewLabor!WorkorderID
itmx.SubItems(2) = rsViewLabor!EmployeeID
itmx.SubItems(3) = rsViewLabor!FullName
itmx.SubItems(4) = rsViewLabor!BillableHours
itmx.SubItems(5) = Format$(rsViewLabor!BillingRate, "$#,##0.00;(#,##0.00)")
itmx.SubItems(6) = Format$(rsViewLabor!Total, "$#,##0.00;(#,##0.00)")
rsViewLabor.MoveNext
Wend
DoEvents
Exit Sub
EH: MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmTree.FillListView
frmTree.TvwCustomer.SetFocus
End Sub

Private Sub Close_Click()
SndClick
Set rsViewLabor = Nothing
Set rsLabor = Nothing
Unload Me
End Sub

