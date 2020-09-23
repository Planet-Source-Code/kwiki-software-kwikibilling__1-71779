VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmWorkorderParts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Items For Order"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWorkorderParts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdShowParts 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "&Load Part List"
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
      Image           =   "frmWorkorderParts.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   " &Remove Part From Order"
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
      Image           =   "frmWorkorderParts.frx":045E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAddPart 
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      Top             =   4320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   " &Add Part To Order"
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
      Image           =   "frmWorkorderParts.frx":0D38
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   8040
      TabIndex        =   24
      Top             =   4320
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
      Image           =   "frmWorkorderParts.frx":1612
      cBack           =   -2147483633
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
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   9255
      Begin lvButton.lvButtons_H cmdEditStock 
         Height          =   375
         Left            =   2880
         TabIndex        =   29
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Caption         =   "&Edit Quantity"
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
         Image           =   "frmWorkorderParts.frx":200C
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
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
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2400
         Width           =   1695
      End
      Begin MSComctlLib.ListView LvwParts 
         Height          =   1815
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin lvButton.lvButtons_H cmdUpdate 
         Height          =   375
         Left            =   2880
         TabIndex        =   30
         Top             =   2400
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Caption         =   "&Update Quantity"
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
         Image           =   "frmWorkorderParts.frx":28E6
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Viewing Order > "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         DataField       =   "CompanyName"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Total:"
         Height          =   255
         Left            =   6720
         TabIndex        =   14
         Top             =   2400
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtTotal 
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
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   1215
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
         Height          =   360
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         DataField       =   "WorkorderPartID"
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
         Left            =   8040
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text8 
         DataField       =   "WorkorderID"
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
         Left            =   6720
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text10 
         DataField       =   "PartID"
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
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H80000018&
         DataField       =   "UnitPrice"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H80000018&
         DataField       =   "Quantity"
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "You can load the part list by double clicking this box"
         Top             =   720
         Width           =   1095
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Amount To Be Added To Order >"
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label18 
         Caption         =   "UnitPrice:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblpartName 
         Caption         =   "PartName:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPartDescription 
         Caption         =   "Description:"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   8040
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
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
      Image           =   "frmWorkorderParts.frx":31C0
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmWorkorderParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParts As Recordset
Dim rsViewParts As Recordset
Dim sqlParts As String
Dim sqlViewParts As String
Dim sqlName As String
Dim rsName As Recordset
Dim itmx As ListItem
Private Mode As String

Private Sub cmdCancel_Click()
cmdClose.Visible = True
cmdCancel.Visible = False
cmdEditStock.Visible = True
cmdEditStock.Enabled = False
cmdUpdate.Visible = False
cmdDelete.Enabled = True
cmdAddPart.Enabled = True
cmdShowParts.Enabled = True
cmdClose.Enabled = True
Text11.Locked = True
ClearPartFields

End Sub

Private Sub cmdEditStock_Click()
cmdEditStock.Visible = False
cmdUpdate.Visible = True
cmdDelete.Enabled = False
cmdAddPart.Enabled = False
cmdShowParts.Enabled = False
cmdClose.Visible = False
'cmdCancel.Visible = True
Text11.Locked = False
Text11.SetFocus
ReStock
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next

If Text11.Text = Text7.Text + 1 Then
MsgBox ("Incorrect Value, Or There is not enough stock for your entered quantity")
Exit Sub
Else
Mode = "UPDATE STOCK"
End If

If Text11.Text = "" Then
MsgBox ("Quantity cannot be left empty")
Exit Sub
Else
Mode = "UPDATE STOCK"
End If

If Text11.Text = 0 Then
MsgBox ("Please enter a value above 0 for quantity")
Exit Sub
Else
Mode = "UPDATE STOCK"
End If

If Mode = "UPDATE STOCK" Then
DB.Execute "UPDATE [Workorder Parts] SET Quantity=" & Text11.Text & " WHERE WorkorderPartID=" & Text9.Text & " "
ViewParts
frmTree.FillListView
AddTotal

        If frmTree.LvwOrders.SelectedItem.SubItems(7) = "" Then
frmTree.Label7.Caption = " " & "(0.00)"
Else
frmTree.Label7.Caption = " " & frmTree.LvwOrders.SelectedItem.SubItems(7)
End If
        If frmTree.LvwOrders.SelectedItem.SubItems(8) = "" Then
frmTree.Label8.Caption = " " & "(0.00)"
Else
frmTree.Label8.Caption = " " & frmTree.LvwOrders.SelectedItem.SubItems(8)
End If

cmdUpdate.Visible = False
frmTree.GetTax
UpdateStock

'ENABLE CONTROLS
cmdEditStock.Visible = True
cmdEditStock.Enabled = False
cmdUpdate.Visible = False
cmdDelete.Enabled = True
cmdAddPart.Enabled = True
cmdShowParts.Enabled = True
cmdClose.Visible = True
cmdCancel.Visible = False
Text11.Locked = True

End If
Exit Sub
End Sub

Private Sub Form_Load()
sqlParts = "Select * FROM [Workorder Parts]"
Set rsParts = DB.OpenRecordset(sqlParts)

SetUpPartList
ClearPartFields
GetPartsWorkorderID
ViewParts
AddTotal

End Sub

Private Sub Form_Activate()
Text11.SetFocus
Label4.Caption = " Customer > " & frmTree.lblName.Caption
End Sub

' Initialize the SetUpPartList control
Public Sub SetUpPartList()
Dim clmHdr As ColumnHeader
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "WPID", 0, lvwColumnLeft)
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "WOID", 0, lvwColumnLeft)
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "PID", 0, lvwColumnLeft)
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "Item Name", 2200, lvwColumnLeft)
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "Description", 2600, lvwColumnLeft)
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "Quantity", 1200, lvwColumnLeft)
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "UnitPrice", 1400, lvwColumnLeft)
Set clmHdr = lvwParts.ColumnHeaders. _
Add(, , "Total", 1300, lvwColumnLeft)
lvwParts.View = lvwReport
Exit Sub
End Sub

Private Sub AddTotal()
On Error Resume Next
Dim i As Integer
Dim cTotal As Currency
With lvwParts
For i = 1 To .ListItems.count
cTotal = cTotal + CCur(.ListItems(i).SubItems(7))
Next
End With
Text1.Text = Format$(cTotal, "$#,##0.00;($#,##0.00)")
End Sub

Private Sub MinusTotal()
On Error Resume Next
Dim i As Integer
Dim cTotal As Currency
With lvwParts
For i = 1 To .ListItems.count
cTotal = cTotal - CCur(.ListItems(i).SubItems(7))
Next
End With
Text1.Text = Format$(cTotal, "($#,##0.00);$#,##0.00")
End Sub

Private Sub ClearPartFields()
Text11.Text = ""
Text12.Text = ""
Text3.Text = ""
Text2.Text = ""
txtTotal.Text = ""
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
If lvwParts.ListItems.count = 0 Then
MsgBox "There's no parts to remove"
Exit Sub
ElseIf Text6.Text = "" Or Text11.Text = "" Then
MsgBox "You must select a part to delete"
Exit Sub
ElseIf MsgBox("Are you sure you want to Remove ? " & lvwParts.SelectedItem.SubItems(3), vbYesNo, "Confirm") = vbYes Then
DB.Execute "DELETE FROM [Workorder Parts] WHERE WorkorderPartID=" & lvwParts.SelectedItem.Text & " "
lvwParts.ListItems.Remove lvwParts.SelectedItem.Index

MinusTotal
'DelStatus
frmTree.FillListView

RMVPart
frmTree.FillInfo
frmTree.GetTax

ElseIf vbNo Then
MsgBox "No Record Deleted"
Exit Sub
End If

ClearPartFields

cmdUpdate.Visible = False
If lvwParts.ListItems.count = 0 Then
cmdEditStock.Enabled = False
Else
cmdEditStock.Enabled = True
End If

End Sub

Private Sub cmdClose_Click()
SndClick
If frmTree.LvwOrders.ListItems.count = 0 Then
frmWorkorderLabor.AddLabor_Click  'ADD Default Labor
End If
'If MsgBox("Would you like to apply a payment now ? ", vbYesNo, "Confirm") = vbYes Then
'frmPayments.Show
'End If
Set rsViewParts = Nothing
Set rsParts = Nothing
Set rsName = Nothing
Unload Me
End Sub

Private Sub LvwParts_Click()
On Error Resume Next
If lvwParts.ListItems.count = 0 Then
ClearPartFields
Else

cmdEditStock.Enabled = True


sqlViewParts = "Select [Workorder Parts].WorkorderPartID, [Workorder Parts].WorkorderID, [Workorder Parts].PartID, Parts.PartName, Parts.PartDescription, [Workorder Parts].Quantity, [Workorder Parts].UnitPrice, [Quantity]*[Workorder Parts].[UnitPrice] AS Total "
sqlViewParts = sqlViewParts & "FROM Parts INNER JOIN [Workorder Parts] ON Parts.PartID = [Workorder Parts].PartID "
sqlViewParts = sqlViewParts & "WHERE [Workorder Parts].WorkorderPartID =" & lvwParts.SelectedItem.Text & " "
Set rsViewParts = DB.OpenRecordset(sqlViewParts)

Text9.Text = rsViewParts!WorkorderPartID
Text8.Text = rsViewParts!WorkorderID
Text10.Text = rsViewParts!PartID
Text2.Text = rsViewParts!PartName
Text3.Text = rsViewParts!PartDescription
Text11.Text = rsViewParts!Quantity
Text12.Text = Format$(rsViewParts!UnitPrice, "$#,##0.00;(#,##0.00)")
txtTotal.Text = Format$(rsViewParts!Total, "$#,##0.00;(#,##0.00)")

SPStock 'See Previos Stock
End If
End Sub

Private Sub lvwParts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the listview on the column clicked
    With lvwParts
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
    If Not lvwParts.SelectedItem Is Nothing Then
        lvwParts.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub cmdShowParts_Click()
ClearPartFields
cmdUpdate.Visible = False
frmAddPart.Show
End Sub

Private Sub Text11_Change()
On Error Resume Next
txtTotal.Text = Format$(Text11.Text * Text12.Text, "$#,##0.00;(#,##0.00)")
End Sub

Private Sub cmdAddPart_Click()
On Error Resume Next
GetPartsWorkorderID

If Text8.Text = "(Null)" Then
MsgBox ("You must select an order first")
Exit Sub
End If

If Text11.Text = "" Then
MsgBox ("Quantity cannot be left empty")
Exit Sub
ElseIf Text11.Text = 0 Then
MsgBox ("Please enter a value above 0 for quantity")
Exit Sub
Else
Mode = "ADDPART"
End If

'LETS CHECK USER HAS NOT ENTERED THE SAME PART TWICE
If lvwParts.SelectedItem.SubItems(2) = "" Then
Mode = "ADDPART"

ElseIf Text10.Text = lvwParts.SelectedItem.SubItems(2) Then
MsgBox ("This part already exist for this order, you can update the quantity")
Exit Sub
Else
Mode = "ADDPART"
End If


'LETS CHECK USER HAS NOT ENTERED A NUMBER GREATER THAN THE STOCK TOTAL

If Text11.Text = Text4.Text + 1 Then
MsgBox ("The value you entered for quanity is greater than the stock amount")
Exit Sub
ElseIf lvwParts.SelectedItem.SubItems(3) = "" Then
Mode = "ADDPART"
End If

If Mode = "ADDPART" Then
With rsParts
    .AddNew

    ![WorkorderID] = Text8.Text
    ![PartID] = Text10.Text
    ![Quantity] = Text11.Text
    ![UnitPrice] = Text12.Text
    
    .Update
    .MoveLast
    
End With
    AdjustStock
    ViewParts
    AddTotal
    ClearPartFields
    
'SaveStatus
frmTree.FillListView
frmTree.FillInfo
frmTree.GetTax

Text11.Locked = True
cmdClose.Visible = True
cmdCancel.Visible = False
cmdDelete.Enabled = True
End If

Exit Sub
EH:
End Sub

Private Sub Text11_DblClick()
If cmdUpdate.Visible = True Then
Exit Sub
Else
cmdShowParts_Click
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
cmdAddPart_Click
End If
End Sub

Private Sub GetPartsWorkorderID()
If frmWorkorders.Visible = True Then
Text8.Text = " " & frmWorkorders.Text2.Text
Label8.Caption = " " & frmWorkorders.Text2.Text
Else
Text8.Text = frmTree.txtWorkorderID.Text
Label8.Caption = " " & frmTree.txtWorkorderID.Text
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmTree.FillListView
frmTree.TvwCustomer.SetFocus
End Sub

Private Sub ViewParts()
lvwParts.ListItems.Clear

sqlViewParts = "SELECT [Workorder Parts].WorkorderPartID, [Workorder Parts].WorkorderID, [Workorder Parts].PartID, Parts.PartName, Parts.PartDescription, [Workorder Parts].Quantity, [Workorder Parts].UnitPrice, [Quantity]*[Workorder Parts].[UnitPrice] AS Total "
sqlViewParts = sqlViewParts & "FROM Parts INNER JOIN [Workorder Parts] ON Parts.PartID = [Workorder Parts].PartID "
sqlViewParts = sqlViewParts & "WHERE [Workorder Parts].WorkorderID = " & Label8.Caption
Set rsViewParts = DB.OpenRecordset(sqlViewParts)

While Not rsViewParts.EOF
Set itmx = lvwParts.ListItems.Add(, , _
rsViewParts!WorkorderPartID)
itmx.SubItems(1) = rsViewParts!WorkorderID
itmx.SubItems(2) = rsViewParts!PartID
itmx.SubItems(3) = rsViewParts!PartName
itmx.SubItems(4) = rsViewParts!PartDescription
itmx.SubItems(5) = rsViewParts!Quantity
itmx.SubItems(6) = Format$(rsViewParts!UnitPrice, "$#,##0.00;(#,##0.00)")
itmx.SubItems(7) = Format$(rsViewParts!Total, "$#,##0.00;(#,##0.00)")
rsViewParts.MoveNext
Wend
DoEvents
End Sub

Private Sub GetName()
On Error GoTo EH:
sqlName = "Select CompanyName, CustomerID From Customers "
sqlName = sqlName & "Where CustomerID = " & Text8.Text
Set rsName = DB.OpenRecordset(sqlName)

Label1.Caption = rsName!CompanyName
EH: MsgBox "Cannot retrieve customers name"
End Sub

'----------------------------------------------------------------------------
'                            MAIN STOCK SECTION
'----------------------------------------------------------------------------

Public Sub CStock() 'SEE TOTAL CURRENT STOCK
Dim rsStock As Recordset
Dim sqlStock As String

sqlStock = "Select PartID, UnitsStock From Parts Where PartID = " & Text10
Set rsStock = DB.OpenRecordset(sqlStock)

Text4.Text = rsStock!UnitsStock
End Sub

Public Sub AdjustStock() 'ADDING NEW PART TO ORDER - REMOVE FROM TOTAL STOCK
On Error Resume Next
Dim rs As Recordset
Dim sql As String

sql = "Select PartID, Quantity From [Workorder Parts] Where PartID = " & Text10
Set rs = DB.OpenRecordset(sql)

Text5.Text = Format$(Text11 - Text4.Text, "0;0")

'----------------------------------------
'----------------------------------------
Dim rsStock As Recordset
Dim sqlStock As String

sqlStock = "Select PartID, UnitsStock From Parts Where PartID = " & Text10
Set rsStock = DB.OpenRecordset(sqlStock)

With rsStock
.Edit
!UnitsStock = Text5.Text
.Update
.MoveNext
End With

End Sub

Public Sub RMVPart() ' REMOVING PART FROM ORDER -- ADD THE STOCK BACK
'Restock Parts Stock Quantity
Dim rsStock As Recordset
Dim sqlStock As String

sqlStock = "Select PartID, UnitsStock From Parts Where PartID = " & Text10
Set rsStock = DB.OpenRecordset(sqlStock)

Text5.Text = rsStock!UnitsStock + Text11.Text

With rsStock
.Edit
!UnitsStock = Text5.Text
.Update
End With
End Sub

'----------------------------------------------------------------------------
'                MAIN SECTION QUANTITY UPDATE SECTION
'----------------------------------------------------------------------------

Public Sub SPStock() 'FOR Update Quantity
'Lets see what the stock was before adding part

Dim rsStock As Recordset
Dim sqlStock As String

sqlStock = "Select PartID, UnitsStock From Parts Where PartID = " & Text10
Set rsStock = DB.OpenRecordset(sqlStock)

'Add it back
Text4.Text = rsStock!UnitsStock
Text6.Text = Format$(rsStock!UnitsStock - Text11.Text, "0;0")
'-------------------------------------------------------------------
'Current Workorder Parts Quantity
Dim rs As Recordset
Dim sql As String

sql = "Select WorkorderPartID, Quantity From [Workorder Parts] Where WorkorderPartID = " & Text9
Set rs = DB.OpenRecordset(sql)

Text5.Text = Format$(rs!Quantity, "0;0") 'For use later

Text7.Text = Format$(rsStock!UnitsStock + Text5.Text, "0;0")

End Sub

Public Sub ReStock()
'--------------------------------------------------------------------
'RESTOCK THE PARTS STOCK
'ADD it back to what it was before adding part
'If you do not do this it will update the current stock instead of
'the overall total stock
'--------------------------------------------------------------------
Dim rsStock As Recordset
Dim sqlStock As String

sqlStock = "Select PartID, UnitsStock From Parts Where PartID = " & Text10
Set rsStock = DB.OpenRecordset(sqlStock)

Text7.Text = Format$(rsStock!UnitsStock + Text5.Text, "0;0")

With rsStock
.Edit
!UnitsStock = Text7.Text
.Update
End With

End Sub

Public Sub UpdateStock()
Dim rsStock As Recordset
Dim sqlStock As String

'Add the new updated quantity

sqlStock = "Select PartID, UnitsStock From Parts Where PartID = " & Text10
Set rsStock = DB.OpenRecordset(sqlStock)

Text6.Text = Format$(rsStock!UnitsStock - Text11.Text, "0;0")

With rsStock
.Edit
!UnitsStock = Text6.Text
.Update
End With
End Sub

Private Sub CancelUpdateStock()
Dim rsStock As Recordset
Dim sqlStock As String

sqlStock = "Select PartID, UnitsStock From Parts Where PartID = " & Text10
Set rsStock = DB.OpenRecordset(sqlStock)

Text7.Text = Format$(rsStock!UnitsStock - Text5.Text, "0;0")

With rsStock
.Edit
!UnitsStock = Text7.Text
.Update
End With
End Sub

Private Sub SaveStatus()
Dim rsStatus As Recordset
Dim sqlStatus As String
sqlStatus = "Select Workorders.WorkorderID, Workorders.Status FROM Workorders "
sqlStatus = sqlStatus & "WHERE Workorders.WorkorderID = " & frmTree.txtWorkorderID.Text
Set rsStatus = DB.OpenRecordset(sqlStatus)
frmTree.PB.Value = frmTree.PB.Value - 7
With rsStatus
.Edit
!Status = frmTree.PB.Value
.Update
End With
End Sub

Private Sub DelStatus()
Dim rsStatus As Recordset
Dim sqlStatus As String
sqlStatus = "Select Workorders.WorkorderID, Workorders.Status FROM Workorders "
sqlStatus = sqlStatus & "WHERE Workorders.WorkorderID = " & frmTree.txtWorkorderID.Text
Set rsStatus = DB.OpenRecordset(sqlStatus)
frmTree.PB.Value = frmTree.PB.Value + 7
With rsStatus
.Edit
!Status = frmTree.PB.Value
.Update
End With
End Sub
