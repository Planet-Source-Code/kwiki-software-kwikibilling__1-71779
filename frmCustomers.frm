VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCustomers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Customers"
   ClientHeight    =   6135
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   9840
   Icon            =   "frmCustomers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   8400
      TabIndex        =   27
      Top             =   5640
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
      Image           =   "frmCustomers.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   6000
      TabIndex        =   26
      Top             =   5640
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
      Image           =   "frmCustomers.frx":0A06
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   5640
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
      Image           =   "frmCustomers.frx":12E0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   5640
      Width           =   1335
      _ExtentX        =   2355
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
      Image           =   "frmCustomers.frx":1CDA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   5640
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
      Image           =   "frmCustomers.frx":25B4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5640
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
      Image           =   "frmCustomers.frx":2E8E
      cBack           =   -2147483633
   End
   Begin VB.Frame fraCurrentRec 
      Caption         =   "Current Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   9615
      Begin VB.TextBox txtLast 
         Height          =   285
         Left            =   720
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtFirst 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtState 
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
         Height          =   375
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   10
         Top             =   2040
         Width           =   555
      End
      Begin VB.TextBox txtCity 
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
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtZip 
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
         Height          =   375
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   8
         Top             =   2040
         Width           =   915
      End
      Begin VB.TextBox txtAddr 
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
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtCompanyName 
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
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtPhone 
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
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6960
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   18
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   17
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   16
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Street Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   1080
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Full Customers Name"
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
         TabIndex        =   14
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Account Number (Assign Account Number)"
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
         Left            =   4680
         TabIndex        =   13
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   9000
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.TextBox txtFind 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
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
      Height          =   285
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwCustomer 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlLVIcons"
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
   Begin MSComctlLib.ImageList imlLVIcons 
      Left            =   7440
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomers.frx":3768
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Search Customers Account  # :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjConn                As ADODB.Connection
Private mobjCmd                 As ADODB.Command
Private mobjRst                 As ADODB.Recordset

Private mstrMaintMode           As String
Private mblnFormActivated       As Boolean
Private mblnUpdateInProgress    As Boolean

' Customer LV SubItem Indexes ...
Private Const mlngCO_CO_IDX              As Long = 1
Private Const mlngCUST_ADDR_IDX          As Long = 2
Private Const mlngCUST_CITY_IDX          As Long = 3
Private Const mlngCUST_ST_IDX            As Long = 4
Private Const mlngCUST_ZIP_IDX           As Long = 5
Private Const mlngCUST_PHONE_IDX         As Long = 6
Private Const mlngCUST_ACCT_IDX          As Long = 7
Private Const mlngCUST_FIRST_IDX         As Long = 8
Private Const mlngCUST_LAST_IDX          As Long = 9
Private Const mlngCUST_ID_IDX            As Long = 10

'*****************************************************************************
'*                          General Form Events                              *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub Form_Load()
'-----------------------------------------------------------------------------
    
    ConnectToDB
    
    SetupCustLVCols
    
    LoadCustomerListView
    
End Sub

'-----------------------------------------------------------------------------
Private Sub Form_Activate()
'-----------------------------------------------------------------------------
    
    If mblnFormActivated Then Exit Sub
    
    Refresh
    
    SetFormState True
    
    mblnFormActivated = True

End Sub

'-----------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------
    
    Dim objRst  As ADODB.Recordset
    
    If mblnUpdateInProgress Then
        MsgBox "You must save or cancel the current action before " _
             & "closing this window.", _
               vbInformation, _
               "Cannot Close"
        Cancel = 1
        Exit Sub
    End If
    frmTree.UpdateTree
    DisconnectFromDB
    
    Set frmCustomers = Nothing
    
frmTree.TvwCustomer.SetFocus
End Sub

'*****************************************************************************
'*                        Command Button Events                              *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub cmdAdd_Click()
'-----------------------------------------------------------------------------

    mstrMaintMode = "ADD"
    mblnUpdateInProgress = True
    
    ClearCurrRecControls
    
    SetFormState False
    
    txtCompanyName.SetFocus
End Sub


'-----------------------------------------------------------------------------
Private Sub cmdUpdate_Click()
'-----------------------------------------------------------------------------
    
    If lvwCustomer.SelectedItem Is Nothing Then
        MsgBox "No Customer selected to update.", _
               vbExclamation, _
               "Update"
        Exit Sub
    End If
    
    mstrMaintMode = "EDIT"
    mblnUpdateInProgress = True
    frmTree.UpdateTree
    SetFormState False
    
    txtCompanyName.SetFocus

End Sub


'-----------------------------------------------------------------------------
Private Sub cmdDelete_Click()
'-----------------------------------------------------------------------------

    Dim strName    As String
    Dim lngCustID       As Long
    Dim lngNewSelIndex  As Long
    
    If lvwCustomer.SelectedItem Is Nothing Then
        MsgBox "No Customer selected to delete.", _
               vbExclamation, _
               "Delete"
        Exit Sub
    End If
    
    With lvwCustomer.SelectedItem
        strName = .SubItems(mlngCO_CO_IDX)
        lngCustID = CLng(.SubItems(mlngCUST_ID_IDX))
    End With
    
    If MsgBox("Are you sure that you want to delete Customer > " _
    & strName & " ?", vbYesNo + vbQuestion, _
    "Confirm Delete") = vbNo Then Exit Sub
    
    
    mobjCmd.CommandText = "DELETE FROM Customers WHERE CustomerID = " & lngCustID
    mobjCmd.Execute

    With lvwCustomer
        If .SelectedItem.Index = .ListItems.count Then
            lngNewSelIndex = .ListItems.count - 1
        Else
            lngNewSelIndex = .SelectedItem.Index
        End If
        .ListItems.Remove .SelectedItem.Index
        If .ListItems.count > 0 Then
            Set .SelectedItem = .ListItems(lngNewSelIndex)
            lvwCustomer_ItemClick .SelectedItem
        Else
            ClearCurrRecControls
        End If
    End With
frmTree.Label1.Caption = ""
frmTree.Label2.Caption = ""
frmTree.Label7.Caption = ""
frmTree.Label8.Caption = ""
frmTree.Label10.Caption = ""
frmTree.Label17.Caption = ""
frmTree.Label18.Caption = ""
frmTree.Label24.Caption = ""
frmTree.lblName.Caption = ""
frmTree.txtWorkorderID.Text = "(Null)"

frmTree.UpdateTree
frmCustProg.Show

End Sub

'-----------------------------------------------------------------------------
Private Sub cmdClose_Click()
'-----------------------------------------------------------------------------
SndClick
frmWorkorders.rsCustName.Refresh
frmTree.UpdateTree
Unload Me
End Sub


'-----------------------------------------------------------------------------
Private Sub cmdSave_Click()
'-----------------------------------------------------------------------------
On Error GoTo EH:

    Dim strPhone        As String
    Dim objNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strSQL          As String

    If Not ValidateFormFields Then Exit Sub
    
    If mstrMaintMode = "ADD" Then
    
        lngIDField = GetNextCustID()
        
        strSQL = "INSERT INTO Customers(  CustomerID"
        strSQL = strSQL & "            , ContactFirstName"
        strSQL = strSQL & "            , ContactLastName"
        strSQL = strSQL & "            , BillingAddress"
        strSQL = strSQL & "            , City"
        strSQL = strSQL & "            , StateOrProvince"
        strSQL = strSQL & "            , PostalCode"
        strSQL = strSQL & "            , PhoneNumber"
        strSQL = strSQL & "            , CompanyName"
        strSQL = strSQL & "            , AccountNum"
        strSQL = strSQL & "         ) VALUES ("
        strSQL = strSQL & lngIDField
        strSQL = strSQL & ", '" & Replace$(txtFirst.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtLast.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtAddr.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtCity.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & txtState.Text & "'"
        strSQL = strSQL & ", '" & txtZip.Text & "'"
        strSQL = strSQL & ", '" & txtPhone.Text & "'"
        strSQL = strSQL & ", '" & Replace$(txtCompanyName.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text1.Text, "'", "''") & "'"
        strSQL = strSQL & ")"
        
        Set objNewListItem = lvwCustomer.ListItems.Add(, , Text1.Text, , "Custs")
        PopulateListItem objNewListItem
        With objNewListItem
           .SubItems(mlngCUST_ID_IDX) = CStr(lngIDField)
           .EnsureVisible
        End With
        Set lvwCustomer.SelectedItem = objNewListItem
        Set objNewListItem = Nothing
    Else
        lngIDField = CLng(lvwCustomer.SelectedItem.SubItems(mlngCUST_ID_IDX))
        
        strSQL = "UPDATE Customers SET "
        strSQL = strSQL & "  ContactFirstName   = '" & Replace$(txtFirst.Text, "'", "''") & "'"
        strSQL = strSQL & ", ContactLastName    = '" & Replace$(txtLast.Text, "'", "''") & "'"
        strSQL = strSQL & ", BillingAddress     = '" & Replace$(txtAddr.Text, "'", "''") & "'"
        strSQL = strSQL & ", City               = '" & Replace$(txtCity.Text, "'", "''") & "'"
        strSQL = strSQL & ", StateOrProvince    = '" & txtState.Text & "'"
        strSQL = strSQL & ", PostalCode         = '" & txtZip.Text & "'"
        strSQL = strSQL & ", PhoneNumber        = '" & txtPhone.Text & "'"
        strSQL = strSQL & ", CompanyName        = '" & Replace$(txtCompanyName.Text, "'", "''") & "'"
        strSQL = strSQL & ", AccountNum        = '" & Text1.Text & "'"
        strSQL = strSQL & " WHERE CustomerID = " & lngIDField
        
        'lvwCustomer.SelectedItem.Text = Text1.Text
        PopulateListItem lvwCustomer.SelectedItem
    End If
    
    mobjCmd.CommandText = strSQL
    mobjCmd.Execute
    SetFormState True

    
    mblnUpdateInProgress = False
    
   frmCustProg.Show
   
Exit Sub
EH:
If mstrMaintMode = "ADD" Then
Label9.Caption = "Account Not Updated...Account Number Assigned"
lvwCustomer.ListItems.Remove lvwCustomer.SelectedItem.Index
Text1.SetFocus
frmCusErr.Show
ElseIf mstrMaintMode = "EDIT" Then
Label9.Caption = "Account Not Updated...Account Number Assigned"
Text1.SetFocus
frmCusErr.Show
End If
End Sub


'-----------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------------
LoadCustomerListView
mblnUpdateInProgress = False
SetFormState True
lvwCustomer_ItemClick lvwCustomer.SelectedItem
    
End Sub


'*****************************************************************************
'*                          ListView Events                                  *
'*****************************************************************************

'-------------------------------------------------------------------------
Private Sub lvwCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'-------------------------------------------------------------------------
    
    ' sort the listview on the column clicked
    With lvwCustomer
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
    If Not lvwCustomer.SelectedItem Is Nothing Then
        lvwCustomer.SelectedItem.EnsureVisible
    End If

End Sub

'-----------------------------------------------------------------------------
Private Sub lvwCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
'-----------------------------------------------------------------------------
On Error Resume Next
    gblnPopulating = True
    
    With Item
        Text1.Text = .Text
        txtCompanyName.Text = .SubItems(mlngCO_CO_IDX)
        txtAddr.Text = .SubItems(mlngCUST_ADDR_IDX)
        txtCity.Text = .SubItems(mlngCUST_CITY_IDX)
        txtState.Text = .SubItems(mlngCUST_ST_IDX)
        txtZip.Text = .SubItems(mlngCUST_ZIP_IDX)
        txtPhone.Text = .SubItems(mlngCUST_PHONE_IDX)
        txtFirst.Text = .SubItems(mlngCUST_FIRST_IDX)
        txtLast.Text = .SubItems(mlngCUST_LAST_IDX)
    End With
    
    gblnPopulating = False
    
End Sub

'*****************************************************************************
'*                      Other Control Events                                 *
'*****************************************************************************

Private Sub txtCompanyName_GotFocus()
    SelectTextboxText txtFirst
End Sub

Private Sub Text1_GotFocus()
    SelectTextboxText Text1
End Sub

Private Sub txtFirst_GotFocus()
    SelectTextboxText txtFirst
End Sub
Private Sub txtLast_GotFocus()
    SelectTextboxText txtLast
End Sub
Private Sub txtAddr_GotFocus()
    SelectTextboxText txtAddr
End Sub
Private Sub txtCity_GotFocus()
    SelectTextboxText txtCity
End Sub
Private Sub txtState_GotFocus()
    SelectTextboxText txtState
End Sub
Private Sub txtState_Change()
    TabToNextTextBox txtState, txtZip
End Sub

Private Sub txtZip_GotFocus()
    SelectTextboxText txtZip
End Sub
Private Sub txtZip_Change()
    TabToNextTextBox txtZip, txtPhone
End Sub

'*****************************************************************************
'*               Programmer-Defined Subs & Functions                         *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub ConnectToDB()
'-----------------------------------------------------------------------------

    Set mobjConn = New ADODB.Connection
    mobjConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Kwikidat\db2.mdb" & ";Persist Security Info=False"
    mobjConn.Open

    Set mobjCmd = New ADODB.Command
    Set mobjCmd.ActiveConnection = mobjConn
    mobjCmd.CommandType = adCmdText

End Sub

'-----------------------------------------------------------------------------
Private Sub DisconnectFromDB()
'-----------------------------------------------------------------------------

    Set mobjCmd = Nothing
    
    mobjConn.Close
    Set mobjConn = Nothing

End Sub


'-----------------------------------------------------------------------------
Private Sub ClearCurrRecControls()
'-----------------------------------------------------------------------------
    
    gblnPopulating = True
    
    txtFirst.Text = " "
    txtLast.Text = " "
    txtAddr.Text = ""
    txtCity.Text = ""
    txtState.Text = ""
    txtZip.Text = ""
    txtPhone.Text = ""
    txtCompanyName.Text = ""
    Text1.Text = ""
    gblnPopulating = False
    
End Sub

'-----------------------------------------------------------------------------
Private Sub SetFormState(pblnEnabled As Boolean)
'-----------------------------------------------------------------------------

    lvwCustomer.Enabled = pblnEnabled
    cmdAdd.Enabled = pblnEnabled
    cmdUpdate.Enabled = pblnEnabled
    cmdDelete.Enabled = pblnEnabled
    cmdClose.Enabled = pblnEnabled
    txtFind.Enabled = pblnEnabled
        
    txtFirst.Enabled = Not pblnEnabled
    txtLast.Enabled = Not pblnEnabled
    txtAddr.Enabled = Not pblnEnabled
    txtCity.Enabled = Not pblnEnabled
    txtState.Enabled = Not pblnEnabled
    txtZip.Enabled = Not pblnEnabled
    txtPhone.Enabled = Not pblnEnabled
    txtCompanyName.Enabled = Not pblnEnabled
    Text1.Enabled = Not pblnEnabled
    
    cmdSave.Enabled = Not pblnEnabled
    cmdCancel.Enabled = Not pblnEnabled

End Sub

'-----------------------------------------------------------------------------
Private Function ValidateFormFields() As Boolean
'-----------------------------------------------------------------------------
    
    If Not ValidateRequiredField(Text1, "Account Number") Then
    ValidateFormFields = False
    Exit Function
    End If
    
    
    If Not ValidateRequiredField(txtCompanyName, "Company Name") Then
    ValidateFormFields = False
    Exit Function
    End If
    
    
    If Not ValidateRequiredField(txtFirst, "First Name") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtLast, "Last Name") Then
        ValidateFormFields = False
        Exit Function
    End If
   
    If Not ValidateZipCode(txtZip) Then
        ValidateFormFields = False
        Exit Function
    End If
    
    'If Not ValidatePhoneNumber(txtPhone) Then
        'ValidateFormFields = False
        'Exit Function
    'End If
        
    ValidateFormFields = True
    
End Function

'-----------------------------------------------------------------------------
Private Sub PopulateListItem(pobjListItem As ListItem)
'-----------------------------------------------------------------------------

    With pobjListItem
        .SubItems(mlngCO_CO_IDX) = txtCompanyName.Text
        .SubItems(mlngCUST_ADDR_IDX) = txtAddr.Text
        .SubItems(mlngCUST_CITY_IDX) = txtCity.Text
        .SubItems(mlngCUST_ST_IDX) = txtState.Text
        .SubItems(mlngCUST_ZIP_IDX) = txtZip.Text
        .SubItems(mlngCUST_PHONE_IDX) = txtPhone.Text
        .SubItems(mlngCUST_FIRST_IDX) = txtFirst.Text
        .SubItems(mlngCUST_LAST_IDX) = txtLast.Text
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub SetupCustLVCols() 'SET COLUMN SIZES
'-----------------------------------------------------------------------------
                                 
    With lvwCustomer
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Account #", .Width * 0.15
        .ColumnHeaders.Add , , "Customer Name", .Width * 0.21
        .ColumnHeaders.Add , , "Address", .Width * 0.2
        .ColumnHeaders.Add , , "City", .Width * 0.15
        .ColumnHeaders.Add , , "St", .Width * 0.06
        .ColumnHeaders.Add , , "Zip", .Width * 0.1
        .ColumnHeaders.Add , , "Phone #", .Width * 0.12
        .ColumnHeaders.Add , , "FN", .Width * 0
        .ColumnHeaders.Add , , "LN", .Width * 0
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "ACCT", 0
    End With
End Sub

'-----------------------------------------------------------------------------
Public Sub LoadCustomerListView()
'-----------------------------------------------------------------------------
                                 
    Dim strSQL      As String
    Dim objCurrLI   As ListItem
    Dim strZip      As String
    Dim strPhone    As String
                                 
    strSQL = "SELECT ContactFirstName" _
           & "     , ContactLastName" _
           & "     , BillingAddress" _
           & "     , City" _
           & "     , StateOrProvince" _
           & "     , PostalCode" _
           & "     , PhoneNumber" _
           & "     , CompanyName" _
           & "     , AccountNum" _
           & "     , CustomerID" _
           & "  FROM Customers " _
           & " ORDER BY ContactLastName" _
           & "        , ContactFirstName"
    
    mobjCmd.CommandText = strSQL
    Set mobjRst = mobjCmd.Execute
    
    lvwCustomer.ListItems.Clear
    
    With mobjRst
        Do Until .EOF
            'strPhone = !PhoneNumber & ""
            'If Len(strPhone) > 0 Then
                'strPhone = "(" & Left$(strPhone, 3) & ") " _
                         '& Mid$(strPhone, 4, 3) & "-" _
                         '& Right$(strPhone, 4)
            'End If
            Set objCurrLI = lvwCustomer.ListItems.Add(, , !AccountNum & "", , "Custs")
            objCurrLI.SubItems(mlngCO_CO_IDX) = !CompanyName & ""
            objCurrLI.SubItems(mlngCUST_ADDR_IDX) = !BillingAddress & ""
            objCurrLI.SubItems(mlngCUST_CITY_IDX) = !City & ""
            objCurrLI.SubItems(mlngCUST_ST_IDX) = !StateOrProvince & ""
            objCurrLI.SubItems(mlngCUST_ZIP_IDX) = !PostalCode & ""
            objCurrLI.SubItems(mlngCUST_PHONE_IDX) = Format$(!PhoneNumber, "###-###-####;(###-###-####)") & ""
            objCurrLI.SubItems(mlngCUST_FIRST_IDX) = !ContactFirstName & ""
            objCurrLI.SubItems(mlngCUST_LAST_IDX) = !ContactLastName & ""
            objCurrLI.SubItems(mlngCUST_ACCT_IDX) = (!AccountNum) & ""
            objCurrLI.SubItems(mlngCUST_ID_IDX) = CStr(!CustomerID)
            .MoveNext
        Loop
    End With
    With lvwCustomer
        If .ListItems.count > 0 Then
            Set .SelectedItem = .ListItems(1)
            lvwCustomer_ItemClick .SelectedItem
        End If
    End With
    
    Set objCurrLI = Nothing
    Set mobjRst = Nothing
End Sub

'------------------------------------------------------------------------
Private Function GetNextCustID() As Long
'------------------------------------------------------------------------

    mobjCmd.CommandText = "SELECT MAX(CustomerID) AS MaxID FROM Customers"
    Set mobjRst = mobjCmd.Execute

    If mobjRst.EOF Then
        GetNextCustID = 1
    ElseIf IsNull(mobjRst!MaxID) Then
        GetNextCustID = 1
    Else
        GetNextCustID = mobjRst!MaxID + 1
    End If

    Set mobjRst = Nothing

End Function

Private Function UpdateTree()
frmTree.UpdateTree
End Function

Private Sub SearchList()
On Error Resume Next
Dim itm As ListItem
With lvwCustomer
Set itm = .FindItem(txtFind.Text, lvwSubItem, lvwPartial)
Label10.Caption = "Searched Customer Not Found"
If Not itm Is Nothing Then
Label10.Caption = "Searched Customer Found"
.ListItems(itm.Index).Selected = True
.SetFocus
End If
End With
frmCusMsg.Show
Set itm = Nothing
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SearchList
End If
End Sub
