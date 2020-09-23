VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00808080&
   Caption         =   "KWIKI BILLING v1.0.4 - System"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12240
   Icon            =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0442
            Key             =   "Vendor"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0D1C
            Key             =   "UpdateStock"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   741
      ButtonWidth     =   767
      ButtonHeight    =   741
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Update Tree"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Update"
            Object.ToolTipText     =   "Live Update"
            ImageKey        =   "Update"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Setup"
            Object.ToolTipText     =   "Company Setup"
            ImageKey        =   "Setup"
         EndProperty
      EndProperty
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9360
         TabIndex        =   5
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8040
         TabIndex        =   4
         Text            =   "Search Customer :"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   6600
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   420
         Left            =   3480
         TabIndex        =   2
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
         ButtonWidth     =   767
         ButtonHeight    =   741
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Vendor"
               Object.ToolTipText     =   "Add Vendor"
               ImageKey        =   "Vendor"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UpdateStock"
               Object.ToolTipText     =   "Apply Received Stock"
               ImageKey        =   "UpdateStock"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15F6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1ED0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2DAA
            Key             =   "Update"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31FC
            Key             =   "Setup"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   960
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4762
            MinWidth        =   4762
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/27/2008"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:47 PM"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
(ByVal hwndCaller As Long, ByVal pszFile As String, _
ByVal uCommand As Long, ByVal dwData As Long) As Long

Private Const HH_DISPLAY_TOC = &H1
Dim HTMLHelpFilePath As String

Private Sub MDIForm_Resize()
On Error Resume Next
Height = 9120
Width = 12360
Left = 1400
Top = 800
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim Reply As Variant

Reply = MsgBox("Would you like to backup your data now before terminating?", vbQuestion + vbYesNoCancel, "Confirm")

Select Case Reply
Case vbYes:
frmDB2.m_strType = "Backup"
frmDB2.Show

Case vbNo:
frmClose.Show

Case vbCancel:
Exit Sub

End Select
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuLiveUpdate_Click()
VerifyUpdates
End Sub

Private Sub mnuMaintainace_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmMaintain.Show
End Sub

Private Sub mnuParts_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmParts.Show
End Sub

Private Sub mnuRegister_Click()
ShellExecute 3, "open", "http://invoice.x10hosting.com/Register.htm", vbNullString, vbNullString, 1
End Sub

Private Sub mnuVerifyDB_Click()
On Error GoTo EH:
With frmWorkorders.CRInvoice
.DataFiles(0) = App.Path & "\KwikiDat\db2.mdb"
.ReportFileName = App.Path & "\KwikiDat\Invoice.rpt"
.WindowTitle = "Verify"
.RetrieveDataFiles
End With
EH:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    Select Case Button.Key
        Case "Refresh"
        Call UpdateTree
            
        Case "Help"
        SndPlayEx App.Path & "\Sounds\Start.wav"
        Call Help
            
        Case "Update"
        VerifyUpdates
        
        Case "Setup"
        SndPlayEx App.Path & "\Sounds\Start.wav"
        frmCompanySetup.Show
        
    End Select
End Sub

Private Sub mnuCalc_Click()
On Error GoTo EH
Shell "C:\Windows\system32\Calc.exe"
Exit Sub
EH:
MsgBox "Calculator not found on your system"
End Sub

Private Sub mnuPaymentAdd_Click()
frmPayments.Show
frmPayments.SetFocus
End Sub

Private Sub mnuPrint_Click()
On Error Resume Next
With CD
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        .Flags = .Flags + cdlPDSelection
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
        End If
    End With
End Sub

Private Sub MDIForm_Load()
WindowState = 2
Screen.MousePointer = vbHourglass
App.HelpFile = App.Path & "\Help\Kwiki_Help.chm"
HTMLHelpFilePath = App.Path & "\Help\Kwiki_Help.chm"
Load frmTree
Screen.MousePointer = vbDefault
sbStatus.Panels.Item(1).Text = "  Ready"
Text1.Text = frmSplash.Label6.Caption
End Sub

'Private Sub mnuTrayPopClose_Click()
'm_blnAllowClose = True
'Call RemoveFromTray
'frmClose.Show
'End Sub

'Private Sub mnuTrayPopRestore_Click()
'MDIForm1.Show
'frmTree.Show
'Call RemoveFromTray
'End Sub

'Private Sub mnuTrayPopCancel_Click()
'mnuTrayPop.Visible = False
'End Sub

'Private Sub mnuAddEmp_Click()
'SndPlayEx App.Path & "\Sounds\Start.wav"
'frmEmployees.Show
'End Sub

Private Sub mnuAddNewWO_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmWorkorders.Show
End Sub

Private Sub mnuAddPayment_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmPayments.Show
End Sub

Private Sub mnuCompSetep_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmCompanySetup.Show
End Sub

Private Sub mnuOpenParts_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmWorkorderParts.Show
End Sub

Private Sub mnuOpenLabor_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmWorkorderLabor.Show
End Sub

Private Sub mnuExit_Click()
Dim Reply As Variant

Reply = MsgBox("Would you like to backup your data now before closing?", vbQuestion + vbYesNo, "Confirm")

Select Case Reply
Case vbYes:
frmDB2.m_strType = "Backup"
frmDB2.Show

Case vbNo:
frmClose.Show
End

End Select
End Sub

Private Sub mnuNewCustomer_Click()
frmCustomers.Show
End Sub

Private Sub mnuOpenCat_Click()
frmCategories.Show
End Sub

Private Sub mnuPaymentMeth_Click()
frmPaymentMethod.Show
End Sub

Private Sub mnuTree_Click()
frmTree.Show
End Sub

Private Sub Help()
    Dim hwndHelp As Long
    hwndHelp = HtmlHelp(hWnd, HTMLHelpFilePath, HH_DISPLAY_TOC, 0)
End Sub

Public Sub UpdateTree()
frmTree.UpdateTree
'frmTree.FillListView
frmTree.lblName.Caption = ""
frmTree.Label2 = ""
frmTree.Label7 = ""
frmTree.Label8 = ""
frmTree.Label10 = ""
frmTree.Label17 = ""
frmTree.Label18 = ""
frmTree.Label24.Caption = ""
'frmTree.PB.Visible = True
End Sub

Public Sub FindCustomer()
    Dim i As Integer
    Dim Clear As String
    Dim FindPos As Integer
    Dim Found As Boolean
    Found = False
    
    'Searcher = InputBox("Enter Customers First & Last Name", "Search Customer")


    For i = 1 To frmTree.TvwCustomer.Nodes.count
        FindPos = InStr(1, frmTree.TvwCustomer.Nodes(i).Text, txtSearch, vbTextCompare)
        
        If txtSearch = Clear Then Exit Sub

        If FindPos <> 0 Then
            frmTree.TvwCustomer.SetFocus
            frmTree.TvwCustomer.Nodes(i).Selected = True
            FindPos = 0
            Found = True
            If Found = True Then
            frmTree.Label25.Caption = ("Searched Customer Has Been Found...")
            frmSrchMsg.Show
            End If
        Else

            If i = frmTree.TvwCustomer.Nodes.count And Found = False Then
            frmTree.Label25.Caption = ("Searched Customer Could Not Be Found...")
            frmSrchMsg.Show
            End If
        End If
    Next i

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    Select Case Button.Key
        Case "Vendor"
        SndPlayEx App.Path & "\Sounds\Start.wav"
        frmVendors.Show
        
        Case "UpdateStock"
        SndPlayEx App.Path & "\Sounds\Start.wav"
        frmRecieveStock.Show
    End Select
End Sub

Public Sub VerifyUpdates()
If Text1.Text = "Unregistered" Then
If MsgBox("Live Updates are not available in demo mode, Would you like to register now ?", vbYesNo, "Register") = vbYes Then
ShellExecute 3, "open", "http://invoice.x10hosting.com", vbNullString, vbNullString, 1
ElseIf vbNo Then
Exit Sub
End If
ElseIf Text1.Text = "Registered" Then
SndPlayEx App.Path & "\Sounds\Start.wav"
frmCheckUpdates.Show
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
FindCustomer
End If
End Sub
