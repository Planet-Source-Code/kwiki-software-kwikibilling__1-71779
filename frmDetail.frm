VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Detail"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView LvwOrders 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1296
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   65280
      BackColor       =   4210752
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Detail.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Detail.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Detail.frx":3B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Detail.frx":5F44
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
setUpListView
End Sub

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "WOID", 800, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "CID", 600, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "DateRec'd", 1200, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "DateRequired", 1200, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "TaxRate", 1200, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "PartsTotal", 1300, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "LaborTotal", 1300, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "Parts+Labor Total", 1600, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "TotalPayments", 1300, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "AmountDue", 1200, lvwColumnLeft)
LvwOrders.View = lvwReport
End Sub

