VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmParts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Products"
   ClientHeight    =   7440
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParts.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFind 
      BackColor       =   &H8000000E&
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
      Left            =   6840
      TabIndex        =   34
      Top             =   120
      Width           =   2175
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   7680
      TabIndex        =   27
      Top             =   6960
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
      Image           =   "frmParts.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   5880
      TabIndex        =   26
      Top             =   6960
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
      Image           =   "frmParts.frx":0A06
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   6960
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
      Image           =   "frmParts.frx":12E0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "frmParts.frx":1BBA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "frmParts.frx":2494
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "frmParts.frx":2D6E
      cBack           =   -2147483633
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3960
      TabIndex        =   16
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bar Label"
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   8895
      Begin lvButton.lvButtons_H cmdPrint 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Print Bar Code"
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
         Image           =   "frmParts.frx":3648
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNewLabel 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Create Bar Code"
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
         Image           =   "frmParts.frx":4522
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdClearBar 
         Height          =   375
         Left            =   7080
         TabIndex        =   30
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Clear Bar Code"
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
         Image           =   "frmParts.frx":4DFC
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdSaveBar 
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Save Bar Code"
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
         Image           =   "frmParts.frx":56D6
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdAddLabel 
         Height          =   375
         Left            =   7080
         TabIndex        =   28
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Add Bar Code"
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
         Image           =   "frmParts.frx":5FB0
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.PictureBox imgPhoto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   2400
         ScaleHeight     =   1335
         ScaleWidth      =   4575
         TabIndex        =   18
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.Frame fraCurrentRec 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   8895
      Begin lvButton.lvButtons_H cmdCats 
         Height          =   375
         Left            =   6720
         TabIndex        =   33
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Open Categories"
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
      Begin VB.TextBox Text6 
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
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000012&
         DataField       =   "PartID"
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
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
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "This field is locked while addind a new part you can update it only in edit mode"
         Top             =   1800
         Width           =   1395
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   2
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   1
         Top             =   1080
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmParts.frx":688A
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         BackColor       =   12648384
         ForeColor       =   -2147483640
         ListField       =   "CategoryName"
         BoundColumn     =   "CategoryID"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "SKU :"
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Total Products Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6360
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
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
         TabIndex        =   9
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " Unit Price"
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
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Description"
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
         Index           =   0
         Left            =   4560
         TabIndex        =   7
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Product Category"
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
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
   End
   Begin MSComctlLib.ListView lvwParts 
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlLVIcons"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6480
      Top             =   6720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Categories"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   7200
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   "Search by SKU:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   7080
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents HO As cBinaryDBObject
Attribute HO.VB_VarHelpID = -1
Dim sFileName As String
Dim Clear As String

Dim rsPart As Recordset
Dim sqlPart As String

Private mobjConn                As ADODB.Connection
Private mobjCmd                 As ADODB.Command
Private mobjRst                 As ADODB.Recordset

Private mstrMaintMode           As String
Private mblnFormActivated       As Boolean
Private mblnUpdateInProgress    As Boolean

' Customer LV SubItem Indexes ...
Private Const mlngPART_ID_IDX            As Long = 1
Private Const mlngPART_CODE_IDX          As Long = 2
Private Const mlngPART_NAME_IDX          As Long = 3
Private Const mlngPART_DESC_IDX          As Long = 4
Private Const mlngPART_UPRICE_IDX        As Long = 5
Private Const mlngPART_STOCK_IDX         As Long = 6
Private Const mlngPART_CAT_IDX           As Long = 7

Private Sub cmdCats_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmCategories.Show
End Sub

Private Sub ChangeCombo()
On Error Resume Next
OpenDatabase
If (Not OpenDatabase()) Then
  MsgBox "Database could not be openend !"
End If

sqlPart = "Select * From Parts Where PartID = " & lvwParts.SelectedItem.Text
Set rsPart = DB.OpenRecordset(sqlPart)

If rsPart.RecordCount > 0 Then
rsPart.MoveFirst
End If
DataCombo1.ListField = rsPart!CategoryID
End Sub

Private Sub GetPic()
Dim rsPic As Recordset
Dim sqlPic As String

sqlPic = "Select * From Parts Where PartID = " & lvwParts.SelectedItem.Text
Set rsPic = DB.OpenRecordset(sqlPic)

imgPhoto.Picture = LoadPicture(rsPic!FileName)
End Sub


Private Sub cmdAddLabel_Click()
On Error Resume Next
With CD2
.DialogTitle = "Add Bar label"
.CancelError = False
.Filter = "Bitmap Files (*.bmp*)|*.BMP|" _
& "Gif Files (*.gif)|*.GIF|" _
& "Jpeg Files (*.jpg)|*.JPG|" _
& "Windows Meta Files (*.wmf)|*.WMF|" _
& "All Files (*.*)|*.*"
.ShowSave
If Len(.FileName) = 0 Then
Exit Sub
End If
sFileName = .FileName
End With
imgPhoto.Picture = LoadPicture(sFileName)
Text9.Text = sFileName
SetAddPic False

End Sub

Private Sub cmdClearBar_Click()
On Error Resume Next
Dim rsPart As Recordset
Dim sqlPart As String
Text9.Text = App.Path & "\KwikiDat\BitMaps\Default.jpg"

sqlPart = "Select PartID, FileName From Parts Where PartID = " & lvwParts.SelectedItem.Text
Set rsPart = DB.OpenRecordset(sqlPart)

With rsPart
.Edit
rsPart!FileName = Text9.Text
.Update
End With

SaveBinaryObject

imgPhoto.Picture = LoadPicture(rsPart!FileName)
SetAddPic False
End Sub

Private Sub cmdNewLabel_Click()
SndClick
frmBar.Show
End Sub

Private Sub cmdPrint_Click()
SndClick
frmPrint.Show
End Sub

Private Sub cmdSaveBar_Click()
Text9.Text = sFileName
SetFormState True

Dim rsPart As Recordset
Dim sqlPart As String
sqlPart = "Select PartID, FileName From Parts Where PartID = " & lvwParts.SelectedItem.Text
Set rsPart = DB.OpenRecordset(sqlPart)

If Text9.Text = "" Then
Text9.Text = rsPart!FileName
End If

With rsPart
.Edit
rsPart!FileName = Text9.Text
.Update
End With
cmdAddLabel.Enabled = False
cmdSaveBar.Enabled = False
cmdClearBar.Enabled = False
mblnUpdateInProgress = False
End Sub

Private Sub DataCombo1_Change()
'SndPlayEx App.Path & "\Sounds\OpenMenu.wav"
End Sub

'*****************************************************************************
'*                          General Form Events                              *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub Form_Load()
'-----------------------------------------------------------------------------

    ConnectToDB
    
    SetupPartLVCols
    
    LoadPartListView
    
   If (IsNull(Adodc1.Recordset!CategoryID)) Then
   DataCombo1.BoundText = Adodc1.Recordset!CategoryID
   End If
   
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
    
    'Dim objRst  As ADODB.Recordset
    
    If mblnUpdateInProgress Then
        MsgBox "You must save or cancel the current action before " _
             & "closing this window.", _
               vbInformation, _
               "Cannot Close"
        Cancel = 1
        Exit Sub
    End If

    DisconnectFromDB
    
    Set frmParts = Nothing

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
    
    cmdAddLabel.Enabled = False
    cmdSaveBar.Enabled = False
    cmdClearBar.Enabled = False
    imgPhoto.Picture = LoadPicture(Clear)
    Text8.Text = Clear
    Text2.SetFocus
    Text9.Text = App.Path & "\KwikiDat\BitMaps\Default.jpg"
End Sub


'-----------------------------------------------------------------------------
Private Sub cmdUpdate_Click()
'-----------------------------------------------------------------------------
    
    If lvwParts.SelectedItem Is Nothing Then
        MsgBox "No Product selected to update.", _
               vbExclamation, _
               "Update"
        Exit Sub
    End If
    
    mstrMaintMode = "EDIT"
    mblnUpdateInProgress = True
    SetFormState False
    cmdAddLabel.Enabled = True
    cmdSaveBar.Enabled = True
    cmdClearBar.Enabled = True
    Text5.Locked = False
    Text2.SetFocus

End Sub


'-----------------------------------------------------------------------------
Private Sub cmdDelete_Click()
'-----------------------------------------------------------------------------

    'Dim strPartName     As String
    'Dim strPartDesc     As String
    'Dim lngPartID       As Long
    Dim lngNewSelIndex  As Long
    
    If lvwParts.SelectedItem Is Nothing Then
        MsgBox "No Product selected to delete.", _
               vbExclamation, _
               "Delete"
        Exit Sub
    End If
    
    
    'With lvwParts.SelectedItem
        '.strPartName = .Text
        '.strPartDesc = .SubItems(mlngPART_NAME_IDX)
        'lngPartID = CLng(lvwParts.SelectedItem.SubItems(mlngPART_ID_IDX))
    'End With
    
    If MsgBox("Are you sure that you want to delete Product " _
            & Text2.Text & "?", _
              vbYesNo + vbQuestion, _
              "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    mobjCmd.CommandText = "DELETE FROM Parts WHERE PartID = " & lvwParts.SelectedItem.Text
    mobjCmd.Execute

    With lvwParts
        If .SelectedItem.Index = .ListItems.count Then
            lngNewSelIndex = .ListItems.count - 1
        Else
            lngNewSelIndex = .SelectedItem.Index
        End If
        .ListItems.Remove .SelectedItem.Index
        If .ListItems.count > 0 Then
            Set .SelectedItem = .ListItems(lngNewSelIndex)
            lvwParts_ItemClick .SelectedItem
        Else
            ClearCurrRecControls
        End If
    End With
'Unload Me
frmPartProg.Show
End Sub

'-----------------------------------------------------------------------------
Private Sub cmdClose_Click()
SndClick
Unload Me
End Sub


'-----------------------------------------------------------------------------
Private Sub cmdSave_Click()
'-----------------------------------------------------------------------------
On Error GoTo EH:

    Dim objNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strSQL          As String

    If Not ValidateFormFields Then Exit Sub
        
    If mstrMaintMode = "ADD" Then
    
        lngIDField = GetNextPartID()
        
        strSQL = "INSERT INTO Parts(  PartID"
        strSQL = strSQL & "            , CategoryID"
        strSQL = strSQL & "            , PartName"
        strSQL = strSQL & "            , PartDescription"
        strSQL = strSQL & "            , UnitPrice"
        strSQL = strSQL & "            , UnitsStock"
        strSQL = strSQL & "            , PartCode"
        strSQL = strSQL & "            , FileName"
        strSQL = strSQL & "            , PartImage"
        strSQL = strSQL & "         ) VALUES ("
        strSQL = strSQL & lngIDField
        strSQL = strSQL & ", '" & Replace$(DataCombo1.BoundText, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text2.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text3.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text4.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text5.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text6.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text9.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(imgPhoto.Picture = imgPhoto.Picture, "'", "''") & "'"
        strSQL = strSQL & ")"
        
        Set objNewListItem = lvwParts.ListItems.Add(, , Text1.Text)
        PopulateListItem objNewListItem
        With objNewListItem
            .SubItems(mlngPART_ID_IDX) = CStr(lngIDField)
            .EnsureVisible
        End With
        
        Set lvwParts.SelectedItem = objNewListItem
        Set objNewListItem = Nothing

    Else
    'On Error Resume Next
        lngIDField = lvwParts.SelectedItem.Text
        strSQL = "UPDATE Parts SET "
        strSQL = strSQL & "  CategoryID = '" & Replace$(DataCombo1.BoundText, "'", "''") & "'"
        strSQL = strSQL & ", PartName    = '" & Replace$(Text2.Text, "'", "''") & "'"
        strSQL = strSQL & ", PartDescription     = '" & Replace$(Text3.Text, "'", "''") & "'"
        strSQL = strSQL & ", UnitPrice     = '" & Replace$(Text4.Text, "'", "''") & "'"
        strSQL = strSQL & ", UnitsStock    = '" & Replace$(Text5.Text, "'", "''") & "'"
        strSQL = strSQL & ", PartCode    = '" & Replace$(Text6.Text, "'", "''") & "'"
        strSQL = strSQL & ", PartImage    = '" & Replace$(imgPhoto.Picture, "'", "''") & "'"
        strSQL = strSQL & " WHERE PartID = " & lngIDField
        
        lvwParts.SelectedItem.Text = Text1.Text
        PopulateListItem lvwParts.SelectedItem
        
        SaveBinaryObject
        GetTotal
        Text5.Locked = True
    End If
    
    mobjCmd.CommandText = strSQL
    mobjCmd.Execute
    SetFormState True
    cmdAddLabel.Enabled = False
    
    mblnUpdateInProgress = False
Unload Me
frmPartProg.Show

Exit Sub
EH:
Label8.Caption = "Product Not Updated...SKU Number Already Exist"
frmPartErr.Show
End Sub

Private Sub HO_Error(ID As Long, Msg As String)
MsgBox ID & ":  " & Msg
End Sub

Private Sub HO_Status(ID As Long, Msg As String)
lblStatus.Caption = CStr(ID) & ":  " & Msg
Exit Sub
End Sub

'-----------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------------
    
    mblnUpdateInProgress = False
    SetFormState True
    lvwParts_ItemClick lvwParts.SelectedItem
    cmdAddLabel.Enabled = False
    Text5.Locked = True
End Sub


'*****************************************************************************
'*                          ListView Events                                  *
'*****************************************************************************

'-------------------------------------------------------------------------
Private Sub lvwParts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'-------------------------------------------------------------------------
    
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

'-----------------------------------------------------------------------------
Private Sub lvwParts_ItemClick(ByVal Item As MSComctlLib.ListItem)
'-----------------------------------------------------------------------------
On Error Resume Next
    gblnPopulating = True

    With Item
        Text1.Text = .Text
        DataCombo1.BoundText = .SubItems(mlngPART_CAT_IDX)
        Text2.Text = .SubItems(mlngPART_NAME_IDX)
        Text3.Text = .SubItems(mlngPART_DESC_IDX)
        Text4.Text = .SubItems(mlngPART_UPRICE_IDX)
        Text5.Text = .SubItems(mlngPART_STOCK_IDX)
        Text6.Text = .SubItems(mlngPART_CODE_IDX)
    End With
        
        GetTotal
        GetPic

        
    gblnPopulating = False

End Sub

'*****************************************************************************
'*                      Other Control Events                                 *
'*****************************************************************************

Private Sub Text2_GotFocus()
    SelectTextboxText Text2
End Sub

Private Sub Text3_GotFocus()
    SelectTextboxText Text3
End Sub
Private Sub Text4_GotFocus()
    SelectTextboxText Text4
End Sub
Private Sub Text5_GotFocus()
    SelectTextboxText Text5
End Sub


'*****************************************************************************
'*               Programmer-Defined Subs & Functions                         *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub ConnectToDB()
'-----------------------------------------------------------------------------

    Set mobjConn = New ADODB.Connection
    mobjConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\KwikiDat\db2.mdb" & ";Persist Security Info=False"
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
    
    'DataCombo1.ListField = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = "0"
    Text6.Text = ""
    
    gblnPopulating = False
    
End Sub

'-----------------------------------------------------------------------------
Private Sub SetFormState(pblnEnabled As Boolean)
'-----------------------------------------------------------------------------

    lvwParts.Enabled = pblnEnabled
    cmdAdd.Enabled = pblnEnabled
    cmdUpdate.Enabled = pblnEnabled
    cmdDelete.Enabled = pblnEnabled
    cmdClose.Enabled = pblnEnabled
    cmdCats.Enabled = pblnEnabled
    txtFind.Enabled = pblnEnabled
    'cmdOrder.Enabled = pblnEnabled
        
    DataCombo1.Enabled = Not pblnEnabled
    Text2.Enabled = Not pblnEnabled
    Text3.Enabled = Not pblnEnabled
    Text4.Enabled = Not pblnEnabled
    Text5.Enabled = Not pblnEnabled
    Text6.Enabled = Not pblnEnabled
    cmdSave.Enabled = Not pblnEnabled
    cmdCancel.Enabled = Not pblnEnabled
    cmdSaveBar.Enabled = Not pblnEnabled
    cmdClearBar.Enabled = Not pblnEnabled
End Sub

'-----------------------------------------------------------------------------
Private Sub SetAddPic(pblnEnabled As Boolean)
'-----------------------------------------------------------------------------

    lvwParts.Enabled = pblnEnabled
    cmdAdd.Enabled = pblnEnabled
    cmdUpdate.Enabled = pblnEnabled
    cmdDelete.Enabled = pblnEnabled
    cmdClose.Enabled = pblnEnabled
    cmdCats.Enabled = pblnEnabled
    cmdSave.Enabled = pblnEnabled
        
    DataCombo1.Enabled = Not pblnEnabled
    Text2.Enabled = Not pblnEnabled
    Text3.Enabled = Not pblnEnabled
    Text4.Enabled = Not pblnEnabled
    Text5.Enabled = Not pblnEnabled
    cmdCancel.Enabled = Not pblnEnabled

End Sub
'-----------------------------------------------------------------------------
Private Function ValidateFormFields() As Boolean
'-----------------------------------------------------------------------------
    
    If Not ValidateRequiredField(Text2, "Part Name") Then
    ValidateFormFields = False
    Exit Function
    End If
    
    
    If Not ValidateRequiredField(Text3, "Part Description") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(Text4, "Unit Price") Then
        ValidateFormFields = False
        Exit Function
    End If
    
        If Not ValidateRequiredField(Text6, "Part Code") Then
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
    
    On Error Resume Next
        .SubItems(mlngPART_STOCK_IDX) = DataCombo1.BoundText
        .SubItems(mlngPART_ID_IDX) = Text1.Text
        .SubItems(mlngPART_NAME_IDX) = Text2.Text
        .SubItems(mlngPART_DESC_IDX) = Text3.Text
        .SubItems(mlngPART_UPRICE_IDX) = Text4.Text
        .SubItems(mlngPART_STOCK_IDX) = Text5.Text
        .SubItems(mlngPART_CODE_IDX) = Text6.Text
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub SetupPartLVCols()
'-----------------------------------------------------------------------------
                                 
    With lvwParts
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "PID", .Width * 0
        .ColumnHeaders.Add , , "CID", .Width * 0
        .ColumnHeaders.Add , , "SKU", .Width * 0.15
        .ColumnHeaders.Add , , "Item Name", .Width * 0.28
        .ColumnHeaders.Add , , "Item Description", .Width * 0.31
        .ColumnHeaders.Add , , "Unit Price", .Width * 0.15
        .ColumnHeaders.Add , , "Stock", .Width * 0.1
        .ColumnHeaders.Add , , "Cat ID", .Width * 0
        
    End With

End Sub


'-----------------------------------------------------------------------------
Public Sub LoadPartListView()
'-----------------------------------------------------------------------------
    On Error Resume Next
    Dim strSQL      As String
    Dim objCurrLI   As ListItem
    
strSQL = "SELECT CategoryID, StockTotal.StockTotal, Parts.PartID, Parts.UnitsStock, Parts.PartName, Parts.PartDescription, Parts.UnitPrice, Parts.PartCode "
strSQL = strSQL & "FROM Parts, StockTotal "
strSQL = strSQL & "WHERE Parts.PartID = StockTotal.PartID"
    
    mobjCmd.CommandText = strSQL
    Set mobjRst = mobjCmd.Execute
    
    lvwParts.ListItems.Clear
    
    With mobjRst
        Do Until .EOF
            Set objCurrLI = lvwParts.ListItems.Add(, , CStr(!PartID))
            objCurrLI.SubItems(mlngPART_CAT_IDX) = !CategoryID & ""
            objCurrLI.SubItems(mlngPART_CODE_IDX) = !PartCode & ""
            objCurrLI.SubItems(mlngPART_NAME_IDX) = !PartName & ""
            objCurrLI.SubItems(mlngPART_DESC_IDX) = !PartDescription & ""
            objCurrLI.SubItems(mlngPART_UPRICE_IDX) = Format$(!UnitPrice, "$#,##0.00;(#,##0.00)") & ""
            objCurrLI.SubItems(mlngPART_STOCK_IDX) = !UnitsStock & ""
            .MoveNext
        Loop
    End With
    With lvwParts
        If .ListItems.count > 0 Then
            Set .SelectedItem = .ListItems(1)
            lvwParts_ItemClick .SelectedItem
        End If
    End With
    
    Set objCurrLI = Nothing
    Set mobjRst = Nothing
End Sub

'------------------------------------------------------------------------
Private Function GetNextPartID() As Long
'------------------------------------------------------------------------

    mobjCmd.CommandText = "SELECT MAX(PartID) AS MaxID FROM Parts"
    Set mobjRst = mobjCmd.Execute

    If mobjRst.EOF Then
        GetNextPartID = 1
    ElseIf IsNull(mobjRst!MaxID) Then
        GetNextPartID = 1
    Else
        GetNextPartID = mobjRst!MaxID + 1
    End If

    Set mobjRst = Nothing

End Function

Private Sub SearchList()
On Error Resume Next
Dim itm As ListItem
With lvwParts
Set itm = .FindItem(txtFind.Text, lvwSubItem, lvwPartial)
Label8.Caption = "Searched Part Not Found"
If Not itm Is Nothing Then
Label8.Caption = "Searched Part Found"
.ListItems(itm.Index).Selected = True
.SetFocus
End If
End With
frmMsg.Show
Set itm = Nothing
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SearchList
End If
End Sub

Public Sub GetTotal()
On Error Resume Next
OpenDatabase
Dim rsPart As Recordset
Dim sqlPart As String
sqlPart = "Select UnitPrice, UnitsStock From Parts Where PartID = " & lvwParts.SelectedItem.Text
Set rsPart = DB.OpenRecordset(sqlPart)

Text8.Text = " " & Format$((rsPart!UnitPrice) * (rsPart!UnitsStock), "$#,##0.00;(#,##0.00)")

End Sub

'Private Sub ShowTotal()
'On Error Resume Next
'Dim i As Integer
'Dim cTotal As Currency
'With lvwParts
'For i = 1 To .ListItems.count
'cTotal = cTotal + CCur(.ListItems(i).SubItems(mlngPART_STOCK_IDX))
'Next
'End With
'Text7.Text = Format$(cTotal, "$#,##0.00;(#,##0.00)")
'End Sub

Private Sub GetBinaryObject()
'OpenDatabase
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim fn As String                       'Binary file name to use as storage
Dim i As Integer

    If Text1.Text = "" Then
        Set HO = New cBinaryDBObject       'create the new bd object

        FieldNames(0) = "PartID"               'return the ID field
        FieldNames(1) = "FileName"         'return the filename

        With HO
            .KillFile = True                        'kill the filename if it exists
             Set .DB = DB                        'pass the database
            .ObjectKeyFieldName = "PartID"      'the key/index field is
            .ObjectKey = Text1.Text                 'the value to search for is
            .ObjectFieldName = "PartImage"              'name of the field that contains the binary file
            .ObjectTableName = "Parts"          'table that contains the binary files
            .SubFieldNames = FieldNames             'pass in the field names to return
            .FileName = "FileName"                  'file name to use"
            .GetObject                              'get the file from the database
            .ReturnData RD()                        'return any aditional data
            fn = .FileName                          'actual file name used - if default was used
        End With
        Set HO = Nothing

        imgPhoto.Picture = LoadPicture(fn)

        For i = 0 To UBound(RD)
            Debug.Print RD(i)                      'print aditional info returned
        Next

    End If
    'DB.Close
End Sub

Private Sub SaveBinaryObject()
'OpenDatabase
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim FieldData(1) As Variant            'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim fn As String                       'Binary file name to use as storage
Dim i As Integer

    If sFileName = "" Then
    Exit Sub
    End If

    Set HO = New cBinaryDBObject         'create the new bd object

     FieldNames(0) = "PartID"        'return the ID field
     FieldNames(1) = "FileName"          'return the filename
     FieldData(0) = Null                 'return the ID field
     FieldData(1) = sFileName            'return the filename

    With HO
        .KillFile = False                       'kill the filename if it exists
         Set .DB = DB                    'pass the database
        .ObjectKeyFieldName = "PartID"    'the key/index field is
        .ObjectKey = Text1.Text               'the value to search for is
        .ObjectFieldName = "PartImage"           'name of the field that contains the binary file
        .ObjectTableName = "Parts"        'table that contains the binary files
        .SubFieldNames = FieldNames            'pass in the field names to return
        .SubFieldData = FieldData
        .FileName = sFileName                  'file name to use
        .SaveObject                            'get the file from the database
        .ReturnData RD()                       'return any aditional data
        fn = .FileName                         'actual file name used - if default was used
    End With
    Set HO = Nothing


    For i = 0 To UBound(RD)
        Debug.Print RD(i)                      'print aditional info returned
    Next
    'DB.Close
End Sub

Private Function SndPlayEx(ByVal FileName As String, Optional ByVal lmodule As Long = 0, Optional ByVal options As Long = (SND_FILENAME Or SND_ASYNC)) As Long
SndPlayEx = PlaySound(FileName, lmodule, options)
End Function
