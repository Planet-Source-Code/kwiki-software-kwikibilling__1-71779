VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmWorkorders 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    "
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9510
   ForeColor       =   &H00000000&
   Icon            =   "frmWorkorders.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWorkorders.frx":000C
   ScaleHeight     =   6360
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Order Section"
      Height          =   3135
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         DataField       =   "SalesTaxRate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ".0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "rsWorkorder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         DataField       =   "SerialNumber"
         DataSource      =   "rsWorkorder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         DataField       =   "PurchaseOrderNumber"
         DataSource      =   "rsWorkorder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         DataField       =   "ProblemDescription"
         DataSource      =   "rsWorkorder"
         Height          =   1215
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1800
         Width           =   3735
      End
      Begin VB.ComboBox cmbTerms 
         BackColor       =   &H80000018&
         DataField       =   "PaymentTerms"
         DataSource      =   "rsWorkorder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmWorkorders.frx":04FE
         Left            =   1560
         List            =   "frmWorkorders.frx":0511
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Pay Terms"
         Top             =   1440
         Width           =   1695
      End
      Begin lvButton.lvButtons_H Command1 
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         ToolTipText     =   "Open Preloaded Text File"
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmWorkorders.frx":054C
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin MSDataListLib.DataCombo dcbEmpName 
         Bindings        =   "frmWorkorders.frx":1426
         DataField       =   "EmployeeID"
         DataSource      =   "rsWorkorder"
         Height          =   360
         Left            =   1200
         TabIndex        =   28
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         ListField       =   "FullName"
         BoundColumn     =   "EmployeeID"
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dcbCustName 
         Bindings        =   "frmWorkorders.frx":143E
         DataField       =   "CustomerID"
         DataSource      =   "rsWorkorder"
         Height          =   360
         Left            =   1200
         TabIndex        =   29
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         ListField       =   "CompanyName"
         BoundColumn     =   "CustomerID"
         Text            =   ""
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
      Begin MSComCtl2.DTPicker DTPicker3 
         Bindings        =   "frmWorkorders.frx":1457
         DataField       =   "DateFinished"
         DataSource      =   "rsWorkorder"
         Height          =   375
         Left            =   3960
         TabIndex        =   32
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483628
         CheckBox        =   -1  'True
         Format          =   45219841
         CurrentDate     =   39661
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Bindings        =   "frmWorkorders.frx":1462
         DataField       =   "DateRequired"
         DataSource      =   "rsWorkorder"
         Height          =   375
         Left            =   4440
         TabIndex        =   33
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483628
         CheckBox        =   -1  'True
         Format          =   45219841
         CurrentDate     =   39661
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Bindings        =   "frmWorkorders.frx":146D
         DataField       =   "DateReceived"
         DataSource      =   "rsWorkorder"
         Height          =   375
         Left            =   4440
         TabIndex        =   34
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483646
         CalendarTitleForeColor=   -2147483628
         CheckBox        =   -1  'True
         Format          =   45219841
         CurrentDate     =   39661
      End
      Begin Crystal.CrystalReport CRInvoice 
         Left            =   4200
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowExportBtn=   0   'False
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.TextBox txtFields 
         BorderStyle     =   0  'None
         DataField       =   "WorkorderID"
         DataSource      =   "rsWorkorder"
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tracking # :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   6240
         TabIndex        =   55
         Top             =   720
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   9120
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblLabels 
         Caption         =   "SalesTaxRate :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   52
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   " DateRequired :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   51
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   " DateReceived :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   50
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "PONumber :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Entered By :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Customer :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "(This rate will apply to this order only)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   45
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "(This Payment Term will apply  to this order only)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Payment Terms :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "(This section will appear at the bottom of your invoice for warranties, terms, conditions, etc.)"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5400
         TabIndex        =   42
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label lblstar 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   41
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblstar 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   40
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblstar 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   39
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblstar 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   38
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblstar 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   37
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label lblstar 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   36
         Top             =   2280
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estimate Section"
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   9255
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "MakeAndModel"
         DataSource      =   "rsWorkorder"
         Height          =   1245
         Index           =   8
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFC0&
         DataField       =   "EstimateFooter"
         DataSource      =   "rsWorkorder"
         Height          =   1245
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   600
         Width           =   3735
      End
      Begin lvButton.lvButtons_H Command2 
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         ToolTipText     =   "Open Preloaded Text File"
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmWorkorders.frx":1478
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H Command3 
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         ToolTipText     =   "Open Preloaded Text File"
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmWorkorders.frx":2352
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Caption         =   "(Estimate Header)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "(Estimate Footer)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdViewEstimate 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "&Print Estimate"
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
      Image           =   "frmWorkorders.frx":322C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPreview 
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "&Print Invoice"
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
      Image           =   "frmWorkorders.frx":4106
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAddPart 
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Add Part"
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
      Image           =   "frmWorkorders.frx":4FE0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAddLabor 
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Add Labor"
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
      Image           =   "frmWorkorders.frx":58BA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "frmWorkorders.frx":6194
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmWorkorders.frx":6B8E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmWorkorders.frx":7468
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmWorkorders.frx":7D42
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmWorkorders.frx":861C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5880
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
      Image           =   "frmWorkorders.frx":8EF6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmWorkorders.frx":97D0
      cBack           =   -2147483633
   End
   Begin MSAdodcLib.Adodc rsCustName 
      Height          =   330
      Left            =   360
      Top             =   4920
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
      CommandType     =   8
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
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select CustomerID, CompanyName From Customers"
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc rsEmpName 
      Height          =   330
      Left            =   1560
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT [EmployeeID], [FullName] FROM Employees"
      Caption         =   ""
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
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmWorkorders.frx":A0AA
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog CD3 
      Left            =   720
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      DataField       =   "WorkorderID"
      DataSource      =   "rsWorkorder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblStat 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   5400
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   2880
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmWorkorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjConn As ADODB.Connection
Private mobjCmd As ADODB.Command
Private mobjRst As ADODB.Recordset

Private mstrMaintMode As String

Dim sql As String
Public rsOrder As Recordset
Dim sFile As String

Private Sub cmdAddLabor_Click()
If Text2.Text = "" Then
MsgBox "No workorder to apply labor to"
Exit Sub
Else
SndClick
frmWorkorderLabor.Show
End If
End Sub

Private Sub cmdAddPart_Click()
If Text2.Text = "" Then
MsgBox "No workorder to apply products to"
Exit Sub
Else
SndClick
frmWorkorderParts.Show
End If
End Sub


Private Sub cmdPreview_Click()
On Error GoTo EH:
SndClick
With frmWorkorders.CRInvoice
.DataFiles(0) = App.Path & "\KwikiDat\db2.mdb"
.ReportFileName = App.Path & "\KwikiDat\Invoice.rpt"
.WindowTitle = "Invoice"
.DiscardSavedData = True
.ReplaceSelectionFormula ("{Workorders.PurchaseOrderNumber} =" & "'" & txtFields(4).Text & "'")
.WindowShowExportBtn = True
.Action = 1
End With
Exit Sub
EH:
If MsgBox("The Invoice Cannot Be Viewed At This Time, Please make sure you have setup up your company, Would You like to do this now?", vbYesNo, "Yes") = vbYes Then
frmCompanySetup.Show
ElseIf vbNo Then
Exit Sub
End If
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo UpdateErr:

If Not ValidateFormFields Then Exit Sub

With rsOrder
.Edit
'!WorkorderID = Text2.Text
!CustomerID = dcbCustName.BoundText
!EmployeeID = dcbEmpName.BoundText
!DateReceived = DTPicker1.Value
!DateRequired = DTPicker2.Value
!DateFinished = DTPicker3.Value
!PurchaseOrderNumber = txtFields(4).Text
!SerialNumber = txtFields(10).Text
!PaymentTerms = cmbTerms.Text
!SalesTaxRate = txtFields(11).Text
!ProblemDescription = Text1.Text
!MakeAndModel = txtFields(8).Text
!EstimateFooter = Text4.Text
.Update
.MoveLast
End With

  SetButtons True
  DisableFields
  

'Unload Me
frmProgress.Show
  Exit Sub
UpdateErr:
  MsgBox "All fields marked with the red asterisk are required"
End Sub

Private Sub cmdViewEstimate_Click()
On Error GoTo EH:
SndClick
With frmWorkorders.CRInvoice
.DataFiles(0) = App.Path & "\KwikiDat\db2.mdb"
.ReportFileName = App.Path & "\KwikiDat\Estimate.rpt"
.WindowTitle = "Estimate"
.DiscardSavedData = True
.ReplaceSelectionFormula ("{Workorders.PurchaseOrderNumber} =" & "'" & txtFields(4).Text & "'")
.WindowShowExportBtn = True
.Action = 1
End With
Exit Sub
EH:
If MsgBox("The Estimate Cannot Be Viewed At This Time, Please make sure you have setup up your company, Would You like to do this now?", vbYesNo, "Yes") = vbYes Then
frmCompanySetup.Show
ElseIf vbNo Then
Exit Sub
End If
End Sub

Public Function ImportText()
On Error Resume Next
Close #1
With CD3
.DialogTitle = "Import Text File"
        .CancelError = False
        .Filter = "Text Files (*.txt*)|*.TXT|" _
        & "Rich Text Files (*.rtf*)|*.RTF" _
        & "Microsoft Word (*.doc*)|*.DOC"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Function
        End If
        sFile = .FileName
    End With
  
Open sFile For Input As #1
Text1.Text = Input$(LOF(1), #1)
Close #1
End Function

Public Function EstimateTop()
On Error Resume Next
Close #1
With CD3
.DialogTitle = "Import Text File"
        .CancelError = False
        .Filter = "Text Files (*.txt*)|*.TXT|" _
        & "Rich Text Files (*.rtf*)|*.RTF" _
        & "Microsoft Word (*.doc*)|*.DOC"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Function
        End If
        sFile = .FileName
    End With
Open sFile For Input As #1
txtFields(8).Text = Input$(LOF(1), #1)
Close #1
End Function

Public Function EstimateFoot()
On Error Resume Next
Close #1
With CD3
.DialogTitle = "Import Text File"
        .CancelError = False
        .Filter = "Text Files (*.txt*)|*.TXT|" _
        & "Rich Text Files (*.rtf*)|*.RTF" _
        & "Microsoft Word (*.doc*)|*.DOC"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Function
        End If
        sFile = .FileName
    End With
Open sFile For Input As #1
Text4.Text = Input$(LOF(1), #1)
Close #1
End Function


Private Sub ExportText()
Close #1
With CD3
.DialogTitle = "Export Text File"
        .CancelError = False
        .Filter = "Text Files (*.txt*)|*.TXT|" _
        & "Rich Text Files (*.rtf*)|*.RTF|" _
        & "Microsoft Word (*.doc*)|*.DOC"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
Open sFile For Output As #1
Write #1, Text1.Text
Close #1
End Sub

Private Sub Command1_Click()
ImportText
End Sub

Private Sub Command2_Click()
EstimateTop
End Sub

Private Sub Command3_Click()
EstimateFoot
End Sub

Private Sub Form_Activate()
On Error GoTo Err:
LOrder

If dcbCustName.Text = "" Then
lblStat.Caption = " " & "Pending Order"
lblStat.ForeColor = vbRed
End If

If Text5.Text = "" Then
lblStat.Caption = " " & "Pending Order"
lblStat.ForeColor = vbRed
Else
lblStat.Caption = " " & "Active Order"
lblStat.ForeColor = vbGreen
End If

If Text3.Text = "0" Then
lblStat.Caption = " " & "Order Complete"
lblStat.ForeColor = vbYellow
End If

frmWorkorders.Caption = "  Orders " & "  Viewing Order " & Text2.Text

If Text2.Text = "" Then
cmdEdit.Enabled = False
Else
cmdEdit.Enabled = True
End If
Exit Sub
Err:
End Sub

Private Sub Form_Load()
On Error GoTo Err:
ConnectToDB
sql = "Select * From Workorders"
Set rsOrder = DB.OpenRecordset(sql)

dcbCustName.Refresh
rsEmpName.Refresh
DisableFields
Text3.Text = CCur(frmTree.Label18.Caption)
Text5.Text = frmTree.LvwOrders.SelectedItem.SubItems(8)
Err:
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmTree.UpdateTree
DisconnectFromDB
Set rsOrder = Nothing
Screen.MousePointer = vbDefault
frmTree.TvwCustomer.SetFocus
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  mstrMaintMode = "ADD"
  
  SetButtons False
  cmdUpdate.Visible = False
  EnableFields
  ClearFields
  dcbCustName.SetFocus
  DTPicker1 = Date
  DTPicker2 = Date
  DTPicker3 = Date
  lblStat.Caption = ""
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error Resume Next
  If Text2.Text = "" Then
  MsgBox ("There are no orders to remove")
  Exit Sub
  Else
  If MsgBox("If you have assigned products to this order and are voiding it you will need to remove each item for this order to replenish the stock before deleting this order, Are you sure you want to Remove ? " & "This Order", vbYesNo, "Confirm") = vbYes Then
  ClearFields
  With rsOrder
    .Delete
    .MoveFirst
  End With
  
  lblStat.Caption = ""
  LOrder
  
  frmPayments.Label7.Caption = " Total Amount > " & ""
  frmPayments.Label10.Caption = " Remaining > " & ""
  
  frmTree.Label2.Caption = ""
  frmTree.Label24.Caption = ""
  frmTree.lblName.Caption = ""
  frmTree.Label1.Caption = ""
  frmTree.Label7.Caption = ""
  frmTree.Label8.Caption = ""
  frmTree.Label10.Caption = ""
  frmTree.Label17.Caption = ""
  frmTree.txtWorkorderID.Text = "(Null)"
 
  frmProgress.Show
  
  ElseIf vbNo Then
  Exit Sub
  End If
  End If
End Sub

Private Sub cmdRefresh_Click()
frmProgress.Show
End Sub

Private Sub cmdSave_Click()
  On Error GoTo SaveErr:

    Dim lngIDField      As Long
    Dim strSQL          As String

    If Not ValidateFormFields Then Exit Sub
        
    If mstrMaintMode = "ADD" Then
    
        lngIDField = GetNextOrderID()
        
        strSQL = "INSERT INTO Workorders(  WorkorderID"
        strSQL = strSQL & "            , CustomerID"
        strSQL = strSQL & "            , EmployeeID"
        strSQL = strSQL & "            , DateReceived"
        strSQL = strSQL & "            , DateRequired"
        strSQL = strSQL & "            , DateFinished"
        strSQL = strSQL & "            , PurchaseOrderNumber"
        strSQL = strSQL & "            , SerialNumber"
        strSQL = strSQL & "            , PaymentTerms"
        strSQL = strSQL & "            , SalesTaxRate"
        strSQL = strSQL & "            , ProblemDescription"
        strSQL = strSQL & "            , MakeAndModel"
        strSQL = strSQL & "            , EstimateFooter"
        strSQL = strSQL & "         ) VALUES ("
        strSQL = strSQL & lngIDField
        strSQL = strSQL & ", '" & Replace$(dcbCustName.BoundText, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(dcbEmpName.BoundText, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(DTPicker1.Value, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(DTPicker2.Value, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(DTPicker3.Value, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtFields(4).Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtFields(10).Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(cmbTerms.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Format$(txtFields(11).Text, "0.0;(0.0)") & "'"
        strSQL = strSQL & ", '" & Replace$(Text1.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtFields(8).Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text4.Text, "'", "''") & "'"
        strSQL = strSQL & ")"

    Else
        lngIDField = Text2.Text
        strSQL = "UPDATE Workorders SET "
        strSQL = strSQL & "  CustomerID = '" & Replace$(dcbCustName.BoundText, "'", "''") & "'"
        strSQL = strSQL & ", EmployeeID    = '" & Replace$(dcbEmpName.BoundText, "'", "''") & "'"
        strSQL = strSQL & ", DateReceived     = '" & Replace$(DTPicker1.Value, "'", "''") & "'"
        strSQL = strSQL & ", DateRequired     = '" & Replace$(DTPicker2.Value, "'", "''") & "'"
        strSQL = strSQL & ", DateFinished    = '" & Replace$(DTPicker3.Value, "'", "''") & "'"
        strSQL = strSQL & ", PurchaseOrderNumber    = '" & Replace$(txtFields(4).Text, "'", "''") & "'"
        strSQL = strSQL & ", SerialNumber    = '" & Replace$(txtFields(10).Text, "'", "''") & "'"
        strSQL = strSQL & ", PaymentTerms    = '" & Replace$(cmbTerms.Text, "'", "''") & "'"
        strSQL = strSQL & ", SalesTaxRate    = '" & Format$(txtFields(11).Text, "0.0;(0.0)") & "'"
        strSQL = strSQL & ", ProblemDescription    = '" & Replace$(Text1.Text, "'", "''") & "'"
        strSQL = strSQL & ", MakeAndModel    = '" & Replace$(txtFields(8).Text, "'", "''") & "'"
        strSQL = strSQL & ", EstimateFooter    = '" & Replace$(Text4.Text, "'", "''") & "'"
        strSQL = strSQL & " WHERE WorkorderID = " & lngIDField
    End If
    
    mobjCmd.CommandText = strSQL
    mobjCmd.Execute
    
SetButtons True
DisableFields

'Unload Me
frmTree.UpdateTree
frmTree.FillListView

frmProgress.Show

Exit Sub
SaveErr:
  MsgBox "All fields marked with the red asterisk are required"
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  mstrMaintMode = "EDIT"
  SetButtons False
  cmdSave.Visible = False
  EnableFields
  dcbCustName.SetFocus
  
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
  SetButtons True
  DisableFields
  ClearFields
  rsOrder.CancelUpdate
  LOrder
End Sub

Private Sub cmdClose_Click()
SndClick
Set rsOrder = Nothing
Unload Me
End Sub

'------------------------------------------------------------------------
Private Function GetNextOrderID() As Long
'------------------------------------------------------------------------
    mobjCmd.CommandText = "SELECT MAX(WorkorderID) AS MaxID FROM Workorders"
    Set mobjRst = mobjCmd.Execute

    If mobjRst.EOF Then
        GetNextOrderID = 1
    ElseIf IsNull(mobjRst!MaxID) Then
        GetNextOrderID = 1
    Else
        GetNextOrderID = mobjRst!MaxID + 1
    End If

    Set mobjRst = Nothing
End Function

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdSave.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdAddPart.Enabled = bVal
  cmdAddLabor.Enabled = bVal
  cmdViewEstimate.Enabled = bVal
  cmdPreview.Enabled = bVal
  Command1.Enabled = Not bVal
  Command2.Enabled = Not bVal
  Command3.Enabled = Not bVal
End Sub

Private Sub EnableFields()
  Text1.Locked = False
  cmbTerms.Locked = False
  dcbCustName.Locked = False
  dcbEmpName.Locked = False
  DTPicker1.Enabled = True
  DTPicker2.Enabled = True
  DTPicker3.Enabled = True
  txtFields(4).Locked = False
  txtFields(8).Locked = False
  txtFields(10).Locked = False
  txtFields(11).Locked = False
  Text4.Locked = False
End Sub

Private Sub DisableFields()
  Text1.Locked = True
  cmbTerms.Locked = True
  dcbCustName.Locked = True
  dcbEmpName.Locked = True
  DTPicker1.Enabled = False
  DTPicker2.Enabled = False
  DTPicker3.Enabled = False
  txtFields(4).Locked = True
  txtFields(8).Locked = True
  txtFields(10).Locked = True
  txtFields(11).Locked = True
  Text4.Locked = True
End Sub

Public Sub LOrder()
On Error Resume Next

sql = "Select * From Workorders Where WorkorderID = " & frmTree.txtWorkorderID
Set rsOrder = DB.OpenRecordset(sql)

If Not (IsNull(rsOrder!WorkorderID)) Then Text2.Text = " " & rsOrder!WorkorderID
If Not (IsNull(rsOrder!CustomerID)) Then dcbCustName.BoundText = rsOrder!CustomerID
If Not (IsNull(rsOrder!EmployeeID)) Then dcbEmpName.BoundText = rsOrder!EmployeeID
If Not (IsNull(rsOrder!DateReceived)) Then DTPicker1.Value = rsOrder!DateReceived
If Not (IsNull(rsOrder!DateRequired)) Then DTPicker2.Value = rsOrder!DateRequired
If Not (IsNull(rsOrder!DateFinished)) Then DTPicker3.Value = rsOrder!DateFinished
If Not (IsNull(rsOrder!PurchaseOrderNumber)) Then txtFields(4).Text = rsOrder!PurchaseOrderNumber
If Not (IsNull(rsOrder!SerialNumber)) Then txtFields(10).Text = rsOrder!SerialNumber
If Not (IsNull(rsOrder!PaymentTerms)) Then cmbTerms.Text = rsOrder!PaymentTerms
If Not (IsNull(rsOrder!SalesTaxRate)) Then txtFields(11).Text = Format$(rsOrder!SalesTaxRate, "0.0;(0.0)")
If Not (IsNull(rsOrder!ProblemDescription)) Then Text1.Text = rsOrder!ProblemDescription
If Not (IsNull(rsOrder!MakeAndModel)) Then txtFields(8).Text = rsOrder!MakeAndModel
If Not (IsNull(rsOrder!EstimateFooter)) Then Text4.Text = rsOrder!EstimateFooter

End Sub

Private Sub ClearFields()
  Text1.Text = " "
  Text2.Text = ""
  cmbTerms.Text = ""
  dcbCustName.Text = ""
  dcbEmpName.Text = ""
  txtFields(4).Text = ""
  txtFields(8).Text = " "
  txtFields(10).Text = ""
  txtFields(11).Text = ""
  Text4.Text = " "
End Sub

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
Private Function ValidateFormFields() As Boolean
'-----------------------------------------------------------------------------
    
    If Not ValidateRequiredField(dcbCustName, "Customer Name") Then
    ValidateFormFields = False
    Exit Function
    End If
    
    
    If Not ValidateRequiredField(dcbEmpName, "Employee Name") Then
    ValidateFormFields = False
    Exit Function
    End If
    
    
    If Not ValidateRequiredField(txtFields(4), "PONumber") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtFields(10), "Tracking Number") Then
        ValidateFormFields = False
        Exit Function
    End If
   
    'If Not ValidateRequiredField(PaymentTerms) Then
        'ValidateFormFields = False
        'Exit Function
    'End If
    
    'If Not ValidatePhoneNumber(txtPhone) Then
        'ValidateFormFields = False
        'Exit Function
    'End If
        
    ValidateFormFields = True
    
End Function

