VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPayments 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Payments"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7575
   Icon            =   "frmPayments.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   5880
      TabIndex        =   46
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
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
      Image           =   "frmPayments.frx":000C
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   609
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
      TabCaption(0)   =   "Payments   "
      TabPicture(0)   =   "frmPayments.frx":0A06
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "View Payments"
      TabPicture(1)   =   "frmPayments.frx":0A22
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "LvwPaymentList"
      Tab(1).Control(3)=   "Text10"
      Tab(1).Control(4)=   "Text11"
      Tab(1).ControlCount=   5
      Begin VB.TextBox Text11 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73680
         TabIndex        =   32
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Text10 
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
         Left            =   -69360
         TabIndex        =   30
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7095
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   3360
            TabIndex        =   48
            Top             =   4080
            Visible         =   0   'False
            Width           =   615
         End
         Begin lvButton.lvButtons_H cmdPaySum 
            Height          =   375
            Left            =   4440
            TabIndex        =   45
            Top             =   3960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "&Print Payment Summary"
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
            Image           =   "frmPayments.frx":0A3E
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H Command1 
            Height          =   375
            Left            =   4440
            TabIndex        =   44
            Top             =   3480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "&Add Payment Method"
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
            Image           =   "frmPayments.frx":1918
            cBack           =   -2147483633
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H80000018&
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   4080
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin Crystal.CrystalReport CRViewer 
            Left            =   3480
            Top             =   3360
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
         End
         Begin MSMask.MaskEdBox Text4 
            Height          =   360
            Left            =   1680
            TabIndex        =   27
            Top             =   3600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   12648384
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            DataField       =   "CardholdersName"
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
            Left            =   1680
            TabIndex        =   12
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000018&
            DataField       =   "CreditCardExpDate"
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
            Left            =   1680
            TabIndex        =   11
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000018&
            DataField       =   "CreditCardNumber"
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
            Left            =   1680
            TabIndex        =   10
            Top             =   2880
            Width           =   2415
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H80000018&
            DataField       =   "PaymentDate"
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
            DataField       =   "CheckNumber"
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
            Left            =   1680
            TabIndex        =   8
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H80000018&
            DataField       =   "CreditCardAuthorizationNumber"
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
            Left            =   1680
            TabIndex        =   7
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H80000018&
            DataField       =   "PaymentTime"
            Enabled         =   0   'False
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
            Left            =   1680
            TabIndex        =   6
            Top             =   1440
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcbPaymentMeth 
            Bindings        =   "frmPayments.frx":21F2
            Height          =   360
            Left            =   1680
            TabIndex        =   13
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   -2147483624
            ForeColor       =   -2147483640
            ListField       =   "PaymentMethod"
            BoundColumn     =   "PaymentMethodID"
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
         Begin VB.TextBox Text9 
            DataField       =   "PaymentID"
            Height          =   285
            Left            =   2640
            TabIndex        =   14
            Text            =   "Text9"
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Frame Frame2 
            Height          =   3135
            Left            =   4920
            TabIndex        =   15
            Top             =   240
            Width           =   1935
            Begin lvButton.lvButtons_H cmdAdd 
               Height          =   375
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
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
               Image           =   "frmPayments.frx":2208
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdEdit 
               Height          =   375
               Left            =   120
               TabIndex        =   42
               Top             =   2640
               Width           =   1695
               _ExtentX        =   2990
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
               Image           =   "frmPayments.frx":2AE2
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdRequery 
               Height          =   375
               Left            =   120
               TabIndex        =   41
               Top             =   2160
               Width           =   1695
               _ExtentX        =   2990
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
               Image           =   "frmPayments.frx":33BC
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdDelete 
               Height          =   375
               Left            =   120
               TabIndex        =   40
               Top             =   1680
               Width           =   1695
               _ExtentX        =   2990
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
               Image           =   "frmPayments.frx":3C96
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdCancel 
               Height          =   375
               Left            =   120
               TabIndex        =   39
               Top             =   1200
               Width           =   1695
               _ExtentX        =   2990
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
               Image           =   "frmPayments.frx":4570
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdSave 
               Height          =   375
               Left            =   120
               TabIndex        =   38
               Top             =   720
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               Caption         =   "&Apply"
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
               Image           =   "frmPayments.frx":4E4A
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdUpdate 
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   2640
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
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
               Image           =   "frmPayments.frx":5724
               cBack           =   -2147483633
            End
         End
         Begin VB.Label Label11 
            Caption         =   "Customer Name:"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "CardholdersName:"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   26
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "CreditCardNumber:"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "PaymentAmount:"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   24
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Caption         =   "CreditCardExpDate:"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   23
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "PaymentDate:"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "PaymentMethod:"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Order >"
            Height          =   255
            Left            =   960
            TabIndex        =   20
            Top             =   4080
            Width           =   615
         End
         Begin VB.Label Label4 
            BackColor       =   &H000000FF&
            Caption         =   "  "
            DataField       =   "WorkorderID"
            DataSource      =   "datPrimaryRS"
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
            Height          =   270
            Left            =   1680
            TabIndex        =   19
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Check #"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "CreditCard-Auth #"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Recorded Time:"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Width           =   1815
         End
      End
      Begin MSComctlLib.ListView LvwPaymentList 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16056314
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   35
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Total Payments:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Search By Customer First && Last Name:"
         Height          =   255
         Left            =   -72240
         TabIndex        =   31
         Top             =   4320
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   ">> |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "| <<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   5160
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   4800
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
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\KwikiBilling\KwikiDat\db2.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "[Payment Methods]"
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
End
Attribute VB_Name = "frmPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPayment As Recordset
Dim sqlPayment As String

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwPaymentList.ColumnHeaders. _
Add(, , "WOID", 0, lvwColumnLeft)
Set clmHdr = LvwPaymentList.ColumnHeaders. _
Add(, , "CustomerName", 2100, lvwColumnLeft)
Set clmHdr = LvwPaymentList.ColumnHeaders. _
Add(, , "Date", 1200, lvwColumnLeft)
Set clmHdr = LvwPaymentList.ColumnHeaders. _
Add(, , "Payment Method", 1800, lvwColumnLeft)
Set clmHdr = LvwPaymentList.ColumnHeaders. _
Add(, , "Paid Amount", 1800, lvwColumnLeft)
LvwPaymentList.View = lvwReport
Exit Sub
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errlog
ClearPaymentFields
DisableFields
EnableCont
cmdEdit.Visible = True
cmdUpdate.Visible = False
cmdSave.Enabled = True
rsPayment.CancelUpdate
Exit Sub
errlog:
End Sub

Private Sub cmdEdit_Click()
EnableFields
DisableCont
cmdEdit.Visible = False
cmdUpdate.Visible = True
cmdSave.Enabled = False
Text14.Text = Trim$(CCur(frmTree.Label10.Caption))
End Sub

Private Sub cmdPaySum_Click()
On Error GoTo EH:
SndPlayEx App.Path & "\Sounds\start.wav"

With CRViewer
.DataFiles(0) = App.Path & "\KwikiDat\db2.mdb"
.ReportFileName = App.Path & "\KwikiDat\Payments.rpt"
.WindowTitle = "Payment Summary"
.DiscardSavedData = True
.ReplaceSelectionFormula ("{Workorders.PurchaseOrderNumber} =" & "'" & Text12.Text & "'")
.WindowShowPrintSetupBtn = True
.WindowShowExportBtn = True
.Action = 1
End With
Exit Sub
EH:
If MsgBox("The Summary Cannot Be Viewed At This Time, Please make sure you have setup up your company, Would You like to do this now?", vbYesNo, "Yes") = vbYes Then
frmCompanySetup.Show
ElseIf vbNo Then
Exit Sub
End If
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo EH
If Label4.Caption = "(Null)" Then
MsgBox "No Workorder ID Provided"
Exit Sub
Else
If frmTree.LvwOrders.SelectedItem.SubItems(1) = "" Then
MsgBox "This customer has no amount to apply payment to"
Exit Sub
Else
If Validate = False Then Exit Sub
On Error Resume Next
With rsPayment
    .Edit
    'add record to the table
    '![WorkorderID] = Label4.Caption
    ![PaymentMethodID] = dcbPaymentMeth.BoundText
    ![CustomerName] = Text13.Text
    ![PaymentTime] = Text8.Text
    ![PaymentDate] = Text5.Text
    ![PaymentAmount] = Text4.Text
    ![CheckNumber] = Text6.Text
    ![CreditCardNumber] = Text3.Text
    ![CardholdersName] = Text2.Text
    ![CreditCardExpDate] = Text1.Text
    ![CreditCardAuthorizationNumber] = Text7.Text
    
    .Update
    .MoveLast
    End With
    End If
    End If
    
    AdjStatus
    frmTree.FillListView
    FillPayment
    frmTree.FillInfo
    frmTree.GetTax
    
    If CCur(Text4) < CCur(Text14.Text) Then
    frmTree.PB.Value = frmTree.PB.Value - 10
    ElseIf CCur(Text4) > CCur(Text14.Text) Then
    frmTree.PB.Value = frmTree.PB.Value + 10
    ElseIf CCur(Text4) = CCur(Text14.Text) Then
    frmTree.PB.Value = frmTree.PB.Value
    End If
    
    GetCount
    EnableCont
    DisableFields
    rsPayment.Requery
    LoadPaymentList
    
    cmdEdit.Visible = True
    cmdUpdate.Visible = False
    cmdSave.Enabled = True
    
    Exit Sub
EH: MsgBox ("This customer has no amount to apply payment to")
End Sub

Private Sub cmdRequery_Click()
On Error GoTo errlog
rsPayment.Requery
frmTree.FillListView
LoadPaymentList
Exit Sub
errlog:
End Sub

Private Sub Command1_Click()
SndPlayEx App.Path & "\Sounds\start.wav"
frmPaymentMethod.Show
End Sub

Private Sub dcbPaymentMeth_Change()
'SndPlayEx App.Path & "\Sounds\OpenMenu.wav"
End Sub

Private Sub Form_Load()
GetWorkorderID
dcbPaymentMeth.Refresh

sqlPayment = "SELECT * FROM Payments WHERE WorkorderID = " & Label4.Caption
Set rsPayment = DB.OpenRecordset(sqlPayment)

GetPO
GetCount
DisableFields
setUpListView
LoadPaymentList
ShowTotal
Text13 = Trim$(frmTree.lblName.Caption)
FillPayment

If Text4.Text = "" Then
cmdEdit.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmTree.FillListView

If frmTree.LvwOrders.SelectedItem.SubItems(8) = "" Then
frmTree.Label18.Caption = ""
frmTree.Label8.Caption = " " & "(0.00)"
Else
frmTree.Label8.Caption = " " & frmTree.LvwOrders.SelectedItem.SubItems(8)
End If

frmTree.TvwCustomer.SetFocus
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        rsPayment.MoveFirst
        FillPayment        'calling the subprocedure
        GetPO
    Case 1
        rsPayment.MovePrevious
        If rsPayment.BOF Then
        rsPayment.MoveFirst
        Beep
        End If
        FillPayment
        GetPO
    Case 2
        rsPayment.MoveNext
        If rsPayment.EOF Then
        rsPayment.MoveLast
        Beep
        End If
        FillPayment
        GetPO
    Case 3
        rsPayment.MoveLast
        FillPayment
        GetPO
End Select
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddErr
ClearPaymentFields
EnableFields
DisableCont
Text13 = Trim$(frmTree.lblName.Caption)
Text5 = Date
Text8 = Time
Text14.Text = Trim$(CCur(frmTree.Label10.Caption))
Exit Sub
AddErr:
MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr:
If Text4.Text = "" Then
MsgBox "No record selected to delete"
Exit Sub
End If

If MsgBox("Are you sure you want to Remove ? " & "This Payment", vbYesNo, "Confirm") = vbYes Then

ClearPaymentFields
With rsPayment
.Delete
.MoveNext
End With

rsPayment.Requery
FillPayment
DisableFields
LoadPaymentList

DelStatus
frmTree.FillListView
frmTree.FillInfo
frmTree.GetTax

Label7.Caption = " Total Amount > " & frmTree.Label10.Caption
Label10.Caption = " Remaining > " & frmTree.Label18.Caption
End If
Exit Sub
DeleteErr:
End Sub

Private Sub cmdSave_Click()
On Error GoTo EH:
If Label4.Caption = "(Null)" Then
MsgBox "No Workorder ID Provided"
Exit Sub
Else
If frmTree.LvwOrders.SelectedItem.SubItems(1) = "" Then
MsgBox "This customer has no amount to apply payment to"
Exit Sub
Else
If Validate = False Then Exit Sub
On Error Resume Next
With rsPayment
    .AddNew
    'add record to the table
    ![WorkorderID] = Label4.Caption
    ![PaymentMethodID] = dcbPaymentMeth.BoundText
    ![CustomerName] = Text13.Text
    ![PaymentTime] = Text8.Text
    ![PaymentDate] = Text5.Text
    ![PaymentAmount] = Text4.Text
    ![CheckNumber] = Text6.Text
    ![CreditCardNumber] = Text3.Text
    ![CardholdersName] = Text2.Text
    ![CreditCardExpDate] = Text1.Text
    ![CreditCardAuthorizationNumber] = Text7.Text
    .Update
    .MoveLast
    End With
    End If
End If
    
    SaveStatus
    frmTree.FillListView
    frmTree.FillInfo
    frmTree.GetTax
    
    GetCount
    FillPayment
    EnableCont
    DisableFields
    rsPayment.Requery
    LoadPaymentList
    Exit Sub
EH:
MsgBox Err.Description
'MsgBox ("This customer has no amount to apply payment to")
End Sub

Private Sub cmdClose_Click()
SndClick
Set rsPayment = Nothing
Unload Me
End Sub

Private Sub LvwPaymentList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' sort the listview on the column clicked
    With LvwPaymentList
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
    If Not LvwPaymentList.SelectedItem Is Nothing Then
        LvwPaymentList.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub GetWorkorderID()
Label4.Caption = "  " & frmTree.txtWorkorderID.Text
End Sub

Private Sub GetCount()
On Error Resume Next
Label7.Caption = " Total Amount > " & frmTree.Label10.Caption
Label10.Caption = " Remaining > " & frmTree.Label18.Caption
End Sub

Private Sub FillPayment()
On Error Resume Next
    Text9.Text = rsPayment![PaymentID] & ""
    Label4.Caption = rsPayment![WorkorderID] & ""
    dcbPaymentMeth.BoundText = rsPayment![PaymentMethodID] & ""
    Text13.Text = rsPayment![CustomerName] & ""
    Text8.Text = rsPayment![PaymentTime] & ""
    Text5.Text = rsPayment![PaymentDate] & ""
    Text4.Text = rsPayment![PaymentAmount] & ""
    Text6.Text = rsPayment![CheckNumber] & ""
    Text3.Text = rsPayment![CreditCardNumber] & ""
    Text2.Text = rsPayment![CardholdersName] & ""
    Text1.Text = rsPayment![CreditCardExpDate] & ""
    Text7.Text = rsPayment![CreditCardAuthorizationNumber] & ""
End Sub

Private Sub ClearPaymentFields()
    Text9.Text = ""
    Label4.Caption = frmTree.txtWorkorderID.Text
    dcbPaymentMeth.Text = ""
    Text8.Text = ""
    Text5.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    Text3.Text = ""
    Text2.Text = ""
    Text1.Text = ""
    Text7.Text = ""
End Sub

Private Sub DisableCont()
cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdRequery.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdBrowse(0).Enabled = False
cmdBrowse(1).Enabled = False
cmdBrowse(2).Enabled = False
cmdBrowse(3).Enabled = False
End Sub

Private Sub EnableCont()
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdRequery.Enabled = True
cmdDelete.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdBrowse(0).Enabled = True
cmdBrowse(1).Enabled = True
cmdBrowse(2).Enabled = True
cmdBrowse(3).Enabled = True
End Sub

Private Sub DisableFields()
dcbPaymentMeth.Locked = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Enabled = False
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Enabled = False
End Sub

Private Sub EnableFields()
dcbPaymentMeth.Locked = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Enabled = True
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Enabled = True
End Sub

Public Function IsValidCreditCardNumber(ByVal pCardNumber As String) As Boolean
  Dim CharPos  As Integer
  Dim CheckSum As Integer
  Dim tChar As String

  For CharPos = Len(pCardNumber) To 2 Step -2
    CheckSum = CheckSum + CInt(Mid$(pCardNumber, CharPos, 1))
    tChar = CStr((Mid$(pCardNumber, CharPos - 1, 1)) * 2)
    CheckSum = CheckSum + CInt(Left$(tChar, 1))

    If Len(tChar) > 1 Then CheckSum = CheckSum + CInt(Right$(tChar, 1))
  Next

  If Len(pCardNumber) Mod 2 = 1 Then CheckSum = CheckSum + CInt(Left$(pCardNumber, 1))

  If CheckSum Mod 10 = 0 Then
    IsValidCreditCardNumber = True
  Else
    IsValidCreditCardNumber = False
  End If
End Function

Private Sub SearchList()
On Error Resume Next
Dim itm As ListItem
With LvwPaymentList
Set itm = .FindItem(Text10.Text, lvwSubItem, lvwPartial)
Label8.Caption = "Customer Payment Not Found"
If Not itm Is Nothing Then
Label8.Caption = "Customer Payment Found"
.ListItems(itm.Index).Selected = True
.SetFocus
End If
End With
Set itm = Nothing
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SearchList
End If
End Sub

Private Sub ShowTotal()
On Error Resume Next
Dim i As Integer
Dim cTotal As Currency
With LvwPaymentList
For i = 1 To .ListItems.count
cTotal = cTotal + CCur(.ListItems(i).SubItems(4))
Next
End With
Text11.Text = Format$(cTotal, "$#,##0.00;(#,##0.00)")
End Sub

Public Sub LoadPaymentList()
Dim rsPaymentList As Recordset
Dim sqlList As String
Dim sqlItem As ListItem

LvwPaymentList.ListItems.Clear

'SELECT Payments.WorkorderID, Workorders.CustomerID, Customers.CompanyName, Payments.PaymentDate , Payments.PaymentAmount, [Payment Methods].PaymentMethodID, [Payment Methods].PaymentMethod "
'FROM Payments, Workorders, Customers, [Payment Methods] "
'WHERE Payments.WorkorderID = Workorders.WorkorderID "
'AND Workorders.CustomerID = Customers.CustomerID
'AND Payments.PaymentMethodID = [Payment Methods].PaymentMethodID

sqlList = "SELECT Payments.WorkorderID, Workorders.CustomerID, Customers.CompanyName, Payments.PaymentDate , Payments.PaymentAmount, [Payment Methods].PaymentMethodID, [Payment Methods].PaymentMethod "
sqlList = sqlList & "FROM Payments, Workorders, Customers, [Payment Methods] "
sqlList = sqlList & "WHERE Payments.WorkorderID = Workorders.WorkorderID "
sqlList = sqlList & "AND Workorders.CustomerID = Customers.CustomerID "
sqlList = sqlList & "AND Payments.PaymentMethodID = [Payment Methods].PaymentMethodID "

Set rsPaymentList = DB.OpenRecordset(sqlList)

If (Not OpenDatabase()) Then
MsgBox "Could not connect database"
End If

While Not rsPaymentList.EOF
Set sqlItem = LvwPaymentList.ListItems.Add(, , _
rsPaymentList!WorkorderID)
sqlItem.SubItems(1) = rsPaymentList!CompanyName
sqlItem.SubItems(2) = rsPaymentList!PaymentDate
sqlItem.SubItems(3) = rsPaymentList!PaymentMethod
sqlItem.SubItems(4) = Format$(rsPaymentList!PaymentAmount, "$#,##0.00;(#,##0.00)")
rsPaymentList.MoveNext
Wend
End Sub

Private Sub GetPO()
On Error Resume Next
Dim sqlPO As String
Dim rsPO As Recordset
sqlPO = "Select * From Workorders Where WorkorderID = " & Label4.Caption
Set rsPO = DB.OpenRecordset(sqlPO)
Text12.Text = rsPO!PurchaseOrderNumber
End Sub

Private Function Validate() As Boolean
If dcbPaymentMeth.Text = "" Or Text4.Text = "" Then
MsgBox "You must enter the type of payment and the amount"
Validate = False
Exit Function
End If
Validate = True
End Function

Private Sub SaveStatus()
Dim rsStatus As Recordset
Dim sqlStatus As String
sqlStatus = "Select Workorders.WorkorderID, Workorders.Status FROM Workorders "
sqlStatus = sqlStatus & "WHERE Workorders.WorkorderID = " & frmTree.txtWorkorderID.Text
Set rsStatus = DB.OpenRecordset(sqlStatus)

If CCur(Text4) > CCur(Text14.Text - 10) Then
frmTree.PB.Value = frmTree.PB.Value + 19
ElseIf CCur(Text4) < CCur(Text14.Text - 10) Then
frmTree.PB.Value = frmTree.PB.Value + 6
End If

With rsStatus
.Edit
!Status = frmTree.PB.Value
.Update
End With
'Next
End Sub

Private Sub DelStatus()
Dim rsStatus As Recordset
Dim sqlStatus As String
sqlStatus = "Select Workorders.WorkorderID, Workorders.Status FROM Workorders "
sqlStatus = sqlStatus & "WHERE Workorders.WorkorderID = " & frmTree.txtWorkorderID.Text
Set rsStatus = DB.OpenRecordset(sqlStatus)

frmTree.PB.Value = frmTree.PB.Value - 9
With rsStatus
.Edit
!Status = frmTree.PB.Value
.Update
End With
End Sub

Private Sub AdjStatus()
Dim rsStatus As Recordset
Dim sqlStatus As String
sqlStatus = "Select Workorders.WorkorderID, Workorders.Status FROM Workorders "
sqlStatus = sqlStatus & "WHERE Workorders.WorkorderID = " & frmTree.txtWorkorderID.Text
Set rsStatus = DB.OpenRecordset(sqlStatus)

If CCur(Text4) > CCur(Text14.Text - 10) Then
frmTree.PB.Value = frmTree.PB.Value + 49
ElseIf CCur(Text4) < CCur(Text14.Text - 10) Then
frmTree.PB.Value = frmTree.PB.Value + 6
End If

With rsStatus
.Edit
!Status = frmTree.PB.Value
.Update
End With
End Sub
