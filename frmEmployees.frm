VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEmployees 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Employees"
   ClientHeight    =   5985
   ClientLeft      =   1095
   ClientTop       =   210
   ClientWidth     =   9825
   FillStyle       =   3  'Vertical Line
   ForeColor       =   &H00A2A2A2&
   Icon            =   "frmEmployees.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   8520
      TabIndex        =   98
      Top             =   5520
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmEmployees.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   7200
      TabIndex        =   97
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "frmEmployees.frx":0A06
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   6000
      TabIndex        =   96
      Top             =   5520
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
      Image           =   "frmEmployees.frx":12E0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   2520
      TabIndex        =   93
      Top             =   5520
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
      Image           =   "frmEmployees.frx":1BBA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   2520
      TabIndex        =   94
      Top             =   5520
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
      Image           =   "frmEmployees.frx":2494
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   1320
      TabIndex        =   92
      Top             =   5520
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
      Image           =   "frmEmployees.frx":2D6E
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   91
      Top             =   5520
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
      Image           =   "frmEmployees.frx":3648
      cBack           =   -2147483633
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
      Left            =   5280
      TabIndex        =   22
      Top             =   5520
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
      Left            =   4800
      TabIndex        =   21
      Top             =   5520
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
      Left            =   4320
      TabIndex        =   20
      Top             =   5520
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
      Left            =   3720
      TabIndex        =   19
      Top             =   5520
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   679
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Employees"
      TabPicture(0)   =   "frmEmployees.frx":3F22
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dcbEmp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Employee TimeSheet"
      TabPicture(1)   =   "frmEmployees.frx":3F3E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(1)=   "Label25"
      Tab(1).Control(2)=   "Label28"
      Tab(1).Control(3)=   "Label29"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Employee Profile"
      TabPicture(2)   =   "frmEmployees.frx":3F5A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "imgPhoto"
      Tab(2).Control(2)=   "Label12"
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(5)=   "lblStatus"
      Tab(2).Control(6)=   "cmdUpdateNote"
      Tab(2).Control(7)=   "Text9"
      Tab(2).Control(8)=   "Text1"
      Tab(2).Control(9)=   "cmdAddPhoto"
      Tab(2).Control(10)=   "cmdClearPic"
      Tab(2).Control(11)=   "cmdNotes"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "View Employees"
      TabPicture(3)   =   "frmEmployees.frx":3F76
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(1)=   "LvwAddEmployees"
      Tab(3).Control(2)=   "Text8"
      Tab(3).ControlCount=   3
      Begin lvButton.lvButtons_H cmdNotes 
         Height          =   375
         Left            =   -68160
         TabIndex        =   111
         Top             =   4200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Add Employee Note"
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
      Begin lvButton.lvButtons_H cmdClearPic 
         Height          =   495
         Left            =   -74520
         TabIndex        =   110
         Top             =   4320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "&Clear Picture"
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
      Begin lvButton.lvButtons_H cmdAddPhoto 
         Height          =   495
         Left            =   -74520
         TabIndex        =   109
         Top             =   3720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "&Add \ Change Picture"
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
      Begin VB.Frame Frame2 
         Caption         =   "Add Employee Weekly Hours"
         Height          =   4215
         Left            =   -74760
         TabIndex        =   34
         Top             =   960
         Width           =   8895
         Begin lvButton.lvButtons_H cmdPrintCheck 
            Height          =   375
            Left            =   5280
            TabIndex        =   107
            Top             =   3600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            Caption         =   "&Print Check"
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
            Image           =   "frmEmployees.frx":3F92
            Enabled         =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCancelHours 
            Height          =   375
            Left            =   5280
            TabIndex        =   105
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            Caption         =   "&Cancel Entries"
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
            Image           =   "frmEmployees.frx":4E6C
            Enabled         =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdAddOT 
            Height          =   375
            Left            =   5280
            TabIndex        =   103
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            Caption         =   "&Adjust Hours"
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
            Image           =   "frmEmployees.frx":5746
            Enabled         =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdEditHours 
            Height          =   375
            Left            =   2520
            TabIndex        =   101
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            Caption         =   "&Change Hours"
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
            Image           =   "frmEmployees.frx":6020
            Enabled         =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdEnterHours 
            Height          =   375
            Left            =   120
            TabIndex        =   99
            Top             =   960
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Caption         =   "&Add New Time Sheet"
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
            Image           =   "frmEmployees.frx":68FA
            Enabled         =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.TextBox Text28 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   300
            Left            =   7320
            TabIndex        =   81
            Text            =   "0.00"
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Frame Frame3 
            Caption         =   "Payroll Deductions"
            Height          =   2055
            Left            =   240
            TabIndex        =   70
            Top             =   2040
            Width           =   4695
            Begin lvButton.lvButtons_H cmdAddDed 
               Height          =   375
               Left            =   2640
               TabIndex        =   108
               Top             =   1560
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               Caption         =   "&Edit Deductions"
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
               Image           =   "frmEmployees.frx":71D4
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin VB.TextBox Text32 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   89
               Text            =   "0"
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox Text31 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   84
               Text            =   "0.00"
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox Text30 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   83
               Text            =   "0.00"
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox Text29 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   82
               Text            =   "0.00"
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox Text27 
               Appearance      =   0  'Flat
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
               Height          =   285
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   79
               Text            =   "0.00"
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox Text26 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   77
               Text            =   "0.00"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text25 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   75
               Text            =   "0.00"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text24 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   73
               Text            =   "0.00"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text23 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   71
               Text            =   "0.00"
               Top             =   480
               Width           =   1095
            End
            Begin lvButton.lvButtons_H cmdSaveDed 
               Height          =   375
               Left            =   2640
               TabIndex        =   106
               Top             =   1560
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               Caption         =   "&Save Changes"
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
               Image           =   "frmEmployees.frx":7AAE
               cBack           =   -2147483633
            End
            Begin VB.CommandButton cmdUpdateDed 
               Caption         =   "&Enter Deductions"
               Enabled         =   0   'False
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
               Left            =   2640
               TabIndex        =   86
               Top             =   1680
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox Text33 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2400
               TabIndex        =   113
               Text            =   "Custom1"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox Text34 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1320
               TabIndex        =   114
               Text            =   "Custom"
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label40 
               Caption         =   "Exemptions:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label36 
               Caption         =   "Advance"
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label35 
               Caption         =   "Total Ded"
               Height          =   255
               Left            =   3480
               TabIndex        =   80
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label34 
               Caption         =   "FICA"
               Height          =   255
               Left            =   3600
               TabIndex        =   78
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label33 
               Caption         =   "SSI"
               Height          =   255
               Left            =   2520
               TabIndex        =   76
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label32 
               Caption         =   "Federal Tax"
               Height          =   255
               Left            =   1320
               TabIndex        =   74
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label31 
               Caption         =   "State Tax"
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear Time Fields"
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
            Left            =   6720
            TabIndex        =   69
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   7320
            TabIndex        =   64
            Text            =   "0.00"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   7320
            TabIndex        =   62
            Text            =   "0"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "0"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7320
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   36
            Text            =   "0"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Height          =   285
            Left            =   7320
            TabIndex        =   35
            Text            =   "0"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdEnterTotal 
            Caption         =   "&Enter Hours"
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
            Left            =   3840
            TabIndex        =   38
            Top             =   1080
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Left            =   7320
            TabIndex        =   61
            Text            =   "0.00"
            Top             =   2880
            Width           =   1455
         End
         Begin lvButton.lvButtons_H cmdAddNewTime 
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   960
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Caption         =   "Enter New Time Sheet"
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
            Image           =   "frmEmployees.frx":8388
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdUpdateHours 
            Height          =   375
            Left            =   2520
            TabIndex        =   102
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "frmEmployees.frx":866F
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdAdjustOT 
            Height          =   375
            Left            =   5280
            TabIndex        =   104
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            Image           =   "frmEmployees.frx":8956
            cBack           =   -2147483633
         End
         Begin VB.Label Label39 
            Caption         =   "Net Pay"
            Height          =   255
            Left            =   7320
            TabIndex        =   87
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            X1              =   7200
            X2              =   7200
            Y1              =   120
            Y2              =   4200
         End
         Begin VB.Label Label30 
            Caption         =   "If this employee was newly added you must add a new timesheet for this employee by clicking Add New Time Sheet Button"
            Height          =   495
            Left            =   240
            TabIndex        =   68
            Top             =   1440
            Width           =   4695
         End
         Begin VB.Label Label27 
            Caption         =   "Regular Pay / 40hrs"
            Height          =   255
            Left            =   7320
            TabIndex        =   65
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "Regular Hours"
            Height          =   255
            Left            =   7320
            TabIndex        =   63
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Monday"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Tuesday"
            Height          =   255
            Left            =   840
            TabIndex        =   56
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Wednesday"
            Height          =   255
            Left            =   1560
            TabIndex        =   55
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "  Thursday"
            Height          =   255
            Left            =   2520
            TabIndex        =   54
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Friday"
            Height          =   255
            Left            =   3480
            TabIndex        =   53
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Saturday"
            Height          =   255
            Left            =   4200
            TabIndex        =   52
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Sunday"
            Height          =   255
            Left            =   5040
            TabIndex        =   51
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Total Hours"
            Height          =   255
            Left            =   5880
            TabIndex        =   50
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Gross Pay"
            Height          =   255
            Left            =   7320
            TabIndex        =   49
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Overtime Hours"
            Height          =   255
            Left            =   7320
            TabIndex        =   48
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "Over Time Pay"
            Height          =   255
            Left            =   7320
            TabIndex        =   47
            Top             =   2040
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -71640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   1320
         Width           =   5655
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -71640
         TabIndex        =   27
         Text            =   "C:\Program Files\KwikiBilling\KwikiDat\BitMaps\Default.jpg"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text8 
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
         Left            =   -67920
         TabIndex        =   25
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   240
         TabIndex        =   1
         Top             =   825
         Width           =   9135
         Begin MSMask.MaskEdBox Text5 
            DataField       =   "BillingRate"
            DataSource      =   "rsEmp"
            Height          =   375
            Left            =   2160
            TabIndex        =   18
            Top             =   2520
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483624
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   7560
            Top             =   1200
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H80000018&
            DataField       =   "EmployeeID"
            DataSource      =   "rsEmp"
            Height          =   285
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H80000018&
            DataField       =   "ContactPhone"
            DataSource      =   "rsEmp"
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
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            DataField       =   "FullName"
            DataSource      =   "rsEmp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   840
            Width           =   3135
         End
         Begin MSMask.MaskEdBox mskSSN 
            DataField       =   "SSNNumber"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "   -   -"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "rsEmp"
            Height          =   375
            Left            =   6240
            TabIndex        =   12
            Top             =   2520
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483624
            Enabled         =   0   'False
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###-##-####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000018&
            DataField       =   "Address"
            DataSource      =   "rsEmp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2160
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   1320
            Width           =   3135
         End
         Begin VB.ComboBox cmbTitle 
            BackColor       =   &H80000018&
            DataField       =   "Title"
            DataSource      =   "rsEmp"
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
            ItemData        =   "frmEmployees.frx":8C3D
            Left            =   6120
            List            =   "frmEmployees.frx":8C4A
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtpHireDate 
            DataField       =   "HireDate"
            DataSource      =   "rsEmp"
            Height          =   375
            Left            =   2160
            TabIndex        =   2
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483634
            CheckBox        =   -1  'True
            Format          =   45219841
            CurrentDate     =   39661
         End
         Begin VB.Label Label41 
            Height          =   255
            Left            =   360
            TabIndex        =   90
            Top             =   3240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Enter SSN # (No Dashes)"
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
            Left            =   6240
            TabIndex        =   17
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "City\State\Zip:"
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
            Left            =   360
            TabIndex        =   14
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "SSN Number:"
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
            Left            =   4800
            TabIndex        =   11
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Street Address:"
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
            Left            =   360
            TabIndex        =   10
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Title:"
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
            Left            =   5520
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Full Name:"
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
            Left            =   360
            TabIndex        =   7
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Caption         =   "HomePhone:"
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
            Index           =   4
            Left            =   360
            TabIndex        =   6
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Caption         =   "SalaryRate:"
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
            Index           =   6
            Left            =   360
            TabIndex        =   5
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "HireDate:"
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
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView LvwAddEmployees 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   23
         Top             =   960
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
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
      Begin MSDataListLib.DataCombo dcbEmp 
         Bindings        =   "frmEmployees.frx":8C5F
         CausesValidation=   0   'False
         DataField       =   "EmployeeID"
         DataSource      =   "rsEmp"
         Height          =   360
         Left            =   6840
         TabIndex        =   33
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "FullName"
         BoundColumn     =   "EmployeeID"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin lvButton.lvButtons_H cmdUpdateNote 
         Height          =   375
         Left            =   -68160
         TabIndex        =   112
         Top             =   4200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Save Employee Note"
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
      Begin VB.Label Label29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72240
         TabIndex        =   67
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "Employees Salary Rate >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   66
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label25 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68160
         TabIndex        =   60
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label24 
         Caption         =   "TimeSheet For Employee >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70680
         TabIndex        =   59
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   -71640
         TabIndex        =   58
         Top             =   4200
         Width           =   4695
      End
      Begin VB.Label Label11 
         Caption         =   "Select Employee:"
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Employee Notes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71640
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69600
         TabIndex        =   29
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label12 
         Caption         =   "Viewing Employee >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71640
         TabIndex        =   28
         Top             =   600
         Width           =   1935
      End
      Begin VB.Image imgPhoto 
         Height          =   2175
         Left            =   -74520
         Picture         =   "frmEmployees.frx":8C73
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Employee Photo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   26
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Search By Employee Name:"
         Height          =   255
         Left            =   -71160
         TabIndex        =   24
         Top             =   600
         Width           =   3015
      End
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   375
      Left            =   2520
      TabIndex        =   95
      Top             =   5520
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
      Image           =   "frmEmployees.frx":8EF6
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents HO As cBinaryDBObject
Attribute HO.VB_VarHelpID = -1

Private mobjConn As ADODB.Connection
Private mobjCmd  As ADODB.Command
Private mobjRst  As ADODB.Recordset

Private mstrMaintMode As String
Private mblnUpdateInProgress    As Boolean
Private mblnFormActivated       As Boolean

Private WithEvents rsEmp As ADODB.Recordset
Attribute rsEmp.VB_VarHelpID = -1
Private CNN As ADODB.Connection

Dim iniFile As String
Dim sFileName As String
Dim Clear As String

Private Sub cmdAddDed_Click()
cmdSaveDed.Visible = True
cmdAddDed.Visible = False
'-------------------------
EnableDEDFields
End Sub

Private Sub cmdAddPhoto_Click()
On Error Resume Next
With CD1
.DialogTitle = "Add Photo"
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
Text9.Text = (sFileName)
End Sub

Private Sub cmdAdjustOT_Click()
AdjustOT
cmdAddOT.Visible = True
cmdAdjustOT.Visible = False
GetWages
ShowOT
ShowDED
If Text18.Text = "0" Then
Text17.Visible = False
Text20.Visible = True
Else
Text17.Visible = True
Text20.Visible = False
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
FillFields
dcbEmp.ReFill
rsEmp.MoveFirst

cmdEdit.Visible = True
cmdSave.Visible = False
cmdAdd.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdRefresh.Enabled = True
cmdClose.Enabled = True
cmdUpdate.Visible = False
cmdClearPic.Enabled = False
cmdAddPhoto.Enabled = False
cmdNotes.Enabled = False

cmdBrowse(0).Enabled = True
cmdBrowse(1).Enabled = True
cmdBrowse(2).Enabled = True
cmdBrowse(3).Enabled = True
dcbEmp.Enabled = True
DisableFields
Frame2.Enabled = True


DisableFields
GetTime
mblnUpdateInProgress = False
End Sub

Private Sub cmdCancelHours_Click()
cmdAddOT.Visible = True
cmdAdjustOT.Visible = False

cmdEnterHours.Visible = True
cmdAddNewTime.Visible = False
cmdEditHours.Enabled = True
cmdEnterHours.Enabled = False
cmdEditHours.Visible = True
cmdUpdateHours.Visible = False
cmdEnterTotal.Enabled = True
cmdSaveDed.Visible = False
cmdAddDed.Visible = True

DisableTimeFields
DisableDEDFields
GetTime
ShowOT
ShowDED
End Sub

Private Sub cmdClear_Click()
ClearHourFields
ClearDedFields
End Sub

Private Sub cmdClearPic_Click()
On Error GoTo EH:
sFileName = App.Path & "\KwikiDat\BitMaps\Default.jpg"

imgPhoto.Picture = LoadPicture(sFileName)

Exit Sub
EH: MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
On Error GoTo EditErr

If dcbEmp.Text = "" Then
MsgBox "Select an employee from the dropdown box to edit"
Exit Sub
End If

cmdAdd.Enabled = False
cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdRefresh.Enabled = False
cmdClose.Enabled = False
cmdEdit.Visible = False
cmdSave.Visible = False
cmdUpdate.Visible = True
cmdClearPic.Enabled = True
cmdNotes.Enabled = True
cmdAddPhoto.Enabled = True
cmdBrowse(0).Enabled = False
cmdBrowse(1).Enabled = False
cmdBrowse(2).Enabled = False
cmdBrowse(3).Enabled = False
dcbEmp.Enabled = False
EnableFields
EditErr:
End Sub

Private Sub cmdNotes_Click()
Text1.Locked = False
cmdNotes.Visible = False
cmdUpdateNote.Visible = True
End Sub

Private Sub cmdPrintCheck_Click()
SndPlayEx App.Path & "\Sounds\start.wav"
Me.Hide
frmCheck.cmdProcCheck.Visible = True
frmCheck.cmdClose.Visible = True
frmCheck.Show
End Sub

Private Sub cmdSave_Click()
On Error GoTo EH
Dim lngIDField As Long
Dim strSQL As String


    If mstrMaintMode = "ADD" Then
        lngIDField = GetNextEmpID()
        
        strSQL = "INSERT INTO Employees(  EmployeeID"
        strSQL = strSQL & "            , Notes"
        strSQL = strSQL & "            , FullName"
        strSQL = strSQL & "            , Address"
        strSQL = strSQL & "            , ContactPhone"
        strSQL = strSQL & "            , BillingRate"
        strSQL = strSQL & "            , SSNNumber"
        strSQL = strSQL & "            , HireDate"
        strSQL = strSQL & "            , Title"
        strSQL = strSQL & "            , FileName"
        strSQL = strSQL & "            , Photo"
        strSQL = strSQL & "         ) VALUES ("
        strSQL = strSQL & lngIDField
        strSQL = strSQL & ", '" & Replace$(Text1.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text2.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text3.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text4.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text5.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(mskSSN.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(dtpHireDate.Value, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(cmbTitle.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(Text9.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(imgPhoto.Picture = imgPhoto.Picture, "'", "''") & "'"
        strSQL = strSQL & ")"
    
    mobjCmd.CommandText = strSQL
    mobjCmd.Execute
    End If

    Frame2.Enabled = True
    EnableCont
    DisableFields
    mblnUpdateInProgress = False
    
Unload Me
frmEmpProg.Show

'Else
'lngIDField = Text7.Text
'strSQL = "UPDATE Employees SET "
'strSQL = strSQL & "  Notes    = '" & Replace$(Text1.Text, "'", "''") & "'"
'strSQL = strSQL & ", FirstName    = '" & Replace$(Text2.Text, "'", "''") & "'"
'strSQL = strSQL & ", Address     = '" & Replace$(Text3.Text, "'", "''") & "'"
'strSQL = strSQL & ", ContactPhone     = '" & Replace$(Text4.Text, "'", "''") & "'"
'strSQL = strSQL & ", BillingRate    = '" & Replace$(Text5.Text, "'", "''") & "'"
'strSQL = strSQL & ", LastName    = '" & Replace$(Text6.Text, "'", "''") & "'"
'strSQL = strSQL & ", SSNNumber    = '" & Replace$(mskSSN.Text, "'", "''") & "'"
'strSQL = strSQL & ", HireDate    = '" & Replace$(dtpHireDate.Value, "'", "''") & "'"
'strSQL = strSQL & ", Title    = '" & Replace$(cmbTitle.Text, "'", "''") & "'"
'strSQL = strSQL & " WHERE PartID = " & lngIDField
'SaveBinaryObject
Exit Sub
EH:
Label41.Caption = "This SSN # Already Exist In The System"
frmEmpErr.Show
End Sub

Private Function GetNextEmpID() As Long
'------------------------------------------------------------------------

    mobjCmd.CommandText = "SELECT MAX(EmployeeID) AS MaxID FROM Employees"
    Set mobjRst = mobjCmd.Execute

    If mobjRst.EOF Then
        GetNextEmpID = 1
    ElseIf IsNull(mobjRst!MaxID) Then
        GetNextEmpID = 1
    Else
        GetNextEmpID = mobjRst!MaxID + 1
    End If

    Set mobjRst = Nothing

End Function

Private Sub cmdSaveDed_Click()
EditDed
'-------------------
cmdSaveDed.Visible = False
cmdAddDed.Visible = True

DisableDEDFields

End Sub

Private Sub cmdUpdate_Click()
Dim rs As Recordset
Dim sql As String
On Error GoTo EH:
'rsEmp.UpdateBatch adAffectAll

If dcbEmp.Text = "" Then
MsgBox "You must select the employee to edit from the dropdown list first"
rs.CancelUpdate
Exit Sub
Else

sql = "Select * From Employees Where EmployeeID = " & Text7.Text
Set rs = DB.OpenRecordset(sql)

With rs
.Edit
rs!FullName = Text2.Text
rs!Title = cmbTitle.Text
rs!Address = Text3.Text
rs!Address = Text3.Text
rs!ContactPhone = Text4.Text
rs!BillingRate = Text5.Text
rs!SSNNumber = mskSSN.Text
.Update
End With


SaveBinaryObject

EnableCont
DisableFields
End If

frmEmpProg.Show
Exit Sub
EH:
Label41.Caption = Err.Description

'Label41.Caption = "An Error Occured Updating Employee Record..."
FillFields
frmEmpErr.Show
End Sub

Private Sub cmdUpdateDed_Click()
AddDed
End Sub

Private Sub cmdUpdateNote_Click()
SaveNote
cmdNotes.Visible = True
cmdUpdateNote.Visible = False
End Sub

Private Sub dcbEmp_Change()
On Error Resume Next
'SndPlayEx App.Path & "\Sounds\OpenMenu.wav"

If Not rsEmp.BOF Then rsEmp.MoveFirst
rsEmp.Find "EmployeeID = " & dcbEmp.BoundText, 0, adSearchForward, 0

ChangeFields
imgPhoto.Picture = LoadPicture(rsEmp!FileName)
Label10 = Text2.Text
Label25 = Text2.Text
Label29 = Format$(Text5.Text, "$#,##0.00;(#,##0.00)")
GetTime
GetWages
ShowOT
ShowNote
ShowDED
cmdAddDed.Enabled = True
cmdEnterHours.Enabled = True
cmdEditHours.Enabled = True
cmdAddOT.Enabled = True
cmdCancelHours.Enabled = True
cmdPrintCheck.Enabled = True
End Sub

Private Sub Form_Activate()
If mblnFormActivated Then Exit Sub
Refresh
mblnFormActivated = True
ShowDED
cmdEnterHours.Enabled = False
End Sub


Private Sub Form_Load()
'On Error Resume Next
ConnectToDB

'----------------------------------------------
Set CNN = New ADODB.Connection
CNN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\KwikiDat\db2.mdb"
CNN.Open
  
Set rsEmp = New ADODB.Recordset
rsEmp.Open "Select * from Employees", CNN, adOpenStatic, adLockOptimistic

Set dcbEmp.RowSource = rsEmp
dcbEmp.ListField = "FullName"
dcbEmp.BoundColumn = "EmployeeID"
'----------------------------------------------

'FillFields
DisableFields
cmbTitle.ListIndex = 0
setUpListView
LoadEmpList

If LvwAddEmployees.SelectedItem Is Nothing Then
dcbEmp.Enabled = False
Else
dcbEmp.Enabled = True
End If

Label25 = Text2.Text
Label29 = Format$(Text5.Text, "$#,##0.00;(#,##0.00)")

If Text18.Text > 0 Then
Text17.Visible = True
Text20.Visible = False
Else
Text17.Visible = False
Text20.Visible = True
End If

'GetWages
'GetTime
'ShowOT
End Sub

Private Sub ConnectToDB()
'-----------------------------------------------------------------------------

    Set mobjConn = New ADODB.Connection
    mobjConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\KwikiDat\db2.mdb;Persist Security Info=False"
    mobjConn.Open

    Set mobjCmd = New ADODB.Command
    Set mobjCmd.ActiveConnection = mobjConn
    mobjCmd.CommandType = adCmdText

End Sub

Private Sub DisconnectFromDB()
    Set mobjCmd = Nothing
    mobjConn.Close
    Set mobjConn = Nothing
End Sub

Private Sub HO_Error(ID As Long, Msg As String)
MsgBox ID & ":  " & Msg
End Sub

Private Sub HO_Status(ID As Long, Msg As String)
lblStatus.Caption = CStr(ID) & ":  " & Msg
Exit Sub
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        rsEmp.MoveFirst
        Beep
        FillFields        'call subprocedure
        Label10 = Text2.Text
        Label25 = Text2.Text
        Label29 = Format$(Text5.Text, "$#,##0.00;(#,##0.00)")
        GetTime
        GetWages
        ShowOT
        ShowNote
        ShowDED
    Case 1
        rsEmp.MovePrevious
        If rsEmp.BOF Then
        rsEmp.MoveFirst
        End If
        FillFields
        Label10 = Text2.Text
        Label25 = Text2.Text
        Label29 = Format$(Text5.Text, "$#,##0.00;(#,##0.00)")
        GetTime
        GetWages
        ShowOT
        ShowNote
        ShowDED
    Case 2
        rsEmp.MoveNext
        If rsEmp.EOF Then
        rsEmp.MoveLast
        End If
        FillFields
        Label10 = Text2.Text
        Label25 = Text2.Text
        Label29 = Format$(Text5.Text, "$#,##0.00;(#,##0.00)")
        GetTime
        GetWages
        ShowOT
        ShowNote
        ShowDED
    Case 3
        rsEmp.MoveLast
        Beep
        FillFields
        Label10 = Text2.Text
        Label25 = Text2.Text
        Label29 = Format$(Text5.Text, "$#,##0.00;(#,##0.00)")
        GetTime
        GetWages
        ShowOT
        ShowNote
        ShowDED
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
If mblnUpdateInProgress Then
MsgBox "You must save or cancel the current action before " _
& "closing this window.", _
vbInformation, _
"Cannot Close"
Cancel = 1
Exit Sub
End If
frmWorkorders.rsEmpName.Refresh
DisconnectFromDB
frmTree.TvwCustomer.SetFocus
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddErr
mstrMaintMode = "ADD"
mblnUpdateInProgress = True
DisableCont
EnableFields
ClearFields
dtpHireDate = Date
Frame2.Enabled = False
ClearHourFields
ClearDedFields
Exit Sub
AddErr:
MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
If Text7.Text = "" Then
MsgBox "No employee selected to delete"
Else
If MsgBox("Are you sure you want to Remove ? " & Text2.Text, vbYesNo, "Confirm") = vbYes Then

mobjCmd.CommandText = "DELETE FROM Employees WHERE EmployeeID = " & Text7.Text
mobjCmd.Execute
'ClearHourFields
'ChangeTime

Unload Me
frmEmpProg.Show

End If
End If
Exit Sub
DeleteErr:
MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo RefreshErr:
frmEmpProg.Show

rsEmp.Requery
mobjRst.Requery

LoadEmpList
Exit Sub
RefreshErr:
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
SndClick
frmWorkorders.rsEmpName.Refresh
Set rsEmp = Nothing
Unload Me
End Sub

Private Sub DisableCont()
cmdAdd.Enabled = False
cmdCancel.Enabled = True
cmdEdit.Visible = False
cmdSave.Visible = True
cmdDelete.Enabled = False
cmdRefresh.Enabled = False
cmdClose.Enabled = False
cmdUpdate.Visible = False
cmdNotes.Enabled = True
cmdAddPhoto.Enabled = False
dcbEmp.Enabled = False
cmdBrowse(0).Enabled = False
cmdBrowse(1).Enabled = False
cmdBrowse(2).Enabled = False
cmdBrowse(3).Enabled = False
End Sub

Private Sub EnableCont()
cmdAdd.Enabled = True
cmdCancel.Enabled = False
cmdEdit.Visible = True
cmdEdit.Enabled = True
cmdSave.Visible = False
cmdDelete.Enabled = True
cmdRefresh.Enabled = True
cmdClose.Enabled = True
cmdClearPic.Enabled = False
cmdAddPhoto.Enabled = False
cmdNotes.Enabled = False
dcbEmp.Enabled = True
cmdBrowse(0).Enabled = True
cmdBrowse(1).Enabled = True
cmdBrowse(2).Enabled = True
cmdBrowse(3).Enabled = True
End Sub

Private Sub DisableFields()
cmbTitle.Locked = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Enabled = False
Text7.Locked = True ' Auto inc
imgPhoto.Enabled = False
cmbTitle.Enabled = False
mskSSN.Enabled = False
dtpHireDate.Enabled = False
End Sub

Private Sub EnableFields()
cmbTitle.Locked = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Enabled = True
Text7.Locked = True ' Auto inc
imgPhoto.Enabled = True
cmbTitle.Enabled = True
mskSSN.Enabled = True
dtpHireDate.Enabled = True
End Sub

Private Sub EnableDEDFields()
Text23.Locked = False
Text24.Locked = False
Text25.Locked = False
Text26.Locked = False
Text29.Locked = False
Text30.Locked = False
Text31.Locked = False
Text32.Locked = False
End Sub

Private Sub DisableDEDFields()
Text23.Locked = True
Text24.Locked = True
Text25.Locked = True
Text26.Locked = True
Text29.Locked = True
Text30.Locked = True
Text31.Locked = True
Text32.Locked = True
End Sub

Private Sub FillFields()
On Error GoTo EH:
iniFile = App.Path & "\Settings.ini"
Text1.Text = rsEmp!Notes
Text2.Text = rsEmp!FullName
Text3.Text = rsEmp!Address
Text4.Text = rsEmp!ContactPhone
Text5.Text = rsEmp!BillingRate
Text7.Text = rsEmp!EmployeeID
mskSSN.Text = rsEmp!SSNNumber
dtpHireDate.Value = rsEmp!HireDate
cmbTitle.Text = rsEmp!Title
imgPhoto.Picture = LoadPicture(rsEmp!FileName)
Label10 = Text2.Text


Text33.Text = ReadINI("Employees", "CustomDED", iniFile)
Text34.Text = ReadINI("Employees", "CustomDED1", iniFile)

Exit Sub
EH:
End Sub

Private Sub ChangeFields()
On Error Resume Next
iniFile = App.Path & "\Settings.ini"
Set Text1.DataSource = rsEmp
Set Text2.DataSource = rsEmp
Set Text3.DataSource = rsEmp
Set Text4.DataSource = rsEmp
Set Text5.DataSource = rsEmp
Set Text7.DataSource = rsEmp
Set mskSSN.DataSource = rsEmp
Set dtpHireDate.DataSource = rsEmp
Set cmbTitle.DataSource = rsEmp
Set imgPhoto.DataSource = rsEmp

Text33.Text = ReadINI("Employees", "CustomDED", iniFile)
Text34.Text = ReadINI("Employees", "CustomDED1", iniFile)
End Sub

Private Sub ClearFields()
Text1.Text = Clear & " "
Text2.Text = Clear
Text3.Text = "Required"
Text4.Text = "0"
Text5.Text = "0"
Text7.Text = Clear
mskSSN.Text = "0"
dtpHireDate.Value = Date
cmbTitle.Text = "Mr."
imgPhoto.Picture = LoadPicture(Clear)

End Sub

Private Sub LoadDefaultPic() '-- For Clearing --
On Error Resume Next
sFileName = App.Path & "\KwikiDat\BitMaps\Default.jpg"
imgPhoto.Picture = LoadPicture(sFileName)
SaveBinaryObject
End Sub

Private Sub GetBinaryObject()
OpenDatabase
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim fn As String                       'Binary file name to use as storage
Dim i As Integer

    If Text7.Text = "" Then
        Set HO = New cBinaryDBObject       'create the new bd object

        FieldNames(0) = "EmployeeID"               'return the ID field
        FieldNames(1) = "FileName"         'return the filename

        With HO
            .KillFile = True                        'kill the filename if it exists
             Set .DB = DB                        'pass the database
            .ObjectKeyFieldName = "EmployeeID"      'the key/index field is
            .ObjectKey = Text7.Text                 'the value to search for is
            .ObjectFieldName = "Photo"              'name of the field that contains the binary file
            .ObjectTableName = "Employees"          'table that contains the binary files
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
    DB.Close
End Sub

Private Sub SaveBinaryObject()
OpenDatabase
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim FieldData(1) As Variant            'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim fn As String                       'Binary file name to use as storage
Dim i As Integer

    If sFileName = "" Then
    Exit Sub
    End If

    Set HO = New cBinaryDBObject         'create the new bd object

     FieldNames(0) = "EmployeeID"        'return the ID field
     FieldNames(1) = "FileName"          'return the filename
     FieldData(0) = Null                 'return the ID field
     FieldData(1) = sFileName            'return the filename

    With HO
        .KillFile = False                       'kill the filename if it exists
         Set .DB = DB                         'pass the database
        .ObjectKeyFieldName = "EmployeeID"    'the key/index field is
        .ObjectKey = Text7.Text               'the value to search for is
        .ObjectFieldName = "Photo"           'name of the field that contains the binary file
        .ObjectTableName = "Employees"        'table that contains the binary files
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
    DB.Close
    
End Sub

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "EmpID", 0, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "FullName", 1800, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "W", 0, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "Address", 2200, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "DateHired", 1800, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "ContactPhone", 1500, lvwColumnLeft)
Set clmHdr = LvwAddEmployees.ColumnHeaders. _
Add(, , "SalaryRate", 1500, lvwColumnLeft)


LvwAddEmployees.View = lvwReport
End Sub

Public Sub LoadEmpList()
'Dim sqlstring As String
Dim sqlEmp As ListItem

LvwAddEmployees.ListItems.Clear

If (rsEmp.RecordCount > 0) Then
rsEmp.MoveFirst
End If
On Error Resume Next
While Not rsEmp.EOF
Set sqlEmp = LvwAddEmployees.ListItems.Add(, , _
rsEmp!EmployeeID)
sqlEmp.SubItems(1) = rsEmp!FullName
sqlEmp.SubItems(2) = rsEmp!LastName
sqlEmp.SubItems(3) = rsEmp!Address
sqlEmp.SubItems(4) = rsEmp!HireDate
sqlEmp.SubItems(5) = rsEmp!ContactPhone
sqlEmp.SubItems(6) = Format$(rsEmp!BillingRate, "$#,##0.00;(#,##0.00)")
rsEmp.MoveNext
Wend
End Sub

Private Sub SearchList()
On Error Resume Next
Dim itm As ListItem
With LvwAddEmployees
Set itm = .FindItem(Text8.Text, lvwSubItem, lvwPartial)
Label7.Caption = "Searched Employee Not Found"
If Not itm Is Nothing Then
Label7.Caption = "Searched Employee Found"
.ListItems(itm.Index).Selected = True
.SetFocus
End If
End With
Set itm = Nothing
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

Private Sub Text33_Change()
iniFile = App.Path & "\Settings.ini"
WriteINI "Employees", "CustomDED", Text33.Text, iniFile
End Sub

Private Sub Text34_Change()
iniFile = App.Path & "\Settings.ini"
WriteINI "Employees", "CustomDED1", Text34.Text, iniFile
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SearchList
End If
End Sub

'---------------------------------------------------------------------------
'EMPLOYEES TIME SECTION
'---------------------------------------------------------------------------

Private Sub cmdAddNewTime_Click()
cmdEnterHours.Visible = True
cmdAddNewTime.Visible = False
cmdEditHours.Enabled = True
cmdEnterTotal.Enabled = True
cmdEnterHours.Enabled = False
cmdEditHours.Enabled = True
cmdAddDed.Enabled = True

Call cmdEnterTotal_Click
AddTime
AddDed

GetTime
GetWages

DisableTimeFields
End Sub

Private Sub cmdAddOT_Click()
EnableTimeFields
Text21.SetFocus
cmdAddOT.Visible = False
cmdAdjustOT.Visible = True
End Sub

Private Sub cmdEditHours_Click()
EnableTimeFields
Text6.SetFocus
cmdEditHours.Visible = False
cmdUpdateHours.Visible = True
cmdAddNewTime.Visible = False
cmdEnterHours.Enabled = False
cmdEnterTotal.Enabled = False
End Sub

Private Sub cmdEnterHours_Click()
EnableTimeFields
ClearHourFields
ClearDedFields
cmdEnterHours.Visible = False
cmdAddNewTime.Visible = True
cmdEditHours.Enabled = False
cmdEnterTotal.Enabled = False
Text6.SetFocus
End Sub

Private Sub cmdEnterTotal_Click()
EnterTotalHours
GetWages
End Sub

Private Sub cmdUpdateHours_Click()
cmdEditHours.Visible = True
cmdUpdateHours.Visible = False
cmdAddNewTime.Visible = False
cmdEnterHours.Enabled = False
cmdEditHours.Enabled = True
cmdEnterTotal.Enabled = True
ChangeTime
DisableTimeFields
GetTime
GetWages
GetOT
Call cmdEnterTotal_Click

End Sub

Private Sub EnterTotalHours()
On Error Resume Next
OpenDatabase
Dim sqlTime As String
Dim rsTime As Recordset

If Text16.Text > 40 Then
MsgBox ("This employee has hours over 40, You can adjust the remaining hours to over time")


sqlTime = "Select EmployeeID, [Total Hours], OTHours  From Employees "
sqlTime = sqlTime & "Where EmployeeID = " & Text7.Text
Set rsTime = DB.OpenRecordset(sqlTime)

If rsTime.RecordCount > 0 Then
rsTime.MoveFirst

With rsTime
.Edit
rsTime![Total Hours] = Text16.Text

.Update
End With


End If
End If
'MsgBox ("Total Hours Successfully Entered")
Exit Sub

End Sub

Private Sub AddTime()
On Error Resume Next
OpenDatabase
Dim sqlATime As String
Dim rsATime As Recordset

sqlATime = "Select EmployeeID, Mon, Tues, Wed, Thurs, Fri, Sat, Sun From EmpsTimeTable "
sqlATime = sqlATime & "Where EmployeeID = " & Text7.Text
Set rsATime = DB.OpenRecordset(sqlATime)

With rsATime
.AddNew
rsATime!EmployeeID = Text7.Text
rsATime!Mon = Text6.Text
rsATime!Tues = Text10.Text
rsATime!Wed = Text11.Text
rsATime!Thurs = Text12.Text
rsATime!Fri = Text13.Text
rsATime!Sat = Text14.Text
rsATime!Sun = Text15.Text
.Update
End With

'Call cmdEnterTotal_Click

AddDed

MsgBox ("Time Successfully Updated")

Exit Sub

End Sub

Private Sub GetTime()
On Error Resume Next
OpenDatabase
Dim sqlATime As String
Dim rsATime As Recordset

sqlATime = "Select EmployeeID, Mon, Tues, Wed, Thurs, Fri, Sat, Sun From EmpsTimeTable "
sqlATime = sqlATime & "Where EmployeeID = " & Text7.Text
Set rsATime = DB.OpenRecordset(sqlATime)

Text7.Text = rsATime!EmployeeID
Text6.Text = rsATime!Mon
Text10.Text = rsATime!Tues
Text11.Text = rsATime!Wed
Text12.Text = rsATime!Thurs
Text13.Text = rsATime!Fri
Text14.Text = rsATime!Sat
Text15.Text = rsATime!Sun

Text16.Text = (rsATime!Mon) + (rsATime!Tues) + (rsATime!Wed) + (rsATime!Thurs) + (rsATime!Fri) + (rsATime!Sat) + (rsATime!Sun)

If rsATime!EmployeeID = 0 Then
ClearHourFields
cmdEnterHours.Enabled = True
cmdEditHours.Enabled = False
cmdAddDed.Enabled = False
End If
Exit Sub

End Sub

Private Sub ChangeTime()
On Error GoTo EH:
OpenDatabase
Dim sqlATime As String
Dim rsATime As Recordset

sqlATime = "Select EmployeeID, Mon, Tues, Wed, Thurs, Fri, Sat, Sun From EmpsTimeTable "
sqlATime = sqlATime & "Where EmployeeID = " & Text7.Text
Set rsATime = DB.OpenRecordset(sqlATime)

With rsATime
.Edit
'rsATime!EmployeeID = Text7.Text
rsATime!Mon = Text6.Text
rsATime!Tues = Text10.Text
rsATime!Wed = Text11.Text
rsATime!Thurs = Text12.Text
rsATime!Fri = Text13.Text
rsATime!Sat = Text14.Text
rsATime!Sun = Text15.Text
.Update
End With
GetWages
ShowOT
MsgBox ("Time Successfully Updated")
cmdEnterHours.Enabled = False
cmdEditHours.Enabled = True
Exit Sub

EH: MsgBox ("This employee has not received a timesheet, You must enter a new timesheet")

cmdEnterHours.Enabled = True
cmdEditHours.Enabled = False
Call cmdEnterHours_Click
End Sub

Private Sub GetWages()
On Error Resume Next

OpenDatabase

Dim sqlWage As String
Dim rsWage As Recordset

sqlWage = "Select EmployeeID, [Total Hours], THours From Emp_Hours "
sqlWage = sqlWage & "Where EmployeeID = " & Text7.Text
Set rsWage = DB.OpenRecordset(sqlWage)

Text20.Text = Format$(rsWage!THours, "$#,##0.00;(#,##0.00)")
Text21.Text = rsWage![Total Hours]


End Sub

Private Sub GetOT()
On Error Resume Next
OpenDatabase

Dim sqlOT As String
Dim rsOT As Recordset

sqlOT = "Select EmployeeID, OTHours, OTotal From OTime "
sqlOT = sqlOT & "Where EmployeeID = " & Text7.Text
Set rsOT = DB.OpenRecordset(sqlOT)

With rsOT
.Edit
rsOT!OTHours = Text18.Text
.Update
End With

Text19.Text = rsOT!OTotal

End Sub

Private Sub AdjustOT()
On Error Resume Next
OpenDatabase
Dim sqlTime As String
Dim rsTime As Recordset


sqlTime = "Select EmployeeID,[Total Hours], OTHours  From Employees "
sqlTime = sqlTime & "Where Employees.EmployeeID = " & Text7.Text
Set rsTime = DB.OpenRecordset(sqlTime)

'Text21.Text = "40"

With rsTime
.Edit
rsTime![Total Hours] = Text21.Text
rsTime!OTHours = Text18.Text
.Update
End With

'----------------------------------------------------------

Dim rsOTime As Recordset
Dim sqlOTime As String

sqlOTime = "Select EmployeeID, OTotal  From OTime "
sqlOTime = sqlOTime & "Where EmployeeID = " & Text7.Text
Set rsOTime = DB.OpenRecordset(sqlOTime)

If rsOTime.RecordCount > 0 Then
rsOTime.MoveFirst
End If

Text19.Text = Format$(rsOTime!OTotal, "$#,##0.00;(#,##0.00)")
End Sub

Private Sub ShowOT()
On Error Resume Next
Dim rsOTime As Recordset
Dim sqlOTime As String

sqlOTime = "Select EmployeeID, OTotal, RHours, TWages  From OTime "
sqlOTime = sqlOTime & "Where EmployeeID = " & Text7.Text
Set rsOTime = DB.OpenRecordset(sqlOTime)

If rsOTime.RecordCount > 0 Then
rsOTime.MoveFirst
End If

Text19.Text = Format$(rsOTime!OTotal, "$#,##0.00;(#,##0.00)")
Text22.Text = Format$(rsOTime!RHours, "$#,##0.00;(#,##0.00)")
Text17.Text = Format$(rsOTime!TWages, "$#,##0.00;(#,##0.00)")

'-------------------------------------------------------------
Dim sqlTime As String
Dim rsTime As Recordset


sqlTime = "Select EmployeeID,[Total Hours], OTHours  From Employees "
sqlTime = sqlTime & "Where Employees.EmployeeID = " & Text7.Text
Set rsTime = DB.OpenRecordset(sqlTime)

Text18.Text = rsTime!OTHours

If Text18.Text = 0 Then
Text17.Visible = False
Text20.Visible = True
Else
Text17.Visible = True
Text20.Visible = False
End If
Exit Sub

End Sub


Private Sub ClearHourFields()
Text6.Text = "0"
Text10.Text = "0"
Text11.Text = "0"
Text12.Text = "0"
Text13.Text = "0"
Text14.Text = "0"
Text15.Text = "0"
Text16.Text = "0"
Text17.Text = "0"
Text18.Text = "0"
Text19.Text = "0"
Text20.Text = "0"
Text21.Text = "0"
Text22.Text = "0"
Text28.Text = "0"
End Sub

Private Sub ClearDedFields()
Text23.Text = "0.00"
Text24.Text = "0.00"
Text25.Text = "0.00"
Text26.Text = "0.00"
Text27.Text = "0.00"
Text28.Text = "0.00"
Text29.Text = "0.00"
Text30.Text = "0.00"
Text31.Text = "0.00"
Text32.Text = "0"

End Sub


Private Sub EnableTimeFields()
Text6.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Text16.Locked = True
End Sub

Private Sub DisableTimeFields()
Text6.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
End Sub

Private Sub SaveNote()
OpenDatabase
Dim sqlNote As String
Dim rsNote As Recordset

sqlNote = "Select EmployeeID, Notes From Employees "
sqlNote = sqlNote & "Where EmployeeID = " & Text7.Text
Set rsNote = DB.OpenRecordset(sqlNote)

With rsNote
.Edit
rsNote!Notes = Text1.Text
.Update
End With
End Sub

Private Sub ShowNote()
Dim sqlNote As String
Dim rsNote As Recordset

sqlNote = "Select EmployeeID, Notes From Employees "
sqlNote = sqlNote & "Where EmployeeID = " & Text7.Text
Set rsNote = DB.OpenRecordset(sqlNote)

Text1.Text = rsNote!Notes
End Sub

Private Sub EditDed()
On Error Resume Next
Dim sqlDED As String
Dim rsDED As Recordset

sqlDED = "Select * From Tax "
sqlDED = sqlDED & "Where EmpID = " & Text7.Text
Set rsDED = DB.OpenRecordset(sqlDED)

'Text27.Text <-- Total -->
With rsDED
.Edit
rsDED!statetax = Text23.Text
rsDED!fedtax = Text24.Text
rsDED!socialsec = Text25.Text
rsDED!fica = Text26.Text
rsDED!advance = Text29.Text
rsDED!garn = Text30.Text
rsDED!childsupport = Text31.Text
rsDED!exemptions = Text32.Text
.Update
End With

ShowDED
End Sub

Private Sub AddDed()
On Error Resume Next
Dim sqlDED As String
Dim rsDED As Recordset

sqlDED = "Select * From Tax "
sqlDED = sqlDED & "Where EmpID = " & Text7.Text
Set rsDED = DB.OpenRecordset(sqlDED)

'Text27.Text <-- Total -->
With rsDED
.AddNew
rsDED!EmpID = Text7.Text
rsDED!statetax = Text23.Text
rsDED!fedtax = Text24.Text
rsDED!socialsec = Text25.Text
rsDED!fica = Text26.Text
rsDED!advance = Text29.Text
rsDED!garn = Text30.Text
rsDED!childsupport = Text31.Text
rsDED!exemptions = Text32.Text
.Update
End With

ShowDED
End Sub

Private Sub ShowDED()
On Error Resume Next
Dim sqlDED As String
Dim rsDED As Recordset

sqlDED = "Select * From Tax "
sqlDED = sqlDED & "Where EmpID = " & Text7.Text
Set rsDED = DB.OpenRecordset(sqlDED)

Text23.Text = Format$(rsDED!statetax, "$#,##0.00;(#,##0.00)")
Text24.Text = Format$(rsDED!fedtax, "$#,##0.00;(#,##0.00)")
Text25.Text = Format$(rsDED!socialsec, "$#,##0.00;(#,##0.00)")
Text26.Text = Format$(rsDED!fica, "$#,##0.00;(#,##0.00)")
Text29.Text = Format$(rsDED!advance, "$#,##0.00;(#,##0.00)")
Text30.Text = Format$(rsDED!garn, "$#,##0.00;(#,##0.00)")
Text31.Text = Format$(rsDED!childsupport, "$#,##0.00;(#,##0.00)")
Text32.Text = rsDED!exemptions

If rsDED!EmpID = 0 Then
ClearDedFields
cmdEnterHours.Enabled = True
cmdEditHours.Enabled = False
cmdAddDed.Enabled = False
End If
'-----------------------------------------------------

Dim sqlTDED As String
Dim rsTDED As Recordset

sqlTDED = "Select TotalDED From Tax_Query "
sqlTDED = sqlTDED & "Where EmpID = " & Text7.Text
Set rsTDED = DB.OpenRecordset(sqlTDED)

Text27.Text = Format$(rsTDED!TotalDED, "$#,##0.00;(#,##0.00)")

If Text17.Visible = True Then
Text28.Text = Format$(Text17.Text - (rsTDED!TotalDED), "$#,##0.00;(#,##0.00)")
Else
Text28.Text = Format$(Text20.Text - (rsTDED!TotalDED), "$#,##0.00;(#,##0.00)")
End If

If rsDED!EmpID = 0 Then
ClearDedFields
cmdEnterHours.Enabled = True
cmdEditHours.Enabled = False
cmdAddDed.Enabled = False
End If

Exit Sub
End Sub
