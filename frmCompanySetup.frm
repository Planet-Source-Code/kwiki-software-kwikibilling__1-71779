VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCompanySetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "     Setup Configuration"
   ClientHeight    =   4785
   ClientLeft      =   1095
   ClientTop       =   210
   ClientWidth     =   8340
   Icon            =   "frmCompanySetup.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   6720
      TabIndex        =   33
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
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
      Image           =   "frmCompanySetup.frx":000C
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   3
      Left            =   4680
      Picture         =   "frmCompanySetup.frx":0A06
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   2
      Left            =   3960
      Picture         =   "frmCompanySetup.frx":0D48
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   1
      Left            =   3240
      Picture         =   "frmCompanySetup.frx":108A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Index           =   0
      Left            =   2520
      Picture         =   "frmCompanySetup.frx":13CC
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   661
      TabMaxWidth     =   2646
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Company Setup  "
      TabPicture(0)   =   "frmCompanySetup.frx":170E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Company Logo  "
      TabPicture(1)   =   "frmCompanySetup.frx":172A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSaveImage"
      Tab(1).Control(1)=   "cmdClearPhoto"
      Tab(1).Control(2)=   "cmdAddPhoto"
      Tab(1).Control(3)=   "lblStatus"
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "imgLogo"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Configuration"
      TabPicture(2)   =   "frmCompanySetup.frx":1746
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Configuration"
         Height          =   3255
         Left            =   -74760
         TabIndex        =   35
         Top             =   600
         Width           =   7575
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   37
            Top             =   840
            Width           =   4335
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   36
            Top             =   360
            Width           =   4335
         End
         Begin lvButton.lvButtons_H cmdOpen 
            Height          =   375
            Left            =   6360
            TabIndex        =   40
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            CapAlign        =   2
            BackStyle       =   4
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
            Image           =   "frmCompanySetup.frx":1762
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdOpen1 
            Height          =   375
            Left            =   6360
            TabIndex        =   41
            Top             =   840
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            CapAlign        =   2
            BackStyle       =   4
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
            Image           =   "frmCompanySetup.frx":1BB4
            cBack           =   -2147483633
         End
         Begin VB.Label Label3 
            Caption         =   "Click Sound Path :"
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
            TabIndex        =   39
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Database Path :"
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
            TabIndex        =   38
            Top             =   360
            Width           =   1575
         End
      End
      Begin lvButton.lvButtons_H cmdSaveImage 
         Height          =   375
         Left            =   -69240
         TabIndex        =   34
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Save Logo"
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
         Image           =   "frmCompanySetup.frx":2006
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdClearPhoto 
         Height          =   375
         Left            =   -69240
         TabIndex        =   32
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Clear Logo"
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
         Image           =   "frmCompanySetup.frx":28E0
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdAddPhoto 
         Height          =   375
         Left            =   -69240
         TabIndex        =   31
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Add Logo"
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
         Image           =   "frmCompanySetup.frx":31BA
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Frame Frame1 
         Height          =   3495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7695
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   960
            TabIndex        =   23
            Text            =   "Cash"
            Top             =   2880
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSMask.MaskEdBox Text6 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   ".0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   22
            Top             =   2880
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   -2147483630
            ForeColor       =   65280
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
            Format          =   ".0"
            PromptChar      =   "_"
         End
         Begin MSComDlg.CommonDialog CD3 
            Left            =   120
            Top             =   2880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000018&
            DataField       =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2160
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H80000018&
            DataField       =   "DefaultInvoiceDescription"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   2160
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   2520
            Width           =   2895
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H80000018&
            DataField       =   "FaxNumber"
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
            TabIndex        =   12
            Top             =   2040
            Width           =   2895
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H80000018&
            DataField       =   "PhoneNumber"
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
            TabIndex        =   11
            Top             =   1560
            Width           =   2895
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            DataField       =   "CompanyName"
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
            TabIndex        =   10
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000012&
            DataField       =   "SetupID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2880
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Frame Frame2 
            Height          =   3255
            Left            =   5640
            TabIndex        =   2
            Top             =   120
            Width           =   1935
            Begin lvButton.lvButtons_H cmdRefresh 
               Height          =   375
               Left            =   240
               TabIndex        =   30
               Top             =   2640
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               Caption         =   "&Requery"
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
               Image           =   "frmCompanySetup.frx":3A94
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdDelete 
               Height          =   375
               Left            =   240
               TabIndex        =   29
               Top             =   2160
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               Caption         =   "&Delete"
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
               Image           =   "frmCompanySetup.frx":436E
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdEdit 
               Height          =   375
               Left            =   240
               TabIndex        =   27
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               Caption         =   "&Edit"
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
               Image           =   "frmCompanySetup.frx":4C48
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdCancel 
               Height          =   375
               Left            =   240
               TabIndex        =   26
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               Caption         =   "&Cancel"
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
               Image           =   "frmCompanySetup.frx":5522
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdUpdate 
               Height          =   375
               Left            =   240
               TabIndex        =   25
               Top             =   720
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               Caption         =   "&Update"
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
               Image           =   "frmCompanySetup.frx":5DFC
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdAdd 
               Height          =   375
               Left            =   240
               TabIndex        =   24
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               Caption         =   "&Add"
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
               Image           =   "frmCompanySetup.frx":60E3
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdSave 
               Height          =   375
               Left            =   240
               TabIndex        =   28
               Top             =   1680
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               Caption         =   "&Save"
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
               Image           =   "frmCompanySetup.frx":69BD
               cBack           =   -2147483633
            End
         End
         Begin VB.Label lblLabels 
            Caption         =   "FaxNumber:"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   8
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "PhoneNumber:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Street Address:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "CompanyName:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Invoice Exit Message:"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   4
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "City\State\Zip:"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   -71760
         TabIndex        =   17
         Top             =   3240
         Width           =   4095
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   3720
         Width           =   7695
      End
      Begin VB.Label Label6 
         Caption         =   $"frmCompanySetup.frx":7297
         Height          =   495
         Left            =   -74760
         TabIndex        =   15
         Top             =   720
         Width           =   7095
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   1515
         Left            =   -71760
         Top             =   1560
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmCompanySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents HO As cBinaryDBObject
Attribute HO.VB_VarHelpID = -1
Public rsSetup As Recordset
Dim sqlSetup As String
Dim sFileName As String
Dim Clear As String

Private Sub cmdAddPhoto_Click()
On Error Resume Next
With CD3
.DialogTitle = "Add Logo"
.CancelError = False
'.Filter = "JPEG Files (*.jpg*)|*.JPG"
.ShowSave
If Len(.FileName) = 0 Then
Exit Sub
End If
sFileName = .FileName
End With
imgLogo.Picture = LoadPicture(sFileName)
Label7.Caption = "File Path : " & sFileName
End Sub

Private Sub cmdClearPhoto_Click()
LoadDefaultPic
End Sub


Private Sub cmdOpen_Click()
SaveDBPath
End Sub

Private Sub cmdOpen1_Click()
SaveSndPath
End Sub

Private Sub cmdSave_Click()
On Error GoTo EditErr
With rsSetup
.Edit
rsSetup!CompanyName = Text2.Text
rsSetup!Address = Text3.Text
rsSetup!PhoneNumber = Text4.Text
rsSetup!FaxNumber = Text5.Text
rsSetup!SalesTaxRate = Text6.Text
rsSetup!DefaultInvoiceDescription = Text7.Text
rsSetup!DefaultPaymentTerms = Text8.Text
.Update
.MoveLast
End With

SaveBinaryObject
FillFields
DisableFields

cmdAdd.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdRefresh.Enabled = True
cmdClose.Enabled = True
cmdAddPhoto.Enabled = False
cmdSaveImage.Enabled = False
cmdClearPhoto.Enabled = False
cmdEdit.Visible = True
cmdSave.Visible = False

frmCoAddProg.Show
Exit Sub
EditErr:
End Sub

Private Sub cmdSaveImage_Click()
SaveBinaryObject
End Sub

Private Sub Form_Load()
OpenDatabase
If (Not OpenDatabase()) Then
  MsgBox "Database could not be openend !"
End If

sqlSetup = "Select * From CompanySetup"
Set rsSetup = DB.OpenRecordset(sqlSetup)

FillFields
DisableFields
EnableCont

'GET DB PATH IF CONFIGURED
ImportDBPath
'Sound
ImportSndPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmTree.TvwCustomer.SetFocus
End Sub

Private Sub HO_Error(ID As Long, Msg As String)
MsgBox ID & ":  " & Msg
End Sub

Private Sub HO_Status(ID As Long, Msg As String)
lblStatus.Caption = CStr(ID) & ":  " & Msg
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        rsSetup.MoveFirst
        Beep
        FillFields        'call subprocedure
        
    Case 1
        rsSetup.MovePrevious
        If rsSetup.BOF Then
        rsSetup.MoveFirst
        End If
        FillFields

    Case 2
        rsSetup.MoveNext
        If rsSetup.EOF Then
        rsSetup.MoveLast
        End If
        FillFields
    
    Case 3
        rsSetup.MoveLast
        Beep
        FillFields
        
End Select
End Sub

Private Sub cmdCancel_Click()
On Error GoTo CancelErr
rsSetup.Requery
FillFields
DisableFields
EnableCont
cmdSave.Enabled = False

If Text2.Text = "" Then
cmdEdit.Enabled = False
cmdEdit.Visible = False
Else
cmdEdit.Enabled = True
cmdEdit.Visible = True
End If

Text2.SetFocus

cmdAddPhoto.Enabled = False
cmdClearPhoto.Enabled = False
cmdSaveImage.Enabled = False
Exit Sub
CancelErr:
End Sub

Private Sub cmdEdit_Click()
DisableCont
cmdEdit.Visible = False
cmdUpdate.Enabled = False
cmdSave.Visible = True
cmdSave.Enabled = True
cmdAddPhoto.Enabled = True
cmdSaveImage.Enabled = True
cmdClearPhoto.Enabled = True
cmdSaveImage.Enabled = True
EnableFields
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddErr
ClearFields
EnableFields
DisableCont
Text2.SetFocus
Exit Sub
AddErr:
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
If MsgBox("Are you sure you want to Remove ? " & Text2.Text, vbYesNo + vbQuestion, "Confirm") = vbYes Then
Unload Me
frmCoMsg.Show
ElseIf vbNo Then
Me.Show
End If

Exit Sub
DeleteErr:
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo RefreshErr
frmCoAddProg.Show
rsSetup.Requery
rsSetup.MoveFirst
FillFields
Exit Sub
RefreshErr:
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo UpdateErr
With rsSetup
.AddNew
rsSetup!CompanyName = Text2.Text
rsSetup!Address = Text3.Text
rsSetup!PhoneNumber = Text4.Text
rsSetup!FaxNumber = Text5.Text
rsSetup!SalesTaxRate = Text6.Text
rsSetup!DefaultInvoiceDescription = Text7.Text
rsSetup!DefaultPaymentTerms = Text8.Text
.Update
.MoveLast
End With

FillFields
DisableFields
EnableCont
LoadDefaultPic
frmCoAddProg.Show
Exit Sub
UpdateErr:
MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
SndClick
Set rsSetup = Nothing
Unload Me
End Sub

Private Sub DisableCont()
cmdAdd.Enabled = False
cmdUpdate.Enabled = True
cmdCancel.Enabled = True

If Text1.Text = "" Then
cmdEdit.Enabled = False
Else
cmdEdit.Enabled = True
End If

cmdDelete.Enabled = False
cmdRefresh.Enabled = False
cmdClose.Enabled = False
End Sub

Private Sub EnableCont()
cmdAdd.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False

If Text1.Text = "" Then
cmdEdit.Enabled = False
cmdDelete.Enabled = False
Else
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End If

cmdRefresh.Enabled = True
cmdClose.Enabled = True
End Sub

Private Sub DisableFields()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Enabled = False
Text7.Locked = True
Text8.Locked = True
End Sub

Private Sub EnableFields()
Text1.Locked = True
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Enabled = True
Text7.Locked = False
Text8.Locked = False
End Sub

Private Sub FillFields()
On Error GoTo EH
Text1.Text = rsSetup!SetupID
Text2.Text = rsSetup!CompanyName
Text3.Text = rsSetup!Address
Text4.Text = rsSetup!PhoneNumber
Text5.Text = rsSetup!FaxNumber
Text6.Text = rsSetup!SalesTaxRate
Text7.Text = rsSetup!DefaultInvoiceDescription
Text8.Text = rsSetup!DefaultPaymentTerms
imgLogo.Picture = LoadPicture(rsSetup!FileName)
Exit Sub
EH:
End Sub

Private Sub ClearFields()
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
Text6.Text = "0"
Text7.Text = Clear
Text8.Text = "Cash"
imgLogo.Picture = LoadPicture(Clear)
End Sub

Private Sub LoadDefaultPic()
On Error Resume Next
sFileName = App.Path & "\Kwikidat\BitMaps\BlankLogo.jpg"
imgLogo.Picture = LoadPicture(sFileName)
Label7.Caption = "Logo Cleared"
SaveBinaryObject
End Sub

Private Sub GetBinaryObject()
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim fn As String                       'Binary file name to use as storage
Dim i As Integer

    If Text1.Text = "" Then
        Set HO = New cBinaryDBObject       'create the new bd object

        FieldNames(0) = "SetupID"               'return the ID field
        FieldNames(1) = "FileName"         'return the filename

        With HO
            .KillFile = True                            'kill the filename if it exists
             Set .DB = DB                        'pass the database
            .ObjectKeyFieldName = "SetupID"   'the key/index field is
            .ObjectKey = Text1.Text          'the value to search for is
            .ObjectFieldName = "Logo"              'name of the field that contains the binary file
            .ObjectTableName = "CompanySetup"          'table that contains the binary files
            .SubFieldNames = FieldNames                 'pass in the field names to return
            .FileName = "FileName"     'file name to use"
            .GetObject                                  'get the file from the database
            .ReturnData RD()                            'return any aditional data
            fn = .FileName                              'actual file name used - if default was used
        End With
        Set HO = Nothing

        imgLogo.Picture = LoadPicture(fn)

        For i = 0 To UBound(RD)
            Debug.Print RD(i)                      'print aditional info returned
        Next

    End If
End Sub

Private Sub SaveBinaryObject()
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim FieldData(1) As Variant            'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim fn As String                       'Binary file name to use as storage
Dim i As Integer

    If sFileName = "" Then
    Exit Sub
    End If

    Set HO = New cBinaryDBObject       'create the new bd object

     FieldNames(0) = "SetupID"       'return the ID field
     FieldNames(1) = "FileName"         'return the filename
     FieldData(0) = Null                 'return the ID field
     FieldData(1) = sFileName           'return the filename

    With HO
        .KillFile = False                       'kill the filename if it exists
         Set .DB = DB                   'pass the database
        .ObjectKeyFieldName = "SetupID" 'the key/index field is
        .ObjectKey = Text1.Text         'the value to search for is
        .ObjectFieldName = "Logo"         'name of the field that contains the binary file
        .ObjectTableName = "CompanySetup"      'table that contains the binary files
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
End Sub



Public Function ImportDBPath()
Dim iniFile As String, ItsThere As Boolean
iniFile = App.Path & "\Settings.ini"
'CREATE THE INI IF NOT EXIST

ItsThere = FileExists(iniFile)
If ItsThere = False Then
Else

Text9.Text = ReadINI("DBPath", "Path", iniFile)

End If
End Function

Public Function SaveDBPath()
Dim iniFile As String
Dim sFile As String
iniFile = App.Path & "\Settings.ini"

With CD3
.DialogTitle = "Set Database"
        .CancelError = False
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Function
        End If
        sFile = .FileName
    End With
Close #1
Text9.Text = sFile

WriteINI "DBPath", "Path", Text9.Text, iniFile

End Function

Public Function ImportSndPath()
Dim iniFile As String, ItsThere As Boolean
iniFile = App.Path & "\Settings.ini"
Dim SndPath As String
'CREATE THE INI IF NOT EXIST

ItsThere = FileExists(iniFile)
If ItsThere = False Then
WriteINI "SndPath", "Path", App.Path & "\Sounds\start.wav", iniFile
Text10.Text = ReadINI("SndPath", "Path", iniFile)

Else

Text10.Text = ReadINI("SndPath", "Path", iniFile)

End If
End Function

Public Function SaveSndPath()
Dim iniFile As String
Dim sFile As String
iniFile = App.Path & "\Settings.ini"

With CD3
.DialogTitle = "Set Database"
        .CancelError = False
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Function
        End If
        sFile = .FileName
    End With
Close #1
Text10.Text = sFile

WriteINI "SndPath", "Path", Text10.Text, iniFile
End Function
