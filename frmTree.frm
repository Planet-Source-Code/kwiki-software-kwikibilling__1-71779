VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmTree 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Active"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   990
   ClientWidth     =   12360
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmTree.frx":0000
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin KwikiBilling.XP_ProgressBar PB 
      Height          =   225
      Left            =   10440
      TabIndex        =   50
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   49152
      Scrolling       =   5
      ShowText        =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   12300
      TabIndex        =   47
      Top             =   7575
      Visible         =   0   'False
      Width           =   12360
      Begin VB.PictureBox picTray 
         Height          =   375
         Left            =   11040
         Picture         =   "frmTree.frx":0442
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin lvButton.lvButtons_H cmdVendors 
      Height          =   1095
      Left            =   10320
      TabIndex        =   46
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      Caption         =   "&Open Vendors"
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
      ImgAlign        =   2
      Image           =   "frmTree.frx":1324
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox picExpandAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      MouseIcon       =   "frmTree.frx":19D9
      MousePointer    =   99  'Custom
      Picture         =   "frmTree.frx":1B2B
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   45
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      MouseIcon       =   "frmTree.frx":1F29
      MousePointer    =   99  'Custom
      Picture         =   "frmTree.frx":207B
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   44
      Top             =   5760
      Width           =   255
   End
   Begin lvButton.lvButtons_H cmdClear 
      Height          =   1095
      Left            =   10320
      TabIndex        =   43
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      Caption         =   "&Clear Views"
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
      ImgAlign        =   2
      Image           =   "frmTree.frx":246E
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command7 
      Height          =   1095
      Left            =   10320
      TabIndex        =   42
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      Caption         =   "&Add Labor To Order"
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
      ImgAlign        =   2
      Image           =   "frmTree.frx":2D48
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command6 
      Height          =   1095
      Left            =   10320
      TabIndex        =   41
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      Caption         =   "&Add Product To Order"
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
      ImgAlign        =   2
      Image           =   "frmTree.frx":3622
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command9 
      Height          =   615
      Left            =   10320
      TabIndex        =   40
      Top             =   6120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "&Exit"
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
      Image           =   "frmTree.frx":3EFC
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command5 
      Height          =   615
      Left            =   8160
      TabIndex        =   39
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   " &Open Payments"
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
      Image           =   "frmTree.frx":48F6
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command4 
      Height          =   615
      Left            =   6240
      TabIndex        =   38
      Top             =   6120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Caption         =   "&Open Products"
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
      Image           =   "frmTree.frx":5828
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command3 
      Height          =   615
      Left            =   4200
      TabIndex        =   37
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "&Open Orders"
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
      Image           =   "frmTree.frx":6102
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command2 
      Height          =   615
      Left            =   2160
      TabIndex        =   36
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "&Open Employees"
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
      Image           =   "frmTree.frx":6554
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H Command1 
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "&Open Customers"
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
      Image           =   "frmTree.frx":6E2E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdViewRPT 
      Height          =   495
      Left            =   3360
      TabIndex        =   34
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "&Print Invoice"
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
      Image           =   "frmTree.frx":71C8
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdViewEstimate 
      Height          =   495
      Left            =   3360
      TabIndex        =   33
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "  &Print Estimate"
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
      Image           =   "frmTree.frx":80A2
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComctlLib.TreeView TvwCustomer 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8916
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   450
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":8F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":9316
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":9768
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtWorkorderID 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "(Null)"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView LvwOrders 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1508
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
      ForeColor       =   6618980
      BackColor       =   4210752
      BorderStyle     =   1
      Appearance      =   0
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
   Begin VB.TextBox txtPONumber 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "(Null)"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   10320
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
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
      Left            =   9600
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
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
      Left            =   8880
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
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
      Left            =   8160
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   2880
      TabIndex        =   49
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label25 
      Height          =   135
      Left            =   1680
      TabIndex        =   32
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   6840
      TabIndex        =   24
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   8160
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   7080
      TabIndex        =   31
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   2880
      TabIndex        =   30
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   2880
      TabIndex        =   28
      Top             =   2880
      Width           =   7215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Height          =   15
      Left            =   2880
      TabIndex        =   14
      Top             =   2280
      Width           =   7215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   7080
      TabIndex        =   19
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Height          =   15
      Left            =   2880
      TabIndex        =   18
      Top             =   3240
      Width           =   7215
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      Caption         =   "------------------------------------------------------------------------------"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   6240
      TabIndex        =   27
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Height          =   15
      Left            =   2880
      TabIndex        =   16
      Top             =   4200
      Width           =   7215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   6960
      TabIndex        =   7
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "                                       Total Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   3000
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Height          =   15
      Left            =   2880
      TabIndex        =   13
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "                                            Total Payments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   3960
      Width           =   7215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   3120
      Width           =   7215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  Customer Name                                                  Order ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   7080
      TabIndex        =   6
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   2880
      TabIndex        =   23
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Amount Remaining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   7080
      TabIndex        =   26
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   2880
      TabIndex        =   15
      Top             =   4680
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   4200
      Width           =   7215
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2880
      TabIndex        =   29
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "   Order Amount                                                 Tax Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   7215
   End
   Begin VB.Menu mnucp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuaddOrder 
         Caption         =   "Open Customers Order"
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddPart 
         Caption         =   "Add Item To Order"
      End
      Begin VB.Menu mnnuAddLabor 
         Caption         =   "Add Labor To Order"
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddPayment 
         Caption         =   "Apply Payment"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu bar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu bar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Minimize To Tray"
      End
   End
   Begin VB.Menu mnuCompSetup1 
      Caption         =   "Setup"
      Begin VB.Menu mnuCompSetup 
         Caption         =   "Setup Configuration"
      End
   End
   Begin VB.Menu mnuNewCustomer1 
      Caption         =   "Customers"
      Begin VB.Menu mnuNewCustomer 
         Caption         =   "Open Customers"
      End
   End
   Begin VB.Menu mnuAddEmp1 
      Caption         =   "Employees"
      Begin VB.Menu mnuAddEmp 
         Caption         =   "Open Employees"
      End
   End
   Begin VB.Menu mnuAddNewWO1 
      Caption         =   "Orders"
      Begin VB.Menu mnuAddNewWO 
         Caption         =   "Open Orders"
      End
      Begin VB.Menu mnubar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenParts 
         Caption         =   "Add Item To Order"
      End
      Begin VB.Menu mnuOpenLabor 
         Caption         =   "Add Labor To Order"
      End
      Begin VB.Menu mnubar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddPayment1 
         Caption         =   "Apply Payment To Order"
      End
   End
   Begin VB.Menu mnuParts1 
      Caption         =   "Products"
      Begin VB.Menu mnuParts 
         Caption         =   "Open Products"
      End
      Begin VB.Menu mnuOpenCat 
         Caption         =   "Product Categories"
      End
   End
   Begin VB.Menu mnuPaymentAdd1 
      Caption         =   "Payments"
      Begin VB.Menu mnuPaymentAdd 
         Caption         =   "Open Payments"
      End
      Begin VB.Menu mnuPaymentMeth 
         Caption         =   "Payment Methods"
      End
   End
   Begin VB.Menu mnuDB1 
      Caption         =   "Database Tools"
      Begin VB.Menu mnuDB 
         Caption         =   "Open Maintaince"
      End
      Begin VB.Menu bar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadTree 
         Caption         =   "Refresh Tree"
      End
   End
   Begin VB.Menu mnuSys 
      Caption         =   "Software Info"
      Begin VB.Menu mnuLiveUpdate 
         Caption         =   "Live Update"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Register Software"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Kwiki Billing"
      End
   End
   Begin VB.Menu mnuTrayPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuTrayPopRestore 
         Caption         =   "Restore Kwiki Billing"
      End
      Begin VB.Menu mnubar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayPopClose 
         Caption         =   "Exit Kwiki Billing"
      End
      Begin VB.Menu mnubar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayPopCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public StartingAddress As String
Dim rsCustDetail As Recordset
Dim rsCustomer As Recordset
Dim rsWorkorder As Recordset
Dim sOrderNr  As String
Dim lOrderKey As String

Private Sub ExpandTV()
    Dim i As Integer

    For i = 1 To TvwCustomer.Nodes.count
        TvwCustomer.Nodes(i).Expanded = True
    Next i

End Sub

Private Sub CollapseTV()
    Dim i As Integer

    For i = 1 To TvwCustomer.Nodes.count
        TvwCustomer.Nodes(i).Expanded = False
    Next i
End Sub

Private Sub cmdClear_Click()
CollapseTV
LvwOrders.ListItems.Clear
'FillInfo
lblName.Caption = ""
Label2 = ""
Label7 = ""
Label8 = ""
Label10 = ""
Label17 = ""
Label18 = ""
Label24.Caption = ""
PB.Visible = False
End Sub

Private Sub cmdVendors_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
frmVendors.Show
End Sub

Private Sub cmdViewEstimate_Click()
On Error GoTo EH:
SndPlayEx App.Path & "\Sounds\Start.wav"

With frmWorkorders.CRInvoice
.DataFiles(0) = App.Path & "\KwikiDat\db2.mdb"
.ReportFileName = App.Path & "\KwikiDat\Estimate.rpt"
.WindowTitle = "Invoice"
.ReplaceSelectionFormula ("{Workorders.PurchaseOrderNumber} =" & "'" & txtPONumber.Text & "'")
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

Private Sub cmdViewRPT_Click()
On Error GoTo EH:
SndPlayEx App.Path & "\Sounds\Start.wav"

With frmWorkorders.CRInvoice
.DataFiles(0) = App.Path & "\KwikiDat\db2.mdb"
.ReportFileName = App.Path & "\KwikiDat\Invoice.rpt"
.WindowTitle = "Invoice"
.ReplaceSelectionFormula ("{Workorders.PurchaseOrderNumber} =" & "'" & txtPONumber.Text & "'")
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

Private Sub Command1_Click()
SndClick
frmCustomers.Show
End Sub

Private Sub Command2_Click()
SndClick
frmEmployees.Show
End Sub

Private Sub Command3_Click()
SndClick
frmWorkorders.Show
End Sub

Private Sub Command4_Click()
SndClick
frmParts.Show
End Sub

Private Sub Command5_Click()
SndClick
frmPayments.Show
End Sub

Private Sub Command6_Click()
SndClick
frmWorkorderParts.Show
End Sub

Private Sub Command7_Click()
SndClick
frmWorkorderLabor.Show
End Sub

Private Sub Command9_Click()
SndClick

Dim Reply As Variant

Reply = MsgBox("Would you like to backup your data now before closing?", vbQuestion + vbYesNoCancel, "Confirm")

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

Private Sub Form_Activate()
On Error GoTo Err:
Set rsWorkorder = DB.OpenRecordset("Workorders", dbOpenTable)
Set rsCustomer = DB.OpenRecordset("Customers", dbOpenTable)
LoadTree
Err:
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Width = 12000
Height = 9000
Left = 0
Top = 0

If (Not OpenDatabase()) Then
  MsgBox "Database could not be openend !"
End If

setUpListView
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
WindowState = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsWorkorder = Nothing
Set rsCustomer = Nothing
Set rsCustDetail = Nothing
End Sub

Public Sub UpdateTree()
On Error Resume Next
Dim nods As MSComctlLib.Nodes
Set nods = TvwCustomer.Nodes

Dim sqlCustomer As String
Dim rsCustomer As Recordset

Dim sqlWorkorder As String
Dim rsWorkorder As Recordset

TvwCustomer.Nodes.Clear

sqlCustomer = "Select CustomerID, CompanyName From Customers "
Set rsCustomer = DB.OpenRecordset(sqlCustomer)

sqlWorkorder = "SELECT WorkorderID, CustomerID, PurchaseOrderNumber FROM Workorders "
Set rsWorkorder = DB.OpenRecordset(sqlWorkorder)

If (rsCustomer.RecordCount > 0) Then
rsCustomer.MoveFirst
End If

If (rsWorkorder.RecordCount > 0) Then
rsWorkorder.MoveFirst
End If

Do While Not rsCustomer.EOF
nods.Add , , "N1" & (rsCustomer!CustomerID), _
(rsCustomer!CompanyName), 1
rsCustomer.MoveNext
Loop
Do While Not rsWorkorder.EOF
nods.Add "N1" & (rsWorkorder!CustomerID), _
tvwChild, "ID" & (rsWorkorder!WorkorderID), "Order " & (rsWorkorder!WorkorderID), 2, 3
rsWorkorder.MoveNext
Loop

MDIForm1.sbStatus.Panels.Item(2).Text = "Currently > " & _
rsCustomer.RecordCount & " Customers"
MDIForm1.sbStatus.Panels.Item(3).Text = "Currently > " & _
rsWorkorder.RecordCount & " Workorders"

DoEvents
Exit Sub
End Sub

Private Sub LvwOrders_DblClick()
On Error GoTo EH:
With frmWorkorders.CRInvoice
.DataFiles(0) = App.Path & "\KwikiDat\db2.mdb"
.ReportFileName = App.Path & "\KwikiDat\invoice.rpt"
.WindowTitle = "Invoice"
.ReplaceSelectionFormula ("{Workorders.PurchaseOrderNumber} =" & "'" & txtPONumber.Text & "'")
.WindowShowExportBtn = True
.Action = 1
End With
Exit Sub
EH:
MsgBox Err.Description
End Sub

Private Sub picHide_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
CollapseTV
End Sub

Private Sub picExpandAll_Click()
SndPlayEx App.Path & "\Sounds\Start.wav"
ExpandTV
End Sub

'---------------------------------------------------------------------
'TREEVIEW POPUP MENU SECTION
'---------------------------------------------------------------------
Private Sub TvwCustomer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
If TvwCustomer.SelectedItem.Child.Selected Then
PopupMenu mnucp
End If
End If
End Sub

Private Sub mnuaddOrder_Click()
SndClick
frmWorkorders.Show
End Sub

Private Sub mnuAddPart_Click()
SndClick
frmWorkorderParts.Show
End Sub

Private Sub mnnuAddLabor_Click()
SndClick
frmWorkorderLabor.Show
End Sub

Private Sub mnuAddPayment_Click()
SndClick
frmPayments.Show
End Sub

'---------------------------------------------------------------------
'END TREEVIEW POPUP MENU
'---------------------------------------------------------------------

'---------------------------------------------------------------------
'MAIN MENU SECTION
'---------------------------------------------------------------------

Private Sub mnuCompSetup_Click()
SndClick
frmCompanySetup.Show
End Sub

Private Sub mnuNewCustomer_Click()
SndClick
frmCustomers.Show
End Sub

Private Sub mnuAddEmp_Click()
SndClick
frmEmployees.Show
End Sub

Private Sub mnuAddNewWO_Click()
SndClick
frmWorkorders.Show
End Sub

Private Sub mnuOpenParts_Click()
SndClick
frmWorkorderParts.Show
End Sub

Private Sub mnuOpenLabor_Click()
SndClick
frmWorkorderLabor.Show
End Sub

Private Sub mnuAddPayment1_Click()
SndClick
frmPayments.Show
End Sub

Private Sub mnuParts_Click()
SndClick
frmParts.Show
End Sub

Private Sub mnuOpenCat_Click()
SndClick
frmCategories.Show
End Sub

Private Sub mnuPaymentAdd_Click()
SndClick
frmPayments.Show
frmPayments.SetFocus
End Sub

Private Sub mnuPaymentMeth_Click()
SndClick
frmPaymentMethod.Show
End Sub

Private Sub mnuDB_Click()
SndClick
frmMaintain.Show
End Sub

Private Sub mnuLoadTree_Click()
UpdateTree
End Sub

Private Sub mnuLiveUpdate_Click()
MDIForm1.VerifyUpdates
End Sub

Private Sub mnuRegister_Click()
On Error GoTo EH:
ShellExecute 3, "open", "http://invoice.x10hosting.com/Register.htm", vbNullString, vbNullString, 1
Exit Sub
EH:
End Sub

Private Sub mnuAbout_Click()
SndClick
frmAbout.Show
End Sub

Private Sub mnuCalc_Click()
On Error GoTo EH
Shell "C:\Windows\system32\Calc.exe"
Exit Sub
EH:
MsgBox "Calculator not found on your system"
End Sub

Private Sub mnuExit_Click()
Dim Reply As Variant

Reply = MsgBox("Would you like to backup your data now before closing?", vbQuestion + vbYesNoCancel, "Confirm")

Select Case Reply
Case vbYes:
frmDB2.m_strType = "Backup"
frmDB2.Show

Case vbNo:
frmClose.Show
End

Case vbCancel:
Exit Sub

End Select
End Sub

Private Sub mnuTray_Click()
Call AddToTray
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'for popup menu at system tray
   '///adapted from example by Gary Lantz///
   'Const sMOD_NAME As String = "MDIForm1.picTray_MouseMove"
   On Error GoTo Error_Handler
   
   Dim lRet As Long
   
   If picTray.ScaleMode = vbPixels Then
      lRet = X
   Else
      lRet = X / Screen.TwipsPerPixelX
   End If
   
   Select Case lRet
      Case WM_RBUTTONUP
         PopupMenu mnuTrayPop
   End Select
   
   Exit Sub
Error_Handler:
End Sub

'----------------------------------------------------------------
Private Sub mnuTrayPopClose_Click()
'm_blnAllowClose = True
Call RemoveFromTray
frmClose.Show
End Sub

Private Sub mnuTrayPopRestore_Click()
MDIForm1.Show
frmTree.Show
Call RemoveFromTray
End Sub

Private Sub mnuTrayPopCancel_Click()
mnuTrayPop.Visible = False
End Sub

'---------------------------------------------------------------

'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then
'PopupMenu mnuMain
'End If
'End Sub


'--------------------------------------------------------------------
'END POPUP MENU SECTION
'--------------------------------------------------------------------

Private Sub TvwCustomer_NodeClick(ByVal Node As MSComctlLib.Node)
'Debug.Print "Node : " & Node.Key & Node.Tag
SndClick

lOrderKey = Mid$(Node.Key, 3)

With rsWorkorder
   .Index = "PrimaryKey"
   .Seek "=", lOrderKey

If Not .NoMatch Then
sOrderNr = TvwCustomer.SelectedItem
Else
Exit Sub
End If
End With

FillListView
FillFields
FillInfo
GetName
GetTax
End Sub

Public Sub FillListView()
On Error GoTo Log:
Dim WOAdd As ListItem
Dim SQLwo As String
LvwOrders.ListItems.Clear


'SELECT DISTINCTROW [Workorders].[WorkorderID], [Workorders].[CustomerID], [Workorders].[DateReceived], [Workorders].[DateRequired], [Workorders].[SalesTaxRate], [Sum Of Payments Query].[Total Payments], [Parts Totals by Workorder].[Parts Total], [Labor Totals by Workorder].[Labor Total], [Workorders].[DateFinished] "
'FROM ((Workorders LEFT JOIN [Sum Of Payments Query]
'ON [Workorders].[WorkorderID]=[Sum Of Payments Query].[WorkorderID]) LEFT JOIN [Parts Totals by Workorder]
'ON [Workorders].[WorkorderID]=[Parts Totals by Workorder].[WorkorderID]) LEFT JOIN [Labor Totals by Workorder]
'ON [Workorders].[WorkorderID]=[Labor Totals by Workorder].[WorkorderID] "


SQLwo = "SELECT DISTINCTROW [Workorders].[WorkorderID], [Workorders].[CustomerID], [Workorders].[PurchaseOrderNumber], [Workorders].[DateReceived], [Workorders].[DateRequired], [Workorders].[SalesTaxRate], [Sum Of Payments Query].[Total Payments], [Parts Totals by Workorder].[Parts Total], [Labor Totals by Workorder].[Labor Total], "
SQLwo = SQLwo & "([Parts Total]+[Labor Total]) As Total, "
SQLwo = SQLwo & "[Parts Total]+[Labor Total]-[Total Payments] AS AmountDue "
SQLwo = SQLwo & "FROM ((Workorders INNER JOIN [Sum Of Payments Query] "
SQLwo = SQLwo & "ON [Workorders].[WorkorderID]=[Sum Of Payments Query].[WorkorderID]) INNER JOIN [Parts Totals by Workorder] "
SQLwo = SQLwo & "ON [Workorders].[WorkorderID]=[Parts Totals by Workorder].[WorkorderID]) INNER JOIN [Labor Totals by Workorder] "
SQLwo = SQLwo & "ON [Workorders].[WorkorderID]=[Labor Totals by Workorder].[WorkorderID] "
SQLwo = SQLwo & "WHERE [Workorders].[WorkorderID] = " & lOrderKey

Set rsCustDetail = DB.OpenRecordset(SQLwo)
If (rsCustDetail.RecordCount > 0) Then
rsCustDetail.MoveFirst

While Not rsCustDetail.EOF
Set WOAdd = LvwOrders.ListItems.Add(, , _
rsCustDetail!WorkorderID, , 3)
WOAdd.SubItems(1) = rsCustDetail![CustomerID]
WOAdd.SubItems(2) = rsCustDetail![DateReceived]
WOAdd.SubItems(3) = rsCustDetail![DateRequired]
WOAdd.SubItems(4) = Format$(rsCustDetail![SalesTaxRate], "0.0;(0.0)")
WOAdd.SubItems(5) = Format$(rsCustDetail![Parts Total], "$#,##0.00;(#,##0.00)")
WOAdd.SubItems(6) = Format$(rsCustDetail![Labor Total], "$#,##0.00;(#,##0.00)")
WOAdd.SubItems(7) = Format$(rsCustDetail![Total], "$#,##0.00;(#,##0.00)")
WOAdd.SubItems(8) = Format$(rsCustDetail![Total Payments], "$#,##0.00;(#,##0.00)")
WOAdd.SubItems(9) = Format$(rsCustDetail![AmountDue], "$#,##0.00;(#,##0.00)")

rsCustDetail.MoveNext
Wend
DoEvents

Else
LvwOrders.ListItems.Clear

'SELECT [Labor Totals by Workorder].WorkorderID, [Labor Totals by Workorder].[Labor Total], [Parts Totals by Workorder].[Parts Total], [Labor Total]+[Parts Total] AS AmountDue
'FROM (Workorders INNER JOIN [Parts Totals by Workorder] ON Workorders.WorkorderID = [Parts Totals by Workorder].WorkorderID) INNER JOIN [Labor Totals by Workorder]
'ON Workorders.WorkorderID = [Labor Totals by Workorder].WorkorderID

SQLwo = "SELECT  [Workorders].[CustomerID], [Workorders].[PurchaseOrderNumber], [Workorders].[DateReceived], [Workorders].[DateRequired], [Workorders].[SalesTaxRate], [Labor Totals by Workorder].[WorkorderID], [Labor Totals by Workorder].[Labor Total], [Parts Totals by Workorder].[Parts Total], [Labor Total]+[Parts Total] AS AmountDue "
SQLwo = SQLwo & "FROM (Workorders INNER JOIN [Parts Totals by Workorder] "
SQLwo = SQLwo & "ON [Workorders].[WorkorderID]=[Parts Totals by Workorder].[WorkorderID]) INNER JOIN [Labor Totals by Workorder] "
SQLwo = SQLwo & "ON [Workorders].[WorkorderID]=[Labor Totals by Workorder].[WorkorderID] "
SQLwo = SQLwo & "WHERE [Workorders].[WorkorderID] = " & lOrderKey

Set rsCustDetail = DB.OpenRecordset(SQLwo)

While Not rsCustDetail.EOF
Set WOAdd = LvwOrders.ListItems.Add(, , _
rsCustDetail!WorkorderID, , 2)
WOAdd.SubItems(1) = rsCustDetail![CustomerID]
WOAdd.SubItems(2) = rsCustDetail![DateReceived]
WOAdd.SubItems(3) = rsCustDetail![DateRequired]
WOAdd.SubItems(4) = Format$(rsCustDetail![SalesTaxRate], "0.0;(0.0)")
WOAdd.SubItems(5) = Format$(rsCustDetail![Parts Total], "$#,##0.00;(#,##0.00)")
WOAdd.SubItems(6) = Format$(rsCustDetail![Labor Total], "$#,##0.00;(#,##0.00)")
WOAdd.SubItems(7) = Format$(rsCustDetail![AmountDue], "$#,##0.00;(#,##0.00)")
rsCustDetail.MoveNext
Wend
DoEvents
'---------------------------------------------------------------
End If

Set rsCustDetail = Nothing

Exit Sub
Log:
End Sub

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "", 300, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "CID", 0, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "Received", 1500, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "Required", 1500, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "TaxRate %", 1400, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "PartsTotal", 1400, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "LaborTotal", 1400, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "Order Amount", 1600, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "Payments", 1300, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "Sub Total", 0, lvwColumnLeft)
Set clmHdr = LvwOrders.ColumnHeaders. _
Add(, , "Status", 1300, lvwColumnLeft)
LvwOrders.View = lvwReport
End Sub

Public Sub FillFields()
With rsWorkorder
If (Not IsNull(!WorkorderID)) Then txtWorkorderID = !WorkorderID
If (Not IsNull(!PurchaseOrderNumber)) Then txtPONumber = !PurchaseOrderNumber
DoEvents
End With
End Sub

Public Sub FillInfo()
On Error Resume Next
Dim rsStatus As Recordset
Dim sqlStatus As String
sqlStatus = "Select Workorders.WorkorderID, Workorders.Status FROM Workorders "
sqlStatus = sqlStatus & "WHERE Workorders.WorkorderID = " & txtWorkorderID.Text
Set rsStatus = DB.OpenRecordset(sqlStatus)

GetTax

Label2.Caption = lOrderKey
Label24.Caption = "Order"

If LvwOrders.SelectedItem.SubItems(7) = "" Then
Label7.Caption = " " & "(0.00)"
Else
Label7.Caption = " " & LvwOrders.SelectedItem.SubItems(7)
End If

If LvwOrders.SelectedItem.SubItems(8) = "" Then
Label8.Caption = " " & "(0.00)"
Text1.Text = ""
Label17.Caption = " " & Text1.Text
Text2.Text = ""
Text4.Text = ""
Label18.Caption = " " & Text4.Text
PB.Visible = True
With rsStatus
.Edit
rsStatus!Status = 10
.Update
End With
PB.Value = rsStatus!Status
PB.Color = vbRed
Else
Label8.Caption = " " & LvwOrders.SelectedItem.SubItems(8)
PB.Visible = True
PB.Value = rsStatus!Status
PB.Color = &HC000&
End If

If LvwOrders.SelectedItem.SubItems(7) = "" Then
Label10.Caption = " " & "(0.00)"
End If

If Text4.Text < "(0.00)" Then
Label18.ForeColor = &H80000002
Label20.Caption = "Amount Remaining"
Else
Label20.Caption = "Change Due"
Label18.ForeColor = vbRed
PB.Color = vbYellow
PB.Value = 100
With rsStatus
.Edit
rsStatus!Status = 100
.Update
End With
End If
'--------------------------------------------------------
If LvwOrders.SelectedItem.SubItems(1) = "" Then
PB.Visible = False
End If

If Text4.Text = "$0.00" Then
PB.Visible = True
PB.Color = vbYellow
PB.Value = 100

With rsStatus
.Edit
rsStatus!Status = 100
.Update
End With
End If
End Sub

Private Sub LoadTree()
Screen.MousePointer = vbHourglass
UpdateTree
DoEvents
Screen.MousePointer = vbDefault
End Sub

Public Sub GetTax()
On Error Resume Next
Dim rsTax As Recordset
Dim sqlTax As String

sqlTax = "Select DISTINCTROW [Total], [Grand Total] FROM [TotalTax] "
sqlTax = sqlTax & "WHERE Workorders.WorkorderID = " & lOrderKey
Set rsTax = DB.OpenRecordset(sqlTax)

If (rsTax.RecordCount > 0) Then
rsTax.MoveFirst

Text1.Text = Format$(rsTax![Total], "$#,##0.00;(#,##0.00)")
Label17.Caption = " " & Text1.Text

Text2.Text = Format$(rsTax![Grand Total], "$#,##0.00;(#,##0.00)")
Label10.Caption = " " & Text2.Text
End If

'------------------------------------------------------------------

Dim rsTotal As Recordset
Dim sqlTotal As String
sqlTotal = "Select DISTINCTROW [Amount Due] FROM [Total] "
sqlTotal = sqlTotal & "WHERE Workorders.WorkorderID = " & lOrderKey
Set rsTotal = DB.OpenRecordset(sqlTotal)

If (rsTotal.RecordCount > 0) Then
rsTotal.MoveFirst
End If

Text4.Text = Format$(rsTotal![Amount Due], "$#,##0.00;(#,##0.00)")
Label18.Caption = " " & Text4.Text

End Sub

Private Sub GetName()
On Error Resume Next
Dim rsName As Recordset
Dim sqlName As String

sqlName = "Select Workorders.CustomerID, Workorders.WorkorderID, Customers.CustomerID, Customers.CompanyName FROM Customers, Workorders "
sqlName = sqlName & "WHERE Customers.CustomerID = " & TvwCustomer.SelectedItem.Index
Set rsName = DB.OpenRecordset(sqlName)

lblName.Caption = "  " & rsName!CompanyName
lblName.Caption = "  " & TvwCustomer.SelectedItem.Parent
End Sub
