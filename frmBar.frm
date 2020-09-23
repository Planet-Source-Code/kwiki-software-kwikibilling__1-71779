VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmBar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    BarCode Label Generator"
   ClientHeight    =   4890
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   9555
   Icon            =   "frmBar.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   4440
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
      Image           =   "frmBar.frx":000C
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   4440
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
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
      Height          =   1275
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   9225
      TabIndex        =   14
      Top             =   3000
      Width           =   9225
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   9225
      TabIndex        =   13
      Top             =   1680
      Width           =   9225
   End
   Begin VB.Frame Frame2 
      Caption         =   "BarCode 128 Properties"
      Height          =   825
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   9285
      Begin VB.ComboBox cboTextStyle 
         Height          =   315
         ItemData        =   "frmBar.frx":0A06
         Left            =   5640
         List            =   "frmBar.frx":0A16
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   150
         Width           =   975
      End
      Begin VB.ComboBox cboBarSize 
         Height          =   315
         ItemData        =   "frmBar.frx":0A3F
         Left            =   7710
         List            =   "frmBar.frx":0A4C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.CheckBox chkTextAlignment 
         Appearance      =   0  'Flat
         Caption         =   "Bottom Align Caption"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox checkBarCaption 
         Appearance      =   0  'Flat
         Caption         =   "Bar With Caption"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   210
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox ChkExecute 
         Appearance      =   0  'Flat
         Caption         =   "Display after Save"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4620
         TabIndex        =   12
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6780
         TabIndex        =   11
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type of Bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   3165
      Begin VB.CheckBox chkBar128 
         Caption         =   "Bar 128"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkBar39 
         Caption         =   "Bar 39"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   2
         Top             =   180
         Value           =   1  'Checked
         Width           =   1395
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaxLength       =   21
      TabIndex        =   0
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right Click Over Each Bar to Get Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Part Code For Bar Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   2565
   End
   Begin VB.Menu mnucp 
      Caption         =   "Copy Image"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyImage1 
         Caption         =   "Copy  Bar of Code 128 to Clip Board"
      End
      Begin VB.Menu mnuSaveImage1 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu mnucp1 
      Caption         =   "Copy Image2"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyImage2 
         Caption         =   "Copy This Bar of Code 39 to Clip Board"
      End
      Begin VB.Menu mnuSaveImage2 
         Caption         =   "Save As"
      End
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Target As Object
Dim LastSave As String 'To Store last Saved Directory
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cboBarSize_Click()
'---------------------------------------------------------------------------------------
' Procedure : cboBarSize_Click
' DateTime  : 10/7/2002 19:40
' Author    : Joyprakash
' Purpose   : This Method is basically used to resize the BarCode
'---------------------------------------------------------------------------------------

Picture2.ScaleMode = 3
Picture1.ScaleMode = 3
Select Case cboBarSize.ListIndex
    Case 0
        Picture2.Height = Picture2.Height * (1.4 * 40 / Picture2.ScaleHeight)
        Picture2.FontSize = 8
    Case 1
        Picture2.Height = Picture2.Height * (2.4 * 40 / Picture2.ScaleHeight)
        Picture2.FontSize = 10
    Case 2
        Picture2.Height = Picture2.Height * (3 * 40 / Picture2.ScaleHeight)
        Picture2.FontSize = 14
End Select
    Text1_Change
End Sub

Private Sub cboTextStyle_Click()
    Text1_Change
End Sub

Private Sub checkBarCaption_Click()
If checkBarCaption.Value = 1 Then
    checkBarCaption.Caption = "Show Text Caption"
    chkTextAlignment.Visible = True
Else
    checkBarCaption.Caption = "No Caption"
    chkTextAlignment.Value = 0
    chkTextAlignment.Visible = False
End If
End Sub

Private Sub chkBar128_Click()
    If chkBar128.Value = 0 Then
        Picture1.Visible = False
    Else
        Picture1.Visible = True
    End If

End Sub

Private Sub chkBar39_Click()
   
    If chkBar39.Value = 0 Then
        Picture2.Visible = False
    Else
        Picture2.Visible = True
    End If
End Sub

Private Sub chkTextAlignment_Click()
If chkTextAlignment.Value = 1 Then
    chkTextAlignment.Caption = "Bottom Align Caption"
Else
        chkTextAlignment.Caption = "Top Align Caption"
End If
End Sub

Private Sub printBar128()


    'Combination of bar
    'Start Character    3 character (Fixed)
    'Data
    'Check Character    3 character (Depends upon then value of the bar)
    'Stop Character     4 character (Fixed)
    
    
    '//######################################################
'    //PARAMETERS AND THIER MEANING
'    //a=LEFT
'    //b=TOP
'    //hgt=Height of the Barcode
'    //width=width of the thin Barcode in pixel
'    //r1=ratio of the thick barcode and thin barcode
'    //str=Value of the barcode
'    //align=alignment ofthe text i.e 1=left,2=center,3=right,4=justify
'    //textdisp= text position with respect to barcode i.e 2=TOP or 1=BOTTOM
'    //extra=distance of the text from the barcode
'    //ln=device context of the out put device
'//######################################################

Dim Dl As Long
Dim MinWidth  As Long
Picture1.Cls

    
Dim RT_VAL As RET_VAL
    
    With bar
        .crBack = RGB(255, 255, 255)
        .crFore = RGB(0, 0, 0)
        .lalign = 1 'Alignment of the text
        .lExtra = 4   'Distance between the barcode and the text
        .lheight = 40  'Height of the bar
        .lLeft = 20     'Left Position of the bar in the specified device (here e.g. Pictire1)
        .lR1 = 1        'Ratio between thin and thick bar (Standard all over world)
        .lR2 = 1        'Not necessary
        .lRetHeight = 0 'Returns the actual height of the bar code
        .lRetWidth = 0  'Returns the actual width of the bar code
        .lRotation = 0  'to rotate the bar code 0=0degree , 1=90 degree etc.
        .lShowCheck = 1 'Whether check digit will be displayed or not in the bar
        .lstyle = cboTextStyle.ListIndex + 1    'Bold, Italic, Underline or Strikethrough of text
        .lTop = 1      'Top Position of the bar in the specified device (here e.g. Picture1)
        .ltxtdisp = chkTextAlignment.Value + 1 'Whether text displayed at bottom(1) or top(2) of the bar
        .lWidth = cboBarSize.ListIndex + 1    'Width of thin bar in pixel
        .nsize = 10     'Font Size of bar
        .szAdDigit = "" 'Not necessary
        .szBarCaption = IIf(checkBarCaption.Value, Text1.Text, "")
        .szDigit = ""   'Not necessary
        .szReadText = Text1.Text
        .szSymbology = 16
        .TextColor = RGB(255, 0, 0) 'Color of text
        .tiFaceName = "Courier New"       'Font name of text
    End With
    
       Set Target = Picture1
       Dl = Special_128b(bar, Target.hDc)
            
    
    If Dl <> 0 Then MsgBox ErrSpecial_128bMessage(Dl)
End Sub

Private Sub cmdClose_Click()
SndClick
Unload Me
End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
If keycode = vbKeyEscape Then Unload Me
End Sub


Private Sub Form_Load()
cboBarSize.ListIndex = 1
cboTextStyle.ListIndex = 0
End Sub



Private Sub mnuCopyImage1_Click()
    pCopyToClipBoard Picture1
End Sub

Private Sub mnuCopyImage2_Click()
     pCopyToClipBoard Picture2
End Sub

Private Sub mnuSaveImage1_Click()
    'Call FileSaving Routine for BarCode 128
     FileSave Picture1
     If ChkExecute.Value = 1 Then
     'if Asked to Show the exported File
        ShellExecute Me.hWnd, "open", CDialog.FileName, 0&, LastSave, vbNormalFocus
     End If
End Sub

Private Sub mnuSaveImage2_Click()
    'Call FileSaving Routine for BarCode 39
     FileSave Picture2
     If ChkExecute.Value = 1 Then
     'if Asked to Show the exported File
        ShellExecute Me.hWnd, "open", CDialog.FileName, 0&, LastSave, vbNormalFocus
     End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Text1 <> "" Then
         
        PopupMenu mnucp
    End If
End Sub



Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Text1 <> "" Then
    
    PopupMenu mnucp1
End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Text1_Change
' DateTime  : 10/7/2002 19:48
' Author    : Joy
' Purpose   : This Procedure is used to Print Barcode
'            I used Change event to Show you the effects
'---------------------------------------------------------------------------------------
'
Private Sub Text1_Change()

Dim MinWidth, pw
If chkBar128.Value = 1 Then
   printBar128
 End If
 If chkBar39.Value = 1 Then
    Call DrawBarcode(Text1, Picture2)
End If


End Sub
Sub pCopyToClipBoard(ByVal PictureBox As Object)
    ' copy The Image to clipboard
    PictureBox.Picture = PictureBox.Image
    Clipboard.Clear
    Clipboard.SetData PictureBox.Image, 2
End Sub


Sub FileSave(Picbox As PictureBox)
    'This Procedure Saves the Bars to desired Format
    Dim sName, retVal, retSave
    Dim ObjGifImg As GIF
    Screen.MousePointer = 11
    
    CDialog.FilterIndex = 1 'make Dialog as Save
    If LastSave <> "" Then
        CDialog.InitDir = LastSave
    Else
        CDialog.InitDir = App.Path & "\bar_labels"
    End If
    
    On Error GoTo ErrHandler
    CDialog.FileName = Picbox.Name
    CDialog.CancelError = True
    CDialog.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
    CDialog.Filter = "Bitmaps (*.bmp)|*.bmp|Gif (*.gif)|*.gif|Transparent Gif (*.gif)|*.gif"
    CDialog.ShowSave
    'retrive the Folder Name
    retSave = InStrRev(CDialog.FileName, "\")
    LastSave = Mid(CDialog.FileName, 1, retSave)
    DoEvents
    
    Picbox.Picture = Picbox.Image
    Select Case CDialog.FilterIndex
      Case 1: 'if BITMAP is selected
                SavePicture Picbox.Picture, CDialog.FileName
      Case 2: 'if GIF is selected
                Set ObjGifImg = New GIF
                ObjGifImg.SaveGIF Picbox.Image, CDialog.FileName, Picbox.hDc, False, Picbox.Point(0, 0)
                Set ObjGifImg = Nothing
      Case 2: 'if Transperent GIF is selected
                Set ObjGifImg = New GIF
                ObjGifImg.SaveGIF Picbox.Image, CDialog.FileName, Picbox.hDc, True, Picbox.Point(0, 0)
                Set ObjGifImg = Nothing
    End Select

Screen.MousePointer = 0
Exit Sub
ErrHandler:

    If Err.Number = 32755 Then ' Handle the Cancel error
        Screen.MousePointer = 0
        Exit Sub
    Else
            If Err.Number <> 0 Then MsgBox "Error saving file: " & Err.Number & " - " & Err.Description
            Screen.MousePointer = 0
    End If
    
End Sub
