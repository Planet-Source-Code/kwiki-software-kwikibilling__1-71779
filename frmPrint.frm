VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printing"
   ClientHeight    =   1875
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4845
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Bar Code"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.PictureBox imgPhoto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdPrint_Click()
cmdPrint.Visible = False
Me.PrintForm
End Sub

Private Sub Form_Load()
imgPhoto.Picture = frmParts.imgPhoto.Picture
End Sub
