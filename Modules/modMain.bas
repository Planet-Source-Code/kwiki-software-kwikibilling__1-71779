Attribute VB_Name = "modMain"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Public Const SND_SYNC = &H0

Public DB As Database
Public gblnPopulating   As Boolean

Public KD As Database


' Database Config
Public Function OpenDatabase() As Boolean
On Error GoTo DBErrors:
Dim iniFile As String, ItsThere As Boolean
Dim DBPath As String
iniFile = App.Path & "\Settings.ini"

ItsThere = FileExists(iniFile)
If ItsThere = False Then

WriteINI "DBPath", "Path", App.Path & "\Kwikidat\db2.mdb", iniFile
DBPath = ReadINI("DBPath", "Path", iniFile)
Else
DBPath = ReadINI("DBPath", "Path", iniFile)

Set DB = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\KwikiDat\db2.mdb", False)
End If

OpenDatabase = True
Exit Function
DBErrors:
OpenDatabase = False
MsgBox (Err.Description)
End Function

Public Function CreateKeyDat()
On Error GoTo KeyErrors:

Set KD = DBEngine.Workspaces(0).CreateDatabase(App.Path & "\KwikiDat\key.mdb", dbLangGeneral)
KD.Close

KeyErrors:
End Function

Public Function OpenKeyDat() As Boolean
On Error GoTo KeyErrors:

Set KD = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\KwikiDat\key.mdb", False)

OpenKeyDat = True
Exit Function
KeyErrors:
OpenKeyDat = False
MsgBox (Err.Description)
End Function

'************************************************************************
'*          Global Variable and Constant Declarations                   *
'************************************************************************
Function ReadINI(Section As String, KeyName As String, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim R
    R = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Public Function SndPlayEx(ByVal FileName As String, Optional ByVal lmodule As Long = 0, Optional ByVal options As Long = (SND_FILENAME Or SND_ASYNC)) As Long
SndPlayEx = PlaySound(FileName, lmodule, options)
End Function

Public Function SndClick()
Dim iniFile As String, ItsThere As Boolean
iniFile = App.Path & "\Settings.ini"

If frmCompanySetup.Text10 = "" Then
WriteINI "SndPath", "Path", App.Path & "\Sounds\Start.wav", iniFile
SndPlayEx App.Path & "\Sounds\Start.wav"
Else

SndPlayEx ReadINI("SndPath", "Path", iniFile)

End If
End Function
'=============================================================================
'                     General Routines
'=============================================================================

'-----------------------------------------------------------------------------
Public Function GetAppPath() As String
'-----------------------------------------------------------------------------

    Dim strAppPath As String
    
    strAppPath = App.Path
    
    GetAppPath = strAppPath

End Function

'-----------------------------------------------------------------------------
Public Sub SelectTextboxText(pobjTextbox As TextBox)
'-----------------------------------------------------------------------------
    With pobjTextbox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'-----------------------------------------------------------------------------
Public Sub TabToNextTextBox(pobjTextBox1 As TextBox, pobjTextBox2 As TextBox)
'-----------------------------------------------------------------------------

    If gblnPopulating Then Exit Sub
    
    If pobjTextBox2.Enabled = False Then Exit Sub
    
    If Len(pobjTextBox1.Text) = pobjTextBox1.MaxLength Then
        pobjTextBox2.SetFocus
    End If

End Sub

Public Function ValidateRequiredField(pobjCtl As Control, _
                                      pstrFieldDesc As String) _
As Boolean
'-----------------------------------------------------------------------------

    If (pobjCtl.Text) = "" Then
        MsgBox pstrFieldDesc & " must not be blank.", _
               vbExclamation, _
               pstrFieldDesc
        pobjCtl.SetFocus
        ValidateRequiredField = False
    Else
        ValidateRequiredField = True
    End If

End Function

'-----------------------------------------------------------------------------
Public Function ValidateZipCode(pobjZip5 As Control, _
                                Optional pobjZip4 As Control = Nothing) _
As Boolean
'-----------------------------------------------------------------------------
    
    Dim strErrorMsg     As String
    Dim objFocusControl As Control
    
    If Not pobjZip4 Is Nothing Then
        If (pobjZip5.Text) = "" _
        And (pobjZip4.Text) <> "" _
        Then
            strErrorMsg = "First part of Zip must be valued when '+4' part " _
                        & "is valued."
            Set objFocusControl = pobjZip5
            GoTo ValidateZipCode_Error
        End If
    End If
    
    If Len(pobjZip5.Text) = 0 _
    Or Len(pobjZip5.Text) = 5 _
    Then
        ' it's OK
    Else
        strErrorMsg = "Invalid length for Zip."
        Set objFocusControl = pobjZip5
        GoTo ValidateZipCode_Error
    End If
    
    If Not pobjZip4 Is Nothing Then
        If Len(pobjZip4.Text) = 0 _
        Or Len(pobjZip4.Text) = 4 _
        Then
            ' it's OK
        Else
            strErrorMsg = "Invalid length for Zip '+4' part."
            Set objFocusControl = pobjZip4
            GoTo ValidateZipCode_Error
        End If
    End If
    
    ValidateZipCode = True
    Exit Function
    
ValidateZipCode_Error:
    MsgBox strErrorMsg, _
           vbExclamation, _
           "Zip"
    objFocusControl.SetFocus
    ValidateZipCode = False

End Function

Public Function ValidateDate(pobjMonth As Control, _
                             pobjDay As Control, _
                             pobjYear As Control, _
                             Optional pblnBlankOK As Boolean = True) _
As Boolean
'-----------------------------------------------------------------------------

    Dim strErrorMsg     As String
    Dim objFocusControl As Control

    If pobjMonth.Text = "" _
    And pobjDay.Text = "" _
    And pobjYear.Text = "" Then
        If pblnBlankOK Then
            ValidateDate = True
            Exit Function
        Else
            strErrorMsg = "Date must not be blank."
            Set objFocusControl = pobjMonth
            GoTo ValidateDate_Error
        End If
    End If

    If pobjYear.Text <> "" And Len(pobjYear.Text) <> 4 Then
        strErrorMsg = "Four digits must be entered for the year."
        Set objFocusControl = pobjYear
        GoTo ValidateDate_Error
    End If

    If Not IsDate(pobjMonth.Text & "/" _
                & pobjDay.Text & "/" _
                & pobjYear.Text) Then
        strErrorMsg = "Date must be a valid date" & IIf(pblnBlankOK, " or blank", "") & "."
        Set objFocusControl = pobjMonth
        GoTo ValidateDate_Error
    End If

    ValidateDate = True
    Exit Function

ValidateDate_Error:
    MsgBox strErrorMsg, _
           vbExclamation, _
           "Date Error"
    objFocusControl.SetFocus
    ValidateDate = False

End Function

'-----------------------------------------------------------------------------
Public Function ValidateSSN(pobjSSN1 As Control, _
                            pobjSSN2 As Control, _
                            pobjSSN3 As Control) _
As Boolean
'-----------------------------------------------------------------------------

    Dim strErrorMsg         As String
    Dim strErrorField       As String
    Dim blnIncompleteSSN    As Boolean
    Dim objFocusControl     As Control
    
    If (Trim$(pobjSSN1.Text) = "" And _
        Trim$(pobjSSN2.Text) = "" And _
        Trim$(pobjSSN3.Text) = "") _
    Or (Trim$(pobjSSN1.Text) <> "" And _
        Trim$(pobjSSN2.Text) <> "" And _
        Trim$(pobjSSN3.Text) <> "") _
    Then
        If Trim$(pobjSSN1.Text) <> "" Then
            blnIncompleteSSN = False
            If Len(pobjSSN1.Text) <> 3 Then
                Set objFocusControl = pobjSSN1
                blnIncompleteSSN = True
            ElseIf Len(pobjSSN2.Text) <> 2 Then
                Set objFocusControl = pobjSSN2
                blnIncompleteSSN = True
            ElseIf Len(pobjSSN3.Text) <> 4 Then
                Set objFocusControl = pobjSSN3
                blnIncompleteSSN = True
            End If
            If blnIncompleteSSN Then
                strErrorMsg = "SSN entry is incomplete."
                strErrorField = "SSN"
                GoTo ValidateSSN_Error
            End If
        End If
    Else
        strErrorMsg = "A partial SSN is not valid. " _
                    & "Either fill all parts, or leave all parts blank."
        strErrorField = "SSN"
        Set objFocusControl = pobjSSN1
        GoTo ValidateSSN_Error
    End If
    
    ValidateSSN = True
    Exit Function

ValidateSSN_Error:
    MsgBox strErrorMsg, _
           vbExclamation, _
           "SSN Error"
    objFocusControl.SetFocus
    ValidateSSN = False

End Function


Public Function FileExists(sFileName As String) As Boolean
   '** Description:
   '** Check to see if file exists
   On Error GoTo FExistsError
   
   Dim F As String
   
   F = FreeFile
   Open sFileName For Input As #F 'Open file
   Close #F
FExistsError:
   If Err.Number = 53 Then 'If doesn't exists
      FileExists = False 'Set FileExists to False
   ElseIf Err.Number = 0 Then 'else if exists
      FileExists = True 'Set FileExists to True
   End If
End Function

Sub UnloadAllForms()
On Error Resume Next
Dim i As Integer
For i = Forms.count - 1 To 1 Step -1
Unload Forms(i)
Next
End Sub



