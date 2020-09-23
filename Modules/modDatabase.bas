Attribute VB_Name = "modDatabase"
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'for data backup & restore operations
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_SILENT = &H4
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_FILESONLY = &H80

Type SHFILEOPSTRUCT
   hWnd      As Long
   wFunc     As Long
   pFrom     As String
   pTo       As String
   fFlags    As Integer
   fAborted  As Boolean
   hNameMaps As Long
   sProgress As String
End Type

'----------------DATABASE UTILITIES-------------------------

Sub RestoreDatabase()
   'restore the user database
   On Error Resume Next
   
   Dim lFileOp  As Long
   Dim lresult  As Long
   Dim lFlags   As Long
   Dim SHFileOp As SHFILEOPSTRUCT
   
   Screen.MousePointer = vbHourglass
   MsgBar "Restoring Data Files", True
   
   lFileOp = FO_COPY
   lFlags = lFlags And Not FOF_SILENT
   lFlags = lFlags Or FOF_NOCONFIRMATION
   lFlags = lFlags Or FOF_NOCONFIRMMKDIR
   lFlags = lFlags Or FOF_FILESONLY
   
   With SHFileOp
      .wFunc = lFileOp
      .pFrom = App.Path & "\KwikiDat\db2BACKUP.mdb" & vbNullChar
      .pTo = App.Path & "\KwikiDat\db2.mdb" & vbNullChar
      .fFlags = lFlags
   End With
   
   lresult = SHFileOperation(SHFileOp)
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   frmMaintain.Show
End Sub

Sub BackupDatabase()
   'backup the current user databsae
   On Error Resume Next
   
   Dim lFileOp  As Long
   Dim lresult  As Long
   Dim lFlags   As Long
   Dim SHFileOp As SHFILEOPSTRUCT
   
   Screen.MousePointer = vbHourglass
   MsgBar "Backing Up Data Files", True
   
   lFileOp = FO_COPY
   lFlags = lFlags And Not FOF_SILENT
   lFlags = lFlags Or FOF_NOCONFIRMATION
   lFlags = lFlags Or FOF_NOCONFIRMMKDIR
   lFlags = lFlags Or FOF_FILESONLY
   
   With SHFileOp
      .wFunc = lFileOp
      .pFrom = App.Path & "\KwikiDat\db2.mdb" & vbNullChar
      .pTo = App.Path & "\KwikiDat\db2BACKUP.mdb" & vbNullChar
      .fFlags = lFlags
   End With
   
   lresult = SHFileOperation(SHFileOp)
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Sub CompactDB()
   On Error GoTo Error_Handler
   
   Dim sOldName As String
   Dim sNewName As String
   Dim sNewName2 As String
   'Dim nEncrypt As Integer
   
   'the file name to compact
   sOldName = App.Path & "\KwikiDat\db2.mdb"
   
   'the file name to compact to
   sNewName = App.Path & "\KwikiDat\db2COMPACTED.mdb"
   
   Screen.MousePointer = vbHourglass
   MsgBar "Compacting " & " Database.", True
   'we are going to overwrite the same file, so we need to create a new MDB
   'and rename after the compact is successful
   If sOldName = sNewName Then
      sNewName2 = sNewName 'save the new name
      sNewName = Left(sNewName, Len(sNewName) - 1) & "N"
   End If
   
   'unload all forms & close the database
   UnloadAllForms
   DB.Close
   Set DB = Nothing
   
   DBEngine.CompactDatabase sOldName, sNewName, dbLangGeneral, dbVersion60
   
   'check for an overwrite of the original mdb
   If VBA.Right(sNewName, 1) = "N" Then
      Kill sNewName2             'nuke the old one
      Name sNewName As sNewName2 'rename the new one to the original name
      sNewName = sNewName2       'reset to the correct name
   End If
   
   're-open the compacted database
   OpenDatabase
   If (Not OpenDatabase()) Then
   MsgBox "Database could not be opened"
   End If
   Load MDIForm1
   MDIForm1.UpdateTree
   
   MsgBar vbNullString, False
   Screen.MousePointer = vbDefault
   
   MsgBox "The database was sucessfully compacted"
   
   Exit Sub
   
Error_Handler:
   MsgBox "An un-known error occurred while Compacting the Database!" & vbCrLf _
      & "Sorry for the inconvenience"
End Sub

'----------------END DATABASE UTILITIES-------------------------

Sub MsgBar(rsMsg As String, rPauseFlag As Integer)
   If Len(rsMsg) = 0 Then
      Screen.MousePointer = vbDefault
   Else
      If rPauseFlag Then
         frmDB.lblStatus.Caption = rsMsg & ", please wait..."
      Else
         frmDB.lblStatus.Caption = rsMsg
      End If
   End If
End Sub

