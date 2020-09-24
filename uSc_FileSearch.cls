VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "uSc_FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Builded from File Search OCX from www.Planet-Source-Code.com

Option Explicit
    
    Public Enum Search_Files_Or_Folders
        uSc_FOF_Both = 0
        uSc_FOF_Files = 1
        uSc_FOF_Folders = 2
    End Enum
    
    Public Event FileFound(Path As String) ' File Found
    Public Event FolderFound(Path As String) ' Folder Found
    Public Event SearchComplete(NoFiles As Double, NoFolders As Double, Files As String, Folders As String, Canceled As Boolean)  ' Search Completed, returns files and folders in a ", " delimited string
    Public Event CurrentFolder(Path As String)
    
    Dim SearchCancel As Byte
    
    'Dim Files() As String ' Lists all found filenames
    'Dim Folders() As String ' Lists all found Folders
    Dim NoFiles As Double ' Number of founf files
    Dim NoFolders As Double ' Number of found Folders
    
    Dim Srch As String
    
    Dim FOF As Search_Files_Or_Folders 'src argument
    
    Dim fContaining As String ' searched string (src argument)
    
'Default Property Values:
Const def_Depth = -1
Const def_EnableDoEvents = True
Const def_fCompare = vbTextCompare
'Property Variables:
Public Depth As Integer
Public EnableDoEvents As Boolean
Public fCompare As VbCompareMethod
    
Public Sub Src(BaseDir As String, SearchString As String, Optional FileContaining = "", Optional FilesOrFolders As Search_Files_Or_Folders) ' Not set search string to '' to search all drives
    Dim Fsys As New FileSystemObject
    Dim Drive As Drive
        
    SearchCancel = 0
    
    FOF = FilesOrFolders

    'initialise variables
    Srch = LCase(SearchString)
    fContaining = FileContaining
    NoFiles = 0
    NoFolders = 0
    ReDim Files(1 To 1)
    ReDim Folders(1 To 1)
    
    'begin search
    If BaseDir <> "" Then ' Searches from specified location
        Search BaseDir, Depth
    Else ' Searches all drives one by one
        For Each Drive In Fsys.Drives
            If Drive.IsReady Then ' If the drive is ready, search it, otherwise dont
                Search Drive.Path, Depth
            End If
        Next Drive
    End If
    RaiseEvent SearchComplete(NoFiles, NoFolders, Join(Files, ", "), Join(Folders, ", "), SearchCancel = 1)
End Sub
Private Sub Search(BaseDir As String, ToDepth As Integer)

'A Recursive procedure (meaning it calls itself) that performs the actual search
    Dim strFile As String
    Dim strFilePath As String
    Dim strFolder As String
    Dim strFolderPath As String
    
    Dim SubFolders() As String
    Dim NoSubFolders As Long
    
    Dim Fsys As New FileSystemObject
    Dim Buffer As String ' General Purpose variable
    
    BaseDir = AddBackSlash(BaseDir)
    
    RaiseEvent CurrentFolder(BaseDir)
    
If FOF = 0 Or FOF = 1 Then
    strFile = Dir(BaseDir & Srch)
    Do
        If EnableDoEvents Then DoEvents
        If SearchCancel = 1 Or strFile = "" Then Exit Do
        Buffer = LCase(strFile)
            strFilePath = BaseDir & strFile
            If FileCondition(strFilePath) Then
                NoFiles = NoFiles + 1
                'ReDim Preserve Files(1 To NoFiles)
                'Files(NoFiles) = strFilePath
                RaiseEvent FileFound(strFilePath)
            End If
        strFile = Dir
    Loop While 1
End If


If FOF = 0 Or FOF = 2 Then
    strFolder = Dir(BaseDir & Srch, vbDirectory)
    Do While SearchCancel = 0 And strFolder <> ""
        If strFolder = "." Or strFolder = ".." Or Not ((GetAttr(BaseDir & strFolder) And vbDirectory) = vbDirectory) Then GoTo SForceLoop1
        If EnableDoEvents Then DoEvents
        Buffer = LCase(strFolder)
            strFolderPath = BaseDir & strFolder
            NoFolders = NoFolders + 1
            'ReDim Preserve Folders(1 To NoFolders)
            'Folders(NoFolders) = strFolderPath
            RaiseEvent FolderFound(strFolderPath)
SForceLoop1:
        strFolder = Dir
    Loop
End If

If ToDepth Then
    NoSubFolders = 0
    
    strFolder = Dir(BaseDir, vbDirectory)
    Do While SearchCancel = 0 And strFolder <> ""
        If strFolder = "." Or strFolder = ".." Or Not ((GetAttr(BaseDir & strFolder) And vbDirectory) = vbDirectory) Then GoTo SForceLoop2
        If EnableDoEvents Then DoEvents
        strFolderPath = BaseDir & strFolder
        NoSubFolders = NoSubFolders + 1
        ReDim Preserve SubFolders(1 To NoSubFolders)
        SubFolders(NoSubFolders) = strFolderPath
SForceLoop2:
        strFolder = Dir
    Loop
    
    Dim i As Long
    For i = 1 To NoSubFolders
        Search SubFolders(i), ToDepth - 1
    Next
End If
End Sub


Private Function FileCondition(FileName As String) As Boolean
    Dim TheLine As String
    
    If fContaining = "" Then
        FileCondition = True: Exit Function
    Else
        FileCondition = False
    End If
    
    Open FileName For Input As #1
    
    On Error GoTo FCerr
    
    Do While Not EOF(1)
        Input #1, TheLine
        If InStr(1, TheLine, fContaining, fCompare) <> 0 Then
            FileCondition = True
            Exit Do
        End If
        If EOF(1) Then Exit Do
    Loop
    
FCerr:
    Close #1
End Function

Public Sub SrcCancel()
    SearchCancel = 1
End Sub


'extract all parts from a file name (DriveLetter, DirPath, FName and Extension)
Private Sub FileNameExtractAllParts(ByVal AccessName As String, _
                                    ByRef DriveLetter As String, _
                                    ByRef DirPath As String, _
                                    ByRef FName As String, _
                                    ByRef Extension As String)
  Dim PathLength As Integer
  Dim ThisLength As Integer
  Dim Offset As Integer
  Dim FileNameFound As Boolean, DotFound As Boolean

  DriveLetter = ""
  DirPath = ""
  FName = ""
  Extension = ""

  If Mid(AccessName, 2, 1) = ":" Then 'find the drive letter.
    DriveLetter = Left(AccessName, 2)
    AccessName = Mid(AccessName, 3)
  End If

  PathLength = Len(AccessName)

  DotFound = False
  For Offset = PathLength To 1 Step -1 'find the next delimiter.
    Select Case Mid(AccessName, Offset, 1)
      Case "."
        'this indicates either an extension or a . or a ..
        If Not DotFound Then
          ThisLength = Len(AccessName) - Offset
          If ThisLength >= 1 Then ' Extension
            Extension = Mid(AccessName, Offset, ThisLength + 1)
          End If
          AccessName = Left(AccessName, Offset - 1)
          DotFound = True
        End If
      Case "\"
        'this indicates a path delimiter.
        ThisLength = Len(AccessName) - Offset
        If ThisLength >= 1 Then ' Filename
          FName = Mid(AccessName, Offset + 1, ThisLength)
          AccessName = Left(AccessName, Offset)
          FileNameFound = True
          Exit For
        End If
    End Select
  Next Offset
  If FileNameFound = False Then
    FName = AccessName
  Else
    DirPath = AccessName
  End If
  If (Len(Extension) = 0) And (Len(FName) > 0) And (InStr(FName, ".") = 0) Then
    DirPath = DirPath & FName
    FName = ""
  End If
End Sub

'extract path name only from a full file name with path
Private Function ExtractPathName(ByVal FullFileName As String) As String
  Dim DriveLetter As String, DirPath As String, FName As String, Extension As String
  Call FileNameExtractAllParts(FullFileName, DriveLetter, DirPath, FName, Extension)
  ExtractPathName = AddBackSlash(DriveLetter & DirPath)
End Function

'extract file name only from a full file name with path
Private Function ExtractFileName(ByVal FullFileName As String) As String
  Dim DriveLetter As String, DirPath As String, FName As String, Extension As String
  Call FileNameExtractAllParts(FullFileName, DriveLetter, DirPath, FName, Extension)
  ExtractFileName = FName & Extension
End Function


Private Function AddBackSlash(ByVal strPath As String) As String
  If (Len(strPath) > 0) And (Right(strPath, 1) <> "\") Then
    AddBackSlash = strPath & "\"
  Else
    AddBackSlash = strPath
  End If
End Function

Private Sub Class_Initialize()
    fCompare = def_fCompare
    EnableDoEvents = def_EnableDoEvents
    Depth = def_Depth
End Sub
