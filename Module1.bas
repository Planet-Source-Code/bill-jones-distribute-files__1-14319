Attribute VB_Name = "Module1"
Option Explicit

Public CurProj As cCurrentProj
Public qT As String

Public VFile1 As String, VFile2 As String

Public CopyCancel As Boolean

Public SourcePath As String
Public pFileName As String
Public Filename As String, Directory As String, FullFileName As String
Public StrucVer As String, FileVer As String, ProdVer As String
Public FileFlags As String, FileOS As String, FileType As String, FileSubType As String
Public FileDate As String

Dim pFileInfo(2) As VS_FIXEDFILEINFO

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Const VS_FFI_SIGNATURE = &HFEEF04BD
Const VS_FFI_STRUCVERSION = &H10000
Const VS_FFI_FILEFLAGSMASK = &H3F&
Const VS_FF_DEBUG = &H1
Const VS_FF_PRERELEASE = &H2
Const VS_FF_PATCHED = &H4
Const VS_FF_PRIVATEBUILD = &H8
Const VS_FF_INFOINFERRED = &H10
Const VS_FF_SPECIALBUILD = &H20
Const VOS_UNKNOWN = &H0
Const VOS_DOS = &H10000
Const VOS_OS216 = &H20000
Const VOS_OS232 = &H30000
Const VOS_NT = &H40000
Const VOS__BASE = &H0
Const VOS__WINDOWS16 = &H1
Const VOS__PM16 = &H2
Const VOS__PM32 = &H3
Const VOS__WINDOWS32 = &H4
Const VOS_DOS_WINDOWS16 = &H10001
Const VOS_DOS_WINDOWS32 = &H10004
Const VOS_OS216_PM16 = &H20002
Const VOS_OS232_PM32 = &H30003
Const VOS_NT_WINDOWS32 = &H40004
Const VFT_UNKNOWN = &H0
Const VFT_APP = &H1
Const VFT_DLL = &H2
Const VFT_DRV = &H3
Const VFT_FONT = &H4
Const VFT_VXD = &H5
Const VFT_STATIC_LIB = &H7
Const VFT2_UNKNOWN = &H0
Const VFT2_DRV_PRINTER = &H1
Const VFT2_DRV_KEYBOARD = &H2
Const VFT2_DRV_LANGUAGE = &H3
Const VFT2_DRV_DISPLAY = &H4
Const VFT2_DRV_MOUSE = &H5
Const VFT2_DRV_NETWORK = &H6
Const VFT2_DRV_SYSTEM = &H7
Const VFT2_DRV_INSTALLABLE = &H8
Const VFT2_DRV_SOUND = &H9
Const VFT2_DRV_COMM = &HA
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

Const OFS_MAXPATHNAME = 128
Const OF_CREATE = &H1000
Const OF_READ = &H0
Const OF_WRITE = &H1
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type




''BROWSE FOR FILE

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


Public Declare Function SHFileExists Lib "Shell32" Alias "#45" (ByVal szPath As String) As Long


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Function DistribFile(FileToSend As String) As String
Dim tmpMsg As String, x As Boolean
    If CurProj.UseVersion Then
        'check the versioning of the source file
        x = CheckFileCopied(SourcePath, FileToSend)
    End If
    
    If x = False Or CurProj.ForceCopy Then
        'x = CopyFile(SourcePath, FileToSend, 0)
        x = CopyThatFile(SourcePath, FileToSend)
    End If
        
    If x Then
        If CurProj.UseVersion Then
            x = CheckFileCopied(SourcePath, FileToSend, CurProj.UseVersion)
        End If
    End If
    
    tmpMsg = vbTab + GetFileDates(FileToSend) + vbCrLf + tmpMsg
    
    If CurProj.ForceCopy Then
        tmpMsg = vbTab & "File Forced!" & vbCrLf & tmpMsg
    End If
    
    If Not CurProj.UseVersion Then
        tmpMsg = vbTab & "Version UNKNOWN" & vbCrLf & tmpMsg
    Else
        tmpMsg = vbTab & "Version=" & VFile2 & vbCrLf & tmpMsg
    End If
    
    If Not x Then
        'file didn't copy
        tmpMsg = vbTab & "COPY FAILED!!" & vbCrLf + tmpMsg
    Else
        tmpMsg = vbTab & "COPY SUCCESSFUL" & vbCrLf + tmpMsg
    End If
    

DistribFile = tmpMsg

End Function


Public Function ReplaceInString(SourceString As String, StringToReplace As String, ReplaceWith As String, Optional AfterPoint As Long) As String
    Dim FrontPart As String
    Dim BackPart As String
    Dim BeginningPoint As Long
    Dim ResumePoint As Long
    Dim ReplaceLength As Long
    Dim WorkingString As String
    
    WorkingString = SourceString
    
DoItAllAgain:
    Select Case AfterPoint
        Case 0
            BeginningPoint = InStr(WorkingString, StringToReplace)
        Case -1
            'replace all occurrances
            BeginningPoint = InStr(WorkingString, StringToReplace)
        Case Else
            If AfterPoint < 0 Then
                BeginningPoint = InStr(ResumePoint, WorkingString, StringToReplace, vbTextCompare)
                AfterPoint = -1
            Else
                BeginningPoint = InStr(AfterPoint, WorkingString, StringToReplace, vbTextCompare)
            End If
    End Select
    If BeginningPoint <> 0 Then
        ReplaceLength = Len(ReplaceWith) ' - Len(StringToReplace)
        
        ResumePoint = BeginningPoint + ReplaceLength
        
        FrontPart = Left$(WorkingString, BeginningPoint - 1)
        BackPart = Mid$(WorkingString, ResumePoint)
        WorkingString = FrontPart & ReplaceWith & BackPart
        
        If AfterPoint = -1 And InStr(ResumePoint, WorkingString, StringToReplace) <> 0 Then
            AfterPoint = -2
            GoTo DoItAllAgain
        End If
    End If
    ReplaceInString = WorkingString
    'Debug.Print ReplaceInString
End Function

Public Sub DisplayVerInfo(FullFileName As String, Index As Integer)
   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      'txtResponses = txtResponses + FullFileName & ":" & vbCrLf & vbTab & "NO VERSION INFO AVAILABLE!" & vbCrLf
      Exit Sub
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

   '**** Determine File Version number ****
   FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)

   '**** Determine Product Version number ****
   ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)

   '**** Determine Boolean attributes of File ****
   FileFlags = ""
   If udtVerBuffer.dwFileFlags And VS_FF_DEBUG Then FileFlags = "Debug "
   If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE Then FileFlags = FileFlags & "PreRel "
   If udtVerBuffer.dwFileFlags And VS_FF_PATCHED Then FileFlags = FileFlags & "Patched "
   If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD Then FileFlags = FileFlags & "Private "
   If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRED Then FileFlags = FileFlags & "Info "
   If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD Then FileFlags = FileFlags & "Special "
   If udtVerBuffer.dwFileFlags And VFT2_UNKNOWN Then FileFlags = FileFlags + "Unknown "

   '**** Determine OS for which file was designed ****
   Select Case udtVerBuffer.dwFileOS
      Case VOS_DOS_WINDOWS16
        FileOS = "DOS-Win16"
      Case VOS_DOS_WINDOWS32
        FileOS = "DOS-Win32"
      Case VOS_OS216_PM16
        FileOS = "OS/2-16 PM-16"
      Case VOS_OS232_PM32
        FileOS = "OS/2-16 PM-32"
      Case VOS_NT_WINDOWS32
        FileOS = "NT-Win32"
      Case Else
        FileOS = "Unknown"
   End Select
   Select Case udtVerBuffer.dwFileType
      Case VFT_APP
         FileType = "App"
      Case VFT_DLL
         FileType = "DLL"
      Case VFT_DRV
         FileType = "Driver"
         Select Case udtVerBuffer.dwFileSubtype
            Case VFT2_DRV_PRINTER
               FileSubType = "Printer drv"
            Case VFT2_DRV_KEYBOARD
               FileSubType = "Keyboard drv"
            Case VFT2_DRV_LANGUAGE
               FileSubType = "Language drv"
            Case VFT2_DRV_DISPLAY
               FileSubType = "Display drv"
            Case VFT2_DRV_MOUSE
               FileSubType = "Mouse drv"
            Case VFT2_DRV_NETWORK
               FileSubType = "Network drv"
            Case VFT2_DRV_SYSTEM
               FileSubType = "System drv"
            Case VFT2_DRV_INSTALLABLE
               FileSubType = "Installable"
            Case VFT2_DRV_SOUND
               FileSubType = "Sound drv"
            Case VFT2_DRV_COMM
               FileSubType = "Comm drv"
            Case VFT2_UNKNOWN
               FileSubType = "Unknown"
         End Select
      Case VFT_FONT
         FileType = "Font"
      Case VFT_VXD
         FileType = "VxD"
      Case VFT_STATIC_LIB
         FileType = "Lib"
      Case Else
         FileType = "Unknown"
   End Select
   pFileInfo(Index) = udtVerBuffer
End Sub

Public Function CheckFileCopied(SrcName As String, DestName As String, Optional UseVersionCheck As Boolean = True) As Boolean
    
    ProdVer = ""
    If UseVersionCheck Then
        DisplayVerInfo DestName, 1
        VFile2 = ProdVer    'Format$(pFileInfo(1).dwProductVersionMSh) & "." & Format$(pFileInfo(1).dwProductVersionMSl) & "." & Format$(pFileInfo(1).dwProductVersionLSh) & "." & Format$(pFileInfo(1).dwProductVersionLSl)
        
        If VFile1 <> VFile2 Then
            CheckFileCopied = False
        Else
            CheckFileCopied = True
        End If
    Else
        CheckFileCopied = True
    End If
    
    If CurProj.ForceCopy Then CheckFileCopied = False
    
End Function


Public Function GetFileDates(FileNm As String) As String
    Dim FileLng As Long, x&, temp$
    Dim FileOpenStruct As OFSTRUCT
    Dim FileCreateStruct As FILETIME, FileLastAccStruct As FILETIME, FileLastWrite As FILETIME
    Dim lFileCreateStruct As FILETIME, lFileLastAccStruct As SYSTEMTIME, lFileLastWrite As FILETIME
    
    FileLng = OpenFile(FileNm, FileOpenStruct, OF_READ)
    x = GetFileTime(FileLng, FileCreateStruct, FileLastAccStruct, FileLastWrite)
    x = FileTimeToLocalFileTime(FileLastWrite, lFileLastWrite)
    x = FileTimeToSystemTime(lFileLastWrite, lFileLastAccStruct)
    With lFileLastAccStruct
        temp = Format$(Trim$(.wHour & ":" & .wMinute & ":" & .wSecond), "long time")
        temp = .wMonth & "\" & .wDay & "\" & .wYear & " " & temp
    End With
    temp = Format$(temp, "long date")
    GetFileDates = temp
    'Debug.Print FileLastAccStruct.dwHighDateTime, FileLastAccStruct.dwLowDateTime
    CloseHandle FileLng
End Function


Public Function GetFolderPath() As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = frmMain.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("Select Distribution ", "Path")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
        '.lpfnCallback = (addressof BrowseFolderCallBack)
        '.lParam = lstrcat(frmMain.DataList1.Columns(0), "")
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    GetFolderPath = sPath
End Function

Public Sub Main()
    'App.FileDescription = "Application for distributing a file to a defined list of locations."
    qT = Chr$(34)
    Set CurProj = New cCurrentProj
    frmMain.Show
End Sub


Public Function ReQryPaths(IdVal As Long) As String
    Dim temp$
    
    temp = "Select * from MachinePaths where ((Proj = " & IdVal & "));"
    ReQryPaths = temp
End Function

Public Sub UpdateDB(DBPath As String)
    
    Dim cat     As New ADOX.Catalog
    Dim tbl     As New ADOX.Table
    Dim AxDBase     As New ADODB.Recordset
    Dim Colmn As New ADOX.Column
    Dim TestTbl As ADOX.Table, TableExists As Boolean, x%
    
   ' Open the catalog
    cat.ActiveConnection = DBPath
    For Each TestTbl In cat.Tables
        If TestTbl.Name = "ProjectInfo" Then
            TableExists = True
            Exit For
        End If
    Next
    
    If TableExists = False Then
        With tbl
            .Name = "ProjectInfo"
            Set .ParentCatalog = cat
                ' Create fields and append them to the new Table object.
                .Columns.Append "Id", adInteger
                ' Make the ContactId column and auto incrementing column
                .Columns("Id").Properties("AutoIncrement") = True
                .Columns.Append "ProjectName", adWChar, 255&
        End With
        cat.Tables.Append tbl
        
        'now add a field to the Machine table to accomodate Project assignment
        'and update all existing records to 1- also set Project name to 'Default'
        AxDBase.Open "ProjectInfo", cat.ActiveConnection, , adLockBatchOptimistic, adCmdTable
        AxDBase.AddNew
            AxDBase.Fields(1) = "Default"
            x = AxDBase.Fields("ID")
        AxDBase.UpdateBatch
        AxDBase.Close
        
        
       Set tbl = cat.Tables("MachinePaths")
        Colmn.Name = "Proj"
        Colmn.Type = adInteger

        tbl.Columns.Append Colmn
        Set tbl = Nothing
        
        AxDBase.Open "MachinePaths", cat.ActiveConnection, , adLockBatchOptimistic, adCmdTable
        AxDBase.MoveLast
        AxDBase.MoveFirst
        Do While AxDBase.EOF = False
             AxDBase.Fields("Proj") = x
             AxDBase.UpdateBatch
             AxDBase.MoveNext
        Loop
        AxDBase.Close
    End If
   Set cat = Nothing
End Sub

