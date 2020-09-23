VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Forces variable declaration

'Events
'^^^^^^^
Public Event Search(ByVal Path As String, Files() As String) 'Raised whenever the search function is done processing a folder

'APIs
'^^^^^
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long 'Creates a file and returns a handle to it
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long 'Closes an opened handle
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long 'Finds the executable associated with the specified file
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long 'Retrieves information on a file matching the specified criteria
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long 'Retrieve the next file matching the specified criteria for a search handle
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long 'Closes a file handle opened by FindFirstFile
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long 'Performs correct file to system time convertion
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long 'Checks whether the specified process is still running
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long 'Retrieves a file's attributes
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long 'Retrieves all available drives
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long 'Retrieves Drive Space Information
Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal Drive As String) As Long 'Gets the type of the specified file
Private Declare Function GetLastError Lib "kernel32" () As Long 'Retrieves last error information
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long 'Retrieves the long path associated with the specified short path
Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long 'Gets a single file name available in the temporary folder
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long 'Gets the path of the Temporary folder
Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long 'Gets the ms-dos short path from a valid, full path
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long 'Retrieves Drive Information
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long 'Converts the local file time to GMT
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long 'Retrieves a handle to a process so we can query its status
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long 'Sets a file's attributes
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long 'Sets the created/accessed/modified dates of a file
Private Declare Function SetVolumeLabel Lib "kernel32.dll" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long 'Sets a new label for the specified volume
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Opens a file/folder/url with its specified application
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long 'Performs the specified operation in a file
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long 'Retrieves information regarding the specified file
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long 'Resolves the value returned by SHGetSpecialFolderLocation into a valid path
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long 'Gets the path of the specified system folder
Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long 'Converts the system time to the file time formay
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) 'Gets the current system time

'ENUMs
'^^^^^^

'List Items Group

Public Enum ListItemsGroup

LIG_ListFiles
LIG_ListFolders
LIG_ListAllItems

End Enum

'Special Folder Constants

Public Enum SpecialFolderConst

SFC_WindowsFolder = &H24
SFC_SystemFolder = &H25
SFC_Desktop = &H0
SFC_MyDocuments = &H5
SFC_ProgramFiles = &H26
SFC_UserDirectory = &H28
SFC_StartupFolder = &H7
SFC_RecycleBin = &HA
SFC_TempFolder = &HF
CSIDL_RECENT = &H8
CSIDL_HISTORY = &H22
End Enum

'File Data Information

Public Enum FileDateInfo

FDI_DateCreated
FDI_DateAccessed
FDI_DateModified

End Enum

'Drive Type Values

Public Enum DriveTypeValues

DTV_Unknown
DTV_NotFound
DTV_Removable
DTV_Fixed
DTV_Remote
DTV_CDROM
DTV_RAMDisk

End Enum

'File Mode Constants

Public Enum FileModeConst

FM_Binary
FM_Write
FM_Append

End Enum

'Constants
'^^^^^^^^^^

'General

Private Const REM_ATTR As Integer = 7 'vbReadOnly + vbHidden + vbSystem
Private Const DEF_ATTR As Integer = 55 'vbReadOnly + vbHidden + vbSystem + vbDirectory + vbArchive
Private Const FOLDERS_OFFSET As Integer = 20 'Amount of items to allocate per cycle for subfolders
Private Const ITEMS_OFFSET As Integer = 100 'Amount of items to allocate per cycle for normal results
Private Const MAX_PATH As Long = 260 'Specifies the maximum size of a file path

'Get Path Type

Const SHGFI_DISPLAYNAME As Long = &H200
Const SHGFI_TYPENAME As Long = &H400

'Execute

Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const STILL_ACTIVE  As Long = &H103

'Open File

Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
Private Const OPEN_EXISTING As Long = 3
Private Const ERROR_SHARING_VIOLATION As Long = 32

'Move To Recycle Bin

Private Const FO_DELETE As Long = &H3
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_SILENT As Long = &H4
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_RENAMEONCOLLISION As Long = &H8

'Types
'^^^^^^

'File Information

Private Type SHFILEINFO

hIcon As Long
iIcon As Long
dwAttributes As Long
szDisplayName As String * MAX_PATH
szTypeName As String * 80

End Type

'Get Special Folder

Private Type SHITEMID

cb As Long
abID As Byte
    
End Type

Private Type ITEMIDLIST
    
mkid As SHITEMID

End Type

'Get Path Date Info

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

'Find Files

Private Type FILETIME

dwLowDateTime As Long
dwHighDateTime As Long

End Type

Private Type WIN32_FIND_DATA

dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * 260
cAlternate As String * 14
    
End Type

'Move To Recycle Bin

Private Type SHFILEOPSTRUCT

hwnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAborted As Boolean
hNameMaps As Long
sProgress As String

End Type

'General Variables
'^^^^^^^^^^^^^^^^^^

'Properties

Private m_blnBreakOnErrors As Boolean 'Specifies wether to abort the current operation should any errors occur, or to ignore them
Private m_objLastError As ErrObject 'Holds information regarding the last ocurring error

'PROPERTIES
'***********

'Break On Errors
'^^^^^^^^^^^^^^^^
Public Property Get BreakOnErrors() As Boolean

BreakOnErrors = m_blnBreakOnErrors

End Property

Public Property Let BreakOnErrors(ByVal blnBreakOnErrors As Boolean)

m_blnBreakOnErrors = blnBreakOnErrors

End Property

'Last Error
'^^^^^^^^^^^
Public Property Get LastError() As ErrObject

Set LastError = m_objLastError

End Property

'FUNCTIONS AND SUBROUTINES
'**************************

'*******************
'GENERAL FUNCTIONS
'*******************

'Function: Exists
'*****************
'Objective: Determines wether the specified path is valid
'Usage: blnRet = Exists(ByVal Path as String)
'Assumptions:
 'I. Wildcards may be used within Path
'Parameters:
 'I. Path: Path to verify
'Returns: True if the specified path is valid, false otherwise
'---------------------------------------------------------------------------------
Public Function Exists(ByVal Path As String) As Boolean

On Error Resume Next

Exists = (Dir$(Path, DEF_ATTR) <> vbNullString) 'Checks if the file exists

End Function

'Function: FixPath
'******************
'Objective: Appends a backslash to the specified path if needed
'Usage: strRet = FixPath(ByVal Path as String)
'Assumptions: None
'Parameters:
 'I. Path: Path to check
'Returns: String containing the fixed path
'---------------------------------------------------------------------------------
Public Function FixPath(ByVal Path As String) As String

If Len(Path) = 1 Then Path = Path & ":" 'Fixes invalid drive paths

If Not right(Path, 1) = "\" Then FixPath = Path & "\" Else FixPath = Path 'Appends a backslash to the specified path, if needed

End Function

'Function: GetPathName
'**********************
'Objective: Returns the Name (File Name + [Extension]) of the specified Path
'Usage: strRet = GetPathName(ByVal Path as String)
'Assumptions:
 'I. If the specifies path contains an extension it will be returned as well. To get only the File Name of a Path use GetBaseName(Path)
'Parameters:
 'I. Path: Path to extract the name from
'Returns: Name portion of the specified path
'---------------------------------------------------------------------------------
Public Function GetPathName(ByVal Path As String) As String

If Len(Path) < 4 Then GetPathName = Path: Exit Function 'If path's a drive then cancel operation

If right$(Path, 1) = "\" Then Path = left$(Path, Len(Path) - 1) 'Remove last character from the path, if it's a backslash

Dim intRet As Integer: intRet = InStrRev(Path, "\") 'Get the position of the last path delimiter

GetPathName = Mid$(Path, intRet + 1) 'Extract the name portion

End Function

'Function: GetParentFolder
'**************************
'Objective: Retrieves the folder which holds the specified path
'Usage: strRet = GetParentFolder(ByVal Path as String)
'Assumptions: None
'Parameters:
 'I. Path: Path to extract the parent folder from
'Returns: Specified path's parent folder
'---------------------------------------------------------------------------------
Public Function GetParentFolder(ByVal Path As String) As String

If Len(Path) < 4 Then GetParentFolder = Path: Exit Function 'If the specified path is a drive, there's no parent folder

If right$(Path, 1) = "\" Then Path = left$(Path, Len(Path) - 1) 'Remove last backslash, if there's one

Dim intRet As Integer: intRet = InStrRev(Path, "\") 'Get the position of the last path delimiter

GetParentFolder = left$(Path, intRet) 'Strip off the parent folder's path

End Function

'Function: GetBaseName
'**********************
'Objective: Returns the name portion of the specified path, discarding any extension
'Usage: strRet = GetBaseName(ByVal Path as String)
'Assumptions: None
'Parameters:
 'I. Path: Path to extract the name portion from
'Returns: name portion of the specified path
'---------------------------------------------------------------------------------
Public Function GetBaseName(ByVal Path As String) As String

Dim strPathName As String, intPos As Integer

strPathName = GetPathName(Path) 'Retrieve the name of the specified path

intPos = InStrRev(Path, "."): If intPos = 0 Then intPos = Len(strPathName) 'Find the extension delimiter position

GetBaseName = left$(strPathName, intPos) 'Strip off the base name

End Function

'Function: GetLongPath
'**********************
'Objective: Retrieves the long path of the specified short path
'Usage: strRet = GetLongPath(ByVal ShortPath as String)
'Assumptions:
 'I. The specified path must exist
'Parameters:
 'I. ShortPath: Path to retrieve the long version from
'Returns: The long version of the specified path
'---------------------------------------------------------------------
Public Function GetLongPath(ByVal ShortPath As String) As String

Dim lngRet As Long

GetLongPath = String$(MAX_PATH, vbNullChar) 'Allocate space in our buffer

lngRet = GetLongPathName(ShortPath, GetLongPath, Len(GetLongPath)) 'Retrieve the long path version

If lngRet > Len(GetLongPath) Then 'If for some reason our buffer wasn't large enough

    GetLongPath = String$(lngRet, vbNullChar) 'Allocate more space
    
    lngRet = GetLongPathName(ShortPath, GetLongPath, lngRet) 'Repeat the operation
    
End If

If Not lngRet = 0 Then GetLongPath = left$(GetLongPath, lngRet) 'Strip off null terminators

End Function

'Function: GetShortPath
'***********************
'Objective: Gets the MS-DOS 6 characters ID from a specified path
'Usage: strRet = GetShortPath(ByVal Path as String)
'Assumptions: None
'Parameters:
 'I. Path: Path to analyze
'Returns: Corresponding short path from the parameter's value
'---------------------------------------------------------------------------------
Public Function GetShortPath(ByVal Path As String) As String

Dim lngRet As Long

GetShortPath = String$(MAX_PATH, vbNullChar) 'Fill the buffer with empty characters
    
lngRet = GetShortPathNameA(Path, GetShortPath, MAX_PATH) 'Retrieve the short path of the specified path

If Not lngRet = 0 Then GetShortPath = left$(GetShortPath, lngRet) 'Strip off null terminators

End Function

'Function: GetPathAttributes
'****************************
'Objective: Retrieves the specified path's attributes
'Usage: enmRet = GetPathAttributes(ByVal Path as String) as VbFileAttribute
'Assumptions:
 'I. The specified path must exist in order for the function to succeed
'Parameters:
 'I. Path: Path to get the attributes from
'Returns: If the operation succeeds, the specified path's attributes will be returned, otherwise -1 will be the returned
'---------------------------------------------------------------------------------
Public Function GetPathAttributes(ByVal Path As String) As VbFileAttribute

GetPathAttributes = GetFileAttributes(Path) 'Retrieve the file's attributes

If GetPathAttributes = -1 And Exists(Path) Then 'Sometimes the GetFileAttributes function will return -1 eventhough the file exists

    Dim FileData As WIN32_FIND_DATA, lngRet As Long
    
    lngRet = FindFirstFile(Path, FileData) 'So attempt to retrieve it's attributes with FindFirstFile
    
    If Not lngRet = -1 Then GetPathAttributes = FileData.dwFileAttributes 'Return file attributes if found

End If
    
End Function

'Function: GetAttributeNames
'****************************
'Objective: Returns a string containing the specified attributes' names
'Usage: strRet = GetAttributeNames(ByVal Attributes as Integer)
'Assumptions: None
'Parameters:
 'I. Attributes: Attributes code to evaluate
'Returns: Corresponding Attributes names
'---------------------------------------------------------------------------------
Public Function GetAttributeNames(ByVal Attributes As VbFileAttribute) As String
    
If Attributes < 0 Then Exit Function 'If the specified attribute is invalid, break the operation
 
If Attributes = 0 Then GetAttributeNames = "Normal": Exit Function 'If the attribute is 0 then simply return the file as 'Normal'

If Attributes And vbReadOnly Then GetAttributeNames = GetAttributeNames & "Read Only, " 'File is Read Only
If Attributes And vbHidden Then GetAttributeNames = GetAttributeNames & "Hidden, " 'File is Hidden
If Attributes And vbSystem Then GetAttributeNames = GetAttributeNames & "System, " 'File is a System File
If Attributes And vbVolume Then GetAttributeNames = GetAttributeNames & "Volume, " 'File is a Volume
If Attributes And vbDirectory Then GetAttributeNames = GetAttributeNames & "Directory, " 'File is a directory
If Attributes And vbArchive Then GetAttributeNames = GetAttributeNames & "Archive, " 'File is an archive
If Attributes And vbAlias Then GetAttributeNames = GetAttributeNames & "Alias, " 'File is an alias

GetAttributeNames = left$(GetAttributeNames, Len(GetAttributeNames) - 2) 'Strip off the last ', '
          
End Function

'Function: GetPathDate
'**********************
'Objective: Retrieves information about a path's created/accessed/modified dates
'Usage: dtmRet = GetPathDate(ByVal Path as String, DateInfo as FileDateInfo)
'Assumptions: None
'Parameters:
 'I. Path: Path to retrieve the date information from
 'II. DateInfo: Specifies which date attribute to extract
'Returns: Corresponding date for specified DateInfo
'---------------------------------------------------------------------------------
Public Function GetPathDate(ByVal Path As String, DateInfo As FileDateInfo) As Date

Dim lngHandle As Long, udtSystemTime As SYSTEMTIME, udtPathDateInfo As WIN32_FIND_DATA

lngHandle = FindFirstFile(Path, udtPathDateInfo) 'Open a handle to the specified file

Select Case DateInfo

    Case FDI_DateCreated: Call FileTimeToSystemTime(udtPathDateInfo.ftCreationTime, udtSystemTime) 'Retrieve the file's created date
    Case FDI_DateAccessed: Call FileTimeToSystemTime(udtPathDateInfo.ftLastAccessTime, udtSystemTime) 'Retrieve the file's last accessed date
    Case FDI_DateModified: Call FileTimeToSystemTime(udtPathDateInfo.ftLastWriteTime, udtSystemTime) 'Retrieve the file's last modified date
            
End Select

'Re-arrange and place file's date into function's return value

GetPathDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour + AdjustTimeForLocalSettings, udtSystemTime.wMinute, udtSystemTime.wSecond)

Call FindClose(lngHandle) 'Close file handle

End Function

'Function: GetPathType
'**********************
'Objective: Retrieves the description associated with the specified file's type
'Usage: strRet = GetPathType(ByVal Path as String)
'Assumptions:
 'I. The specified file must exist, or the type will be returned according to the file's extension
'Parameters:
 'I. Path: Path to retrieve the file type from
'Returns: Specified file's type description
'---------------------------------------------------------------------
Public Function GetPathType(ByVal Path As String) As String

Path = FixPath(Path) 'Make sure we have a valid path

Dim FileInfo As SHFILEINFO, lngRet As Long
    
lngRet = SHGetFileInfo(Path, 0, FileInfo, Len(FileInfo), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME) 'Fill the FileInfo structure with the specified file's information

If lngRet = 0 Then GetPathType = Trim$(GetFileExtension(Path) & " File"): Exit Function 'If the item was not found, return a generic file type

GetPathType = left$(FileInfo.szTypeName, InStr(1, FileInfo.szTypeName, vbNullChar) - 1) 'Retrieve the specified file's type

End Function

'Function: GetAbsolutePath
'**************************
'Objective: Converts the specified path from relative to absolute in relation to the specified StartPath
'Usage: strRet = GetAbsolutePath(ByVal Path As String, ByVal StartPath As String)
'Assumptions:
 'I. If the specified path is not relative the function will return it
 'II. The specified Path must be relative while the specified StartPath must be in absolute form
'Parameters:
 'I. Path: Path to convert from relative to absolute
 'II. StartPath: Initial path to convert from
'Returns: The absolute path of the specified relative path
'---------------------------------------------------------------------
Public Function GetAbsolutePath(ByVal Path As String, ByVal StartPath As String) As String

Path = FixPath(Path): If InStr(1, Path, "..\") = 0 Then GetAbsolutePath = Path: Exit Function 'If the specified path is not relative then return it

GetAbsolutePath = StartPath 'Set the function result as a buffer for StartPath

Do Until InStrRev(Path, "..\") = 0 'Loop until all parent folders have been processed

    Path = Mid$(Path, 4): GetAbsolutePath = GetParentFolder(GetAbsolutePath) 'Retrieve parent folder of the current relative path

Loop

GetAbsolutePath = GetAbsolutePath & Path 'Return the full path of the specified relative path

End Function

'Function: GetRelativePath
'**************************
'Objective: Returns a relative path from a specified path
'Usage: strRet = GetRelativePath(ByVal Path As String, RelativeTo As String)
'Assumptions
 'I. Both paths must be in the same drive, otherwise Path will be returned
'Parameters:
 'I. Path: Path to check
 'II. RelativeTo: Path to build from
'Returns: Relative path from RelativeTo
'---------------------------------------------------------------------------------
Public Function GetRelativePath(ByVal Path As String, RelativeTo As String) As String

If GetPathDrive(Path) <> GetPathDrive(RelativeTo) Then GetRelativePath = Path: Exit Function 'Both path's must be on the same drive

Dim strParentPath As String, strRelativePath As String, intFoldersAbove As Integer

Path = FixPath(Path): RelativeTo = FixPath(RelativeTo) 'Append backslashes to the path's as needed

strParentPath = Path: strRelativePath = RelativeTo 'Assign path buffers

Do Until strParentPath = strRelativePath 'Keep looping until we have found the relative path
    
    If Len(strRelativePath) <= 3 Then 'If we reached the root drive
    
        intFoldersAbove = intFoldersAbove + 1: strRelativePath = RelativeTo 'Assume we found the relative path
    
        strParentPath = GetParentFolder(strParentPath) 'Retrieve the parent folder of the specified path
        
    End If
                  
    strRelativePath = GetParentFolder(strRelativePath) 'Retrieve the parent folder of the relative path
        
Loop
    
Dim i As Integer

For i = 1 To intFoldersAbove: GetRelativePath = GetRelativePath & "..\": Next i 'Append parent folder symbols accordingly

GetRelativePath = GetRelativePath & Mid$(RelativeTo, Len(strRelativePath) + 1) 'Return relative path

End Function

'Function: BuildPath
'********************
'Objective: Returns a valid, available path to create a file
'Usage: strRet = BuildPath(ByVal Path As String, ByVal Name As String)
'Assumptions: None
'Parameters:
 'I. Path: Path to check
 'II. Name: Name to check
'Returns: Available path, adjusted as necessary (If 'Path & Name' already exists it will return 'Path & Name (2)' and so on.
'---------------------------------------------------------------------------------
Public Function BuildPath(ByVal Path As String, ByVal Name As String) As String

Path = FixPath(Path) 'Adds a backslash to the path if necessary

Dim intFileCount As Integer: intFileCount = 1 'Set initial file's name counter

Dim strFileName As String: strFileName = Name 'Buffer the name of the path

Dim intPos As Integer: intPos = InStrRev(GetPathName(Name), ".") 'Look for the extension delimeter, so that the number gets added before it
    
If intPos = 0 Or IsFolder(Path & Name) Then intPos = Len(Path & Name) + 1 'If the file is a folder or the '.' delimiter was not found, then move the position index

CheckName:

If Exists(Path & strFileName) Then 'If the current file's path and name already exist then
                       
    intFileCount = intFileCount + 1 'Increase the file's name counter
               
    strFileName = left$(Name, intPos - 1) & " (" & intFileCount & ")" & Mid$(Name, intPos) 'Retrieve a valid name for the file
    
    GoTo CheckName 'Check if the new name is valid

End If

BuildPath = Path & strFileName 'Return the new valid file path

End Function

'Function: Execute
'******************
'Objective: Executes the specified file/folder/url with its assigned application
'Usage: intRet = Execute(ByVal Path As String,[WindowStyle As VbAppWinStyle],[ByVal Parameters As String],[WaitForExecution as Boolean],[hWnd as Long])
'Assumptions:
 'I. Path must point to a valid item
 'II. If the hWnd argument is omitted the application's active form's handle will be used (if found)
'Parameters:
 'I. Path: Item's path to execute
 'II. WindowStyle: Specifies how to load the application
 'III. Parameters: Any parameters that should be passed to the application
 'IV. WaitForExecution: Specifies wether to wait until the process is terminated (will halt execution)
 'V. hWnd: Specifies the handle of the module to pass to the function
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'----------------------------------------------------------------------------------------
Public Function Execute(ByVal Path As String, Optional WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal Parameters As String, Optional WaitForExecution As Boolean, Optional hwnd As Long) As Integer

On Error GoTo Execute_Err

If hwnd = 0 Then If Not Screen.ActiveForm Is Nothing Then hwnd = Screen.ActiveForm.hwnd 'Attempts to set the default hWnd to the Active Form's one, in case it's ommitted

If WaitForExecution Then 'If we should wait for the file to be executed

    Dim strExt As String: strExt = GetFileExtension(Path) 'Retrieve its extension to check if it's an executable
    
    If Not (strExt = "EXE" Or strExt = "COM" Or strExt = "BAT" Or strExt = "PIF") Then 'Check whether the extension matches those of executable formats
    
        strExt = GetFileExecutable(Path): If strExt = vbNullString Then Execute = 75: Exit Function 'Attempt to retrieve the executable associated with the extension
        
        Path = strExt & " " & Path 'Append the associated executable to the beginning of the path to shell
        
    End If
    
     'This method for starting the executable isn't very robust, for it completely halts the class until the process has been terminated.
     'A much better approach would be to set a callback function to raise an event when the process is closed.
     
     'The function will only work for files that aren't executables if they should be opened as 'AssociatedExe FileToOpen'
    
     Dim hProc As Long, lngExitCode As Long
     
     hProc = Shell(Path & IIfStr(Parameters = vbNullString, vbNullString, " " & Parameters), WindowStyle) 'Start the executable and grab its PID
     
     hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProc) 'Open a handle to query the process status

     Do
     
         Call GetExitCodeProcess(hProc, lngExitCode): DoEvents 'Check whether the process has been terminated
     
     Loop While lngExitCode = STILL_ACTIVE 'Loop while the process is still running
     
     Exit Function
    
End If

Execute = ShellExecute(hwnd, "open", Path, IIfStr(Parameters = vbNullString, vbNullString, Parameters), vbNullString, WindowStyle) 'Execute File with its associated application

If Execute > 32 Then Execute = 0 'If ShellExecute returns a value lower than 32, an error has occurred

Exit Function

Execute_Err: Set m_objLastError = Err: Execute = Err.Number

End Function

'Function: ParseSize
'********************
'Objective: Returns a string containing the specified size, formatted accordingly
'Usage: strRet = ParseSize(ByVal Size As Currency,[RoundBy as Integer])
'Assumptions: None
'Parameters:
 'I. Size: Size to parse
 'II. RoundBy: Specifies how many decimal numbers should be displayed
'Returns: Formatted and rounded size
'---------------------------------------------------------------------
Public Function ParseSize(ByVal Size As Currency, Optional RoundBy As Integer = 2) As String

'Store the size of each unit in bytes to save up calculation time

Const KB_LEN As Long = &H400 'Number of bytes needed to make a KiloByte
Const MB_LEN As Long = &H100000 'Number of bytes needed to make a MegaByte
Const GB_LEN As Long = &H40000000 'Number of bytes needed to make a GigaByte

If Size < KB_LEN Then 'If the file is smaller than one KB

    ParseSize = Round(Size, RoundBy) & " Bytes" 'Output it's size in bytes

ElseIf Size < MB_LEN Then 'If the file is smaller than one MB

    ParseSize = Round(Size / KB_LEN, RoundBy) & " KB" 'Output it's size in KBs

ElseIf Size < GB_LEN Then 'If the file is smaller than one GB

    ParseSize = Round(Size / MB_LEN, RoundBy) & " MB" 'Output it's size in MBs

Else 'If the file is bigger than than one GB

    ParseSize = Round(Size / GB_LEN, RoundBy) & " GB" 'Output it's size in GBs

End If

End Function

'Function: Rename
'*****************
'Objective: Renames the specified path to the determined name
'Usage: intRet = Rename(ByVal Path as String, ByVal NewPath as String,[ForceValidPath as Boolean]))
'Assumptions:
 'I. If the specified Path is the same as NewPath, no action will be performed
'Parameters:
 'I. Path: Path to rename
 'II. NewPath: Name to rename the specified path to
 'III. ForceValidPath: Specifies wether to loop until a valid path name has been generated
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function Rename(ByVal Path As String, ByVal NewPath As String, Optional ForceValidPath As Boolean) As Integer

On Error GoTo Rename_Err

If Exists(NewPath) And ForceValidPath Then

    NewPath = BuildPath(GetParentFolder(Path), GetPathName(NewPath)) 'Adds like 'file(2).txt', if 'file.txt' already exists
                    
End If

If Not Path = NewPath Then Name Path As NewPath 'No need to rename anything if the path and new path match

Exit Function

Rename_Err: Set m_objLastError = Err: Rename = Err.Number

End Function

'Function: Search
'*****************
'Objective: Searches the specified path for files matching the given criteria
'Usage: lngRet = Search(ByVal Path As String, ByVal Criteria As String,[ListItems As ListItemsGroup],[Recurse As Boolean],[MatchCase As Boolean])
'Assumptions:
 'I. The function will raise the Search event when all matching files of a folder are processed.
 'II. In order to keep track of all files of all subfolders, simply keep track of each Search event, adding the results to a common array.
 'III. The function will return the full path of the files in the results array
'Parameters:
 'I. Path: Specifies the path to search for matching files
 'II. Criteria: Specifies which criteria to use while searching. Wildcards may be used
 'III. ListItems: Specifies which type of items to retrieve, All Items, Files Only or Folders Only
 'IV. Recurse: Specifies whether to perform a recursive search and process subfolders as well
 'V. MatchCase: Specifies whether to perform a case sensitive search operation
'Returns: Returns the total amount of files found
'---------------------------------------------------------------------
Public Function Search(ByVal Path As String, ByVal Criteria As String, Optional ListItems As ListItemsGroup = LIG_ListAllItems, Optional Recurse As Boolean, Optional MatchCase As Boolean) As Long

Dim FileData As WIN32_FIND_DATA, lngHandle As Long, intRet As Integer

Path = FixPath(Path) 'Add a backslash to the current path if needed

Dim strCurrentItem As String, intAddedItems As Integer, intUbound As Integer

ReDim Results(0) As String: ReDim Folders(0) As String

lngHandle = FindFirstFile(Path & "*", FileData) 'Find the first file within the specified path

If lngHandle = -1 Then Exit Function 'If no files were found, break the operation

CheckItems:

strCurrentItem = left$(FileData.cFileName, InStr(1, FileData.cFileName, vbNullChar) - 1) 'Retrieve the first item within the current path

If strCurrentItem = "." Or strCurrentItem = ".." Then intRet = FindNextFile(lngHandle, FileData): If intRet = 0 Then Exit Function Else GoTo CheckItems 'Skip the '.' and '..' items

Do

    strCurrentItem = left$(FileData.cFileName, InStr(1, FileData.cFileName, vbNullChar) - 1) 'Retrieve the file's path
                   
    If intAddedItems > UBound(Results) Then ReDim Preserve Results(intAddedItems + ITEMS_OFFSET) 'If we're running out of space in our results array
    
    If ListItems = LIG_ListAllItems Or CBool((ListItems = Abs(CInt(IsFolder(Path & strCurrentItem))))) Then 'Check whether the item matches the ListItems specification
                        
        Results(intAddedItems) = strCurrentItem 'Add item to the results array
    
        intAddedItems = intAddedItems + 1 'Increase the added items count
        
    End If
        
    If Recurse Then 'If we are set to recurse
    
        If IsFolder(Path & strCurrentItem) Then 'Keep track of the subfolders to recurse later
    
            If intUbound > UBound(Folders) Then ReDim Preserve Folders(intUbound + FOLDERS_OFFSET) 'Allocate more space to place subfolders in, if necessary
        
            Folders(intUbound) = Path & strCurrentItem: intUbound = intUbound + 1 'Include the folder path to search in
    
        End If
        
    End If
        
    intRet = FindNextFile(lngHandle, FileData) 'Retrieve the next file in the folder
                       
Loop While intRet 'Loop while there are files to process

intRet = FindClose(lngHandle) 'Close search handle

If Not intAddedItems = 0 Then 'If at least one item has been found

    ReDim Preserve Results(intAddedItems - 1) 'Remove empty items from the results

    If Criteria <> "*.*" And Criteria <> "*" Then Results = FilterString(Results, Criteria, True, MatchCase) 'Unless we are set to return all files, check for pattern matching
        
    Search = Search + UBound(Results) + 1 'Return the amount of files found
    
    RaiseEvent Search(Path, Results) 'Raise the search event containing the results
    
End If

If Recurse Then 'If the current search session is recursive

    If intUbound = 0 Then Exit Function 'No subfolders were found

    Dim i As Integer
    
    ReDim Preserve Folders(intUbound - 1) 'Remove unused subfolder items
    
    For i = 0 To UBound(Folders) 'Loop for each subfolder found
    
        If Folders(i) <> vbNullString Then Search = Search + Search(Folders(i), Criteria, ListItems, True, MatchCase) 'Search for matches in the current subfolder
        
    Next i
    
End If
    
End Function

'Function: SetPathAttributes
'****************************
'Objective: Sets the specified path's attributes to those passed to the function
'Usage: intRet = SetPathAttributes(ByVal Path As String, Attributes As VbFileAttribute)
'Parameters:
 'I. Path: Path to set the attributes
 'II. Attributes: Attributes to set the path's to
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function SetPathAttributes(ByVal Path As String, Attributes As VbFileAttribute) As Integer

If SetFileAttributes(Path, Attributes) = 0 Then SetPathAttributes = GetLastError Else SetPathAttributes = 0 'Set path's attributes

End Function

'Function: SetPathDate
'**********************
'Objective: Sets the Created/Accessed/Modified date of the specified path
'Usage: blnRet = SetPathDate(ByVal Path as String, NewDate as Date, DateInfo as FileDateInfo)
'Assumptions: None
'Parameters:
 'I. Path: Path of the file to change the date of
 'II. NewDate: Specifies the date to set for the file
 'III. DateInfo: Specifies which date to set for the file
'Returns: True if the function succeeds, returns false otherwise
'---------------------------------------------------------------------
Public Function SetPathDate(ByVal Path As String, NewDate As Date, DateInfo As FileDateInfo) As Boolean

Dim DateSystem As SYSTEMTIME, DateFile As FILETIME

Dim DateModified As FILETIME, DateAccessed As FILETIME, DateCreated As FILETIME, LocalDate As FILETIME

'Set System Time Info

With DateSystem

    .wDay = Day(NewDate)
    .wMonth = Month(NewDate)
    .wYear = Year(NewDate)
    .wHour = Hour(NewDate)
    .wMinute = Minute(NewDate)
    .wSecond = Second(NewDate)
    .wDayOfWeek = Weekday(NewDate) - 1
    
End With

Call SystemTimeToFileTime(DateSystem, LocalDate) 'Convert it to the File Time format

Call LocalFileTimeToFileTime(LocalDate, DateFile) 'Make sure its set to the correct time zone

'Set which file date to modify

If DateInfo = FDI_DateAccessed Then DateAccessed = DateFile
If DateInfo = FDI_DateModified Then DateModified = DateFile
If DateInfo = FDI_DateCreated Then DateCreated = DateFile

Dim hFile As Long

'Get a handle to the specified file

hFile = CreateFileA(Path, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)

If hFile = -1 Then SetPathDate = False: Exit Function 'If the file does not exist or the function failed for any other reason, abort the operation

SetPathDate = SetFileTime(hFile, DateCreated, DateAccessed, DateModified) 'Set the date on the specified file

Call CloseHandle(hFile) 'Close file handle

End Function

'Function: MoveToRecycleBin
'***************************
'Objective: Moves the specified file/folder to the Recycle Bin
'Usage: intRet = MoveToRecycleBin(ByVal Path As String)
'Parameters:
 'I. Path: Path to move to the Recycle Bin
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function MoveToRecycleBin(ByVal Path As String) As Integer

On Error GoTo MoveToRecycleBin_Err

Dim SHDeleteFile As SHFILEOPSTRUCT

'Set UDT Flags

With SHDeleteFile

    .wFunc = FO_DELETE
    .pFrom = Path
    .fFlags = FOF_ALLOWUNDO Or FOF_SILENT Or FOF_NOCONFIRMATION Or FOF_RENAMEONCOLLISION

End With

MoveToRecycleBin = SHFileOperation(SHDeleteFile) 'Move file/folder to the recycle bin
    
Exit Function

MoveToRecycleBin_Err: MoveToRecycleBin = Err.Number: Set m_objLastError = Err

End Function

'****************
'FILE FUNCTIONS
'****************

'Function: GetFileExtension
'***************************
'Objective: Returns the extension portion of the specified path, if available
'Usage: strRet = GetFileExtension(ByVal Path As String)
'Assumptions: None
'Parameters:
 'I. Path: Path to extract the extension from
'Returns: Extension portion of the specified path
'---------------------------------------------------------------------------------
Public Function GetFileExtension(ByVal Path As String) As String

If Exists(Path) And IsFolder(Path) Then GetFileExtension = vbNullString: Exit Function 'Folders won't have extensions

Dim intRet As Integer: intRet = InStrRev(Path, ".") 'Retrieves the position of the last "." indicating the beginning of the file extension

If intRet = 0 Then Exit Function 'Break the operation if the extension delimiter wasn't found

GetFileExtension = UCase(Mid$(Path, intRet + 1)) 'Strips off the estension part from the file path

End Function

'Function: GetFileExecutable
'****************************
'Objective: Returns the executable associated with the specified file
'Usage: strRet = GetFileExecutable(ByVal Path as String)
'Assumptions:
 'I. If the file has no association, an empty string will be returned
'Parameters:
 'I. Path: Path of the file to check
'Returns: The path of the executable associated with the specified file
'---------------------------------------------------------------------
Public Function GetFileExecutable(ByVal Path As String) As String

If Not Exists(Path) Then Exit Function 'If the file doesn't exist then break the operation

Dim strBuffer As String, intRet As Integer

strBuffer = String$(MAX_PATH, vbNullChar) 'Fill the buffer

intRet = FindExecutable(Path, vbNullString, strBuffer) 'Retrieve the path of the executable

If intRet > 32 Then GetFileExecutable = left$(strBuffer, InStr(1, strBuffer, vbNullChar) - 1) 'Strip off null terminators

End Function

'Function: GetFileSize
'**********************
'Objective: Returns the specified file's size
'Usage: curRet = GetFileSize(ByVal Path as String)
'Assumptions
 'I. The function now supports retrieving the size of files of up to approximately 922 GB
'Parameters:
 'I. Path: Path to get the specified file's size from
'Returns: If the function suceeds, the return value is the specified file's size, otherwise 0 is returned
'---------------------------------------------------------------------------------
Public Function GetFileSize(ByVal Path As String) As Currency

Const OFFSET As Currency = 4294967296# 'Maximum size an unsigned long may hold

Dim WFD As WIN32_FIND_DATA, hFile As Long

Dim HiSize As Currency, LoSize As Currency

hFile = FindFirstFile(Path, WFD) 'Open a handle for the specified file

If hFile = -1 Then Exit Function 'If the file wasn't found, break the operation

Call FindClose(hFile) 'Close the handle now we got the information we need

HiSize = WFD.nFileSizeHigh: LoSize = WFD.nFileSizeLow 'Retrieve the high and low size of the file

'Resize variables accordingly

If HiSize < 0 Then HiSize = HiSize + OFFSET
If LoSize < 0 Then LoSize = LoSize + OFFSET

GetFileSize = (HiSize * OFFSET) + LoSize 'Convert from long integer to currency

End Function

'Function: GetFileContents
'**************************
'Objective: Retrieves the contents of the specified file
'Usage: strRet = GetFileContents(ByVal Path As String,[ByVal StartPos As Long],[ByVal Length As Long])
'Assumptions:
 'I. Will only retrieve the contents of a file of up to approx. 2GB (will overflow otherwise)
 'II. If you still want to retrieve the contents of files well over that size, loop while retrieving a couple of MB's on each iteration
 'III. Loading the whole file's contents in memory might take some time for large files
'Parameters:
 'I. Path: Path of the file to retrieve the contents of
 'II. StartPos: Position to being retrieving the contents
 'III. Length: Length of the total data to retrieve. A value of zero or less will retrieve the full file's contents
'Returns: A string containing the data within the specified file
'---------------------------------------------------------------------
Public Function GetFileContents(ByVal Path As String, Optional ByVal StartPos As Long = 1, Optional ByVal Length As Long) As String

On Error GoTo GetFileContents_Err

If Not Exists(Path) Or IsFolder(Path) Then Exit Function 'Make sure a valid file has been passed to the procedure

Dim FileNum As Integer, lngFileSize As Long

FileNum = FreeFile: lngFileSize = FileLen(Path) 'Retrieve a valid handle to open the file and it's size

If StartPos > lngFileSize Then Exit Function: If StartPos < 1 Then StartPos = 1 'Make sure the start position is valid
If Length < 1 Or Length > lngFileSize Then Length = lngFileSize 'Make sure the length is valid

GetFileContents = String$(Length, vbNullChar) 'Allocate space to place the data inside our buffer

Open Path For Binary As #FileNum 'Open the file so we can retrieve the data

    Get #FileNum, StartPos, GetFileContents 'Retrieve file's contents and store it in our buffer

Close #FileNum 'Close the file handle

Exit Function

GetFileContents_Err:

Set m_objLastError = Err: GetFileContents = vbNullString

End Function

'Function: CopyFile
'*******************
'Objective: Copies a file to the specified location
'Usage: intRet = CopyFile(ByVal Source As String, ByVal Destination As String,[MoveInstead As Boolean],[Overwrite As Boolean],[ForceValidPath as Boolean])
'Parameters:
 'I. Source: File to copy
 'II. Destination: Path to copy the file to
 'III. MoveInstead: Moves the file instead of copying it
 'IV. Overwrite: Replaces the file if it already exists in the destination
 'V. ForceValidPath: Specifies wether to create the full tree of folders so the path is valid, if required
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function CopyFile(ByVal Source As String, ByVal Destination As String, Optional MoveInstead As Boolean, Optional Overwrite As Boolean, Optional ForceValidPath As Boolean) As Integer

Const ERR_FILE_ALREADY_EXISTS = 58

On Error GoTo CopyFile_Err

If Source = Destination Then Exit Function 'No point of copying/moving anything if both the source and destination are the same

If Exists(Destination) And Not Overwrite Then 'If the file exists and we're not set to overwrite then

    CopyFile = ERR_FILE_ALREADY_EXISTS: Exit Function 'break the operation and return the error
    
ElseIf Not Exists(GetParentFolder(Destination)) And ForceValidPath Then
        
    CopyFile = CreateFolder(GetParentFolder(Destination), True): If Not CopyFile = 0 Then Exit Function 'If an error occurred then break the operation
    
End If
                   
Dim blnDeleteSource As Boolean
               
If MoveInstead Then 'If we're moving the file instead of copying it

    If GetPathDrive(Source) = GetPathDrive(Destination) Then 'If both the file are in the same drive then

        Name Source As Destination 'Move file
        
    Else 'We can't move files in different drives, so make a copy of it and then delete the original file
    
        blnDeleteSource = True 'Delete original source
        
        GoTo CopyFile 'Copy file
        
    End If
    
Else

CopyFile: Call FileCopy(Source, Destination) 'Copy file
        
End If

If blnDeleteSource Then Call DeleteFile(Source)  'If we're moving a file to another drive, then remove the original source file

Exit Function

CopyFile_Err: Set m_objLastError = Err: CopyFile = Err.Number

End Function

'Function: CreateFile
'*********************
'Objective: Creates a file and optionally places the specified content inside of it
'Usage: intRet = CreateFile(ByVal Path as String,[Content as String],[FileMode as FileModeConst],[ForceValidPath as Boolean])
'Assumptions
 'I. If the specified file's parent folder doesn't exist, it will be created
'Parameters:
 'I. Path: File's path to create
 'II. Content: Content to place inside of the file
 'III. FileMode: Specifies the mode to use when writing to files
 'IV. ForceValidPath: If the file already exists and this variable is set to true, then a new file will be created
'Returns: Zero if successful, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function CreateFile(ByVal Path As String, Optional Content As String, Optional ByVal FileMode As FileModeConst = FM_Binary, Optional ForceValidPath As Boolean) As Integer

On Error GoTo CreateFile_Err

If Exists(Path) And ForceValidPath Then Path = BuildPath(GetParentFolder(Path), GetPathName(Path))
        
Dim FileNum As Integer: FileNum = FreeFile 'Set file number to open

Select Case FileMode 'Open file on the specified mode

    Case FM_Binary: Open Path For Binary As #FileNum
    Case FM_Write: Open Path For Output As #FileNum
    Case FM_Append: Open Path For Append As #FileNum
    
    'Default mode if for some reason the user passes an invalid FileModeConst value
    
    Case Else: Open Path For Binary As #FileNum: FileMode = FM_Binary
        
End Select

If Not Content = vbNullString Then 'If the user specified any content to place on the file, do so

    If FileMode = FM_Binary Then Put #FileNum, , Content Else Print #FileNum, Content 'Put the specified contents inside the file accordingly
        
End If

Close #FileNum 'Close file handle

Exit Function

CreateFile_Err: CreateFile = Err.Number: Set m_objLastError = Err
    
End Function

'Function: DeleteFile
'*********************
'Objective: Deletes the specified file
'Usage: intRet = DeleteFile(ByVal Path as String)
'Parameters:
 'I. Path: File's path to delete
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function DeleteFile(ByVal Path As String) As Integer

On Error GoTo DeleteFile_Err

'Remove any attributes that might stop the file from being deleted

Dim intAttr As Integer: intAttr = GetPathAttributes(Path)

Call SetPathAttributes(Path, intAttr - (intAttr And REM_ATTR))
    
Kill Path 'Delete the specified file

Exit Function

DeleteFile_Err: DeleteFile = Err.Number: Set m_objLastError = Err

End Function

'Function: GetTempFile
'**********************
'Objective: Returns an available file name to open within the system's temporary folder
'Usage: strRet = GetTempFile([ByVal Prefix as String])
'Parameters:
 'I. Prefix: Optional prefix to add to the file name
'Returns: Free file name to open
'---------------------------------------------------------------------------------
Public Function GetTempFile(Optional ByVal Prefix As String) As String

On Error GoTo GetTempFile_Err

Dim strBuffer As String: strBuffer = String$(MAX_PATH, vbNullChar) 'Allocate space in our buffer
  
Call GetTempFileName(GetSpecialFolder(SFC_TempFolder), Prefix, 0, strBuffer) 'Retrieve a valid temp file name
    
GetTempFile = left$(strBuffer, InStr(1, strBuffer, vbNullChar) - 1) 'Strip off null terminators

Exit Function

GetTempFile_Err: GetTempFile = Err.Number: Set m_objLastError = Err

End Function

'Function: IsFileInUse
'**********************
'Objective: Determines if the specified file is currently in use by another process
'Usage: blnRet = IsFileInUse(ByVal Path As String)
'Assumptions:
 'I. The specified file must exist
'Parameters:
 'I. Path: Path of the file to check
'Returns: True if the file is locked by another process, false otherwise
'---------------------------------------------------------------------
Public Function IsFileInUse(ByVal Path As String) As Boolean

Dim hFile As Long

If right$(Path, 1) = "\" Then Path = left$(Path, Len(Path) - 1) 'Remove the last character if it is a backslash
    
hFile = CreateFileA(Path, GENERIC_WRITE, 0&, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_WRITE_THROUGH, 0&) 'Get a handle to try and write to a file

If hFile = -1 Then IsFileInUse = (Err.LastDllError = ERROR_SHARING_VIOLATION) 'Check if the file is already in use

Call CloseHandle(hFile) 'Close file handle

End Function

'******************
'FOLDER FUNCTIONS
'******************

'Function: IsFolder
'*******************
'Objective: Detects wether the specified path is a directory
'Usage: blnRet = IsFolder(ByVal Path as String)
'Assumptions:
 'I. The specified path must exist, or the function will fail
'Parameters:
 'I. Path: Path to check
'Returns: True if the specified path is a folder, false otherwise
'---------------------------------------------------------------------------------
Public Function IsFolder(ByVal Path As String) As Boolean

Dim Attr As Integer: Attr = GetPathAttributes(Path) 'Retrieve the attributes of the selected file

If Attr = -1 Then Exit Function 'Discard invalid attributes

IsFolder = (Attr And vbDirectory) 'Check wether the Directory attribute flag is set

End Function

'Function: CopyFolder
'*********************
'Objective: Copies an entire directory to the specified location
'Usage: intRet = CopyFolder(ByVal Source As String, ByVal Destination As String,[MoveInstead As Boolean],[Overwrite As Boolean],[ForceValidPath as Boolean])
'Parameters:
 'I. Source: Folder to copy
 'II. Destination: Path to copy the folder to
 'III. MoveInstead: Moves the folder instead of copying it
 'IV. Overwrite: Replaces the folder if it already exists in the destination
 'V. ForceValidPath: Specifies whether to create the full tree of folders so the path is valid, if required
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function CopyFolder(ByVal Source As String, ByVal Destination As String, Optional MoveInstead As Boolean, Optional Overwrite As Boolean, Optional ForceValidPath As Boolean) As Integer

On Error GoTo CopyFolder_Err

If Source = Destination Then Exit Function 'No point of copying/moving anything if both the source and destination are the same

Dim blnDeleteSource As Boolean

If MoveInstead Then 'If we're moving the folder instead of copying it

    If Not GetPathDrive(Source) = GetPathDrive(Destination) Then 'We can't move folders in different drives, so make a copy of it and then delete the original folder

        blnDeleteSource = True: GoTo CopyFolder 'Copy Folder

    Else: Name Source As Destination 'Move Folder

    End If
    
    Exit Function

End If

If Not Exists(GetParentFolder(Destination)) And ForceValidPath Then 'Check if the path exists and whether we should create it
        
    CopyFolder = CreateFolder(Destination, True): If Not CopyFolder = 0 Then Exit Function 'If an error occurred then break the operation
    
End If

CopyFolder:

If Not Exists(Destination) Then CopyFolder = CreateFolder(Destination): If CopyFolder <> 0 Then Exit Function Else CopyFolder = 0 'Creates a new folder if it doesn't exist, otherwise break the operation

Dim Contents() As String: Call GetFolderContents(Source, Contents, LIG_ListAllItems, True, True) 'Retrieve the contents of the folder

Dim i As Long

For i = 0 To UBound(Contents) 'Loop for each item within the folder

    If IsFolder(Contents(i)) Then 'If the item is a folder
    
        Call CreateFolder(Destination & Mid$(Contents(i), Len(Source) + 1)) 'Create it
    
    Else: Call CopyFile(Contents(i), Destination & Mid$(Contents(i), Len(Source) + 1), MoveInstead, Overwrite, ForceValidPath) 'Otherwise, copy the file to the new location
    
    End If
    
Next i

If blnDeleteSource Then Call DeleteFolder(Source) 'If we're moving the folder to a different drive, then delete the original source folder

Exit Function

CopyFolder_Err:

CopyFolder = Err.Number: Set m_objLastError = Err

If Not m_blnBreakOnErrors Then Resume Next

End Function

'Function: CreateFolder
'***********************
'Objective: Creates the specified folder, if it doesn't exist
'Usage: intRet = CreateFolder(ByVal Path As String,[CreateFullPath as Boolean])
'Parameters:
 'I. Path: Folder's path to create
 'II. CreateFullPath: Loops until all folders specified by Path are created, if required
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function CreateFolder(ByVal Path As String, Optional CreateFullPath As Boolean) As Integer

On Error GoTo CreateFolder_Err

If CreateFullPath And Not Exists(Path) Then 'If we're set to create a full tree of paths
    
    Dim i As Integer, Folders() As String, strBuffer As String
    
    If right$(Path, 1) = "\" Then Path = left$(Path, Len(Path) - 1) 'Remove last character from the path, if it's a backslash
    
    Folders = Split(GetParentFolder(Path), "\") 'Retrieve each folder to create
    
    On Error Resume Next
    
    For i = 0 To UBound(Folders) 'Loop for each folder
    
        If Not Exists(strBuffer & Folders(i)) Then Call MkDir(strBuffer & Folders(i)) 'If it does not exist then create it
        
        strBuffer = strBuffer & Folders(i) & "\" 'Buffer the current path
        
    Next i
    
    On Error GoTo CreateFolder_Err
    
End If

Call MkDir(Path) 'Create Folder

Exit Function

CreateFolder_Err: CreateFolder = Err.Number: Set m_objLastError = Err

End Function

'Function: DeleteFolder
'***********************
'Objective: Deletes the specified folder and its entire contents
'Usage: intRet = DeleteFolder(ByVal Path as String)
'Parameters:
 'I. Path: Folder's path to delete
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function DeleteFolder(ByVal Path As String) As Integer

On Error GoTo DeleteFolder_Err

'Get Folder Contents

Dim FolderItems() As String, intRet As Integer

If GetFolderContents(Path, FolderItems, LIG_ListAllItems, True, False) = 0 Then GoTo DeleteFolder 'If no subitems were found, skip deleting them

Dim i As Long

For i = 0 To UBound(FolderItems) 'Loop for each file inside the specified folder

    If IsFolder(FolderItems(i)) Then intRet = DeleteFolder(FolderItems(i)) Else intRet = DeleteFile(FolderItems(i)) 'Delete the file accordingly
    
    If Not intRet = 0 Then Call Err.Raise(intRet, "clsFileAnalyzer.DeleteFolder", Error$(intRet)) 'Check whether errors occurred

Next i

DeleteFolder:

'Remove any attributes that might stop the folder from being deleted

Dim intAttr As Integer: intAttr = GetPathAttributes(Path)

Call SetPathAttributes(Path, intAttr - (intAttr And REM_ATTR))

Call RmDir(Path)  'Delete the specified folder (now empty)

Exit Function

DeleteFolder_Err:

DeleteFolder = Err.Number: Set m_objLastError = Err

If Not m_blnBreakOnErrors Then Resume Next

End Function

'Function: PathHasSubfolders
'****************************
'Objective: Checks for the existance of subfolders within the specified path
'Usage: blnRet = PathHasSubfolders(ByVal Directory As String)
'Assumptions:
 'I. Directory must point to a valid folder
'Parameters:
 'I. Directory: Path of the Directory to check for subfolders
'Returns: True if any subfolders were found within the specified path, false otherwise
'----------------------------------------------------------------------------------------
Public Function PathHasSubfolders(ByVal Directory As String) As Boolean

Dim strRet As String: Directory = FixPath(Directory) 'Add a backslash to the path if needed

If Not IsFolder(Directory) Then Exit Function 'If the specified path's not a directory then exit the function

strRet = Dir$(Directory, DEF_ATTR) 'Get folder's items

IgnoreFiles:

If (strRet = "." Or strRet = "..") Or Not IsFolder(Directory & strRet) Then strRet = Dir(): GoTo IgnoreFiles 'Ignore symbolic folders and files

PathHasSubfolders = (strRet <> vbNullString) 'If a subfolder was found then return true

End Function

'Function: IsEmpty
'******************
'Objective: Detects wether the specified directory is empty
'Usage: blnRet = IsEmpty(ByVal Directory as String)
'Assumptions:
 'I. Path must point to a valid Directory
'Parameters:
 'I. Directory: Folder to check
'Returns: True if the specified directory is empty, false otherwise
'---------------------------------------------------------------------------------
Public Function IsEmpty(ByVal Directory As String) As Boolean

Dim strRet As String: Directory = FixPath(Directory) 'Add a backslash to the path if needed

If Not IsFolder(Directory) Then Exit Function 'If the specified path's not a directory then exit the function

strRet = Dir$(Directory, DEF_ATTR) 'Get folder's items

IgnoreSymbolicFolders: If left$(strRet, 1) = "." Then strRet = Dir(): GoTo IgnoreSymbolicFolders 'Ignore symbolic folders

IsEmpty = (strRet = vbNullString) 'If files were found then return true

End Function

'Function: GetFolderContents
'****************************
'Objective: Performs operations concerning the RecurseFolders procedure
'Usage: lngRet = GetFolderContents(ByVal Path As String, ByRef Contents() As String,[ListItems As ListItemsGroup],[FullPath as Boolean],[Recurse As Boolean])
'Assumptions:
 'I. The Specified Path must be valid
'Parameters:
 'I. Path: Path to retrieve the contents from
 'II. Contents: Array to store the contents in
 'III. ListItems: List Files/Folders/All Items
 'IV. FullPath: Specifies whether to retrieve the full path of the contents, of just their names
 'V. Recurse: Specifies wether to loop for each subfolder in the specified path, until all file/folders have been retrieved
'Returns: Total items in the specified folder
'---------------------------------------------------------------------------------
Public Function GetFolderContents(ByVal Path As String, ByRef Contents() As String, Optional ListItems As ListItemsGroup = LIG_ListAllItems, Optional FullPath As Boolean, Optional Recurse As Boolean = False) As Long

Dim lngCount As Long: ReDim Contents(0) 'Initialize Contents Array

Call RecurseFolders(Path, Contents, lngCount, ListItems, FullPath, Recurse) 'Get the specified folder's contents

If lngCount = 0 Then Exit Function 'If no file was found then break the operation

ReDim Preserve Contents(lngCount - 1) 'Dimensionate contents array

If ListItems <> LIG_ListAllItems Then Contents = FilterString(Contents, vbNullChar, False) 'Remove items with chr$(0)

If UBound(Contents) = 0 And Contents(0) = vbNullString Then Exit Function
   
GetFolderContents = UBound(Contents) + 1 'Return the number of files in the folder

End Function

'Function: GetFolderSize
'************************
'Objective: Retrieves the specified folder's size
'Usage: currRet = GetFolderSize(ByVal Path as String)
'Parameters:
 'I. Path: Path to retrieve the size from
'Returns: Specified folder's size
'---------------------------------------------------------------------------------
Public Function GetFolderSize(ByVal Path As String) As Currency

Dim FolderItems() As String

Path = FixPath(Path) 'Add a backslash to the path if needed

Call GetFolderContents(Path, FolderItems, LIG_ListFiles, True, True) 'Get Folder Contents

Dim i As Long

For i = 0 To UBound(FolderItems): GetFolderSize = GetFolderSize + GetFileSize(FolderItems(i)): Next i 'Loop and retrieve the size of each file within the specified folder

End Function

'Function: GetSpecialFolder
'***************************
'Objective: Retrieves the paths of various system folders
'Usage: strRet = GetSpecialFolder(FolderID as SpecialFolderConst)
'Parameters:
 'I. FolderID: Enum specifying the system folder to retrieve
'Returns: Specified system folder's path
'---------------------------------------------------------------------------------
Public Function GetSpecialFolder(FolderID As SpecialFolderConst) As String

Select Case FolderID

    Case SFC_TempFolder 'Get the System Temporary Folder path
    
        GetSpecialFolder = String$(MAX_PATH, vbNullChar) 'Allocate space in our buffer
               
        Dim intRet As Integer: intRet = GetTempPath(MAX_PATH, GetSpecialFolder) 'Retrieve the path of the temp folder
        
        GetSpecialFolder = left$(GetSpecialFolder, intRet)
        
    Case SFC_RecycleBin 'Get the System Recicle Bin path
            
        GetSpecialFolder = GetSystemFolder(FolderID) 'Try to retrieve the path of the system recycle bin
        
        If GetSpecialFolder = vbNullString Then If Exists(GetSystemDrive & "RECYCLER") Then GetSpecialFolder = GetSystemDrive & "RECYCLER" 'Return the default recycle bin path in Windows XP if it wasn't found otherwise
            
    Case Else: GetSpecialFolder = GetSystemFolder(FolderID) 'Retrieve the path of the specified Special Folder accordingly

End Select

End Function

'*******************
'DRIVE FUNCTIONS
'*******************

'Function: DriveExists
'**********************
'Objective: Checks wether the specified drive exists
'Usage: blnRet = DriveExists(ByVal Drive as String)
'Assumptions: None
'Parameters:
 'I. Drive: Drive to check
'Returns: True if the drive is found, false otherwise
'---------------------------------------------------------------------
Public Function DriveExists(ByVal Path As String)

DriveExists = (GetDriveType(Path) <> DTV_NotFound) 'Check whether the specified drive exists

End Function

'Function: GetSystemDrive
'*************************
'Objective: Returns the drive in which the Operating System is running
'Usage: strRet = GetSystemDrive()
'Parameters: None
'Returns: Operating System Drive
'---------------------------------------------------------------------------------
Public Function GetSystemDrive() As String

GetSystemDrive = FixPath(Environ$("SYSTEMDRIVE")) 'Retrieves the path of the System Drive

End Function

'Function: GetPathDrive
'***********************
'Objective:
'Usage: strRet = GetPathDrive(ByVal Path as String)
'Assumptions: None
'Parameters:
 'I. Path: Path to check
'Returns: Drive portion of the specified path
'---------------------------------------------------------------------------------
Public Function GetPathDrive(ByVal Path As String) As String

GetPathDrive = FixPath(left$(Path, 3)) 'Retrieve the first three characters of the specifed path, which should be its drive path
    
End Function

'Function: GetDrives
'********************
'Objective: Retrieves all available drives
'Usage: intRet = GetDrives()
'Parameters:
 'I. Drives: Array in which the results will be placed
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function GetDrives(Drives() As String) As Integer

Dim strBuffer As String

strBuffer = String(255, vbNullChar) 'Create space in the buffer

Call GetLogicalDriveStrings(255, strBuffer) 'Get the available drives in the buffer
    
strBuffer = left$(strBuffer, InStrRev(strBuffer, "\") + 1) 'Strip null terminators

Drives() = Split(strBuffer, vbNullChar) 'Add the available drives to an array

ReDim Preserve Drives(UBound(Drives) - 1) 'Strip off last empty item

End Function

'Function: GetDriveType
'***********************
'Objective: Retrieves an integer specifying the passed Drive's type
'Usage: enmRet = GetDriveType(ByVal Drive as String)
'Parameters:
 'I. Drive: Drive to check
'Returns: Specified's drive type
'---------------------------------------------------------------------------------
Public Function GetDriveType(ByVal Drive As String) As Integer

GetDriveType = GetDriveTypeA(Drive) 'Retrieves the specified drive's type

End Function

'Function: GetDriveTypeName
'***************************
'Objective: Returns a string with the corresponding Drive type association
'Usage: strRet = GetDriveTypeName(ByVal Drive as String)
'Parameters:
 'I. Drive: Drive to check
'Returns: String containing the drive's type
'---------------------------------------------------------------------------------
Public Function GetDriveTypeName(ByVal Drive As String) As String

Select Case GetDriveType(Drive) 'Check the type of the specified drive

    Case DTV_Unknown: GetDriveTypeName = "Unknown" 'Drive Type is Unknown
    Case DTV_NotFound: GetDriveTypeName = "Not Found" 'Drive was not found
    Case DTV_Removable: If left$(Drive, 1) = "A" Or left$(Drive, 1) = "B" Then GetDriveTypeName = "Floppy" Else GetDriveTypeName = "Removable" 'Drive is either Removable or a Floppy Drive
    Case DTV_Fixed: GetDriveTypeName = "Fixed" 'Drive is Fixed
    Case DTV_Remote: GetDriveTypeName = "Remote" 'Drive is Remote
    Case DTV_CDROM: GetDriveTypeName = "CD-ROM" 'Drive is a CD-ROM
    Case DTV_RAMDisk: GetDriveTypeName = "RAM Disk" 'Drive is a RAM Disk

End Select

End Function

'Function: GetDriveInfo
'***********************
'Objective: Retrieves the specified drive's information
'Usage: intRet = GetDriveInfo(ByVal Drive As String, [ByRef Volume As String],[ByRef FileSystem As String],[ByRef SerialNumber As String])
'Parameters:
 'I. Drive: Drive to check
 'II. Volume: Optional variable passed ByRef to store the drive's volume
 'III. FileSystem: Optional variable passed ByRef to store file system of the drive
 'IV. SerialNumber: Optional variable passed ByRef to store the Drive's serial number
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function GetDriveInfo(ByVal Drive As String, Optional ByRef Volume As String, Optional ByRef FileSystem As String, Optional ByRef SerialNumber As String) As Integer

On Error GoTo GetDriveInfo_Err

Dim lngSerial As Long
    
Volume = String$(255, vbNullChar): FileSystem = String$(255, vbNullChar) 'Fill the buffers with null characters

GetDriveInfo = GetVolumeInformation(Drive, Volume, Len(Volume), lngSerial, 0, 0, FileSystem, Len(FileSystem)) 'Retrieve volume information

If GetDriveInfo = 0 Then GetDriveInfo = GetLastError: Exit Function Else GetDriveInfo = 0 'Check if any errors occurred
        
Volume = left$(Volume, InStr(Volume, vbNullChar) - 1) 'Retrieve the volume name

FileSystem = left$(FileSystem, InStr(FileSystem, vbNullChar) - 1) 'Retrieve the drive's file system

SerialNumber = Hex$(lngSerial) 'Retrieve the drive's serial number

Exit Function

GetDriveInfo_Err: GetDriveInfo = Err.Number: Set m_objLastError = Err

End Function

'Function: GetDriveSpace
'************************
'Objective: Retrieves information about the specified drive's size
'Usage: intRet = GetDriveSpace(ByVal Drive As String, [ByRef TotalSpace As Currency],[ByRef FreeSpace As Currency],[ByRef UsedSpace As Currency],[ByRef UsedSpacePercent As Single])
'Parameters:
 'I. Drive: Drive to check
 'II. TotalSpace: Buffer in which the total space information will be placed
 'III. FreeSpace: Buffer in which the free space information will be placed
 'IV. UsedSpace: Buffer in which the used space information will be placed
 'V. UsedSpacePercent: Percent representing the total used space on the drive
'Returns: Zero if sucessfull, otherwise returns the error code raised by the procedure
'---------------------------------------------------------------------------------
Public Function GetDriveSpace(ByVal Drive As String, Optional ByRef TotalSpace As Currency, Optional ByRef FreeSpace As Currency, Optional ByRef UsedSpace As Currency, Optional ByRef UsedSpacePercent As Single) As Integer

On Error GoTo GetDriveSpace_Err

GetDriveSpace = GetDiskFreeSpaceEx(Drive, 0&, TotalSpace, FreeSpace) 'Retrieve drive size information

If GetDriveSpace = 0 Then GetDriveSpace = GetLastError: Exit Function Else GetDriveSpace = 0 'Check if any errors occurred

'Convert the results propertly to a currency value

TotalSpace = TotalSpace * 10000
FreeSpace = FreeSpace * 10000

'Calculate used space and percent

UsedSpace = TotalSpace - FreeSpace
UsedSpacePercent = Round((UsedSpace / TotalSpace) * 100, 2)

Exit Function

GetDriveSpace_Err: GetDriveSpace = Err.Number: Set m_objLastError = Err

End Function

'Function: GetDriveNumber
'*************************
'Objective: Retrieves the system order of the specified drive
'Usage: intRet = GetDriveNumber(ByVal Drive as String)
'Assumptions:
 'I. The specified drive must exist in order for the function to succeed
'Parameters:
 'I. Drive: Drive to check
'Returns: If the function succeeds the drive order index is returned, otherwise zero is returned
'---------------------------------------------------------------------
Public Function GetDriveNumber(ByVal Drive As String) As Integer

Dim i As Integer, Drives() As String

Call GetDrives(Drives) 'Retrieve available drives

For i = 0 To UBound(Drives) 'Loop for each drive

    If Drives(i) = Drive Then GetDriveNumber = i + 1: Exit Function 'Check to find the assigned drive number

Next i

End Function

'Function: SetDriveLabel
'************************
'Objective: Sets the label of the specified drive
'Usage: blnRet = SetDriveLabel(ByVal Drive As String, NewLabel As String)
'Assumptions:
 'I. Only fixed and possibly remote/removable drives can have their label changed
'Parameters:
 'I. Drive: Drive to change the label of
 'II. NewLabel: Label to set for the drive
'Returns: True if the function succeeded, returns false otherwise
'---------------------------------------------------------------------
Public Function SetDriveLabel(ByVal Drive As String, NewLabel As String) As Boolean

SetDriveLabel = SetVolumeLabel(FixPath(Drive) & vbNullChar, NewLabel & vbNullChar)

End Function




'ReadFile

'return readed information as string

Public Function ReadBinaryFile(ByRef File_Read As String) As String
Dim sData As String
Open File_Read For Binary Access Read As #1
        sData = String(LOF(1), Chr(0))
        Get #1, , sData
Close #1
ReadBinaryFile = sData
End Function

Public Function WriteBinaryFile(ByRef File_Write As String, ByRef sData As String) As Boolean
On Error GoTo LblError
Open File_Write For Binary Access Write As #1
    Put #1, , sData
Close #1
WriteBinaryFile = True
Exit Function

LblError:
WriteBinaryFile = False
End Function




















'*******************
'PRIVATE FUNCTIONS
'*******************

'Function: RecurseFolders
'*************************
'Objective: Retrieves all files/folders inside the specified path
'Usage: lngRet = RecurseFolders(ByVal Path As String, ByRef Contents() As String, ByRef Count As Long,[ListItems As ListItemsGroup],[FullPath as Boolean],[Recurse As Boolean])
'Assumptions:
 'I. The Specified Path must be valid
'Parameters:
 'I. Path: Path to retrieve the contents from
 'II. Contents: Array to store the contents in
 'III. Count: Variable for internal files management
 'IV. ListItems: List Files/Folders/All Items
 'V. FullPath: Specifies whether to retrieve the full path of the contents, of just their names
 'VI. Recurse: Specifies wether to loop for each subfolder in the specified path, until all file/folders have been retrieved
'Returns: If the function succeeds, the return value is the total items in the specified folder, otherwise -1 is returned
'---------------------------------------------------------------------------------
Private Function RecurseFolders(ByVal Path As String, ByRef Contents() As String, ByRef count As Long, Optional ListItems As ListItemsGroup = LIG_ListAllItems, Optional FullPath As Boolean, Optional Recurse As Boolean) As Long

On Error GoTo RecurseFolders_Err

Dim FileData As WIN32_FIND_DATA, lngHandle As Long, intRet As Integer

Path = FixPath(Path) 'Add a backslash to the current path if needed

Dim strCurrentItem As String, intAddedItems As Integer

lngHandle = FindFirstFile(Path & "*", FileData) 'Find the first file within the specified path

If lngHandle = -1 Then Exit Function 'If no files were found, break the operation

CheckItems:

strCurrentItem = left$(FileData.cFileName, InStr(1, FileData.cFileName, vbNullChar) - 1) 'Retrieve the first item within the current path

If strCurrentItem = "." Or strCurrentItem = ".." Then intRet = FindNextFile(lngHandle, FileData): If intRet = 0 Then intRet = FindClose(lngHandle): Exit Function Else GoTo CheckItems 'Skip the '.' and '..' items

Do

    strCurrentItem = left$(FileData.cFileName, InStr(1, FileData.cFileName, vbNullChar) - 1) 'Retrieve item's path
                   
    If count > UBound(Contents) Then ReDim Preserve Contents(count + ITEMS_OFFSET) 'Allocate more space for the Contents array if needed
                    
    Contents(count) = strCurrentItem 'Add current item to the Contents array
    
    intAddedItems = intAddedItems + 1 'Increase the currently added items count
    count = count + 1 'Increase the total items added count
    
    intRet = FindNextFile(lngHandle, FileData) 'Retrieve the next item within the current path
                       
Loop While intRet

intRet = FindClose(lngHandle) 'Close search handle
  
Dim i As Long
    
For i = count - intAddedItems To count - 1 'Loop for each item added to the array in this session
       
    If IsFolder(Path & Contents(i)) Then 'If the current item's a folder
    
        If Recurse Then Call RecurseFolders(Path & Contents(i), Contents, count, ListItems, FullPath, True) 'And recurse through its contents if set to
    
        If ListItems = LIG_ListFiles Then Contents(i) = vbNullChar 'If we're set to list files only then remove the reference to the current folder
        
    Else
            
        If ListItems = LIG_ListFolders Then Contents(i) = vbNullChar 'If we're set to list folders only then remove the reference to the current file
        
    End If
    
    If Not Contents(i) = vbNullChar Then If FullPath Then Contents(i) = Path & Contents(i) 'Append the item's path to the start of the current item if set to
    
Next i

Exit Function

RecurseFolders_Err: RecurseFolders = -1: Resume Next

End Function

'Function: FilterString
'***********************
'Objective: Returns an array containing only the items matching a specified criteria of the original array
'Usage: FilterString(ByRef SourceArray() As String, ByVal Criteria As String,[Include As Boolean],[MatchCase As Boolean],[Compare As VbCompareMethod])
'Assumptions:
 'I. The specified Array must contain at least one item
'Parameters:
 'I. SourceArray: Array containing the items to filter
 'II. Criteria: Criteria to filter the items with
 'III. Include: Specifies wether to remove or add the items matching the criteria
 'IV. MatchCase: Specifies wether to return only items matching a specific text case
 'V. Compare: Specifies the item compare method to use
'Returns: An array containing only the items matching the specified criteria
'---------------------------------------------------------------------
Private Function FilterString(ByRef SourceArray() As String, ByVal Criteria As String, Optional Include As Boolean = True, Optional MatchCase As Boolean = False) As String()

Dim lngCount As Long, strBuffer() As String

ReDim strBuffer(UBound(SourceArray) - LBound(SourceArray)) 'Initialize Buffer Array

If Err Then FilterString = SourceArray: Exit Function 'Array not Initialized

If Not MatchCase Then Criteria = UCase(Criteria) 'If we're not matching cases then set the criteria to high case

Dim i As Long

For i = LBound(SourceArray) To UBound(SourceArray) 'Loop for each array item

    'If the current item contains the specified criteria within it, then add/remove it according to the Include parameter
    
    If (IIfStr(MatchCase, SourceArray(i), UCase(SourceArray(i))) Like Criteria) = Include Then
           
        strBuffer(lngCount) = SourceArray(i) 'Add Item to the Buffer Array
        
        lngCount = lngCount + 1 'Increase Items Count
        
    End If
    
Next i

If lngCount = 0 Then ReDim FilterString(0): Exit Function 'If no items were found then return an empty array

ReDim Preserve strBuffer(lngCount - 1) 'Dimensionate final buffer array
    
FilterString = strBuffer 'Return array containing the matching items

End Function

'Function: GetSystemFolder
'**************************
'Objective: Retrieves and strips the system folder path specified by CSIDL
'Usage: strRet = GetSystemFolder()
'Parameters:
 'I. CSIDL: Constant representing a system folder
'Returns: Corresponding system folder's path
'---------------------------------------------------------------------------------
Private Function GetSystemFolder(CSIDL As Long) As String

Dim lngRet As Long
    
Dim IDL As ITEMIDLIST
    
lngRet = SHGetSpecialFolderLocation(100, CSIDL, IDL) 'Retrieve the path of the specified folder ID
    
If Not lngRet = 0 Then Exit Function
      
Dim strBuffer As String: strBuffer = String$(MAX_PATH, vbNullChar) 'Dimensionate the buffer

lngRet = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strBuffer) 'Resolve the ID's Path to an actual system path

GetSystemFolder = left$(strBuffer, InStr(strBuffer, vbNullChar) - 1) 'Strip the return value
   
End Function

'Function: AdjustTimeForLocalSettings
'*************************************
'Objective: Adjusts file's time according to the system configuration
'Usage: lngRet = AdjustTimeForLocalSettings()
'Parameters: None
'Returns: Time difference to adjust
'---------------------------------------------------------------------------------
Private Function AdjustTimeForLocalSettings() As Long

Dim dtmSystemDate As Date, udtSystemTime As SYSTEMTIME

Call GetSystemTime(udtSystemTime) 'Retrieve the current system time

'Parse the time in order to retrieve each parameters separetely
  
dtmSystemDate = DateSerial(udtSystemTime.wYear, udtSystemTime.wMonth, udtSystemTime.wDay) + TimeSerial(udtSystemTime.wHour, udtSystemTime.wMinute, udtSystemTime.wSecond)
    
AdjustTimeForLocalSettings = DateDiff("h", dtmSystemDate, Now) 'Compare the system time with the current time and return changes

End Function

'Function: IIfStr
'*****************
'Objective: Replace the default VB's IIF function for strings, to improve execution speed
'Usage: strRet = IIfStr(ByVal Expression As Boolean, ByVal TruePart As String, ByVal FalsePart As String)
'Assumptions: None
'Parameters:
 'I. Expression: Expression to check
 'II. TruePart: Return value if the expression returns true
 'III. FalsePart: Return value if the expression returns false
'Returns: The return value will be either TruePart of FalsePart depending on Expression
'---------------------------------------------------------------------
Private Function IIfStr(ByVal Expression As Boolean, ByVal TruePart As String, ByVal FalsePart As String) As String

If Expression Then IIfStr = TruePart Else IIfStr = FalsePart 'Check expression

End Function







