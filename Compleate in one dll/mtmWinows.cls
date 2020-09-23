VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmWinows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Private Const SHERB_NOCONFIRMATION = &H1
Private Const SHERB_NOPROGRESSUI = &H2
Private Const SHERB_NOSOUND = &H4
Private Const WITHOUT_ANY = SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long



Private Function GetWindowsPath() As String
    Dim lpBuffer As String * 255
    Dim nSize As Long
    nSize = GetWindowsDirectory(lpBuffer, 255)
    GetWindowsPath = left(lpBuffer, nSize) & "\"
End Function


Public Sub ClearJunkFile()
    On Error Resume Next
    Dim obj As New mtmFile
    Kill GetWindowsPath & "Prefetch\*.*"
    Kill GetWindowsPath & "Temp\*.*"
    Kill obj.GetSpecialFolder(CSIDL_RECENT) & "\*.*"
    Kill obj.GetSpecialFolder(CSIDL_HISTORY) & "\*.*"
    Kill obj.GetSpecialFolder(SFC_UserDirectory) & "\Cookies\*.*"
    Kill obj.GetSpecialFolder(SFC_UserDirectory) & "\Local Settings\Temp\*.*"
    Kill obj.GetSpecialFolder(SFC_UserDirectory) & "\Local Settings\Temporary Internet Files\*.*"
        
End Sub


Public Function RemoveWallpaper() As Long
    ' Comments  :
    ' Parameters:  -
    ' Returns   : Long -
    ' Author    : Chris Gibbs
    ' Modified  :
    ' --------------------------------------------------
    ' this function will remove the wallpaper
    RemoveWallpaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, "(None)", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function

Public Function SetWallpaper(ByVal strBitmap As String) As Long
    ' Comments  :
    ' Parameters: strBitmap -
    ' Returns   : Long -
    ' Author    : Chris Gibbs
    ' Modified  :
    ' --------------------------------------------------
    ' this function will set the wallpaper
    SetWallpaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, strBitmap, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function
Public Function EmptyTheRecycle(hwnd As Long) As Long
    On Error Resume Next
    EmptyTheRecycle = SHEmptyRecycleBin(hwnd, vbNullString, WITHOUT_ANY)
End Function


