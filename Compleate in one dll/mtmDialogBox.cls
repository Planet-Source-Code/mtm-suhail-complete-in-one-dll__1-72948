VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmDialogBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type


Public Function BrowseFolder(DialogText As String, DefaultPath As String, OwnerhWnd As Long, Optional ShowCurrentPath As Boolean = True, Optional RootPath As Variant, Optional NewDialogStyle As Boolean = False, Optional IncludeFiles As Boolean = False) As String
   BrowseFolder = BrowseForFolder(DialogText, DefaultPath, OwnerhWnd, ShowCurrentPath, RootPath, NewDialogStyle, IncludeFiles)
End Function

Public Function BrowseFolderNew(hwndOwner As Long, sPrompt As String) As String
   BrowseFolderNew = BrowseForFolderNew(hwndOwner, sPrompt)
End Function

Public Function ShowOpenWindow(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
        ShowOpenWindow = ShowOpen(hwnd, centerForm)
End Function

Public Function ShowColorWindow(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedColor
        ShowColorWindow = ShowColor(hwnd, centerForm)
End Function

Public Function ShowPrinterWindow(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As Long
            ShowPrinterWindow = ShowPrinter(hwnd, centerForm)
End Function
Public Function ShowSaveWindow(ByVal hwnd As Long, Optional ShowStrFileName As String, Optional ByVal centerForm As Boolean = True) As SelectedFile
           ShowSaveWindow = ShowSave(hwnd, ShowStrFileName, centerForm)
End Function
