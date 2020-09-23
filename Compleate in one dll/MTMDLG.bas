Attribute VB_Name = "MTMDLG"
Option Explicit

Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5

Dim hHook As Long

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256

Public Const LF_FACESIZE = 32

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY



Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260



Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type



Public Type BROWSEINFONEW
     hwndOwner As Long
     pidlRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hdc As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    Flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String          '  return the style field here
    nFontType As Integer          '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

Type PRINTDLGS
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        Flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000

Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Public Const DN_DEFAULTPRN = &H1

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type



Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Public FileDialog As OPENFILENAME
Public ColorDialog As CHOOSECOLORS
Public FontDialog As CHOOSEFONTS
Public PrintDialog As PRINTDLGS
Dim ParenthWnd As Long



Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


Private Const DialogCurrentSelectionText As String = "Auswahl: "






Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function ILCreateFromPath Lib "shell32" Alias "#157" _
    (ByVal Path As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, _
    lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long




Private Const WM_USER = &H400

Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Private Const BIF_NEWDIALOGSTYLE As Long = &H40

Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_STATUSTEXT As Long = &H4

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Public Function BrowseForFolder(DialogText As String, DefaultPath As String, _
    OwnerhWnd As Long, Optional ShowCurrentPath As Boolean = True, _
    Optional RootPath As Variant, Optional NewDialogStyle As Boolean = False, _
    Optional IncludeFiles As Boolean = False) As String
    

    Dim InfoBrowse As BROWSEINFO
    Dim lPIDL As Long
    Dim sBuffer As String
    Dim lpSelPath As Long
    Dim lpPathBuffer As Long

    With InfoBrowse
        ' Handle des übergeordneten Fensters
        .hOwner = OwnerhWnd
        
        ' PIDL des Rootordners
        If Not IsMissing(RootPath) Then .pidlRoot = PathToPIDL(RootPath)
        
        ' Dialogtext
        .lpszTitle = DialogText
        
        ' Stringbuffer für aktuell selektierten Pfad reservieren und
        ' Adresse zuweisen
        If ShowCurrentPath Then
            lpPathBuffer = LocalAlloc(LPTR, MAX_PATH)
            .pszDisplayName = lpPathBuffer
        End If
        
        ' Dialogeinstellungen
        .ulFlags = BIF_RETURNONLYFSDIRS + _
            IIf(ShowCurrentPath, BIF_STATUSTEXT, 0) + _
            IIf(NewDialogStyle, BIF_NEWDIALOGSTYLE, 0) + _
            IIf(IncludeFiles, BIF_BROWSEINCLUDEFILES, 0)
        
        ' Callbackfunktion-Adresse zuweisen
        .lpfnCallback = FARPROC(AddressOf CallbackString)
        
        ' Stringspeicher für vorselektierten Ordner reservieren
        lpSelPath = LocalAlloc(LPTR, Len(DefaultPath) + 1)
        
        ' Vorselektierten Ordnerpfad in den reservierten Speicherbereich
        ' kopieren
        CopyMemory ByVal lpSelPath, ByVal DefaultPath, Len(DefaultPath) + 1
        
        ' Adresse des vorselektierten Ordnerpfades zuweisen (wird im
        ' lpData-Parameter an die Callback-Funktion weitergeleitet)
        .lParam = lpSelPath
    End With

    ' BrowseForFolder-Dialog anzeigen
    lPIDL = SHBrowseForFolder(InfoBrowse)

    If lPIDL Then
        ' Stringspeicher reservieren
        sBuffer = Space$(MAX_PATH)
    
        ' Selektierten Pfad aus der zurückgegebenen PIDL ermitteln
        SHGetPathFromIDList lPIDL, sBuffer
        
        ' Nullterminierungszeichen des Strings entfernen
        sBuffer = left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        
        ' Selektierten Pfad zurückgeben
        BrowseForFolder = sBuffer
        
        ' Reservierten Task-Speicher wieder freigeben
        Call CoTaskMemFree(lPIDL)
    End If

    ' Stringspeicher wieder freigeben
    Call LocalFree(lpSelPath)
    Call LocalFree(lpPathBuffer)
End Function

Private Function CallbackString(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal lParam As Long, ByVal lpData As Long) As Long
    
    ' Callback-Funktion des BrowseForFolder-Dialogs. Wird bei eintretenden Ereignissen des Dialogs aufgerufen.
    
    Dim sBuffer As String
    
    ' Meldungen herausfiltern
    Select Case uMsg
    Case BFFM_INITIALIZED
        ' Dialog wurde initialisiert
        
        ' Zu selektierenden Pfad (dessen Pointer wurde in lpData übergeben) an den Dialog senden
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    Case BFFM_SELCHANGED
        ' Selektierung hat sich geändert
        
        ' Stringspeicher reservieren
        sBuffer = Space$(MAX_PATH)
        
        ' Aktuell selektierten Pfad ermitteln und diesen an den Dialog senden
        If SHGetPathFromIDList(lParam, sBuffer) Then Call SendMessage(hwnd, BFFM_SETSTATUSTEXTA, 0&, ByVal DialogCurrentSelectionText & sBuffer)
    End Select
End Function

Private Function FARPROC(pfn As Long) As Long
    ' Funktion wird benötigt, um Funktions-Adresse ermitteln zu können, dessen Adresse mit AddressOf übergeben und anschließend wieder zurückgegeben wird.
    
    FARPROC = pfn
End Function

Private Function PathToPIDL(ByVal Path As String) As Long
    ' Konvertiert einen Pfad in dessen PIDL.
    
    Dim lRet As Long
    
    lRet = ILCreateFromPath(Path)
    If lRet = 0 Then
        Path = StrConv(Path, VbStrConv.vbUnicode)
        lRet = ILCreateFromPath(Path)
    End If
    
    PathToPIDL = lRet
End Function


'This module contains all the declarations to use the
'Windows 95 Shell API to use the browse for folders
'dialog box.  To use the browse for folders dialog box,
'please call the BrowseForFolders function using the
'syntax: stringFolderPath=BrowseForFolders(Hwnd,TitleOfDialog)
'
'For Aliasing information, see other module


Public Function BrowseForFolderNew(hwndOwner As Long, sPrompt As String) As String
     On Error Resume Next
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BROWSEINFO

    'initialise variables
     With udtBI
        .hOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = left$(sPath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolderNew = sPath

End Function


Public Function ShowOpen(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
'On Error Resume Next
Dim ret As Long
Dim count As Integer
Dim fileNameHolder As String
Dim LastCharacter As Integer
Dim NewCharacter As Integer
Dim tempFiles(1 To 200) As String
Dim hInst As Long
Dim Thread As Long
    
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = "" & Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    'If FileDialog.flags = 0 Then
        FileDialog.Flags = OFS_FILE_OPEN_FLAGS
    'End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetOpenFileName(FileDialog)

    If ret Then
        If Trim$(FileDialog.sFileTitle) = "" Then
            LastCharacter = 0
            count = 0
            
            While ShowOpen.nFilesSelected = 0
            MsgBox count
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)
                If count > 0 Then
                    tempFiles(count) = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                count = count + 1
                If InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) Then
                    tempFiles(count) = Mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = count
                End If
                LastCharacter = NewCharacter
            Wend
            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            For count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(count) = tempFiles(count)
                
            Next
        Else
            
            'Fix by me... to open multipile file....
            ShowOpen.sLastDirectory = left$(FileDialog.sFile, FileDialog.nFileOffset - 1)
            
            Dim sfil As Variant
            Dim sfils As Variant
            Dim cnt As Long
            
            sfils = Split(FileDialog.sFile, Chr$(0))
            For Each sfil In sfils
                ReDim Preserve ShowOpen.sFiles(cnt + 1)
                ShowOpen.sFiles(cnt) = sfil
                cnt = cnt + 1
            Next
            
            If cnt = 4 Then
                ShowOpen.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
                ShowOpen.nFilesSelected = 1
            Else
                ShowOpen.nFilesSelected = cnt - 4
            End If
            
            
            
        End If
        ShowOpen.bCanceled = False
        Exit Function
    Else
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    End If
End Function

Public Function ShowSave(ByVal hwnd As Long, Optional ShowStrFileName As String, Optional ByVal centerForm As Boolean = True) As SelectedFile
On Error Resume Next
Dim ret As Long
Dim hInst As Long
Dim Thread As Long
    
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = ShowStrFileName & Space$(2047) & Chr$(0) 'Space$(2047) & chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    If FileDialog.Flags = 0 Then
        FileDialog.Flags = OFS_FILE_SAVE_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetSaveFileName(FileDialog)
    'MsgBox 1
    ReDim ShowSave.sFiles(1)

    If ret Then
        ShowSave.sLastDirectory = left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        ShowSave.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        ShowSave.bCanceled = False
        Exit Function
    Else
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles
        Exit Function
    End If
End Function

Public Function ShowPrinter(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As Long
On Error Resume Next
Dim hInst As Long
Dim Thread As Long
    
    ParenthWnd = hwnd
    PrintDialog.hwndOwner = hwnd
    PrintDialog.lStructSize = Len(PrintDialog)
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ShowPrinter = PrintDlg(PrintDialog)
End Function
Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        X = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.right - rectMsg.left) / 2
        Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.bottom - rectMsg.top) / 2
        'Debug.Print "Screen " & Screen.Height / 2
        'Debug.Print "MsgBox " & (rectMsg.Right - rectMsg.Left) / 2
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.left + (rectForm.right - rectForm.left) / 2) - ((rectMsg.right - rectMsg.left) / 2)
        Y = (rectForm.top + (rectForm.bottom - rectForm.top) / 2) - ((rectMsg.bottom - rectMsg.top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
End Function


Public Function ShowColor(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedColor
Dim customcolors() As Byte  ' dynamic (resizable) array
Dim i As Integer
Dim ret As Long
Dim hInst As Long
Dim Thread As Long

    
    
    ParenthWnd = hwnd
    If ColorDialog.lpCustColors = "" Then
        ReDim customcolors(0 To 63) As Byte

    '
    '    For i = LBound(customcolors) To UBound(customcolors)
    '      customcolors(i) = 160 ' sets all custom colors to white
    '    Next i
    '
    

           customcolors(0) = 139   'red
           customcolors(1) = 155   'green
           customcolors(2) = 184   'blue
           'box2
           customcolors(4) = 188   'red
           customcolors(5) = 213   'green
           customcolors(6) = 254   'blue
           'box3
           customcolors(8) = 115   'red
           customcolors(9) = 172   'green
           customcolors(10) = 183  'blue
           'box4
           customcolors(12) = 200    'red
           customcolors(13) = 249   'green
           customcolors(14) = 198   'blue
           'box5
           customcolors(16) = 189   'red
           customcolors(17) = 194   'green
           customcolors(18) = 253   'blue
           'box6
           customcolors(20) = 200    'red
           customcolors(21) = 249   'green
           customcolors(22) = 255   'blue
           'box7
           customcolors(24) = 108    'red
           customcolors(25) = 213   'green
           customcolors(26) = 210   'blue
           'box8
           customcolors(28) = 236   'red
           customcolors(29) = 164   'green
           customcolors(30) = 236   'blue
           
           customcolors(32) = 160   'red
           customcolors(33) = 160  'green
           customcolors(34) = 160  'blue
           '10
           customcolors(36) = 160   'red
           customcolors(37) = 160  'green
           customcolors(38) = 160  'blue
           '11
           customcolors(40) = 160   'red
           customcolors(41) = 160  'green
           customcolors(42) = 160  'blue
           '12
           customcolors(44) = 160   'red
           customcolors(45) = 160  'green
           customcolors(46) = 160  'blue
           '13
           customcolors(48) = 160   'red
           customcolors(49) = 160  'green
           customcolors(50) = 160  'blue
           '14
           customcolors(52) = 160   'red
           customcolors(53) = 160  'green
           customcolors(54) = 160  'blue
           '15
           customcolors(56) = 160   'red
           customcolors(57) = 160  'green
           customcolors(58) = 160  'blue
           '16
           customcolors(60) = 160  'red
           customcolors(61) = 160  'green
           customcolors(62) = 160  'blue
           
        ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
        
    End If
    ColorDialog.hwndOwner = hwnd
    ColorDialog.lStructSize = Len(ColorDialog)
    ColorDialog.Flags = COLOR_FLAGS
    ColorDialog.rgbResult = 16777215 'Display wite color
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = ChooseColor(ColorDialog)
    If ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else
        'return -1 if canceled
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = -1 '&H0&
        Exit Function
    End If
End Function






