VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const RGN_OR = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const SWP_SHOWWINDOW = &H40
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2
Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Function MakePictureSkin(ByRef PictureFile As String, Optional MoveOnClick As Boolean = False) As Boolean
On Error GoTo mtm
Dim WindowRegion As Long
Dim Frm As New MakeFrm
    Frm.PicBox.ScaleMode = vbPixels
    Frm.PicBox.AutoRedraw = True
    Frm.PicBox.AutoSize = True
    Frm.PicBox.BorderStyle = vbBSNone
    Frm.BorderStyle = vbBSNone
    Frm.PicBox.Picture = LoadPicture(PictureFile)
    Frm.Width = Frm.PicBox.Width * 15
    Frm.Height = Frm.PicBox.Height * 15
    WindowRegion = MakeRegion(Frm.PicBox)
    SetWindowRgn Frm.hwnd, WindowRegion, True
    Frm.Picture = Frm.PicBox.Picture
    If MoveOnClick = True Then Frm.Movement = True
    Frm.Show
    Frm.Refresh
    MakePictureSkin = True
    Exit Function
mtm:
MakePictureSkin = False
End Function
Public Function CenterForm(ByVal frmForm As Object) As Boolean

    ' Comments  :
    ' Parameters: frmForm -
    ' Returns   : Boolean -
    ' Author    : Chris Gibbs
    ' Modified  :
    ' --------------------------------------------------
    frmForm.Move (Screen.Width - frmForm.Width) \ 2, (Screen.Height - frmForm.Height) \ 2
    CenterForm = True
End Function


Public Function FormOnTop(hWindow As Long, bTopMost As Boolean) As Boolean
    On Error GoTo mtm
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Dim wFlags As Long
    Dim Placement As Long
    
    
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
    FormOnTop = True
mtm:
FormOnTop = False
    
End Function



Private Function MakeRegion(picSkin As PictureBox) As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region.
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable enter the colors as HEX like &H000000FF&
    ' or RGB(100,100,100) or vbBlack, vbWhite etc...
    
    TransparentColor = GetPixel(hdc, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hdc, X, Y) = TransparentColor Or X = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function

Sub Make_Transparent(lhWnd As Long, Percentage As Integer)
 On Error GoTo HELL
 
  Dim OSV As OSVersionInfo
  OSV.OSVSize = Len(OSV)
  If GetVersionEx(OSV) <> 1 Then Exit Sub
       
  If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then Exit Sub '/* Win 98/ME
  If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then '/* Win 2000/XP
    Call SetWindowLong(lhWnd, GWL_EXSTYLE, GetWindowLong(lhWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(lhWnd, 0, (Percentage * 255) / 100, LWA_ALPHA)
  End If
Exit Sub
HELL:
End Sub




