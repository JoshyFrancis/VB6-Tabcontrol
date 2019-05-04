VERSION 5.00
Begin VB.UserControl TabControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0069BDFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "TabControl.ctx":0000
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1560
      Top             =   1920
   End
   Begin VB.PictureBox picSliderDsabled 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "TabControl.ctx":0312
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSliderNormal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "TabControl.ctx":05E8
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSliderDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "TabControl.ctx":08BE
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSliderHover 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      Picture         =   "TabControl.ctx":0B94
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Timer TimerCheckMouseOut 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   480
      Top             =   1560
   End
   Begin VB.PictureBox picInactiveTab 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Picture         =   "TabControl.ctx":0E6A
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picActiveTab 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Picture         =   "TabControl.ctx":2680
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox picBackColor 
      BackColor       =   &H0069BDFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picTabBackColor 
      BackColor       =   &H00D6AA88&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBorder 
      BackColor       =   &H00CF9365&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "TabControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*              I thank Almighty ever living God for all the graces He has showered upon me.
'*              All honour and glory and thanks and praise and worship belong to him, my Lord Jesus Christ.
'-----------------------------------------------------------------------------------------------------
'// Title:    TabControl
'// Author:   Joshy Francis
'// Version:  1.0
'// Copyright: All rights reserved
'-----------------------------------------------------------------------------------------------------

' This software is provided "as-is," without any express or implied warranty.
' In no event shall the author be held liable for any damages arising from the use of this software.
' If you do not agree with these terms, do not use "Grid". Use of the program implicitly means
' you have agreed to these terms.

' Permission is granted to anyone to use this software for any purpose,
' including commercial use, and to alter and redistribute it, provided that
' the following conditions are met:

' 1. All redistributions of source code files must retain all copyright
'    notices that are currently in place, and this list of conditions without
'    any modification.
' 2. All redistributions in binary form must retain all occurrences of the
'    above copyright notice and web site addresses that are currently in
'    place (for example, in the About boxes).
' 3. Modified versions in source or binary form must be plainly marked as
'    such, and must not be misrepresented as being the original software.
'-----------------------------------------------------------------------------------------------------

Option Explicit
'--- for MST subclassing (1)
#Const ImplNoIdeProtection = True ' (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1
Private Const SIGN_BIT                      As Long = &H80000000
Private Const EBMODE_DESIGN                 As Long = 0
 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If
Private m_pSubclass         As IUnknown
Private mUserMode As Boolean
Private mContainerHwnd As Long
Dim mUserControlHwnd As Long
Dim mUserControlHdc As Long
'--- End for MST subclassing (1)

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const HALFTONE = 4
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Const NULL_BRUSH = 5
Private Const NULL_PEN = 8
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Const DT_BOTTOM = &H8&
Private Const DT_CENTER = &H1&
Private Const DT_LEFT = &H0&
Private Const DT_CALCRECT = &H400&
Private Const DT_WORDBREAK = &H10&
Private Const DT_VCENTER = &H4&
Private Const DT_TOP = &H0&
Private Const DT_TABSTOP = &H80&
Private Const DT_SINGLELINE = &H20&
Private Const DT_RIGHT = &H2&
Private Const DT_NOCLIP = &H100&
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000&
Private Const DT_EXTERNALLEADING = &H200&
Private Const DT_EXPANDTABS = &H40&
Private Const DT_CHARSTREAM = 4&
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_END_ELLIPSIS = &H8000&
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function apiTranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Enum TabCaptionAlignment
   TabCaptionAlignLeftTop = DT_LEFT Or DT_TOP
   TabCaptionAlignLeftCenter = DT_LEFT Or DT_VCENTER
   TabCaptionAlignLeftBottom = DT_LEFT Or DT_BOTTOM
   TabCaptionAlignCenterTop = DT_CENTER Or DT_TOP
   TabCaptionAlignCenterCenter = DT_CENTER Or DT_VCENTER
   TabCaptionAlignCenterBottom = DT_CENTER Or DT_BOTTOM
   TabCaptionAlignRightTop = DT_RIGHT Or DT_TOP
   TabCaptionAlignRightCenter = DT_RIGHT Or DT_VCENTER
   TabCaptionAlignRightBottom = DT_RIGHT Or DT_BOTTOM
   TabCaptionSingleLine = DT_SINGLELINE
   TabCaptionWordWrap = DT_WORDBREAK
   TabCaptionEllipsis = DT_WORD_ELLIPSIS
   TabCaptionNoPrefix = DT_NOPREFIX
End Enum
Public Event TabClick(ByVal TabIndex As Long)
Public Event BeforeTabChange(ByVal LastTab As Long, ByRef NewTab As Long)
Public Event TabOrderChanged(ByVal LastTab As Long, ByRef NewTab As Long)

Private Type tTabInfo
    sCaption As String
    sTag  As String
    sKey As String
    ItemData As Long
    Left As Long
    Width As Long
    Image As Long
    Alignment As Long ' TabCaptionAlignment
    Enabled As Boolean
    Visible As Boolean
    Active As Boolean
    Top As Long
    Height As Long
    Right As Long
    Bottom As Long
End Type
Private Type tControlInfo
    ctlName As String
    iTabIndex As Long
End Type

Private m_Tabs() As tTabInfo
Private m_TabCount As Long
Private m_TabOrder() As Long
Private m_Ctls() As tControlInfo
Private m_ctlCount As Long

Private mhDC As Long, hBmp As Long, hBmpOld As Long
Private lScrollX As Long, lScrollWidth As Long
Private lTabIndex As Long
Private cx As Long, cy As Long
Private lTabDragging As Boolean
Private lTabDragged As Long
Private lTabHover As Long
Private mFont As IFont
Private selIndex As Long
Private Enum ssSliderStatus
    ssDisabled
    ssNormal
    ssDown
    ssHover
End Enum
Private SliderBox As RECT
Private bSliderShown As Boolean
Private LeftSliderStatus As ssSliderStatus
Private RightSliderStatus As ssSliderStatus
Private bInFocus As Boolean
Private m_AllowReorder As Boolean

'for contained controls
'   store control name, index and tabIndex: ctl1(0)1 or clt1()1 if there is no index
'Private ctlLst() As New Collection ', visibleCTLs As New Collection

'Autor: wqweto http://www.vbforums.com/showthread.php?872819
'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================
Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As TabControl
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwEB4BV1aLdCQUg8YIgz4AdC+L+oHHABIeAYvCBQgRHgGri8IFRBEeAauLwgVUER4Bq4vCBXwRHgGruQkAAADzpYHCABIeAVJqGP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlg+8YagBX/3IM/3cM/1IYi0QkGIk4Xl+4NBIeAS1wEB4BwhAAZpCLRCQIgzgAdSqDeAQAdSSBeAjAAAAAdRuBeAwAAABGdRKLVCQE/0IEi0QkDIkQM8DCDAC4AkAAgMIMAJCLVCQE/0IEi0IEwgQADx8Ai1QkBP9KBItCBHUYiwpS/3EM/3IM/1Eci1QkBIsKUv9RFDPAwgQAkFWL7ItVGIsKi0EshcB0OFL/0FqJQgiD+AF3VIP4AHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1EYsK/3UU/3UQ/3UM/3UI/1EgXcIYAA==" ' 1.4.2019 11:41:46
    Const THUNK_SIZE    As Long = 452
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitSubclassingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitSubclassingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, 4)
    End If
End Property

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & App.hInstance & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & App.hInstance & "_" & sKey, lValue)
End Property
#End If
 
Private Sub pvSubclass()
        pvUnsubclass
    If mUserControlHwnd <> 0 Then
        Set m_pSubclass = InitSubclassingThunk(mUserControlHwnd, Me, InitAddressOfMethod(Me, 5).SubclassProc(0, 0, 0, 0, 0))
    End If
End Sub

Private Sub pvUnsubclass()
    Set m_pSubclass = Nothing
End Sub
Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
'Dim x As Single
'Dim y As Single
'
'    Select Case wMsg
'        Case WM_LBUTTONDOWN
'            x = (lParam And &HFFFF&)
'            y = (lParam \ &H10000 And &HFFFF&)
'
'    End Select
'    '--- note: performance optimization for design-time subclassing
'    If Not Handled And ThunkPrivateData(m_pSubclass) = EBMODE_DESIGN Then
'        Handled = True
'        SubclassProc = DefSubclassProc(hwnd, wMsg, wParam, lParam)
'    End If
    If mUserMode = True Then
        Handled = True
        SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    End If

Dim lTab As Long
Dim x As Single
Dim y As Single
    Select Case wMsg
         Case WM_LBUTTONDOWN ' UserControl message, only in design mode (Not Ambient.UserMode), to provide change of selected tab by clicking at design time
    '            If TypeOf UserControl.Parent Is Form Then
    '                 UserControl.Parent.Caption = Timer
    '            End If
'                If Not MouseIsOverAContainedControl Then
                    lTab = selIndex
    '                Call ProcessMouseMove(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
    '                    X = (lParam And &HFFFF&) * Screen_TwipsPerPixelX
    '                    Y = (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixely
                        x = (lParam And &HFFFF&)
                        y = (lParam \ &H10000 And &HFFFF&)
    '                    UserControl.Parent.Caption = "X " & X & " Y " & Y
                    Call UserControl_MouseDown(vbLeftButton, -1, x, y)
                    If selIndex <> lTab Then
                        Handled = True
                        SubclassProc = 0
                        Exit Function
                    End If
'                End If
    '            If mChangeControlsBackColor And (mTabBackColor <> vbButtonFace) Then
    '                mLastContainedControlsCount = UserControl.ContainedControls.Count
    '                tmrCheckContainedControlsAdditionDesignTime.Enabled = True
    '            End If
    End Select
        Handled = False
        SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)

End Function
'--- End for MST subclassing (2)
Private Function MouseIsOverAContainedControl() As Boolean
    Dim iPt As POINTAPI
    Dim iSM As Long
    Dim iCtl As Control
    Dim iWidth As Long
    
    iSM = UserControl.ScaleMode
    UserControl.ScaleMode = vbTwips
    GetCursorPos iPt
    ScreenToClient mUserControlHwnd, iPt
    iPt.x = iPt.x * Screen.TwipsPerPixelX
    iPt.y = iPt.y * Screen.TwipsPerPixelY
    
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        iWidth = -1
        iWidth = iCtl.Width
        If iWidth <> -1 Then
            If iCtl.Left <= iPt.x Then
                If iCtl.Left + iCtl.Width >= iPt.x Then
                    If iCtl.Top <= iPt.y Then
                        If iCtl.Top + iCtl.Height >= iPt.y Then
                            MouseIsOverAContainedControl = True
                            Err.Clear
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
    UserControl.ScaleMode = iSM
End Function
Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Private Sub TimerCheckMouseOut_Timer()
    Dim Pos As POINTAPI
    Dim WFP As Long
    
    GetCursorPos Pos
    WFP = WindowFromPoint(Pos.x, Pos.y)
    
    If WFP <> Me.hWnd Then
        UserControl_MouseMove -1, 0, -1, -1
        TimerCheckMouseOut.Enabled = False 'kill that timer at once
    End If
End Sub

Private Sub tmrSlide_Timer()
If tmrSlide.Enabled = False Then
    Exit Sub
End If
    If RightSliderStatus = ssDown Then
        lScrollX = lScrollX + 10
            If lScrollX > lScrollWidth Then
                    lScrollX = lScrollWidth
                RightSliderStatus = ssDisabled
                        tmrSlide.Enabled = False
                    Me.Refresh
                    Exit Sub
            End If
                LeftSliderStatus = ssNormal
    Else
        lScrollX = lScrollX - 10
            If lScrollX < 0 Then
                lScrollX = 0
                LeftSliderStatus = ssDisabled
                    tmrSlide.Enabled = False
                Me.Refresh
                Exit Sub
            End If
                RightSliderStatus = ssNormal
    End If
'Draw
Me.Refresh
    If tmrSlide.Interval > 10 Then tmrSlide.Interval = tmrSlide.Interval - 2
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "UserMode" Then mUserMode = Ambient.UserMode
End Sub

Private Sub UserControl_EnterFocus()
    bInFocus = True
    Me.Refresh
End Sub

Private Sub UserControl_ExitFocus()
    bInFocus = False
    Me.Refresh
End Sub

Private Sub UserControl_Initialize()
            Me.Refresh

End Sub

Private Sub UserControl_InitProperties()
Dim c As Long
    On Error GoTo 0
        mUserMode = Ambient.UserMode
        mUserControlHwnd = UserControl.hWnd
        mContainerHwnd = UserControl.ContainerHwnd
        mUserControlHdc = UserControl.hDC

    pvSubclass
                CreateGraphicsDC
    
        For c = 1 To 4
            AddTab , "Tab " & c
        Next
End Sub
Private Function getTabOrder(ByVal Index As Long) As Long
Dim c As Long
    For c = 0 To m_TabCount - 1
        If m_TabOrder(c) = Index Then
            getTabOrder = c
            Exit For
        End If
    Next
End Function
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lTab As Long, lCount As Long
'    lTab = selIndex
     lTab = getTabOrder(selIndex)
NextTab:
    lCount = lCount + 1
Select Case KeyCode
    Case vbKeyLeft
        If lCount > 1 Then
            lTab = m_TabCount ' + 1
        End If
        If lTab > 0 Then
            lTab = lTab - 1
        End If
Case vbKeyRight
        If lCount > 1 Then
            lTab = 0
        End If
        If lTab < m_TabCount - 1 Then
            lTab = lTab + 1
        End If
End Select
    If lTab >= 0 And lTab < m_TabCount And lCount <= m_TabCount Then
        If m_Tabs(m_TabOrder(lTab)).Enabled = False Or m_Tabs(m_TabOrder(lTab)).Visible = False Then
            GoTo NextTab:
        End If
    End If
        lTab = m_TabOrder(lTab)
If selIndex <> lTab Then

    RaiseEvent BeforeTabChange(selIndex, lTab)
        RaiseEvent TabClick(lTab)
    SelectedItem = lTab
End If
End Sub

Private Sub UserControl_Paint()
mUserControlHdc = hDC
'    Draw
    Me.Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mFont = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = mFont
'    Set mFontActive = PropBag.ReadProperty("FontActive", Ambient.Font)
'    Set mFontHover = PropBag.ReadProperty("FontHover", Ambient.Font)
'    Set mFontDisabled = PropBag.ReadProperty("FontDisabled", Ambient.Font)
    m_AllowReorder = PropBag.ReadProperty("AllowReorder", False)
    selIndex = PropBag.ReadProperty("SelectedItem", 0)
    m_TabCount = PropBag.ReadProperty("ItemCount", 4)
   
    ReDim m_Tabs(m_TabCount - 1)
    ReDim m_TabOrder(m_TabCount - 1)

    ReDim ctlLst(m_TabCount - 1)
        m_ctlCount = 0
    ReDim m_Ctls(m_ctlCount)
        
    Dim i As Long, z As Long
    Dim mCCount As Long, ctlName As String, ItemMax As Long
    For i = 0 To m_TabCount - 1
'            m_TabOrder(I) = I
            m_TabOrder(i) = PropBag.ReadProperty("TabOrder" & i, i)
        m_Tabs(i).Image = PropBag.ReadProperty("TabIcon" & i, 0)
        m_Tabs(i).Enabled = PropBag.ReadProperty("TabEnabled" & i, True)
        m_Tabs(i).sKey = PropBag.ReadProperty("Key" & i, "")
        m_Tabs(i).sTag = PropBag.ReadProperty("TabTag" & i, "")
        m_Tabs(i).Visible = PropBag.ReadProperty("TabVisible" & i, True)
        
        m_Tabs(i).sCaption = PropBag.ReadProperty("Item(" & i & ").Caption", "Tab " & i + 1)
        mCCount = PropBag.ReadProperty("Item(" & i & ").ControlCount", 0)
                
        For z = 0 To mCCount - 1
                ctlName = PropBag.ReadProperty("Item(" & i & ").Control(" & z & ")", "")
            If ctlName <> "" Then
                ReDim Preserve m_Ctls(m_ctlCount)
                   m_Ctls(m_ctlCount).ctlName = ctlName
                   m_Ctls(m_ctlCount).iTabIndex = i
                m_ctlCount = m_ctlCount + 1
            End If
        Next z
        
    Next i
        
        ItemMax = PropBag.ReadProperty("ItemMax", 0)
    For i = m_TabCount To ItemMax
        mCCount = PropBag.ReadProperty("Item(" & i & ").ControlCount", 0)
        For z = 0 To mCCount - 1
                ctlName = PropBag.ReadProperty("Item(" & i & ").Control(" & z & ")", "")
            If ctlName <> "" Then
                ReDim Preserve m_Ctls(m_ctlCount)
                   m_Ctls(m_ctlCount).ctlName = ctlName
                   m_Ctls(m_ctlCount).iTabIndex = i
                m_ctlCount = m_ctlCount + 1
            End If
        Next z
    Next
        If selIndex > m_TabCount - 1 Then
            selIndex = m_TabCount - 1
        End If
''    handleControls 0, selIndex
     SelectedItem = selIndex
        
    On Error GoTo 0
        mUserMode = Ambient.UserMode
        mUserControlHwnd = UserControl.hWnd
        mContainerHwnd = UserControl.ContainerHwnd
        mUserControlHdc = UserControl.hDC

    pvSubclass
        CreateGraphicsDC
'        Draw
    Me.Refresh
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Font", mFont, Ambient.Font
'    PropBag.WriteProperty "FontActive", mFontActive, Ambient.Font
'    PropBag.WriteProperty "FontHover", mFontHover, Ambient.Font
'    PropBag.WriteProperty "FontDisabled", mFontDisabled, Ambient.Font
    PropBag.WriteProperty "AllowReorder", m_AllowReorder, False
    PropBag.WriteProperty "SelectedItem", selIndex, 0
    
    PropBag.WriteProperty "ItemCount", m_TabCount, 4
    '
    Dim i As Long, z As Long, c As Long, MaxIndex As Long
    For i = 0 To m_TabCount - 1
        PropBag.WriteProperty "TabOrder" & i, m_TabOrder(i), i
        PropBag.WriteProperty "TabIcon" & i, m_Tabs(i).Image, 0
        PropBag.WriteProperty "TabEnabled" & i, m_Tabs(i).Enabled, True
        PropBag.WriteProperty "TabTag" & i, m_Tabs(i).sTag, ""
        PropBag.WriteProperty "TabKey" & i, m_Tabs(i).sKey, ""
        PropBag.WriteProperty "TabVisible" & i, m_Tabs(i).Visible, True
        PropBag.WriteProperty "Item(" & i & ").Caption", m_Tabs(i).sCaption, "Tab " & i + 1
                c = 0
            For z = 0 To m_ctlCount - 1
                If m_Ctls(z).iTabIndex = i Then
                    PropBag.WriteProperty "Item(" & i & ").Control(" & c & ")", m_Ctls(z).ctlName, ""
                    c = c + 1
                End If
                If MaxIndex < m_Ctls(z).iTabIndex Then
                    MaxIndex = m_Ctls(z).iTabIndex
                End If
            Next z
            PropBag.WriteProperty "Item(" & i & ").ControlCount", c, 0
    Next i
        PropBag.WriteProperty "ItemMax", MaxIndex, 0
        For i = m_TabCount To MaxIndex
                c = 0
            For z = 0 To m_ctlCount - 1
                If m_Ctls(z).iTabIndex = i Then
                    PropBag.WriteProperty "Item(" & i & ").Control(" & c & ")", m_Ctls(z).ctlName, ""
                    c = c + 1
                End If
            Next z
            PropBag.WriteProperty "Item(" & i & ").ControlCount", c, 0
        Next
End Sub
Private Function TranslateColor(ByVal OLE_COLOR As Long) As Long
        apiTranslateColor OLE_COLOR, 0, TranslateColor
End Function
Private Sub TransBlt(ByVal hdcScreen As Long, ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal clrMask As OLE_COLOR)
    
'one check to see if GdiTransparentBlt is supported
'better way to check if function is suported is using LoadLibrary and GetProcAdress
'than using GetVersion or GetVersionEx
'=====================================================
    Dim Lib As Long
    Dim ProcAdress As Long
    Dim lMaskColor As Long
    lMaskColor = TranslateColor(clrMask)
    Lib = LoadLibrary("gdi32.dll")
    '--------------------->make sure to specify corect name for function
    ProcAdress = GetProcAddress(Lib, "GdiTransparentBlt")
    FreeLibrary Lib
    If ProcAdress <> 0 Then
        'works on XP
        GdiTransparentBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcSrc, xSrc, ySrc, nWidthSrc, nHeightSrc, lMaskColor
        'Debug.Print "Gdi transparent blt"
        Exit Sub 'make it short
    End If
'=====================================================
    Const DSna              As Long = &H220326
    Dim hdcMask             As Long
    Dim hdcColor            As Long
    Dim hbmMask             As Long
    Dim hbmColor            As Long
    Dim hbmColorOld         As Long
    Dim hbmMaskOld          As Long
    Dim hdcScnBuffer        As Long
    Dim hbmScnBuffer        As Long
    Dim hbmScnBufferOld     As Long
    

    
   
   lMaskColor = TranslateColor(clrMask)
   hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hdcScnBuffer = CreateCompatibleDC(hdcScreen)
   hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)

   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy

   hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hbmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)

   hdcColor = CreateCompatibleDC(hdcScreen)
   hbmColorOld = SelectObject(hdcColor, hbmColor)
    
   Call SetBkColor(hdcColor, GetBkColor(hdcSrc))
   Call SetTextColor(hdcColor, GetTextColor(hdcSrc))
'   Call BitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)
   Call StretchBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, nWidthSrc, nHeightSrc, vbSrcCopy)

   hdcMask = CreateCompatibleDC(hdcScreen)
   hbmMaskOld = SelectObject(hdcMask, hbmMask)

   SetBkColor hdcColor, lMaskColor
   SetTextColor hdcColor, vbWhite
   BitBlt hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy
'   StretchBlt hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, nWidthSrc, nHeightSrc, vbSrcCopy
 
   SetTextColor hdcColor, vbBlack
   SetBkColor hdcColor, vbWhite
   BitBlt hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, DSna
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint
   BitBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy
     
     'clear
   DeleteObject SelectObject(hdcColor, hbmColorOld)
   DeleteDC hdcColor
   DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
   DeleteDC hdcScnBuffer
   DeleteObject SelectObject(hdcMask, hbmMaskOld)
   
   DeleteDC hdcMask
   'ReleaseDC 0, hdcScreen
End Sub

Sub Draw()
Dim hFont As Long, hFontOld As Long
Dim hBrush As Long, hBrushOld As Long
Dim hPen As Long, hPenOld As Long
Dim i As Long, c As Long, x As Long
Dim rcCalc As RECT, rc As RECT, lWidth As Long, lHeight As Long
Dim W As Long, H As Long, OldTextColor As Long
        If mFont Is Nothing Then
            Set mFont = UserControl.Font
        End If
        hFontOld = SelectObject(mhDC, mFont.hFont)
'
'''    CopyPic picActiveTab, picActiveTab.hDC
''    BitBlt mhDC, 4, 1, 87, 28, picActiveTab.hdc, 1, 1, vbSrcCopy
'    StretchBlt mhDC, 4, 1, 87 + 45, 25, picInactiveTab.hdc, 1, 1, 87, 25, vbSrcCopy
'            TextOut mhDC, 8, 4, "Jesus", 5
'
'    StretchBlt mhDC, 110, 1, 87 + 45, 25, picActiveTab.hdc, 1, 1, 87, 25, vbSrcCopy
'            TextOut mhDC, 8 + 87 + 45, 4, "Jesus", 5
        lWidth = ScaleX(ScaleWidth, ScaleMode, 3)
        lHeight = ScaleY(ScaleHeight, ScaleMode, 3)
    x = 4 + -lScrollX
For i = 0 To m_TabCount - 1
        c = m_TabOrder(i)
    With m_Tabs(c)
        If .Visible = True Then
                rcCalc.Left = 0
                rcCalc.Right = lWidth
                rcCalc.Top = 0
                rcCalc.Bottom = lHeight
            DrawText mhDC, .sCaption, Len(.sCaption), rcCalc, DT_CALCRECT Or .Alignment
            .Left = x
'            .rc.Right = .rc.Left + 50
'                If rcCalc.right < 60 Then
'                    rcCalc.right = 60
'                End If
            .Right = .Left + (rcCalc.Right + 36) ' 48)
            .Width = .Right - .Left
            .Top = 1
            .Bottom = .Top + rcCalc.Bottom + 8
'            X = .rc.right - 10
            x = x + (.Right - .Left) - 14
        End If
    End With
Next
        Dim b As Boolean
                b = bSliderShown
            bSliderShown = x + lScrollX + 32 >= lWidth
        lScrollWidth = ((x + lScrollX) - lWidth) + 48 '50
            If b = False And bSliderShown = True Then
                RightSliderStatus = ssNormal
            End If
        If lScrollWidth < 0 Then
            lScrollX = 0
        End If
    hBrush = CreateSolidBrush(picTabBackColor.BackColor)
    hBrushOld = SelectObject(mhDC, hBrush)
'        Rectangle mhDC, -1, -1, lWidth + 1, rcCalc.bottom + 7 ' 28
        Rectangle mhDC, -1, -1, lWidth + 1, rcCalc.Bottom + 9 ' 28
    DeleteObject SelectObject(mhDC, hBrushOld)
    
    hBrush = CreateSolidBrush(picBackColor.BackColor)
    hBrushOld = SelectObject(mhDC, hBrush)
    hPen = CreatePen(0, 1, picBorder.BackColor)
    hPenOld = SelectObject(mhDC, hPen)
'        Rectangle mhDC, 0, 24, lWidth, lHeight
        Rectangle mhDC, 0, rcCalc.Bottom + 8, lWidth, lHeight
    DeleteObject SelectObject(mhDC, hBrushOld)
    DeleteObject SelectObject(mhDC, hPenOld)

For i = m_TabCount - 1 To 0 Step -1
        c = m_TabOrder(i)
   If c <> selIndex Then
        With m_Tabs(c)
'            If .Left > 0 And .Right < lWidth Then
            If .Left > -.Width And .Right < lWidth + .Width Then
                DrawTab c
            End If
        End With
    End If
Next
If m_TabCount > 0 Then
    DrawTab selIndex
End If
    If lTabDragging = True And lTabHover <> lTabDragged And lTabHover >= 0 And lTabHover < m_TabCount And m_TabCount > 0 Then
                 c = m_TabOrder(lTabHover)
           OldTextColor = SetTextColor(mhDC, vbBlue)
                DrawTab c
            SetTextColor mhDC, OldTextColor
       With m_Tabs(c)
             rc.Left = .Left: rc.Top = .Top: rc.Right = .Right: rc.Bottom = .Bottom
                rc.Left = rc.Left + 4
                rc.Right = rc.Right - 24
                rc.Top = rc.Top + 4
                rc.Bottom = rc.Bottom - 4
'            DrawFocusRect mhDC, RC
            hPen = CreatePen(2, 1, vbRed)
            hPenOld = SelectObject(mhDC, hPen)
            hBrush = GetStockObject(NULL_BRUSH)
            hBrushOld = SelectObject(mhDC, hBrush)
                Rectangle mhDC, rc.Left, rc.Top, rc.Right, rc.Bottom
            DeleteObject SelectObject(mhDC, hPenOld)
            DeleteObject SelectObject(mhDC, hBrushOld)
        End With
    End If
If bSliderShown Then
        
    With SliderBox
        .Top = 0
        .Bottom = rcCalc.Bottom + 8
        .Left = lWidth - IIf(.Bottom > 32, .Bottom, 32)
        .Right = lWidth
        hBrush = CreateSolidBrush(picTabBackColor.BackColor)
        hBrushOld = SelectObject(mhDC, hBrush)
        hPen = CreatePen(0, 1, picTabBackColor.BackColor)
        hPenOld = SelectObject(mhDC, hPen)
            Rectangle mhDC, .Left, .Top, .Right, .Bottom
        DeleteObject SelectObject(mhDC, hBrushOld)
        DeleteObject SelectObject(mhDC, hPenOld)
            W = ((.Right - .Left) - 4) / 2
            H = ((.Bottom - .Top) - 4) / 2
                If W < 14 Then
                    W = 14
                End If
                If H < 15 Then
                    H = 15
                End If
        Select Case LeftSliderStatus
            Case ssSliderStatus.ssDisabled
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderDsabled.hDC, 0, 0, 14, 15, vbSrcCopy
            Case ssSliderStatus.ssNormal
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderNormal.hDC, 0, 0, 14, 15, vbSrcCopy
            Case ssSliderStatus.ssDown
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderDown.hDC, 0, 0, 14, 15, vbSrcCopy
            Case ssSliderStatus.ssHover
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, -W, H, picSliderHover.hDC, 0, 0, 14, 15, vbSrcCopy
        End Select
        Select Case RightSliderStatus
            Case ssSliderStatus.ssDisabled
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderDsabled.hDC, 0, 0, 14, 15, vbSrcCopy
            Case ssSliderStatus.ssNormal
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderNormal.hDC, 0, 0, 14, 15, vbSrcCopy
            Case ssSliderStatus.ssDown
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderDown.hDC, 0, 0, 14, 15, vbSrcCopy
            Case ssSliderStatus.ssHover
                StretchBlt mhDC, .Left + W + 2, .Top + H \ 2, W, H, picSliderHover.hDC, 0, 0, 14, 15, vbSrcCopy
        End Select
    End With
End If

    BitBlt mUserControlHdc, 0, 0, lWidth, lHeight, mhDC, 0, 0, vbSrcCopy
'    BitBlt mUserControlHdc, 0, 0, 87, 28, picActiveTab.hDC, 0, 0, vbSrcCopy
End Sub
Private Sub DrawTab(ByVal i As Long)
Dim OldTextColor As Long
Dim rc As RECT
If i > -1 And i < m_TabCount And m_TabCount > 0 Then
    With m_Tabs(i)
             rc.Left = .Left: rc.Top = .Top: rc.Right = .Right: rc.Bottom = .Bottom
    '    StretchBlt mhDC, .rc.Left, .rc.Top, .rc.Right, .rc.Bottom, picInactiveTab.hDC, 1, 1, 81, 25, vbSrcCopy
    '    GdiTransparentBlt mhDC, .rc.Left, .rc.Top, .rc.Right, .rc.Bottom, picInactiveTab.hdc, 1, 1, 81, 25, 0 'GetPixel(picInactiveTab.hdc, 1, 1)
        If selIndex = i Then
'            TransBlt mUserControlHdc, mhDC, RC.left, RC.Top, RC.Right - RC.left, RC.Bottom - RC.Top, picActiveTab.hDC, 1, 1, 80, 24, GetPixel(picActiveTab.hDC, 1, 1)
            TransBlt mUserControlHdc, mhDC, rc.Left, rc.Top, 4, rc.Bottom - rc.Top, picActiveTab.hDC, 1, 1, 4, 24, GetPixel(picActiveTab.hDC, 1, 1)
            TransBlt mUserControlHdc, mhDC, rc.Left + 4, rc.Top, (rc.Right - rc.Left) - 28, rc.Bottom - rc.Top, picActiveTab.hDC, 1 + 4, 1, 80 - 24, 24, GetPixel(picActiveTab.hDC, 1, 1)
            TransBlt mUserControlHdc, mhDC, rc.Left + (rc.Right - rc.Left) - 28, rc.Top, 24, rc.Bottom - rc.Top, picActiveTab.hDC, 1 + 80 - 24, 1, 24, 24, GetPixel(picActiveTab.hDC, 1, 1)
        Else
'            TransBlt mUserControlHdc, mhDC, RC.left, RC.Top, RC.Right - RC.left, RC.Bottom - RC.Top, picInactiveTab.hDC, 1, 1, 80, 24, GetPixel(picInactiveTab.hDC, 1, 1)
'            'Only Stretching middle part
            TransBlt mUserControlHdc, mhDC, rc.Left, rc.Top, 4, rc.Bottom - rc.Top, picInactiveTab.hDC, 1, 1, 4, 24, GetPixel(picInactiveTab.hDC, 1, 1)
            TransBlt mUserControlHdc, mhDC, rc.Left + 4, rc.Top, (rc.Right - rc.Left) - 28, rc.Bottom - rc.Top, picInactiveTab.hDC, 1 + 4, 1, 80 - 24, 24, GetPixel(picInactiveTab.hDC, 1, 1)
            TransBlt mUserControlHdc, mhDC, rc.Left + (rc.Right - rc.Left) - 28, rc.Top, 24, rc.Bottom - rc.Top, picInactiveTab.hDC, 1 + 80 - 24, 1, 24, 24, GetPixel(picInactiveTab.hDC, 1, 1)
        End If
'            TextOut mhDC, .rc.left + 8, .rc.top + 4, .sCaption, Len(.sCaption)
            rc.Left = rc.Left + 8
            rc.Right = rc.Right - 8
            rc.Top = rc.Top + 4
            rc.Bottom = rc.Bottom - 4
                If .Enabled = False Then
                    OldTextColor = SetTextColor(mhDC, RGB(128, 128, 128))
                End If
            DrawText mhDC, .sCaption, Len(.sCaption), rc, .Alignment
                If .Enabled = False Then
                    SetTextColor mhDC, OldTextColor
                End If
        If selIndex = i Then
            If bInFocus = True Then
                DrawText mhDC, .sCaption, Len(.sCaption), rc, DT_CALCRECT Or .Alignment
                rc.Left = rc.Left - 2
                rc.Right = rc.Right + 4
                DrawFocusRect mhDC, rc
            End If
        End If
    End With
End If
End Sub
Sub CreateGraphicsDC()
Dim lWidth As Long, lHeight As Long
        lWidth = ScaleX(ScaleWidth, ScaleMode, 3)
        lHeight = ScaleY(ScaleHeight, ScaleMode, 3)
        DestroyGraphicsDC
    mhDC = CreateCompatibleDC(0)
        hBmp = CreateCompatibleBitmap(mUserControlHdc, lWidth, lHeight)
    hBmpOld = SelectObject(mhDC, hBmp)
        SetBkMode mhDC, TRANSPARENT
End Sub
Sub DestroyGraphicsDC()
    If mhDC Then
        DeleteObject SelectObject(mhDC, hBmpOld)
        DeleteDC mhDC
        mhDC = 0
    End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    CreateGraphicsDC
''        Draw
        Me.Refresh
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Sub UserControl_Terminate()
        DestroyGraphicsDC
    pvUnsubclass
Erase m_Tabs
Erase m_TabOrder
    m_TabCount = 0
Erase m_Ctls
    m_ctlCount = 0
End Sub
Public Function FindTabByKey(ByVal sKey As String) As Long
Dim c As Long
    FindTabByKey = -1
        For c = 0 To m_TabCount - 1
            If m_Tabs(c).sKey = sKey Then
                FindTabByKey = c
                Exit For
            End If
        Next
End Function

Public Function AddTab( _
      Optional ByVal sKey As String, _
      Optional ByVal sCaption As String, _
      Optional ByVal lWidth As Long = 64 _
      , Optional ByVal ItemData As Long _
      , Optional ByVal lImage As Long _
   ) As Long
           ReDim Preserve m_Tabs(m_TabCount)
        With m_Tabs(m_TabCount)
                .sCaption = sCaption
                .sKey = sKey
                .Width = lWidth
                .ItemData = ItemData
                .Image = lImage
                .Alignment = DT_VCENTER Or DT_NOPREFIX Or DT_END_ELLIPSIS
                .Visible = True
                .Enabled = True
        End With
                ReDim Preserve m_TabOrder(m_TabCount)
                    m_TabOrder(m_TabCount) = m_TabCount
            m_TabCount = m_TabCount + 1
AddTab = m_TabCount
            
''    Draw
    Me.Refresh
End Function

Public Sub RemoveTab(ByVal nIndex As Long)
Dim j As Long
Dim itemOrder As Long
If nIndex < m_TabCount Then
        itemOrder = m_TabOrder(nIndex)
   '// Reset m_Tabs
   For j = m_TabOrder(nIndex) To m_TabCount - 2
      m_Tabs(j) = m_Tabs(j + 1)
   Next
   '// Adjust m_TabOrder
   For j = nIndex To m_TabCount - 2
'      m_Tabs(j) = m_Tabs(j + 1)
      m_TabOrder(j) = m_TabOrder(j + 1)
   Next
   '// Validate Indexes for Items after deleted Item
   For j = 0 To m_TabCount - 1
      If m_TabOrder(j) > itemOrder Then
         m_TabOrder(j) = m_TabOrder(j) - 1
      End If
   Next

m_TabCount = m_TabCount - 1
    ReDim Preserve m_Tabs(m_TabCount)
    ReDim Preserve m_TabOrder(m_TabCount)
        If selIndex > m_TabCount - 1 Then
'            SelectedItem = m_TabOrder(m_TabCount - 1)
            SelectedItem = m_TabCount - 1
        Else
            Me.Refresh
        End If
End If
End Sub
Public Property Get TabAlignment(ByVal nIndex As Long) As TabCaptionAlignment
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabAlignment = m_Tabs(nIndex).Alignment
End If
End Property

Public Property Let TabAlignment(ByVal nIndex As Long, ByVal Value As TabCaptionAlignment)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    m_Tabs(nIndex).Alignment = Value
    Draw
End If
End Property

Public Function tabCount() As Long
    tabCount = m_TabCount
End Function
Public Property Get TabKey(ByVal nIndex As Long) As String
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabKey = m_Tabs(nIndex).sKey
End If
End Property
Public Property Get TabTag(ByVal nIndex As Long) As String
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabTag = m_Tabs(nIndex).sTag
End If
End Property
Public Property Let TabTag(ByVal nIndex As Long, ByVal sTag As String)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
     m_Tabs(nIndex).sTag = sTag
End If
End Property
Public Property Get TabCaption() As String
Dim nIndex As Long
    nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    TabCaption = m_Tabs(nIndex).sCaption
End If
End Property
Public Property Let TabCaption(ByVal sCaption As String)
Dim nIndex As Long
     nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    m_Tabs(nIndex).sCaption = sCaption
    PropertyChanged "TabCaption"
    Me.Refresh
End If
End Property
Public Sub SetTabCaption(ByVal nIndex As Long, ByVal sCaption As String)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    m_Tabs(nIndex).sCaption = sCaption
    Draw
End If
End Sub
Public Property Get TabEnabled() As Boolean
Dim nIndex As Long
       nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    TabEnabled = m_Tabs(nIndex).Enabled
End If
End Property
Public Property Let TabEnabled(ByVal Newval As Boolean)
Dim nIndex As Long
       nIndex = selIndex
If nIndex > -1 And nIndex < m_TabCount Then
    m_Tabs(nIndex).Enabled = Newval
    PropertyChanged "TabEnabled"
    Me.Refresh
End If
End Property
Public Sub SetTabEnabled(ByVal nIndex As Long, ByVal Newval As Boolean)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    m_Tabs(m_TabOrder(nIndex)).Enabled = Newval
    Me.Refresh
End If
End Sub

Public Property Get TabImage(ByVal nIndex As Long) As Long
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabImage = m_Tabs(nIndex).Image
End If
End Property
Public Property Let TabImage(ByVal nIndex As Long, ByVal lImage As Long)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
     m_Tabs(nIndex).Image = lImage
    Draw
End If
End Property
Public Property Get TabWidth(ByVal nIndex As Long) As Long
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
    TabWidth = m_Tabs(nIndex).Width
End If
End Property
Public Property Let TabWidth(ByVal nIndex As Long, ByVal lWidth As Long)
If nIndex > -1 And nIndex < m_TabCount And m_TabCount > 0 Then
     m_Tabs(nIndex).Width = lWidth
End If
End Property
Public Sub Refresh()
    Draw
    UserControl.Refresh
End Sub
Private Function hitTest(ByVal lX As Long, ByVal lY As Long) As Long
Dim i As Long, c As Long, rc As RECT, lTab As Long
        lTab = -1
    For i = 0 To m_TabCount - 1
            c = m_TabOrder(i)
        With m_Tabs(c)
            If .Enabled = True And .Visible = True Then
                    rc.Left = .Left: rc.Top = .Top: rc.Right = .Right: rc.Bottom = .Bottom
                If lX >= rc.Left And lX <= rc.Right - 10 And lY >= rc.Top And lY <= rc.Bottom Then
                    lTab = i ' C
                    Exit For
                End If
            End If
        End With
    Next
        hitTest = lTab
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift <> -1 Then
        cx = ScaleX(x, UserControl.ScaleMode, 3)
        cy = ScaleY(y, UserControl.ScaleMode, 3)
    Else
        cx = x
        cy = y
    End If
If tmrSlide.Enabled = True Then
    Exit Sub
End If
Dim bDraw As Boolean
Dim rc As RECT, p As POINTAPI
Dim L As Long, lX As Long, lY As Long, W As Long, H As Long
Dim lTab As Long
        If Shift <> -1 Then
            lX = ScaleX(x, UserControl.ScaleMode, 3)
            lY = ScaleY(y, UserControl.ScaleMode, 3)
        Else
            lX = x
            lY = y
        End If
    If bSliderShown = True And Button = 1 Then
        Dim lss As ssSliderStatus, rss As ssSliderStatus
            W = (SliderBox.Right - SliderBox.Left) / 2
                lss = ssDisabled
                rss = ssDisabled
        If lX >= SliderBox.Left And lX <= SliderBox.Right And lY >= SliderBox.Top And lY <= SliderBox.Bottom Then
            If lX >= (SliderBox.Right - W) Then
                rss = ssDown
            Else
                lss = ssDown
            End If
        End If
            If lScrollX < lScrollWidth And rss <> ssDisabled And RightSliderStatus <> rss Then 'And RightSliderStatus <> ssDisabled And RightSliderStatus <> ssDown Then
                RightSliderStatus = rss
                bDraw = True
            End If
            If lScrollX > 0 And lss <> ssDisabled And LeftSliderStatus <> lss Then 'And LeftSliderStatus <> ssDisabled And LeftSliderStatus <> ssDown Then
                LeftSliderStatus = lss
                bDraw = True
            End If
    End If
If bDraw Then
    tmrSlide.Interval = 50
    tmrSlide.Enabled = True
    Me.Refresh
    Exit Sub
End If
        lTab = hitTest(lX, lY)
    
        lTabIndex = lTab
If lTab <> -1 Then
        lTab = m_TabOrder(lTab)
    If selIndex <> lTab Then
        RaiseEvent BeforeTabChange(selIndex, lTab)
            RaiseEvent TabClick(lTab)
        SelectedItem = lTab
    End If
End If
        lTabDragging = False

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim bDraw As Boolean
Dim rc As RECT, p As POINTAPI
Dim c As Long, L As Long, lX As Long, lY As Long, W As Long, H As Long
Dim lTab As Long
'''            GetCursorPos P
'''            ScreenToClient UserControl.hwnd, P
'''                lX = P.X
'''                lY = P.Y
        If Shift <> -1 Then
            lX = ScaleX(x, UserControl.ScaleMode, 3)
            lY = ScaleY(y, UserControl.ScaleMode, 3)
        Else
            lX = x
            lY = y
        End If
        If m_AllowReorder = True And lTabDragging = False And Button = 1 And lTabIndex <> -1 And Abs(cx - lX) > 4 Then
            lTabDragged = lTabIndex
            lTabDragging = True
'            SetCapture UserControl.hWnd
        End If
    If tmrSlide.Enabled = True And lTabDragging = False Then
        Exit Sub
    End If
        If Button = 0 And lTabDragging = False Then
           TimerCheckMouseOut.Enabled = True
        End If
    lTab = -1
If lTabDragging = True Then
    lTab = hitTest(lX, lY)
End If
    If bSliderShown = True And lTabDragging = False Then
        Dim lss As ssSliderStatus, rss As ssSliderStatus
            W = (SliderBox.Right - SliderBox.Left) / 2
                rss = ssDisabled
                lss = ssDisabled
        If lX >= SliderBox.Left And lX <= SliderBox.Right And lY >= SliderBox.Top And lY <= SliderBox.Bottom Then
            If lX >= (SliderBox.Right - W) Then
                rss = ssHover
            Else
                lss = ssHover
            End If
        End If
            If RightSliderStatus = ssHover And rss = ssDisabled Then
                RightSliderStatus = ssNormal
                bDraw = True
            End If
            If LeftSliderStatus = ssHover And lss = ssDisabled Then
                LeftSliderStatus = ssNormal
                bDraw = True
            End If
            If rss <> ssDisabled And RightSliderStatus = ssNormal Then
                RightSliderStatus = rss
                bDraw = True
            End If
            If lss <> ssDisabled And LeftSliderStatus = ssNormal Then
                LeftSliderStatus = lss
                bDraw = True
            End If
    End If
If bDraw Then
    Me.Refresh
    Exit Sub
End If

        If lTabDragging = False And Screen.MousePointer <> 0 Then
            Screen.MousePointer = 0
        End If
            lTabHover = lTab
    If lTabDragging = True Then
        Screen.MousePointer = 5
        If lTab <> -1 Then
            FocusTab m_TabOrder(lTab)
        End If
        If lTabHover <> lTabIndex Then
            lTabIndex = lTab
'            Draw
        End If
            Me.Refresh
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
            If lTabDragging = True Then
'                ReleaseCapture
                Screen.MousePointer = 0
                lTabHover = -1
                lTabDragging = False
                    If lTabIndex <> -1 And lTabIndex <> lTabDragged Then
''                            lTabDragged = lTabDragged + 1
''                            lTabIndex = lTabIndex + 1
                                lTabIndex = m_TabOrder(lTabIndex)
                        RaiseEvent TabOrderChanged(m_TabOrder(lTabDragged), lTabIndex)
''                                lTabDragged = lTabDragged - 1
''                                lTabIndex = lTabIndex - 1
                                lTabIndex = getTabOrder(lTabIndex)
                            MoveTab lTabDragged, lTabIndex
'                        SelectedItem = lTabIndex + 1
                        SelectedItem = m_TabOrder(lTabIndex)
                    Else
                        Me.Refresh
                    End If
                    lTabDragged = -1
                    lTabIndex = -1
                Exit Sub
            End If
If RightSliderStatus = ssDown Or LeftSliderStatus = ssDown Then
        tmrSlide.Enabled = False
    RightSliderStatus = IIf(lScrollX < lScrollWidth, ssNormal, ssDisabled)
    LeftSliderStatus = IIf(lScrollX > 0, ssNormal, ssDisabled)
        Me.Refresh
End If
End Sub
Private Sub MoveTab(ByVal nTab As Long, ByVal toTab As Long)
Dim c As Long, j As Long, tempIndex As Long
Dim nInfo As tTabInfo, nIndex As Long, toIndex As Long
'****** Sorting Controls **************
Dim NoMoreSwaps As Boolean, NumberOfItems As Long, Temp As tControlInfo, bDirection As Boolean
    bDirection = True
        NumberOfItems = UBound(m_Ctls)
    Do Until NoMoreSwaps = True
            NoMoreSwaps = True
         For c = 0 To (NumberOfItems - 1)
            If bDirection = True Then 'Ascending
                 If m_Ctls(c).iTabIndex > m_Ctls(c + 1).iTabIndex Then
                     NoMoreSwaps = False
                     Temp = m_Ctls(c)
                     m_Ctls(c) = m_Ctls(c + 1)
                     m_Ctls(c + 1) = Temp
                 End If
            Else
                 If m_Ctls(c).iTabIndex < m_Ctls(c + 1).iTabIndex Then
                     NoMoreSwaps = False
                     Temp = m_Ctls(c)
                     m_Ctls(c) = m_Ctls(c + 1)
                     m_Ctls(c + 1) = Temp
                 End If
            End If
         Next
            NumberOfItems = NumberOfItems - 1
    Loop
'********* End Sorting *************************************
'        LSet nInfo = m_Tabs(nTab)
            nIndex = m_TabOrder(nTab)
            toIndex = m_TabOrder(toTab)
If toTab > nTab Then
   For c = nTab To toTab - 1
'      m_Tabs(C) = m_Tabs(C + 1)
      m_TabOrder(c) = m_TabOrder(c + 1)
   Next
        For j = 0 To m_ctlCount - 1
            If m_Ctls(j).iTabIndex > nTab And m_Ctls(j).iTabIndex <= toTab Then
                m_Ctls(j).iTabIndex = m_Ctls(j).iTabIndex - 1
            ElseIf m_Ctls(j).iTabIndex = nTab Then
                m_Ctls(j).iTabIndex = toTab
            End If
        Next

Else
   For c = nTab To toTab + 1 Step -1
'      m_Tabs(C) = m_Tabs(C - 1)
      m_TabOrder(c) = m_TabOrder(c - 1)
   Next
        For j = 0 To m_ctlCount - 1
            If m_Ctls(j).iTabIndex >= toTab And m_Ctls(j).iTabIndex < nTab Then
                m_Ctls(j).iTabIndex = m_Ctls(j).iTabIndex + 1
            ElseIf m_Ctls(j).iTabIndex = nTab Then
                m_Ctls(j).iTabIndex = toTab
            End If
        Next
End If
'    m_Tabs(toTab) = nInfo
    m_TabOrder(toTab) = nIndex

End Sub
Public Function IsDir(ByVal sDir As String) As Boolean
    IsDir = ((GetFileAttributes(sDir) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
End Function
Public Function IsFile(ByVal sFile As String) As Boolean
    IsFile = GetFileAttributes(sFile) <> -1
End Function
Sub SaveTabOrder(ByVal sFile As String)
'Dim m_PropBag As New PropertyBag
'Dim b() As Byte
'    UserControl_WriteProperties m_PropBag
'        b = m_PropBag.Contents

    If IsFile(sFile) = True Then
        DeleteFile sFile
    End If
Dim m_tabIndex() As Long, j As Long
Dim m_tabLeft() As Single, m_tabVisible() As Boolean, ctl As Control

        ReDim m_tabIndex(m_ctlCount)
        ReDim m_tabLeft(m_ctlCount)
        ReDim m_tabVisible(m_ctlCount)
    If m_ctlCount > 0 Then
        ReDim m_tabIndex(m_ctlCount - 1)
        ReDim m_tabLeft(m_ctlCount - 1)
        ReDim m_tabVisible(m_ctlCount - 1)
    End If
            For j = 0 To m_ctlCount - 1
                m_tabIndex(j) = m_Ctls(j).iTabIndex
                    For Each ctl In UserControl.ContainedControls
                        If pGetControlId(ctl) = m_Ctls(j).ctlName Then
                             m_tabLeft(j) = ctl.Left
                             m_tabVisible(j) = ctl.Visible
                             Exit For
                        End If
                    Next
            Next
Open sFile For Binary As 1
'    Put #1, , UBound(b)+1
'    Put #1, , b

'    Put #1, , selIndex
'    Put #1, , m_ctlCount
'    Put #1, , m_TabCount
'    Put #1, , m_TabOrder
'    Put #1, , m_tabIndex
'    Put #1, , m_tabLeft
'    Put #1, , m_tabVisible
Close 1
    
End Sub
Sub LoadTabOrder(ByVal sFile As String)
'Dim m_PropBag As New PropertyBag
'Dim b() As Byte, nLen As Long
Dim ctlCount As Long, nIndex As Long, tempOrder() As Long, m_tabIndex() As Long
Dim j As Long, tabCount As Long, k As Long
Dim m_tabLeft() As Single, m_tabVisible() As Boolean, ctl As Control
    If IsFile(sFile) = False Then
        Exit Sub
    End If
Open sFile For Binary As 1
'    Get #1, , nLen
'    If nLen > 0 Then
'        ReDim b(nLen - 1)
'        Get #1, , b
'    End If
    
    Get #1, , nIndex
    Get #1, , ctlCount
    If ctlCount > 0 Then
        Get #1, , tabCount
            ReDim tempOrder(tabCount)
        If tabCount > 0 Then
            ReDim tempOrder(tabCount - 1)
        End If
            ReDim m_tabIndex(ctlCount - 1)
            ReDim m_tabLeft(ctlCount - 1)
            ReDim m_tabVisible(ctlCount - 1)
            Get #1, , tempOrder
            Get #1, , m_tabIndex
            Get #1, , m_tabLeft
            Get #1, , m_tabVisible
    End If
Close 1
'If nLen > 0 Then
'    m_PropBag.Contents = b
'    UserControl_ReadProperties m_PropBag
'End If
    If ctlCount > 0 Then
        If tabCount > 0 Then
            For j = 0 To IIf(m_TabCount > tabCount, tabCount, m_TabCount) - 1
                 m_TabOrder(j) = tempOrder(j)
            Next
        End If
            For j = 0 To IIf(m_ctlCount > ctlCount, ctlCount, m_ctlCount) - 1
                m_Ctls(j).iTabIndex = m_tabIndex(j)
                    For Each ctl In UserControl.ContainedControls
                        If pGetControlId(ctl) = m_Ctls(j).ctlName Then
                             ctl.Left = m_tabLeft(j)
                             ctl.Visible = m_tabVisible(j)
                             Exit For
                        End If
                    Next
            Next
            SelectedItem = nIndex
    End If
End Sub
Public Property Get Font() As StdFont
    Set Font = mFont
End Property
Public Property Set Font(ByVal nV As StdFont)
    Set mFont = nV
    PropertyChanged "Font"
    Me.Refresh
End Property
Public Property Get AllowReorder() As Boolean
    AllowReorder = m_AllowReorder
End Property
Public Property Let AllowReorder(ByVal nV As Boolean)
        m_AllowReorder = nV
    PropertyChanged "AllowReorder"
End Property
Public Property Get ItemCount() As Long
    ItemCount = m_TabCount
End Property
Public Property Let ItemCount(ByVal nV As Long)
Dim i As Long
If nV > m_TabCount Then
    For i = m_TabCount + 1 To nV
        AddTab , "Tab " & (i)
    Next
        Me.Refresh
Else
    If nV > 0 Then
        Dim T As Long
            T = m_TabCount - 1
        For i = T To nV Step -1
            RemoveTab i
        Next
    End If
End If
    PropertyChanged "ItemCount"
End Property

Public Property Get SelectedItem() As Long
    SelectedItem = selIndex
End Property
Public Property Let SelectedItem(ByVal mTabIndex As Long)
    If mTabIndex < 0 Or mTabIndex > m_TabCount Then
'        Err.Raise 380 ' invalid property value
        MsgBox "invalid property value", vbCritical
        Exit Property
    End If
        If mTabIndex > m_TabCount - 1 Then
            mTabIndex = m_TabCount - 1
        End If

    handleControls selIndex, mTabIndex
        selIndex = mTabIndex
    PropertyChanged "SelectedItem"
    Draw
        FocusTab selIndex
    Me.Refresh
End Property
Public Property Get MoveItem() As Long
    MoveItem = getTabOrder(selIndex)
End Property
Public Property Let MoveItem(ByVal mTabIndex As Long)
    If mTabIndex < 0 Or mTabIndex > m_TabCount - 1 Then
'        Err.Raise 380 ' invalid property value
        MsgBox "invalid property value", vbCritical
        Exit Property
    End If

'        MoveTab selIndex, mTabIndex
        MoveTab getTabOrder(selIndex), mTabIndex
    SelectedItem = m_TabOrder(mTabIndex)
End Property
Private Sub FocusTab(ByVal lTab As Long)
Dim lWidth As Long, lHeight As Long
If lTab < m_TabCount And m_TabCount > 0 Then
    If bSliderShown = True Then
        lWidth = ScaleX(ScaleWidth, ScaleMode, 3)
        lHeight = ScaleY(ScaleHeight, ScaleMode, 3)
''            lTab = m_TabOrder(lTab)
        If m_Tabs(lTab).Right + IIf(bSliderShown, (SliderBox.Right - SliderBox.Left) + 16, 0) > (lWidth) Then
            lScrollX = lScrollX + m_Tabs(lTab).Right + (SliderBox.Right - SliderBox.Left) + 16 - lWidth
                LeftSliderStatus = ssNormal
                RightSliderStatus = ssNormal
            If lScrollX > lScrollWidth Then
                lScrollX = lScrollWidth
                RightSliderStatus = ssDisabled
            End If
        ElseIf m_Tabs(lTab).Left < (-4) Then
            lScrollX = 4 + m_Tabs(lTab).Left
                LeftSliderStatus = ssNormal
                RightSliderStatus = ssNormal
            If lScrollX < 0 Then
                lScrollX = 0
                LeftSliderStatus = ssDisabled
            End If
        End If
        Dim lTabOrder As Long
                lTabOrder = getTabOrder(lTab)
            If lTabOrder = 0 Then
                LeftSliderStatus = ssDisabled
            ElseIf lTabOrder = m_TabCount - 1 Then
                RightSliderStatus = ssDisabled
            End If
    End If
End If
End Sub
'handle contained controls
Private Sub handleControls_old(ByVal LastIndex As Long, ByVal nIndex As Long)
'Dim I As Long, z As Integer, mCTL As Control
'Dim lstRemove As New Collection
''    On Error Resume Next
'    'hide controls
'    For Each mCTL In UserControl.ContainedControls
'        If mCTL.left > -35000 Then
'            If lastIndex <= UBound(ctlLst) Then
'                If isInCollection(ctlLst(lastIndex), pGetControlId(mCTL)) <> True Then
'                    ctlLst(lastIndex).Add pGetControlId(mCTL)
'                End If
'            End If
'        End If
'    Next
'    'find controls that we need to show
'    For Each mCTL In UserControl.ContainedControls
'       '
'        If nIndex <= UBound(ctlLst) Then
'
'            If isInCollection(ctlLst(nIndex), pGetControlId(mCTL)) Then ' Or isInCollection(ctlLst, pGetControlId(mCTL) & "-1") Then
''                mCTL.Visible = True
'                If mCTL.left < -35000 Then
'                    mCTL.left = mCTL.left + 70000
'                End If
'            Else
'                If mCTL.left > -35000 Then
''                        mCTL.Visible = False
'                        mCTL.left = mCTL.left - 70000
'                End If
'            End If
'        End If
'        Err.Clear
'    Next
'
'If nIndex <= UBound(ctlLst) Then
'    Set ctlLst(nIndex) = Nothing
'End If
'    For Each mCTL In UserControl.ContainedControls
'        If mCTL.left > -35000 Then
'            If nIndex <= UBound(ctlLst) Then
'                ctlLst(nIndex).Add pGetControlId(mCTL)
'            End If
'        End If
'    Next
End Sub
Private Function isInCollection(ByRef lstCollection As Collection, ByVal vData As Variant) As Boolean
    isInCollection = False
    If lstCollection.count = 0 Then Exit Function
    Dim i As Long
    For i = 1 To lstCollection.count
        If lstCollection.Item(i) = vData Then
            isInCollection = True
            Exit Function
        End If
    Next i
End Function
Public Function HasIndex(ByVal ctl As Control) As Boolean
    'determine if it's a control array
    HasIndex = Not ctl.Parent.Controls(ctl.Name) Is ctl
End Function
Private Function pGetControlId(ByVal oCtl As Control) As String
Dim sCtlName As String
Dim iCtlIndex As Integer
        iCtlIndex = -1
    sCtlName = oCtl.Name
'If VarType(oCtl) = vbObject Then
On Error Resume Next
'    iCtlIndex = oCtl.Index
If HasIndex(oCtl) Then
    iCtlIndex = oCtl.Index
End If
'End If
    pGetControlId = sCtlName & IIf(iCtlIndex <> -1, "(" & iCtlIndex & ")", "")
End Function
Private Function IsInTabControls(ByVal ctlName As String, ByVal lIndex As Long) As Long
Dim j As Long
        IsInTabControls = -1
    For j = 0 To m_ctlCount - 1
        If Trim$(m_Ctls(j).ctlName) = ctlName And m_Ctls(j).iTabIndex = lIndex Then
            IsInTabControls = j
            Exit For
        End If
    Next
End Function
Private Sub AddTabControls(ByVal lIndex As Long, ByVal ctlName As String)
    ReDim Preserve m_Ctls(m_ctlCount)
        m_Ctls(m_ctlCount).ctlName = ctlName
        m_Ctls(m_ctlCount).iTabIndex = lIndex
    m_ctlCount = m_ctlCount + 1
End Sub
 
Private Sub handleControls(ByVal LastIndex As Long, ByVal nIndex As Long)
Dim i As Long, z As Long, mCTL As Control, ctlName As String
If m_TabCount > 0 Then
''''    lastIndex = m_TabOrder(lastIndex - 1)
''''    nIndex = m_TabOrder(nIndex - 1)
'    LastIndex = LastIndex - 1
'    nIndex = nIndex - 1
    LastIndex = getTabOrder(LastIndex)
    nIndex = getTabOrder(nIndex)
End If
            
    For Each mCTL In UserControl.ContainedControls
                ctlName = pGetControlId(mCTL)
        If IsInTabControls(ctlName, nIndex) <> -1 Then
            If mCTL.Left < -35000 Then
                mCTL.Visible = True
                mCTL.Left = mCTL.Left + 70000
'                    If TypeOf mCTL Is Label Then
'                        mCTL.Refresh
'                    ElseIf TypeOf mCTL Is Image Then
'                        mCTL.Refresh
'                    ElseIf TypeOf mCTL Is Shape Then
'                        mCTL.Refresh
'                    End If
            End If
        Else
            If mCTL.Left > -35000 Then
                If IsInTabControls(ctlName, LastIndex) = -1 Then
                    AddTabControls LastIndex, ctlName
                End If
                    mCTL.Left = mCTL.Left - 70000
                    mCTL.Visible = False
            End If
        End If
    Next
Dim j As Long
    For j = 0 To m_ctlCount - 1
        If m_Ctls(j).iTabIndex > (m_TabCount - 1) Then
            For Each mCTL In UserControl.ContainedControls
                    ctlName = pGetControlId(mCTL)
                If Trim$(m_Ctls(j).ctlName) = ctlName Then
'                If IsInTabControls(ctlName, nIndex) = -1 Then
                    If mCTL.Left > -35000 Then
                        mCTL.Left = mCTL.Left - 70000
                        mCTL.Visible = False
                    End If
'                End If
                        Exit For
                End If
            Next
        End If
    Next
End Sub


