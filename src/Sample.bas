Attribute VB_Name = "Win32API"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, _
                    ByVal nHeight As Long) As Long
    Public Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Public Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Public Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Public Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, _
                    ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdc As Long, _
                    ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
    Public Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PicBmp, RefIID As GUID, _
                    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function GetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBmp As Long, ByVal uStartScan As Long, _
                    ByVal cScanLines As Long, lpvBits As Any, lpbi As BITMAPINFO, ByVal uUsage As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare PtrSafe Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, _
                    ByRef lpRect As RECT) As Long
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" _
                    (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
#Else
    Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
                    ByVal nHeight As Long, ByVal hdc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
    Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PicBmp, RefIID As GUID, _
                    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Declare Function GetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBmp As Long, ByVal uStartScan As Long, _
                    ByVal cScanLines As Long, lpvBits As Any, lpbi As BITMAPINFO, ByVal uUsage As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, _
                    ByRef lpRect As RECT) As Long
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function GetForegroundWindow Lib "user32" () As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function GetAsyncKeyState Lib "user32" _
            (ByVal vKey As Long) As Integer
    Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

' https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_0 = &H30

Public Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Const SRCCOPY = &HCC0020
Public Const DIB_RGB_COLORS = 0
Public Const BITSPIXEL = 12
Public Const BI_RGB = 0

Public Const S_OK As Long = &H0
Public Const E_NOINTERFACE As Long = &H80004002
Public Const E_POINTER As Long = &H80004003

Public Const PICTYPE_BITMAP = 1

Public Const SW_HIDE = 0                   ' Hide the window.
Public Const SW_SHOWNORMAL = 1             ' Show the window and activate it (as usual).
Public Const SW_SHOWMINIMIZED = 2          ' Show the window minimized.
Public Const SW_SHOWMAXIMIZED = 3          ' Maximize the window.
Public Const SW_SHOWNOACTIVATE = 4         ' Show the window in its most recent size and position but do not activate it.
Public Const SW_SHOW = 5                   ' Show the window.
Public Const SW_MINIMIZE = 6               ' Minimize the window.
Public Const SW_SHOWMINNOACTIVE = 7        ' Show the window minimized but do not activate it.
Public Const SW_SHOWNA = 8                 ' Show the window in its current state but do not activate it.
Public Const SW_RESTORE = 9                ' Restore the window (not maximized nor minimized).

Type BITMAPFILEHEADER
       bfType       As String * 2
       bfSize       As Long
       bfReserved1  As Integer
       bfReserved2  As Integer
       bfOffBits    As Long
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Enum wiaFormat
    BMP = 0
    GIF = 1
    JPEG = 2
    png = 3
    TIFF = 4
End Enum

