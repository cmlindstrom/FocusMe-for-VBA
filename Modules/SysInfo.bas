Attribute VB_Name = "SysInfo"
' This module includes Private declarations for certain Windows API functions
' plus code for Public Function Screen, which returns metrics for the screen displaying ActiveWindow
' This module requires VBA7 (Office 2010 or later)
' DEVELOPER: J. Woolley (for wellsr.com)
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" _
    (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" _
    (ByVal hMonitor As LongPtr, ByRef lpMI As MONITORINFOEX) As Boolean
Private Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" _
    (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Const SM_CMONITORS              As Long = 80    ' number of display monitors
Private Const MONITOR_CCHDEVICENAME     As Long = 32    ' device name fixed length
Private Const MONITOR_PRIMARY           As Long = 1
Private Const MONITOR_DEFAULTTONULL     As Long = 0
Private Const MONITOR_DEFAULTTOPRIMARY  As Long = 1
Private Const MONITOR_DEFAULTTONEAREST  As Long = 2
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MONITORINFOEX
   cbSize As Long
   rcMonitor As RECT
   rcWork As RECT
   dwFlags As Long
   szDevice As String * MONITOR_CCHDEVICENAME
End Type
Private Enum DevCap     ' GetDeviceCaps nIndex (video displays)
    HORZSIZE = 4        ' width in millimeters
    VERTSIZE = 6        ' height in millimeters
    HORZRES = 8         ' width in pixels
    VERTRES = 10        ' height in pixels
    BITSPIXEL = 12      ' color bits per pixel
    LOGPIXELSX = 88     ' horizontal DPI (assumed by Windows)
    LOGPIXELSY = 90     ' vertical DPI (assumed by Windows)
    COLORRES = 108      ' actual color resolution (bits per pixel)
    VREFRESH = 116      ' vertical refresh rate (Hz)
End Enum

''' Returns the size of the screen and number of monitors
Public Sub GetScreenResolution(ByRef hgt As Integer, _
                                ByRef wid As Integer, _
                        Optional ByRef n As Integer = 1)

    Dim strTrace As String

    wid = Screen("horizontalresolution")
    hgt = Screen("verticalresolution")
    n = Screen("numMonitors")

    strTrace = "Screen is " & wid & " pixels wide and " & hgt & " pixels tall."

End Sub

Private Function Screen(Item As String) As Variant

' Return display screen Item for monitor displaying ActiveWindow
' Patterned after Excel's built-in information functions CELL and INFO
' Supported Item values (each must be a string, but alphabetic case is ignored):

' NumMonitors       ' # of monitors found
' HorizontalResolution or pixelsX
' VerticalResolution or pixelsY
' WidthInches or inchesX
' HeightInches or inchesY
' DiagonalInches or inchesDiag
' PixelsPerInchX or ppiX
' PixelsPerInchY or ppiY
' PixelsPerInch or ppiDiag
' WinDotsPerInchX or dpiX
' WinDotsPerInchY or dpiY
' WinDotsPerInch or dpiWin ' DPI assumed by Windows
' AdjustmentFactor or zoomFac ' adjustment to match actual size (ppiDiag/dpiWin)
' IsPrimary ' True if primary display
' DisplayName ' name recognized by CreateDC
' Update ' update cells referencing this UDF and return date/time
' Help ' display all recognized Item string values
' EXAMPLE: =Screen("pixelsX")
' Function Returns #VALUE! for invalid Item

    Dim xHSizeSq As Double, xVSizeSq As Double, xPix As Double, xDot As Double
    Dim hWnd As LongPtr, hDC As LongPtr, hMonitor As LongPtr
    Dim tMonitorInfo As MONITORINFOEX
    Dim nMonitors As Integer
    Dim vResult As Variant
    Dim sItem As String
    ' Application.Volatile
    nMonitors = GetSystemMetrics(SM_CMONITORS)
    If nMonitors < 2 Then
        nMonitors = 1                                       ' in case GetSystemMetrics failed
        hWnd = 0
    Else
        hWnd = GetActiveWindow()
        hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONULL)
        If hMonitor = 0 Then
            Debug.Print "ActiveWindow does not intersect a monitor"
            hWnd = 0
        Else
            tMonitorInfo.cbSize = Len(tMonitorInfo)
            If GetMonitorInfo(hMonitor, tMonitorInfo) = False Then
                Debug.Print "GetMonitorInfo failed"
                hWnd = 0
            Else
                hDC = CreateDC(tMonitorInfo.szDevice, 0, 0, 0)
                If hDC = 0 Then
                    Debug.Print "CreateDC failed"
                    hWnd = 0
                End If
            End If
        End If
    End If
    If hWnd = 0 Then
        hDC = GetDC(hWnd)
        tMonitorInfo.dwFlags = MONITOR_PRIMARY
        tMonitorInfo.szDevice = "PRIMARY" & vbNullChar
    End If
    sItem = Trim(LCase(Item))
    Select Case sItem
    Case "nummonitors"
        vResult = nMonitors
    Case "horizontalresolution", "pixelsx"                  ' HorizontalResolution (pixelsX)
        vResult = GetDeviceCaps(hDC, DevCap.HORZRES)
    Case "verticalresolution", "pixelsy"                    ' VerticalResolution (pixelsY)
        vResult = GetDeviceCaps(hDC, DevCap.VERTRES)
    Case "widthinches", "inchesx"                           ' WidthInches (inchesX)
        vResult = GetDeviceCaps(hDC, DevCap.HORZSIZE) / 25.4
    Case "heightinches", "inchesy"                          ' HeightInches (inchesY)
        vResult = GetDeviceCaps(hDC, DevCap.VERTSIZE) / 25.4
    Case "diagonalinches", "inchesdiag"                     ' DiagonalInches (inchesDiag)
        vResult = Sqr(GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2 + GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2) / 25.4
    Case "pixelsperinchx", "ppix"                           ' PixelsPerInchX (ppiX)
        vResult = 25.4 * GetDeviceCaps(hDC, DevCap.HORZRES) / GetDeviceCaps(hDC, DevCap.HORZSIZE)
    Case "pixelsperinchy", "ppiy"                           ' PixelsPerInchY (ppiY)
        vResult = 25.4 * GetDeviceCaps(hDC, DevCap.VERTRES) / GetDeviceCaps(hDC, DevCap.VERTSIZE)
    Case "pixelsperinch", "ppidiag"                         ' PixelsPerInch (ppiDiag)
        xHSizeSq = GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2
        xVSizeSq = GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2
        xPix = GetDeviceCaps(hDC, DevCap.HORZRES) ^ 2 + GetDeviceCaps(hDC, DevCap.VERTRES) ^ 2
        vResult = 25.4 * Sqr(xPix / (xHSizeSq + xVSizeSq))
    Case "windotsperinchx", "dpix"                          ' WinDotsPerInchX (dpiX)
        vResult = GetDeviceCaps(hDC, DevCap.LOGPIXELSX)
    Case "windotsperinchy", "dpiy"                          ' WinDotsPerInchY (dpiY)
        vResult = GetDeviceCaps(hDC, DevCap.LOGPIXELSY)
    Case "windotsperinch", "dpiwin"                         ' WinDotsPerInch (dpiWin)
        xHSizeSq = GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2
        xVSizeSq = GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2
        xDot = GetDeviceCaps(hDC, DevCap.LOGPIXELSX) ^ 2 * xHSizeSq + GetDeviceCaps(hDC, DevCap.LOGPIXELSY) ^ 2 * xVSizeSq
        vResult = Sqr(xDot / (xHSizeSq + xVSizeSq))
    Case "adjustmentfactor", "zoomfac"                      ' AdjustmentFactor (zoomFac)
        xHSizeSq = GetDeviceCaps(hDC, DevCap.HORZSIZE) ^ 2
        xVSizeSq = GetDeviceCaps(hDC, DevCap.VERTSIZE) ^ 2
        xPix = GetDeviceCaps(hDC, DevCap.HORZRES) ^ 2 + GetDeviceCaps(hDC, DevCap.VERTRES) ^ 2
        xDot = GetDeviceCaps(hDC, DevCap.LOGPIXELSX) ^ 2 * xHSizeSq + GetDeviceCaps(hDC, DevCap.LOGPIXELSY) ^ 2 * xVSizeSq
        vResult = 25.4 * Sqr(xPix / xDot)
    Case "isprimary"                                        ' IsPrimary
        vResult = CBool(tMonitorInfo.dwFlags And MONITOR_PRIMARY)
    Case "displayname"                                      ' DisplayName
        vResult = tMonitorInfo.szDevice & vbNullChar
        vResult = Left(vResult, (InStr(1, vResult, vbNullChar) - 1))
    Case "update"                                           ' Update
        vResult = Now
    Case "help"                                             ' Help
        vResult = "HorizontalResolution (pixelsX), VerticalResolution (pixelsY), " _
            & "WidthInches (inchesX), HeightInches (inchesY), DiagonalInches (inchesDiag), " _
            & "PixelsPerInchX (ppiX), PixelsPerInchY (ppiY), PixelsPerInch (ppiDiag), " _
            & "WinDotsPerInchX (dpiX), WinDotsPerInchY (dpiY), WinDotsPerInch (dpiWin), " _
            & "AdjustmentFactor (zoomFac), IsPrimary, DisplayName, Update, Help"
    Case Else                                               ' Else
        vResult = CVErr(xlErrValue)                         ' return #VALUE! error (2015)
    End Select
    If hWnd = 0 Then
        ReleaseDC hWnd, hDC
    Else
        DeleteDC hDC
    End If
    Screen = vResult
End Function

