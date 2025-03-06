Attribute VB_Name = "Module1"
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

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1


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

Sub TrueSizeViewing()

    Dim xHSizeSq As Double, xVSizeSq As Double, xPix As Double, xDot As Double
    Dim hWnd As LongPtr, hDC As LongPtr, hMonitor As LongPtr
    Dim tMonitorInfo As MONITORINFOEX
    Dim nMonitors As Integer
    Dim vResult As Variant
    Dim sItem As String
    Application.Volatile
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
        
    Dim oScreenW As Double
    oScreenW = GetDeviceCaps(hDC, 4) / 10
    
    Dim oScreenH As Double
    oScreenH = GetDeviceCaps(hDC, 6) / 10

    Dim oPixelW As Long
    oPixelW = GetDeviceCaps(hDC, 8)
    
    Dim oPixelH As Long
    oPixelH = GetDeviceCaps(hDC, 10)
 
    
    Dim oViewW As Long
    oViewW = ThisApplication.ActiveView.Width
    
    Dim oViewH As Long
    oViewH = ThisApplication.ActiveView.Height
    
    Dim cam As Camera
    Set cam = ThisApplication.ActiveView.Camera
    
    Call cam.SetExtents(oViewW / oPixelW * oScreenW, oViewH / oPixelH * oScreenH)
    cam.Apply
    
End Sub

