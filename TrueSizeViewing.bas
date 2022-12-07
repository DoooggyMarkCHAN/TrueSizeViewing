Attribute VB_Name = "RealSizeDisplay"
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Sub TrueSizeViewing()
    
    Dim hDc As LongPtr
    hDc = GetDC(0)
    
    Dim oPixelW As Long
    oPixelW = GetSystemMetrics(SM_CXSCREEN)
    
    Dim oPixelH As Long
    oPixelH = GetSystemMetrics(SM_CYSCREEN)
 
    Dim oScreenW As Double
    oScreenW = GetDeviceCaps(hDc, 4) / 10
    
    Dim oScreenH As Double
    oScreenH = GetDeviceCaps(hDc, 6) / 10
    
    Dim oViewW As Long
    oViewW = ThisApplication.ActiveView.Width
    
    Dim oViewH As Long
    oViewH = ThisApplication.ActiveView.Height
    
    Dim cam As Camera
    Set cam = ThisApplication.ActiveView.Camera
    
    Call cam.SetExtents(oViewW / oPixelW * oScreenW, oViewH / oPixelH * oScreenH)
    cam.Apply
    
End Sub
