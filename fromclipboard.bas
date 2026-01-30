Attribute VB_Name = "fromclipboard"
'https://makingyenything.wordpress.com/2025/12/18/how-to-convert-an-image-of-a-chart-to-excel/
Option Explicit
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function apiGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hgdiobj As LongPtr, ByVal cbBuffer As Long, lpvObject As Any) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObj As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal xDest As Long, ByVal yDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal rop As Long) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal hDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr

Private Const srccopy = &HCC0020
Private Const CF_BITMAP = 2
Private Const DIB_RGB_COLORS As Long = 0
Private Const BI_RGB As Long = 0

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr ' Must be LongPtr for 64-bit compatibility
End Type

Private Type BITMAPINFOHEADER    '40 bytes
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

Private Type RGBQUAD
 Red As Byte
 Blue As Byte
 Green As Byte
 alpha As Byte
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
 bmiColors As RGBQUAD
End Type
'just to calculate distance between colors
Private Type rgb
    r As Long
    g As Long
    b As Long
End Type
Function getrgb(colorvalue As Long) As rgb
    Dim output As rgb
    output.r = colorvalue Mod 256
    output.g = (colorvalue \ 256) Mod 256
    output.b = (colorvalue \ 65536) Mod 256
    getrgb = output
End Function

Function distancebetweencolors(c1 As Long, c2 As Long) As Long
    'https://www.compuphase.com/cmetric.htm
    Dim colora As rgb, colorb As rgb
    Dim r As Long, g As Long, b As Long
    Dim rmean As Long
    colora = getrgb(c1)
    colorb = getrgb(c2)
    rmean = (colora.r + colorb.r) / 2
    r = colora.r - colorb.r
    g = colora.g - colorb.g
    b = colora.b - colorb.b
    distancebetweencolors = Sqr((2 + (rmean / 256)) * r * r + 4 * g * g + (2 + (255 - rmean) / 256) * b * b)
End Function


Function getPixelsfromBMP(hMemDC As LongPtr, bmp As BITMAP, dh As Long, dw As Long) As Variant
    'creates a copy of the bitmap that's the dimensions we want and then returns an array with the rgb values of the resulting pixels
    Dim memDC     As LongPtr: memDC = CreateCompatibleDC(hMemDC)
    Dim memBMP    As LongPtr: memBMP = CreateCompatibleBitmap(hMemDC, dw, dh)
    
    Dim lpbi As BITMAPINFO
    With lpbi.bmiHeader
        .biSize = 40
        .biWidth = dw
        .biHeight = -dh
        .biPlanes = 1
        .biBitCount = 32
    End With
    
    If SelectObject(memDC, memBMP) <> 0 Then
        If StretchBlt(memDC, 0, 0, dw, dh, hMemDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srccopy) <> 0 Then
            Dim rgbVals() As RGBQUAD: ReDim rgbVals((dh * dw) - 1)
            If GetDIBits(memDC, memBMP, 0, dh, rgbVals(0), lpbi, DIB_RGB_COLORS) <> 0 Then
                Dim result() As Long: ReDim result(0 To dw - 1, 0 To dh - 1)
                Dim i As Long
                Dim j As Long
                For i = 0 To dh - 1
                    For j = 0 To dw - 1
                        With rgbVals((i * dw) + j)
                            result(j, i) = rgb(.Green, .Blue, .Red)
                            
                        End With
                    Next j
                Next i
            getPixelsfromBMP = result
            Else
            End If
        Else
        End If
    End If
    DeleteObject memBMP
    DeleteDC memDC

End Function
Sub convertimagewcf()
    'fills cells with the values used in conditional formatting according to image in clipboard
    Dim width As Long, height As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim lColor As Long
    Dim kcolor() As Long
    Dim x As Integer, y As Integer
    Dim finalcolor As String
    Dim cs As Integer, rs As Integer, n As Integer, c As Integer, r As Integer
    Dim outrange As Range
    Dim dis As Long, temp As Long
    Dim key As Range
    Dim kcount As Integer
    Dim pixels As Variant
    
    Dim hBitmap As LongPtr
    Dim bmp As BITMAP
    Dim hMemDC As LongPtr
    
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo done
    
    Set key = ActiveSheet.ListObjects(1).DataBodyRange
    
    If OpenClipboard(0) Then
        'Get handle to bitmap from clipboard and close it
        hBitmap = GetClipboardData(CF_BITMAP)
        CloseClipboard
        If hBitmap <> 0 Then
            rs = InputBox("Number of rows")
            cs = InputBox("Number of columns")
            kcount = key.Rows.Count
            
            ' Create memory context and select bitmap
            hMemDC = CreateCompatibleDC(0)
            If SelectObject(hMemDC, hBitmap) <> 0 Then
                ' Put in bmp
                If apiGetObject(hBitmap, LenB(bmp), bmp) <> 0 Then
                    width = bmp.bmWidth
                    height = bmp.bmHeight
                    'width of each column, height of each row in pixels, rounded, dimensions of the boxes
                    c = width / cs
                    r = height / rs
                    
                    pixels = getPixelsfromBMP(hMemDC, bmp, rs * r, cs * c)
                End If
                
                DeleteDC hMemDC
                DeleteObject hBitmap
            End If
            
        Else
            MsgBox "No image found in clipboard."
            GoTo done
        End If
        
    End If

    'cells to be colored
    Set outrange = Range(ActiveCell.Address & ":" & ActiveCell.Offset(rs - 1, cs - 1).Address)

    ReDim kcolor(1 To kcount)
    For i = 1 To kcount
        kcolor(i) = key.Cells(i, 2).Interior.color
    Next
    
    'initializing variables, dis represents distance between colors and 500 is just an arbitrary number to reset it
    dis = 500
    n = 1
    finalcolor = key.Cells(1, 1).Value
    
    'starting point at the middle of the first box (upper left) so we should get the midpoints of each box
    'potentially might not want the midpoint and then you'd use a different point to start
    x = c / 2
    y = r / 2
    
    For i = 1 To rs
        For j = 1 To cs
            lColor = pixels(x, y)
            'find closest color in key
            For k = 1 To kcount
                temp = distancebetweencolors(kcolor(k), lColor)
                If dis > temp Then
                    dis = temp
                    finalcolor = key.Cells(k, 1).Value
                    n = k
                End If
            Next
            outrange.Cells(i, j).Value = finalcolor
            dis = 500
            x = x + c
        Next
        y = y + r
        'reset to beginning of row
        x = c / 2
    Next
    Application.ScreenUpdating = True
done:

End Sub

Sub convertimage()
    'changes cell color according to image in clipboard and desired rows and columns
    Dim width As Long, height As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim lColor As Long
    Dim kcolor() As Long
    Dim x As Integer, y As Integer
    Dim finalcolor As Long
    Dim cs As Integer, rs As Integer, n As Integer, c As Integer, r As Integer
    Dim outrange As Range
    Dim dis As Long, temp As Long
    Dim key As Range
    Dim kcount As Integer
    Dim pixels As Variant
    
    Dim hBitmap As LongPtr
    Dim bmp As BITMAP
    Dim hMemDC As LongPtr
    
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo done
    
    Set key = ActiveSheet.ListObjects(1).DataBodyRange
    
    If OpenClipboard(0) Then
        'Get handle to bitmap from clipboard and close it
        hBitmap = GetClipboardData(CF_BITMAP)
        CloseClipboard
        If hBitmap <> 0 Then
            rs = InputBox("Number of rows")
            cs = InputBox("Number of columns")
            kcount = key.Rows.Count

            ' Create memory context and select bitmap
            hMemDC = CreateCompatibleDC(0)
            If SelectObject(hMemDC, hBitmap) <> 0 Then
                ' Put in bmp
                If apiGetObject(hBitmap, LenB(bmp), bmp) <> 0 Then
                    width = bmp.bmWidth
                    height = bmp.bmHeight
                    'width of each column, height of each row in pixels, rounded, dimensions of the boxes
                    c = width / cs
                    r = height / rs
                    
                    pixels = getPixelsfromBMP(hMemDC, bmp, rs * r, cs * c)
                    
                End If
                DeleteDC hMemDC
                DeleteObject hBitmap
            End If
        Else
            MsgBox "No image found in clipboard."
            GoTo done
        End If
    End If

    'cells to be colored
    Set outrange = Range(ActiveCell.Address & ":" & ActiveCell.Offset(rs - 1, cs - 1).Address)
    

    ReDim kcolor(1 To kcount)
    For i = 1 To kcount
        kcolor(i) = key.Cells(i, 2).Interior.color
    Next
    
    'initializing variables, dis represents distance between colors and 500 is just an arbitrary number to reset it
    dis = 500
    n = 1
    finalcolor = kcolor(1)
    
    'starting point at the middle of the first box (upper left) so we should get the midpoints of each box
    'potentially might not want the midpoint and then you'd use a different point to start
    x = c / 2
    y = r / 2
    
    For i = 1 To rs
        For j = 1 To cs
            lColor = pixels(x, y)
            'find closest color in key
            For k = 1 To kcount
                temp = distancebetweencolors(kcolor(k), lColor)
                If dis > temp Then
                    dis = temp
                    finalcolor = kcolor(k)
                    n = k
                End If
            Next
            outrange.Cells(i, j).Interior.color = finalcolor
            dis = 500
            x = x + c
        Next
        y = y + r
        'reset to beginning of row
        x = c / 2
    Next
    Application.ScreenUpdating = True
done:

End Sub


