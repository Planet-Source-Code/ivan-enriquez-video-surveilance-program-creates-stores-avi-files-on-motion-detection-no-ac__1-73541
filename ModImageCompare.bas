Attribute VB_Name = "ModImageCompare"
Option Explicit

Private Type BITMAPINFOHEADER
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

Private Type RGBTRIPLE
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBTRIPLE
End Type

Public Declare Function SendMessageA Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal HDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal HDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Dim CurColors() As RGBTRIPLE, prevColors() As RGBTRIPLE ', PrevColors() As RGBTRIPLE
Dim DifferenceCount As Long, Index As Long, PixelCount As Long
Dim DifRed As Long, DifGreen As Long, DifBlue As Long
Dim aancho As Long, aalto As Long


'Realiza un comparativo entre dos imagenes (en controles picturebox). Regresa el numero de pixeles que
'cambiaron, tomando en cuenta un indice de tolerancia entre cambio de colores (NoiseReductionLevel).

Public Function ImageDifference(Pic_1 As PictureBox, Pic_2 As PictureBox) As Long
    Dim bih_1 As BITMAPINFO
    Dim bih_2 As BITMAPINFO
    With bih_1.bmiHeader
        .biSize = Len(bih_1)
        .biPlanes = 1
        .biWidth = Ancho
        .biHeight = Alto
        .biBitCount = 24
    End With
    bih_2 = bih_1
    PixelCount = Ancho * Alto
    ReDim CurColors(PixelCount)
    ReDim prevColors(PixelCount)
    'ReDim DifColors(PixelCount)
    GetDIBits Pic_1.HDC, Pic_1.Image, 0&, Alto, CurColors(0), bih_1, 0
    GetDIBits Pic_2.HDC, Pic_2.Image, 0&, Alto, prevColors(0), bih_2, 0
    ImageDifference = ComparePixels
    SetDIBits Form1.Picture1.HDC, Form1.Picture1.Image, 0&, Alto, CurColors(0), bih_1, 0
End Function

'Returns true if pixels from both arrays are the same.
Private Function ComparePixels() As Long
    Dim INTERMEDIO As Long
    Dim I As Long
    Dim j As Long
    Dim K As Long
    ComparePixels = 0
        'For lonLoop = 0 To PixelCount - 1 Step GridDensity
        '   INTERMEDIO = Abs(CLng(CurColors(lonLoop).rgbBlue) - CLng(PrevColors(lonLoop).rgbBlue)) + _
        '   Abs(CLng(CurColors(lonLoop).rgbGreen) - CLng(PrevColors(lonLoop).rgbGreen)) + _
        '   Abs(CLng(CurColors(lonLoop).rgbRed) - CLng(PrevColors(lonLoop).rgbRed))
        '   If INTERMEDIO > NoiseReductionLevel Then
        '     ComparePixels = ComparePixels + 1
        '     DifColors(lonLoop) = CurColors(lonLoop)
        '   End If
        'Next lonLoop
         ' For I = 0 To Ancho - 1 Step GridDensity
         '    For J = 0 To Alto - 1 Step GridDensity
         '       K = I * J
         '       INTERMEDIO = Abs(CLng(CurColors(K).rgbBlue) - CLng(PrevColors(K).rgbBlue)) + _
         '       Abs(CLng(CurColors(K).rgbGreen) - CLng(PrevColors(K).rgbGreen)) + _
         '       Abs(CLng(CurColors(K).rgbRed) - CLng(PrevColors(K).rgbRed))
         '       If INTERMEDIO > NoiseReductionLevel Then
         '          ComparePixels = ComparePixels + 1
         '          DifColors(K) = CurColors(K)
         '       End If
         '   Next J
         'Next I
         
         'se simplificó el loop (de ser recorrido por ver luego por hor se dejo uno del largo de la resolucion
         'esto implica que el grid es parejo verticalmente y se distancia horizontalmente
         'se quitó la revisión de cada rayo, ahora solo se checa uno.
         For I = 0 To PixelCount Step ScanGridDensity
                INTERMEDIO = Abs(CLng(CurColors(I).rgbBlue) - CLng(prevColors(I).rgbBlue)) + _
                Abs(CLng(CurColors(I).rgbGreen) - CLng(prevColors(I).rgbGreen)) + _
                Abs(CLng(CurColors(I).rgbRed) - CLng(prevColors(I).rgbRed))
                If INTERMEDIO > NoiseReductionLevel Then
                   ComparePixels = ComparePixels + 1
                   CurColors(I).rgbBlue = 255
                   CurColors(I).rgbRed = 255
                   CurColors(I).rgbGreen = 255
                
                End If
         Next I
End Function


