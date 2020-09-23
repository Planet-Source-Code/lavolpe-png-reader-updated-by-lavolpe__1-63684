VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stdPNGdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' DO NOT CALL THIS CLASS OR INSTANTIATE THIS CLASS FROM YOUR PROJECT
' This class is exposed via the stdPNG object/class. Calling this class' methods
' directly may result in memory leaks or general protection fault errors.
' *** You've been warned ****

' The resources used by this class are described below. Note that when the PNG file
' is completely processed, the only thing that remains is 1 DC and 1 bitmap. The
' bitmap of course is your PNG, the DC can be removed by calling the
' HasOwnDC method of stdPNG. Setting that property to False will remove the DC from
' within this class & setting it to True will have this class create & manage a DC.
' In either case, the bitmap and/or DC are destroyed upon termination of stdPNG.
' Suggest that if you have multiple stdPNG objects, set StdPNG.HasOwnDC=False. You
' can still call any of the stdPNG methods, passing it your own DC.


' When possible, the PNG bit depths are maintained in its BMP conversion.
' Following are exceptions. Note that progressively displaying an interlaced
' image almost always will be done via a 32bpp bitmap. Without it, alphablending
' for a fade in isn't possible short of pixel by pixel alphablending vs API (faster).
' Downside is size: Using a 256x256 4bpp (2 pixels per byte) example:
'   Without fade-in: bitmap will be (256*256*.5 bytes) or 32768 bytes.
'   With the 32bpp fade in: (256*256*4 bytes) or 262144 bytes (8x larger)
'   With a pixelated fade in (non-alpha images): same as without fade in
'   Although pixelated is default (slower) for non-alpha interlaced PNGs,
'       fade-in (faster) is optional

'   Therefore, pixelated will be the default for non-alpha interlaced PNGs

'    -------------   -------------  --- --------------  --- -------------
'    PNG Bit Depth   BMP Bit Depth  DC? 24bpp Work BMP  DC? Refresher BMP
'    -------------   -------------  --- --------------  --- -------------
'    With Transparency or Interlacing using Fade-In (even if transparency is opted out):
'    All bit depths     32bpp       Yes     *Yes        *Yes    *Yes
'                                           * if not progressively displayed then No
'
'    Interlaced: (using pixelated interlacing)
'    48bpp          24bpp           Yes     No          No      No
'    24bpp          24bpp           Yes     No          No      No
'    2bpp           4bpp            Yes     *Yes        No      No
'    All Others     Same as PNG     Yes     *Yes        No      No
'                                           * if not progressively displayed then No
'
'    Non-Interlaced, Non-Transparent, no progressive display or scanner-type display
'    48bpp          24bpp           Yes     No          No      No
'    2bpp           4bpp            Yes     No          No      No
'    All Others     Same as PNG     Yes     No          No      No
'

' PNGs are designed to also use 48bit images (billions of colors). I know
' of no routine out there that can use all 6 bytes (RRGGBB vs RGB); therefore, this
' project uses the high byte and the 48bit is converted to a 24bit image. Supposedly,
' a 48bit image is a 24bit image with 2 bitmap planes. But current APIs only
' support bitmaps with a single plane.

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound As SafeArrayBound
End Type


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_UPDATENOW As Long = &H100
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' BlendOp:
Private Const AC_SRC_OVER = &H0
' AlphaFormat:
Private Const AC_SRC_ALPHA = &H1

Private Declare Function AlphaBlend Lib "MSIMG32.dll" ( _
  ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long

Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Const OBJ_BITMAP As Long = 7

Private Declare Function GetDIBColorTable Lib "gdi32.dll" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, ByRef pRGBQuad As Any) As Long

Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32.dll" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, ByRef pcRGBQuad As Any) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GdiFlush Lib "gdi32.dll" () As Long

Private Type CIEXYZ
    ciexyzX As Long
    ciexyzY As Long
    ciexyzZ As Long
End Type
Private Type CIEXYZTRIPLE
    ciexyzRed As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyBlue As CIEXYZ
End Type
Private Type BITMAPINFOHEADER
    'base: http://msdn.microsoft.com/library/en-us/gdi/bitmaps_1rw2.asp
    'v4: http://msdn.microsoft.com/library/en-us/gdi/bitmaps_2k1e.asp
    'v5: http://msdn.microsoft.com/library/en-us/gdi/bitmaps_7c36.asp
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
    biClrImportant As Long  ' << NT 3.51 and older; 40 bytes
    bV4RedMask As Long
    bV4GreenMask As Long
    bV4BlueMask As Long
    bV4AlphaMask As Long
    bV4CSType As Long
    bV4Endpoints As CIEXYZTRIPLE
    bV4GammaRed As Long
    bV4GammaGreen As Long
    bV4GammaBlue As Long    ' << Win98 and above can use this; 108 bytes
    bV5Intent As Long
    bV5ProfileData  As Long
    bV5ProfileSize As Long
    bV5Reserved As Long     ' << Win98 and above can use this too; 124 bytes
End Type



' Type - GetObjectAPI.lpObject
Private Type BITMAP
    bmType As Long 'LONG
    bmWidth As Long 'LONG
    bmHeight As Long 'LONG
    bmWidthBytes As Long 'LONG
    bmPlanes As Integer 'WORD
    bmBitsPixel As Integer 'WORD
    bmBits As Long 'LPVOID
End Type

' Type - SavePictureEx
Private Type BITMAPFILEHEADER
    bfType As Integer 'WORD
    bfSize As Long 'DWORD
    bfReserved1 As Integer 'WORD
    bfReserved2 As Integer 'WORD
    bfOffBits As Long 'DWORD
End Type

' Type - SavePictureEx
'Private Type BITMAPINFOHEADER
'    biSize As Long 'DWORD
'    biWidth As Long 'LONG
'    biHeight As Long 'LONG
'    biPlanes As Integer 'WORD
'    biBitCount As Integer 'WORD (0,1,4,6,16,24,32)
'    biCompression As Long 'DWORD (BI_RGB,BI_RLE8,BI_RLE4,BI_BITFIELDS,BI_JPEG,BI_PNG)
'    biSizeImage As Long 'DWORD
'    biXPelsPerMeter As Long 'LONG
'    biYPelsPerMeter As Long 'LONG
'    biClrUsed As Long 'DWORD
'    biClrImportant As Long 'DWORD
'End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As Long ' really RGB quads but this way we can do RGB triples too
End Type

' Type - SavePictureEx
Private Type RGBQUAD
    rgbBlue As Byte 'BYTE
    rgbGreenas As Byte 'BYTE
    rgbRedas As Byte 'BYTE
    rgbReservedas As Byte 'BYTE
End Type

' Type - SavePictureEx
Private Type BITMAPINFO_1 ' 1 Bit (2 Colors - Monochrome)
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 1) As RGBQUAD
End Type

' Type - SavePictureEx
Private Type BITMAPINFO_4 ' 4 Bits (16 colors)
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As RGBQUAD
End Type

' Type - SavePictureEx
Private Type BITMAPINFO_8 ' 8 Bits (256 colors)
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

' Constants - Color Depths
Public Enum ColorDepths
    Color_True = 0 ' 24 Bit Color (Default - This is what VB works with)
    Color_256 = 256 ' 8 Bit Color (256 Colors)
    Color_16 = 16 ' 4 Bit Color (16 Colors)
    Color_2 = 2 ' 1 Bit Color (2 Colors - Monochrome)
End Enum

' Constants - BITMAP.bmType & CopyImage.uType
Public Enum PictureTypes
    IMAGE_BITMAP = 0
    IMAGE_CURSOR = 1
    IMAGE_ICON = 2
    IMAGE_ENHMETAFILE = 3
End Enum

' Constants - CopyImage.fuFlags
Private Const LR_COPYDELETEORG = &H8
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_COPYRETURNORG = &H4
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_MONOCHROME = &H1

' Constants - BITMAPINFOHEADER.biCompression
Private Const BI_RGB = 0 ' An uncompressed format.
Private Const BI_RLE8 = 1 ' A run-length encoded (RLE) format for bitmaps with 8 bpp.
Private Const BI_RLE4 = 2 ' An RLE format for bitmaps with 4 bpp.
Private Const BI_JPEG = 4 ' Windows 98, Windows 2000: Indicates that the image is a JPEG image.
Private Const BI_PNG = 5 ' Windows 98, Windows 2000: Indicates that the image is a PNG image.
Private Const BI_BITFIELDS = 3 ' Specifies that the bitmap is not compressed and that the
' color table consists of three DWORD color masks that specify
' the red, green, and blue components, respectively, of each pixel.
' This is valid when used with 16-bpp and 32-bpp bitmaps.

' Constants - GetDIBits.uUsage (RGB_or_PAL)
Private Const DIB_RGB_COLORS = 0 ' The color table should consist of an array of 16-bit
' indexes into the current logical palette.
Private Const DIB_PAL_COLORS = 1 ' The color table should consist of literal red, green,
' blue (RGB) values.

' Win32 API Declarations
Private Declare Function CopyImage Lib "user32" (ByVal hImage As Long, _
    ByVal uType As Long, _
    ByVal OutputWidth As Long, _
    ByVal OutputHeight As Long, _
    ByVal fuFlags As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hGDIObj As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, _
    ByVal hBitmap As Long, _
    ByVal FirstScanLine As Long, _
    ByVal ScanLineCount As Long, _
    ByRef Return_BitmapData As Any, _
    ByRef lpBITMAPINFO As Any, _
    ByVal RGB_or_PAL As Long) As Long
Private Declare Function GetGDIObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
    ByVal nCount As Long, _
    ByRef lpObject As Any) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hGDIObj As Long) As Long


Private Type DCinfo
    hDC As Long
    hBmp As Long
    hPreBmp As Long
    hPtr As Long
End Type

Private PngDat As DCinfo
Private WorkDat As DCinfo
Private hSrc As Long
Private m_BmpInfo As BITMAPINFO
Private m_TgtCX As Long, m_TgtCY As Long

Friend Function CreateBitmapBase(BitCount As Byte, Width As Long, _
                Height As Long, Palette() As Byte, targetDC As Long) As Boolean

' Parameters, none are modified
' ----------------------------
'   BitCount :: Bitmap bit count needed to process PNG
'   Width :: the width of the PNG image
'   Height :: the height of the PNG image
'   Palette() :: array of palette entries provided by the PNG or created locally
'   targetDC :: DC where PNG will be drawn to (may be zero if only reading/caching PNG)

' This function creates a blank DIBSection which will eventually be the png converted to bitmap
' The plus side of using DIBs is that we can read/write to its memory space directly, which
'   prevents us from having to create copies of the byte arrays, manipulating, then
'   copying back to the DIB via SetDIBits. We just simply update the DIB directly

Dim pIndex As Long, pOffset As Integer ' palette counters
Dim tDC As Long                        ' temp DC used to create DIB
Dim tgtHwnd As Long
Dim tgtBmp As BITMAPINFO

CleanUp_ALL     ' clear any existing GDI objects

With m_BmpInfo.bmiHeader  ' set the default header information, regardless of bitdepth
    .biBitCount = BitCount ' take 1st 6 bits of passed integer (bit count)
    .biSizeImage = ScanLineWidth(BitCount, Width) * Height
    .biWidth = Width
    .biHeight = -Height ' top to bottom; easier to process as pngs are top to bottom
    .biCompression = BI_RGB ' no compression
    .biPlanes = 1           ' limitation per MSDN (prevents full 48bpp capability)
    .biSize = 40
End With

Select Case BitCount
    Case 32: ' full color w/transparency
        ' testing until positive full v5 of BitmapHeader unneeded
    Case 24: ' full color - no transparency
       ' testing until positive full v5 of BitmapHeader unneeded
    
    Case 8: ' palettized
        With m_BmpInfo
            If IsArrayEmpty(Not Palette) Then
                ReDim .bmiColors(0 To 255)
                For pIndex = 1 To 255
                    .bmiColors(pIndex) = (pIndex * 65536) Or (pIndex * 256&) Or pIndex
                Next
                .bmiHeader.biClrUsed = 256
            Else
                ReDim .bmiColors(0 To (UBound(Palette) + 1) \ 3 - 1)
                For pIndex = 0 To UBound(Palette) Step 3
                    .bmiColors(pOffset) = Palette(pIndex + 2) Or _
                        (Palette(pIndex + 1) * 256&) Or _
                        (Palette(pIndex + 0) * 65536)
                    pOffset = pOffset + 1
                Next
                .bmiHeader.biClrUsed = (UBound(.bmiColors) + 1)
            End If
        End With
        
        
    Case 4: ' 2 pixels per byte
        With m_BmpInfo
            If IsArrayEmpty(Not Palette) Then
                ReDim .bmiColors(0 To 15)
                For pIndex = 1 To 15
                    pOffset = pOffset + 17
                    .bmiColors(pIndex) = (pOffset * 65536) Or (pOffset * 256&) Or pOffset
                Next
            Else
                ReDim .bmiColors(0 To UBound(Palette) \ 3)
                For pIndex = 0 To UBound(Palette) - 1 Step 3
                    .bmiColors(pOffset) = Palette(pIndex + 2) _
                        Or (Palette(pIndex + 1) * 256&) _
                        Or (Palette(pIndex) * 65536)
                    pOffset = pOffset + 1
                Next
            End If
            .bmiHeader.biClrUsed = UBound(.bmiColors) + 1
        End With
        
    Case 1: ' monochrome (two-color)
        With m_BmpInfo
            ReDim .bmiColors(0 To 1)
            If IsArrayEmpty(Not Palette) Then
                .bmiColors(1) = vbWhite
                .bmiHeader.biClrUsed = 2
            Else
                .bmiColors(0) = Palette(2) Or (Palette(1) * 256&) Or (Palette(0) * 65536)
                If UBound(Palette) > 2 Then
                    .bmiColors(1) = Palette(5) Or (Palette(4) * 256&) Or (Palette(3) * 65536)
                End If
                .bmiHeader.biClrUsed = 2
            End If
        End With
    
    Case Else
        Exit Function
End Select

    If targetDC Then
        tDC = targetDC
    Else
        tDC = GetDC(0)   ' get a DC in order to create DIB & then create the DIB
    End If
    ' note we are passing a ptr placeholder & API will assign it if DIB created
    PngDat.hBmp = CreateDIBSection(tDC, m_BmpInfo, 0, PngDat.hPtr, ByVal 0&, 0)
    
    If PngDat.hPtr <> 0 Then ' no errors
        ' create a DC to display our DIB in
        PngDat.hDC = CreateCompatibleDC(tDC)
        If PngDat.hDC <> 0 Then ' no errors
            If targetDC = 0 Then ReleaseDC 0, tDC  ' go ahead & release the DC we grabbed earlier
            m_BmpInfo.bmiHeader.biHeight = -m_BmpInfo.bmiHeader.biHeight
            m_TgtCX = m_BmpInfo.bmiHeader.biWidth
            m_TgtCY = m_BmpInfo.bmiHeader.biHeight
            
            ' now select the bitmap into our DC
            PngDat.hPreBmp = SelectObject(PngDat.hDC, PngDat.hBmp)
            
            If m_BmpInfo.bmiHeader.biBitCount < 16 Then
                ' uses palettes; so set the palette to the DC
                SetDIBColorTable PngDat.hDC, 0, m_BmpInfo.bmiHeader.biClrUsed, m_BmpInfo.bmiColors(0)
                
            ElseIf m_BmpInfo.bmiHeader.biBitCount = 32 Then
                ' fill image as transparent for now
                FillMemory ByVal PngDat.hPtr, m_BmpInfo.bmiHeader.biSizeImage, 0
            End If
            
            CreateBitmapBase = True
            
            
        Else
            ' error create DC for our png; system resources must be too low
            If targetDC = 0 Then ReleaseDC 0, tDC
            CleanUp_ALL
        
        End If
        
    Else ' error creating png; either my dWord calculations are off or system resources low
        If targetDC = 0 Then ReleaseDC 0, tDC
        CleanUp_ALL
    End If
    Erase m_BmpInfo.bmiColors()


End Function

Friend Function CreateWorkSpace(ByVal targetDC As Long, ByVal X As Long, ByVal Y As Long, _
                    ByVal Interlaced As Boolean, ByVal altBkgColor As Long) As Boolean

' Function creates the needed bitmaps and/or DCs to show progressively displayed PNGs
' Only called if we are reading & displaying PNG on a window-owned DC, not memory DC

' Parameters, none are modified
' ----------------------------
'   targetDC :: DC where PNG will be drawn to
'   X,Y :: coordinates at targetDC to draw to
'   Cx,Cy :: Width & Height of targetDC being drawn to
'   pX,pY :: Left,Top position of PNG to begin drawing from
'   pCx,pCy :: Width & Height of PNG being drawn


    Dim bmpInfo24 As BITMAPINFO         ' bitmap header for 24bit image only
    Dim tgtBmp As BITMAPINFO            ' header for destDC's bitmap
    Dim hBrush As Long
    Dim bkgRect As RECT
    Dim tDC As Long
    
    If m_BmpInfo.bmiHeader.biBitCount = 32 Then
        ' determine min width of DC needed to refresh alpha progressive display
        With m_BmpInfo.bmiHeader
            If targetDC = 0 Then
                m_TgtCX = .biWidth
                m_TgtCY = .biHeight
                tDC = GetDC(0)
            Else
                ' get information from the targetDC's bitmap
                tgtBmp.bmiHeader.biSize = Len(tgtBmp.bmiHeader)
                If GetGDIObjectAPI(GetCurrentObject(targetDC, OBJ_BITMAP), Len(tgtBmp), tgtBmp) = 0 Then Exit Function
                With tgtBmp.bmiHeader
                    If .biWidth < X + m_TgtCX Then m_TgtCX = .biWidth - X
                    If .biHeight < Y + m_TgtCY Then m_TgtCY = .biHeight - Y
                End With
                tDC = targetDC
            End If
        End With
        
        With bmpInfo24.bmiHeader
            .biBitCount = 24
            .biHeight = -m_TgtCY
            .biWidth = m_TgtCX
            .biSize = Len(bmpInfo24.bmiHeader)
            .biPlanes = 1
            .biSizeImage = ScanLineWidth(24, m_TgtCX) * m_TgtCY
        End With
    
        ' create a bitmap to be used to draw on
        WorkDat.hBmp = CreateCompatibleBitmap(tDC, m_TgtCX, m_TgtCY)
        'WorkDat.hBmp = CreateDIBSection(tDC, bmpInfo24, 0, WorkDat.hPtr, ByVal 0&, 0)
        ' create a bitmap that will be used for refreshing btwn progressive scans
        If altBkgColor < 0 Then hSrc = CreateCompatibleBitmap(tDC, m_TgtCX, m_TgtCY)
        
        If targetDC = 0 Then ReleaseDC 0, tDC
        
        If WorkDat.hBmp <> 0 And (hSrc <> 0 Or altBkgColor > -1) Then ' no errors
            
            ' now create a separate DC used to draw in
            WorkDat.hDC = CreateCompatibleDC(targetDC)
            If WorkDat.hDC Then
                
                If altBkgColor < 0 Then
                    WorkDat.hPreBmp = SelectObject(WorkDat.hDC, hSrc)
                    BitBlt WorkDat.hDC, 0, 0, m_TgtCX, m_TgtCY, targetDC, X, Y, vbSrcCopy
                    ' select working copy of bitmap into the work DC
                    ' & get another copy of the destination area
                    SelectObject WorkDat.hDC, WorkDat.hBmp
                    BitBlt WorkDat.hDC, 0, 0, m_TgtCX, m_TgtCY, targetDC, X, Y, vbSrcCopy
                Else
                    WorkDat.hPreBmp = SelectObject(WorkDat.hDC, WorkDat.hBmp)
                    hBrush = CreateSolidBrush(altBkgColor)
                    SetRect bkgRect, 0, 0, m_TgtCX, m_TgtCY
                    FillRect WorkDat.hDC, bkgRect, hBrush
                    DeleteObject hBrush
                End If
            
            Else    ' got troubles here; low on resources maybe?
                CleanUp_Workspace ' destroy any workspace objects
                Exit Function ' PNG can still be processed, but not progressively displayed
            End If
        Else    ' got troubles here; low on resources maybe?
            Exit Function ' PNG can still be processed, but not progressively displayed
        End If
        
    ElseIf Interlaced = True Then
    
        If m_BmpInfo.bmiHeader.biBitCount < 24 Then
            ' non-alpha interlacing; create a 24bit image to work on to display pixelated image
            ' Again, 24bit pngs are the exception & don't need the extra workspace
            With bmpInfo24.bmiHeader
                .biSize = 40
                .biHeight = -m_BmpInfo.bmiHeader.biHeight
                .biWidth = m_BmpInfo.bmiHeader.biWidth
                .biSizeImage = -(ScanLineWidth(24, m_BmpInfo.bmiHeader.biWidth) * m_BmpInfo.bmiHeader.biHeight)
                .biBitCount = 24
                .biPlanes = 1
                .biClrUsed = 0
            End With
            ' create the working DIBsection
            WorkDat.hBmp = CreateDIBSection(targetDC, bmpInfo24, 0, WorkDat.hPtr, ByVal 0&, 0)
            If WorkDat.hBmp = 0 Then Exit Function
        End If
    End If
    CreateWorkSpace = True
End Function

Friend Sub UseWorkingCopy(ByVal globalAlpha As Byte)

    ' when using an alternate background color on transparent images, we use a
    ' 32bpp to get transparency information and blt it over a solid color 24bpp
    ' image to apply the effect. When the image is completely processed we are
    ' left with a 32bpp with transparency & a 24bpp without transparency.
    ' So, we will kill the 32bpp image & use the 24bpp (with the alternate bkg color).

    ' checks. This shouldn't happen as it is only called from the stdPNG class,
    ' but since there is nothing stopping you from calling this from within your
    ' app if you instantiate this class, we will add a few checks.
    If WorkDat.hDC = 0 Then Exit Sub
    If m_BmpInfo.bmiHeader.biBitCount < 32 Then Exit Sub
    
    Dim lBlend As Long          ' used for alpha blending
    Dim bf As BLENDFUNCTION     ' used for alpha blending
    Dim dibArray() As Byte      ' 24bpp DDB bitmap bytes
    Dim tBmp As Long            ' 24bpp DIB pointer
    
    ' build the AlphaBlend structure
    bf.BlendOp = AC_SRC_OVER
    bf.SourceConstantAlpha = globalAlpha
    bf.AlphaFormat = AC_SRC_ALPHA
    CopyMemory lBlend, bf, &H4     ' copy the 4byte structure to a Long value
    
    ' ensure working copy is the most current
    AlphaBlend WorkDat.hDC, 0, 0, m_TgtCX, m_TgtCY, PngDat.hDC, 0, 0, m_TgtCX, m_TgtCY, lBlend
    
    ' we are moving from a 32bpp to 24bpp, adjust our m_BmpInfo header
    With m_BmpInfo.bmiHeader
        .biBitCount = 24
        .biSizeImage = ScanLineWidth(24, m_TgtCX) * m_TgtCY
        ReDim dibArray(0 To .biSizeImage - 1) ' resize array
    End With
    ' create the 24bpp image
    tBmp = CreateDIBSection(PngDat.hDC, m_BmpInfo, 0, PngDat.hPtr, ByVal 0&, 0)
    
    If tBmp Then
        ' remove and delete the 32bpp
        DeleteObject SelectObject(PngDat.hDC, PngDat.hPreBmp)
        ' get the byte info for the 24bpp working copy
        GetDIBits WorkDat.hDC, WorkDat.hBmp, 0, m_TgtCY, dibArray(0), m_BmpInfo, 0&
        ' copy those bytes to the new 24bpp
        CopyMemory ByVal PngDat.hPtr, dibArray(0), UBound(dibArray) + 1
        ' select new 24bpp into our DC
        PngDat.hBmp = tBmp
        PngDat.hPreBmp = SelectObject(PngDat.hDC, PngDat.hBmp)
    End If
    
    Erase dibArray()
    CleanUp_Workspace

End Sub

Friend Function DisplayInterlaced(ByVal m_tgtDC As Long, ByVal m_tgtHwnd As Long, _
        ByVal dcX As Long, ByVal dcY As Long, ByVal rowInitial As Long, ByVal rowInterval As Long, _
        ByVal isFinalRow As Boolean, ByVal isPartial As Boolean, _
        ByVal altBkgColor As Long, ByVal globalAlpha As Byte) As Boolean

' Function displays the image, either in whole or in part

' Parameters, none are modified
' ----------------------------
'   destDC :: the DC to draw the PNG into
'   rowIntitial :: used for interlacing & is the start row of image to draw at
'   rowInterval :: used for interlacing & is the step value for rowInitial
'   isFinalRow :: flag indicating last row of a progressively displayed PNG
'   destHwnd :: Optional. If destDC is non-memory DC, then handle to refresh after drawing
'   isPartial :: Optional. Non-standard interlacing
'       ^^ Some PNGs can be huge and if not interlaced, you can wait for literally seconds
'          before the image is displayed. With an optional property, users can choose
'          to fake interlacing for larger images. This interlacing method gives the
'          appearance of an image being scanned & produces visual confirmation that the
'          routines are working vs displaying a simple hourglass & waiting

' this routine shouldn't be called without workspaces already created
If PngDat.hPtr = 0 Then Exit Function
If m_tgtDC = 0 Then Exit Function    ' ensure DC passed


Dim lBlend As Long          ' used for alpha blending
Dim bf As BLENDFUNCTION     ' used for alpha blending
Dim Row As Long             ' loop counter
Dim iRect As RECT           ' used for refreshing DC after drawing
Dim bkgRect As RECT
Dim hBrush As Long


If m_BmpInfo.bmiHeader.biBitCount = 32 Then ' uses transparency
    
    ' build the AlphaBlend structure
    bf.BlendOp = AC_SRC_OVER
    bf.SourceConstantAlpha = globalAlpha
    bf.AlphaFormat = AC_SRC_ALPHA
    CopyMemory lBlend, bf, &H4     ' copy the 4byte structure to a Long value
    
    If WorkDat.hDC = 0 Or isFinalRow = True Then  ' we can progressively display interlaced png
    
            ' so we will blend the final result to the destination dc
            AlphaBlend WorkDat.hDC, 0, 0, m_TgtCX, m_TgtCY, PngDat.hDC, 0, 0, m_TgtCX, m_TgtCY, lBlend
            BitBlt m_tgtDC, dcX, dcY, m_TgtCX, m_TgtCY, WorkDat.hDC, 0, 0, vbSrcCopy
            ' identify far row to update in destination dc
            iRect.Bottom = m_TgtCY
    
    Else
        If isPartial Then
        
            ' display n blended rows. Used for non-interlaced, larger images
            AlphaBlend WorkDat.hDC, 0, rowInitial, m_TgtCX, rowInterval, PngDat.hDC, 0, rowInitial, m_TgtCX, rowInterval, lBlend
            BitBlt m_tgtDC, dcX, rowInitial + dcY, m_TgtCX, rowInterval, WorkDat.hDC, 0, rowInitial, vbSrcCopy
            ' identify far row to update in destination dc
            iRect.Bottom = rowInitial + rowInterval
        
        Else
            
            Row = rowInitial
            Do While Row < m_TgtCY
                ' we only need to alphablend a row at a time vs every row, this significantly
                ' improves speed when say only 1/8th of the image is being progressively displayed
                AlphaBlend WorkDat.hDC, 0, Row, m_TgtCX, 1, PngDat.hDC, 0, Row, m_TgtCX, 1, lBlend
                BitBlt m_tgtDC, dcX, Row + dcY, m_TgtCX, 1, WorkDat.hDC, 0, Row, vbSrcCopy
                Row = Row + rowInterval
            Loop
            ' identify far row to update in destination dc
            iRect.Bottom = Row - rowInterval + 1
            
        End If
    
'    Else    ' for whatever reasons, we don't have resources to work off screen
        
'        If isFinalRow Then
'            ' so we will blend the final result to the destination dc
'            AlphaBlend m_tgtDC, dcX, dcY, m_TgtCX, m_TgtCY, PngDat.hDC, 0, 0, m_TgtCX, m_TgtCY, lBlend
'            ' identify far row to update in destination dc
'            iRect.Bottom = m_TgtCY
            
'        End If
        
    End If
    
    If Not isFinalRow Then
        ' we need to replace the working copy, otherwise we blend over blend, forever
        If altBkgColor < 0 Then
            SelectObject PngDat.hDC, hSrc
            BitBlt WorkDat.hDC, 0, 0, m_TgtCX, m_TgtCY, PngDat.hDC, 0, 0, vbSrcCopy
            SelectObject PngDat.hDC, PngDat.hBmp
        End If
    End If
    If altBkgColor > -1 Then
        hBrush = CreateSolidBrush(altBkgColor)
        SetRect bkgRect, 0, 0, m_TgtCX, m_TgtCY
        FillRect WorkDat.hDC, bkgRect, hBrush
        DeleteObject hBrush
    End If
    
Else ' PNGs without any transparency

    If WorkDat.hPtr <> 0 And Not isFinalRow Then
        ' non-alpha interlaced images less than 24bit
        
        If WorkDat.hPtr = PngDat.hPtr Then
            ' 24bit interlaced. Uses only one DC; blt the results
            BitBlt m_tgtDC, dcX, dcY, m_TgtCX, m_TgtCY, PngDat.hDC, 0, 0, vbSrcCopy
        Else
            ' non-24bit interlaced. The working copy is 24bit, so blt from that DC
            ' These do not have their own DC, so we toggle them in an out for drawing
            
            ' get the 24bit into our DC & blt results
            SelectObject PngDat.hDC, WorkDat.hBmp
            BitBlt m_tgtDC, dcX, dcY, m_TgtCX, m_TgtCY, PngDat.hDC, 0, 0, vbSrcCopy
            ' now get the 24bit out of the DC
            SelectObject PngDat.hDC, PngDat.hBmp
        End If
        ' identify far row to update in destination dc
        iRect.Bottom = m_TgtCY
        
    Else ' non-interlaced, non-alpha
    
        If isPartial Then   ' custom interlacing; only blt a few rows at a time
            If rowInitial > m_TgtCY Then Exit Function
            BitBlt m_tgtDC, dcX, rowInitial + dcY, m_TgtCX, rowInterval, PngDat.hDC, 0, rowInitial, vbSrcCopy
            ' identify far row to update in destination dc
            iRect.Bottom = rowInitial + rowInterval
        Else
            ' straight blt
            BitBlt m_tgtDC, dcX, dcY, m_TgtCX, m_TgtCY, PngDat.hDC, 0, 0, vbSrcCopy
            ' identify far row to update in destination dc
            iRect.Bottom = m_TgtCY
        End If
    
    End If
    
End If

' now tell the DC to update
If m_tgtHwnd Then
    SetRect iRect, dcX, rowInitial + dcY, m_TgtCX + dcX, iRect.Bottom + dcY
    RedrawWindow m_tgtHwnd, iRect, ByVal 0&, &H100 Or 1
End If

If isFinalRow = True Then
    If altBkgColor < 0 Then
        CleanUp_Workspace ' get rid of any extra workspaces or bitmaps
    Else
        UseWorkingCopy globalAlpha
    End If
End If


End Function


Private Function ScanLineWidth(BitDepth As Byte, Width As Long) As Long
' simple function to align byte range on dWord boundaries
'    Dim tScanWidth As Long
'    tScanWidth = (Width * BitDepth + 7) \ 8
'    ScanLineWidth = (tScanWidth Mod 4 + tScanWidth)
    ScanLineWidth = (((Width * BitDepth) + &H1F) And Not &H1F&) \ &H8
End Function

Private Function IsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
  ' test to see if an array has been initialized
  IsArrayEmpty = (lArrayPointer = -1)
End Function

Friend Sub CleanUp_Workspace()
    ' called when class terminates and when last line of a PNG has been displayed
    ' Note that this does not destroy the processed PNG-converted bitmap
    
    If WorkDat.hDC Then
        ' if a working space was created, destroy it
        DeleteObject SelectObject(WorkDat.hDC, WorkDat.hPreBmp)
        DeleteDC WorkDat.hDC
        WorkDat.hDC = 0
    ElseIf WorkDat.hBmp Then
        ' 24bit interlaced images may have a work bitmap even though it hasn't a DC
        DeleteObject WorkDat.hBmp
    End If
    WorkDat.hBmp = 0
    ' when 32bit alphas are used, we created a refresher bitmap; destroy it
    If hSrc Then DeleteObject hSrc
    hSrc = 0
    WorkDat.hPtr = 0
End Sub

Friend Sub CleanUp_ALL()
    
    ' Function called before a PNG-DIB is created and when the class is terminated.
    ' All memory objects are destroyed:
    
    If PngDat.hDC Then
        DeleteObject SelectObject(PngDat.hDC, PngDat.hPreBmp)
        DeleteDC PngDat.hDC
        PngDat.hDC = 0
    Else
        If PngDat.hBmp Then DeleteObject PngDat.hBmp
    End If
    PngDat.hBmp = 0
    PngDat.hPtr = 0
    m_BmpInfo.bmiHeader.biBitCount = 0
    CleanUp_Workspace

End Sub

Friend Property Get PngDataPtr() As Long
    PngDataPtr = PngDat.hPtr
End Property
Friend Property Get PngDataSize() As Long
    PngDataSize = m_BmpInfo.bmiHeader.biSizeImage
End Property
Friend Property Get PngHandle() As Long
    PngHandle = PngDat.hBmp
End Property
Friend Property Get PngDC() As Long
    PngDC = PngDat.hDC
End Property
Friend Property Get pngDataBitCount() As Byte
    pngDataBitCount = m_BmpInfo.bmiHeader.biBitCount
End Property

Friend Property Get PngDataPtr_Working() As Long
    PngDataPtr_Working = WorkDat.hPtr
End Property
Friend Property Get PngDataSize_Working() As Long
    PngDataSize_Working = ScanLineWidth(24, m_BmpInfo.bmiHeader.biWidth) * m_BmpInfo.bmiHeader.biHeight
End Property
Friend Property Let ManageDC(bManaged As Boolean)
    If PngDat.hBmp = 0 Then Exit Property
    
    If bManaged = False Then
        If PngDat.hDC Then
            SelectObject PngDat.hDC, PngDat.hPreBmp
            DeleteDC PngDat.hDC
            PngDat.hDC = 0
        End If
    Else
        If PngDat.hDC = 0 Then
            Dim tDC As Long
            tDC = GetDC(0)
            PngDat.hDC = CreateCompatibleDC(tDC)
            ReleaseDC 0, tDC
            PngDat.hPreBmp = SelectObject(PngDat.hDC, PngDat.hBmp)
        End If
    End If

End Property
    
Private Sub Class_Terminate()
    CleanUp_ALL
End Sub

Friend Sub DrawToDC(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal Cx1 As Long, ByVal Cy1 As Long, ByVal fromDC As Long, _
        globalAlpha As Byte, Optional hWndRefresh As Long)
        
    If PngDat.hBmp = 0 Then Exit Sub
        
    Dim hOldBitmap As Long, iRect As RECT
        
    ' no negative values (mirroring not supported by AlphaBlend API)
    If hDC = 0 Or cx < 0 Or cy < 0 Then Exit Sub
    
    ' allow a bunch of optional parameters? Gotta provide them as needed though
    If cx = 0 Then cx = m_BmpInfo.bmiHeader.biWidth ' ensure non-zero, positive values
    If cy = 0 Then cy = m_BmpInfo.bmiHeader.biHeight
    If Cx1 < 1 Then Cx1 = m_BmpInfo.bmiHeader.biWidth ' ensure non-zero, positive values
    If Cy1 < 1 Then Cy1 = m_BmpInfo.bmiHeader.biHeight
    
    ' when hDC=0, assume we are managing the DC. Ensure a valid value is passed
    If fromDC = 0 Then fromDC = PngDat.hDC
    If fromDC = 0 Then Exit Sub
    
    ' if we are not managing our own DC, select PNG into the passed DC
    If fromDC <> PngDat.hDC Then hOldBitmap = SelectObject(fromDC, PngDat.hBmp)
    
    On Error Resume Next
    If m_BmpInfo.bmiHeader.biBitCount = 32 Then ' alpha type bitmap
    
        Dim lBlend As Long          ' used for alpha blending
        Dim bf As BLENDFUNCTION     ' used for alpha blending
                
        ' build the AlphaBlend structure
        bf.BlendOp = AC_SRC_OVER
        bf.SourceConstantAlpha = globalAlpha
        bf.AlphaFormat = AC_SRC_ALPHA
        CopyMemory lBlend, bf, &H4     ' copy the 4byte structure to a Long value
                
        AlphaBlend hDC, X, Y, cx, cy, fromDC, X1, Y1, Cx1, Cy1, lBlend
    
    Else
    
        StretchBlt hDC, X, Y, cx, cy, fromDC, X1, Y1, Cx1, Cy1, vbSrcCopy
    
    End If
    
    ' when not managing our own DC, unselect the PNG from the passed DC
    If fromDC <> PngDat.hDC Then SelectObject fromDC, hOldBitmap
    
    If Err Then Err.Clear
        
    If hWndRefresh Then
        SetRect iRect, X, Y, X + cx, Y + cy
        RedrawWindow hWndRefresh, iRect, ByVal 0&, &H100 Or 1
    End If
        
End Sub


Private Sub TileBkgColor(ByVal bkgColor As Long, Height As Long)

If bkgColor < 0 Then ' transparent
    FillMemory ByVal Me.PngDataPtr, Me.PngDataSize, 0
    


Else                 ' non-transparent
    
    Dim X As Long, xOffset As Long
    Dim hColor As Long, ScanWidth As Long, DataSize As Long
    Dim TSA As SafeArray1D, PngBytes() As Byte
    
    DataSize = Me.PngDataSize
    With TSA
        .cDims = 1                          ' Number of dimensions
        .cbElements = 1                     ' Size of data elements
        .pvData = Me.PngDataPtr             ' Data address
        .rgSABound.lLbound = 0              ' Low bound of array
        .rgSABound.cElements = DataSize     ' Nr of elements in array
    End With
    CopyMemory ByVal VarPtrArray(PngBytes), VarPtr(TSA), &H4&
    
    bkgColor = ((bkgColor \ 65536) And &HFF) Or ((bkgColor \ 256) And &HFF) * 256 Or ((bkgColor And &HFF) * 65536)
    CopyMemory PngBytes(0), bkgColor, &H3
    CopyMemory ByVal Me.PngDataPtr + 3, 255, 1
    CopyMemory bkgColor, PngBytes(0), &H4
    Debug.Print "reversed "; bkgColor
    
    ScanWidth = DataSize \ Height
    
    xOffset = 4
    Do While (xOffset + xOffset) < DataSize
        CopyMemory PngBytes(xOffset), PngBytes(0), xOffset
        xOffset = (xOffset + xOffset)
    Loop
    If xOffset < DataSize Then
        CopyMemory PngBytes(xOffset), PngBytes(0), DataSize - xOffset
    End If
    
    CopyMemory ByVal VarPtrArray(PngBytes), 0&, &H4&
    
End If
End Sub


