VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ZLIB Specific Instructions: Don't follow these instructions & project will NOT work!
' 1. Go to http://www.winimage.com/zLibDll/zlib123dll.zip
' 2. Save the Zip file to your desktop
' 3. Open the zip file and extract to your desktop the file named: zlibwapi.dll
' 4. Rename the zlibwapi.dll to zlibVB.dll
'    If you opted to use Folder Names while unzipping, the DLL will be in a
'    folder named Dll32, also on your desktop
' 5. Move that file to your system folder. No DLL registration is needed
' 6. Delete/clean up any desktop items that resulted from the previous 5 steps.
' 7. After playing with project, delete the DLL from your system folder if desired
' 8. If for some reason you cannot get to that website. Email me for a copy: LaVolpeVB@hotmail.com

' CALLING ALL CODERS!!!
' Let's try to make the best PNG project on PSC and share it with everyone. Asking each
' of you to take a look at the code provided & improve and add upon it. Every coder has
' their own area of expertise and could make this a better or more accurate/efficient
' project. Note that many of the png chunks are not handled; up to you to help there.

' I'm thinking that everyone who wants to improve this should do so and repost it
' under their handle. To keep things somewhat organized, suggest following when
' updating/reposting:
' 1. Copy & paste the top portion of this module into your new version
'   so it is carried over to your new version
' 2. Mark your changes with your handle & date/time
' 3. Always include link to the other revisions
'    - this way we can see what was changed & possibly reuse code no longer in current version
'      For example. In my revision here, I removed the vb-only code to decode png compression,
'      but that information is extremely good -- find it in the original: Version 0
' 4. Find a section that fits your expertise. Take a crack at modifying the routines
'    to handle any unhandled chunks; there are many things to do at this point.
' 5. Suggest the name for this project & posting: PNG Reader/Writer
' 6. Add other references as needed to help others improve upon this too.
' 7. Shouldn't have to be asked, but comment your code/improvements as much as possible

'=========================================================================
' PROJECT HISTORY.... When you update & repost, add your's here too

' Version 0. Original Code: 10/5/2004, by alfred.koppold (txtCodeId=56537)
' Version 1. 31 Dec 2005, by LaVolpe, (txtCodeID= ) << fill in please
' Version ?
'=========================================================================
' REFERENCES:
'   PNG links of interest:
'       1: http://www.libpng.org/pub/png/book/toc.html
'       2: http://www.libpng.org/pub/png/spec/1.2/PNG-Contents.html
'       3: http://www.libpng.org/pub/png/png-sitemap.html
'       4: http://www.libpng.org/pub/png/pngsuite.html  <<(png test images)
'       5: http://www.winimage.com/zLibDll/zlib123dll.zip <<(VB6 friendly ZLIB.dll)
'       6: review the ReaderREADME.rtf file
'=========================================================================
' Copy & paste everything above this line to your latest revision before posting on PSC


' // LaVolpe, Dec 2005
' My contributions to the Original Code (Ver.0).
' 1. Speed Improvements, even with all the extra safety checks.
'       When compiled my routines are 2 to 6 times faster than Ver.0 compiled version
'       My routines are up to 10x faster compiled than uncompiled depending on compression/alpha
' 2. Ver.0 didn't handle zero-byte blocks & would error going past EOF
' 3. Ver.0 didn't offer progressive display. Interlacing routines were slow
'       In fact 4 progressive displays are provided:
'       a. Fade in for interlaced images with transparency. Only option for transparent interlaced images
'       b. Fade in for non-alpha interlaced
'       c. Pixelated progression for non-alpha interlaced
'       d. Scanner effect for non-interlaced images (nice effect on very large pngs)
'       e. Prevent interlacing (is this a 5th option?)
'       ^^ Progressive display while stretching an image is not available (too costly in time)
' 4. Ver.0 used pure vb decompression routines, but provided no compression; now ref's zlib dll
'       Note. Pure VB is good, but w/o compression routines; can't write png files
'             Therefore, using zlib to write & read compression & do CRC checks
'       Speed improvement: a non-interlaced 960x860 png. Ver.0=5.2 secs, w/zlib=1.8 secs
' 5. Ver.0 drew some 24bpp pngs upside down and skewed others (alpha tucan.png )
' 6. Ver.0 didn't use advanced APIs like Msimg32.dll if available
'       This version will use them for reading/displaying if possible unless overridden
'       Thus should be faster displaying when system has access to those DLLs
'       Win98 & above can use 32bpp alpha bitmaps & AlphaBlend via Msimg32.dll
'       WinNT4 & lower are screwed: per pixel manual alphablending is required
' 7. Ver.0 automatically assumed the hardware was little endian
'       Now using ws2_32.dll to convert endians if needed (some pcs are compatible with both)
' 8. Ver.0 assumed chunk names would equate to US ASCII characters when read from file.
'       Now using the numerical equivalent of the chunk names
' 9. Ver.0 did not offer CRC checks. Now, by default, critical chunks are CRC checked
'       and non-critical are optional. Using zLib to create CRC codes
' 10. Ver.0 did not handle known png errors. For example, per png specs:
'       http://www.w3.org/TR/PNG/#5ChunkOrdering
'       "When present, the tRNS chunk must precede the first IDAT chunk,
'           and must follow the PLTE chunk, if any.
' 11. Ver.0 didn't recognize all known chunk names
' 12. Tweaked areas that could be optimized. For example, the Ver.0 read the
'       8 byte signature and checked each byte (1 Read & 8 IF statements).
'       Modified to cache the signature & uses only 2 reads and 2 IF statement
' 13. Ver.0 didn't offer ability to StretchBlt, tricky when using transparency
' 14. Reworked many routines to prove to myself I understood what was happening.
'     Majority of routines are now commented and organized.

' // LaVolpe, Dec 1 thru 10
' KNOWN ISSUES WITH THIS PROJECT
' 1. Does not use all chunks when displaying an image. Should you know how to
'    use these chunks, please update project and repost or email me so I can
'    update the posting. The following chunks are not handled at this time,
'    simply cause I don't know 100% for sure how to handle them.
'    - cHRM - Primary chromaticities and white point
'    - oFFs - Image Offset, primarily used for printing; but only for printing?
'    - pCAL - Pixel Calibration
'    - sCAL - Physical Scale
'    - pHYs - Physical Pixel Dimensions
'    - sRGB - sRGB color space as defined by the International Color Consortium
'    - gAMA - relationship between the image & the desired display output intensity
'    - sBIT - Significant bits & bit depth enforcement
'    - hIST - Palette histogram
'    - iTXt - International Text Annotations (unicode)
'    - gIFg - GIF Conversion Info
'    - gIFx - GIF Conversion Info
'    - iCCP - Embedded ICC profile
'    - sPLT - Suggested Palette
'    - fRAc - Fractal Parameters
' 2. Not positive all my color routines are best; afterall, I am color blind :)
'    For 48bpp (2.8 trillion colors) images I am simply selecting the 1st of the
'    two bytes provided; not totally confident it should be handled that way.
'    Playing with small 48bpp pngs, I have seen a difference as much as 100 between
'    the two bytes; however, averaging the two requires more calculations per pixel
'    and didn't look good IMO. Like scaling down 24bpp to 8bpp, color loss would be
'    expected, so would scaling 48bpp to 24bpp. But is there a better way than
'    averaging or just picking one of the two provided bytes?
'    BTW 48bpp could be a 24bpp bitmap with 2 planes vs 1. As MSDN says the bmpinfo
'    header .bmPlanes member cannot be other than one, I am assuming scaling is required
'    and true 48bpp bitmaps with current APIs is not possible.
' 3. Per PNG specs, the width & height of an image can be Long values. This is not possible
'    for usage of API Bitmaps. The .biDataSize member of the BitmapInfoHeader is a Long
'    value & calculating (2^31-1)width * (2^31-1)height will be an overflow. To be safe,
'    the routines are limiting total width and/or height to <= 46,000 or w*h <= 2.116gb
'    Additionally, array/memory pointers are Long values, and image sizes beyond those
'    max values will cause an overflow on the pointers as they are moved along the arrays.
' 4. VB5 users: Look for Replace$() calls or other non-VB5 friendly calls & modify as needed
'           Also replace following line:
'       Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
'           with
'       Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
' 5. Requires the zlib.dll dependency. See ZLIB.DLL REQUIREMENT...
' zip file containing VB-ready DLL: http://www.winimage.com/zLibDll/zlib123dll.zip
' The above URL is for the zip; the webpage is http://www.winimage.com/zLibDll/
' IMPORTANT: Rename the dll to zlibVB.dll & place in your system32 folder; no registration required

'--------------------------------------------------------------------------------------
' TODO:
' 1. Finish parsing chunks where 100% confident can at least parse. These can be passed
'    to user via properties like the sRGB & cHRM chunks currently are
' 2. Process Unknown/Private chunks into array and provide thru a property/function
'   Although these are non-standard, they might mean something to the owner of the PNG
' 3. Prevent unauthorized chunks from being passed at all. These are UCase chunks
'    that are not IEND, IHDR or IDAT. UCase names are reserved and cannot be private
' 4. Should keep track of which chunk came first. This info can be used to determine
'   invalid chunk order and also if PNG modified and rewritten, which chunks get
'   saved in which order. There is documentation on this process in the PNG specs
' 5. Finish writing needed code to use class in Win95/NT4 & below
'    Will be slower but at least doable and will make class all O/S friendly
' 6. The PNG Writer -- get it finalized & tested.
'--------------------------------------------------------------------------------------

' zLIB (v1.2.3) calls, needed to compress/decompress png data
' ///////////// ZLIB.DLL REQUIREMENT \\\\\\\\\\\\
' not all calls are used yet and are validated via ValidateDLLExists function
Private Declare Function Zcompress Lib "zlibVB.dll" Alias "compress" (ByRef Dest As Any, ByRef destLen As Long, ByRef Source As Any, ByVal sourceLen As Long) As Long
Private Declare Function ZcompressBound Lib "zlibVB.dll" Alias "compressBound" (ByVal sourceLen As Long) As Long
Private Declare Function Zuncompress Lib "zlibVB.dll" Alias "uncompress" (ByRef Dest As Any, ByRef destLen As Long, ByRef Source As Any, ByVal sourceLen As Long) As Long
Private Declare Function Zadler32 Lib "zlibVB.dll" Alias "adler32" (ByVal adler As Long, ByRef buf As Any, ByVal Length As Long) As Long
Private Declare Function Zcrc32 Lib "zlibVB.dll" Alias "crc32" (ByVal crc As Long, ByRef buf As Any, ByVal Length As Long) As Long
Private Declare Function Zcompress2 Lib "zlibVB.dll" Alias "compress2" (ByRef Dest As Any, ByRef destLen As Long, ByRef Source As Any, ByVal sourceLen As Long, ByVal level As Long) As Long
Private Const Z_OK As Long = &H0
Private Const Z_NO_COMPRESSION As Long = 0
Private Const Z_BEST_SPEED As Long = 1
Private Const Z_BEST_COMPRESSION As Long = 9
Private Const Z_DEFAULT_COMPRESSION As Long = (-1)

' This next Event is fired after the PNG image is loaded and just before it is about
' to be processed/painted. Use this event to resize your DC if needed, change background
' colors, and change other options as needed (i.e., prevent transparencies)

' If you are simply loading the PNG file and will call this class's Paint method
' after the PNG has been fully processed, do not respond to this event. This event
' is used only for progressive displaying image while simultaneously loading it
' Limitation: Progressive display not available if image needs to be stretched
'   ^^ if this applies, Paint the image after it is completely processed
Public Event ProgessiveDisplay(ByVal Width As Long, ByVal Height As Long, _
        ByRef destinationDC As Long, ByRef destinationHwnd As Long, _
        ByVal WinBkgColor As Long, ByVal AlphaPng As Boolean, _
        ByRef X As Long, ByRef Y As Long, _
        ByVal IsInterlaced As Boolean, ByVal useProgression As Boolean)
' Width, Height are the actual dimensions of the PNG file
' WinBkgColor is the window color suggested in the PNG, -1 indicates no color suggested
' destinationDC is by reference so you can provide the DC to for progressive display
' destinationHwnd is by reference so you can provide the hWnd to refresh during display
' AlphaPng is true if the PNG uses transparency
' X,Y are DC coordinates to being progressive displaying
' IsInterlaced indicates whether encoded PNG is interlaced or not
' useProgression is suggested & is True under the 2 circumstances:
'   - Image is Interlaced
'   - Image is not interlaced & Width*Height > (Screen Width*Height)\3


' This next Event is fired while the PNG file is being processed but after it is read
Public Event Progress(ByVal Percentage As Long)
' The Percentage parameter will be a number between 0 and 100,
' where 100 occurs when the PNG file is completely processed & just before this
' class's LoadFile method returns. Very large (screen size) images can take multiple
' seconds to load and you can provide your user a progress meter. Suggest employing
' that option whenever an image more than 1/2 the total screen size is being loaded.
' However if progressive display is being used, that in itself, is a visual progress meter.


' Following ws2_32 DLL is validated via ValidateDLLExists function
' used to convert network ordered bytes to system ordered bytes (big/little endian)
Private Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
' used to convert system ordered bytes to network ordered bytes
Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long 'also used in clsBarColors
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

' change to msvbbm50.dll for VB5 projects
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
' array mapping structures
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

' following are the actual PNG image properties, exposed via class properties
' (**)Not all are translated until called from the appropriate class property
Private png_Width As Long                 ' image width
Private png_Height As Long                ' image height
Private png_BitDepth As Byte              ' image bit depth/count: 1,2,4,8,16
Private png_ColorType As Byte             ' image color type: 0,2,3,4,6
Private png_Interlacing As Byte           ' interlaced: 0,1
Private png_WinBKG As Long                ' suggested window background color
Private png_Gamma As Long                 ' gamma correction value * 10000
Private png_Palette() As Byte             ' image palette information
Private png_TransSimple() As Byte         ' image simple transparency information
Private png_Chromaticity() As Long          ' image Chromaticity information(**)
Private png_Physical() As Long            ' image aspect ratio/pixel size information(**)
Private png_Offsets() As Long             ' image physical offset information(**)
Private png_Comments() As Byte            ' unprocessed text chunks within png file(**)
Private png_LastModified() As Byte        ' date in 7-byte format(**)
Private png_sRGB As Byte                  ' standard RGB color space: 0,1,2,3

' following are class options/settings.
Private c_Overrides As Integer ' compilation of all options in 1 variable vs checking multiple variables
' 1=no progressive display, 2=proDisplay/Fade, 4=proDisplay/Pixelated
' 8=ProDisplay/Scanner, 16=ForceNoAlpha, 32=Force32Bit
' 2048=class busy/prevent access to functions
' 4096=class locked/prevent access to everything until PNG finished processing
Private c_TransColor As Long  ' translated simple transparency color (BGR or index value)
Private c_DisposableDC As Boolean ' 0=keep DC;1=destroy DC after PNG processed
Private c_ClassValidated As Boolean ' indicates if required DLLs are present
Private c_InterlaceOption As Byte ' enumPgDsply value
Public Enum enumPgDsply
    [pngPixelated] = 8
    pngAuto = 0
    '^^ Alpha Interlaced images are Fade-In, Non-Alpha interlace are pixelated,
    '   non-interlaced is dependent on size of image
    pngFadeIn = 1 ' for non-alpha interlaced pngs
    pngScanner = 2 ' applies for all non-interlaced images regardless of overall size
    pngNeverProgressive = 4 ' no progressive display ever; consider a custom progress meter
End Enum
Private c_AlphaBkg As Long      ' enumAlphaBkg value
Private c_AlphaColor As Long    ' translated color based off of c_AlphaBkg value
Public Enum enumAlphaBkg
    alphaTransparent = -1
    alphaNoBkgNoAlpha = -2
    alphaTransColorBkg = -3
    alphaPNGwindowBkg = -4
End Enum

' constants used to make reading c_Overrides easier
Private Const pdNoInterlace As Integer = 1
Private Const pdFadeIn As Integer = 2
Private Const pdPixelated As Integer = 4
Private Const pdScanner As Integer = 8
Private Const pdNoAlpha As Integer = 16
Private Const pd32bit As Integer = 32
' 64,128,256,512,1024 ' reserved as needed
Private Const pdBUSY As Integer = 2047 ' << purposely off by 1
Private Const pdLocked As Integer = 4096

' matrix/lookup tables
Private pow2x8() As Byte   ' a look up table for bit shifting (1,2,4 bit pixels)
' arrays to help process interlacing more quickly & are in relation to the 8x8 matrix
Private MatrixRow() As Byte       ' row where each pass starts within interlace matrix
Private MatrixCol() As Byte       ' column where each pass starts within interlace matrix
Private MatrixRowAdd() As Byte    ' gaps between each row withiin each pass
Private MatrixColAdd() As Byte    ' gaps between each column within each pass
' following are used for progressive/pixellated display only
Private MatrixCY() As Byte        ' height of each pixel in a scanline
Private MatrixCX() As Byte        ' width of each pixel in a scanline

' known PNG chunk names & their numerical equivalent
' Per png specs; using alpha chars is a no-no should system not support those characters
' http://www.libpng.org/pub/png/spec/1.1/PNG-Chunks.html
Private Const chnk_IHDR As Long = &H52444849 'Image header
Private Const chnk_IDAT As Long = &H54414449 'Image data
Private Const chnk_IEND As Long = &H444E4549 'End of Image
Private Const chnk_PLTE As Long = &H45544C50 'Palette
Private Const chnk_bKGD As Long = &H44474B62 'Background Color
Private Const chnk_cHRM As Long = &H4D524863 'Primary chromaticities and white point
Private Const chnk_gAMA As Long = &H414D4167 'relationship between the image & the desired display output intensity
Private Const chnk_hIST As Long = &H54534968 'Palette histogram
Private Const chnk_pHYs As Long = &H73594870 'Physical Pixel Dimensions
Private Const chnk_sBIT As Long = &H54494273 'Significant bits & bit depth enforcement
Private Const chnk_tEXt As Long = &H74584574 'Text - uncompressed
Private Const chnk_zTXt As Long = &H7458547A 'Text - compressed
Private Const chnk_tIME As Long = &H454D4974 'Timestamp
Private Const chnk_tRNS As Long = &H534E5274 'Simple Transparency
' http://www.libpng.org/pub/png/png-sitemap.html
Private Const chnk_sRGB As Long = &H42475273 'RGB color space as defined by the International Color Consortium
Private Const chnk_iCCP As Long = &H50434369 'Embedded ICC profile
Private Const chnk_sPLT As Long = &H544C5073 'Suggested Palette
Private Const chnk_iTXt As Long = &H74585469 'International Text Annotations (unicode)
Private Const chnk_oFFs As Long = &H7346466F 'Image Offset, primary used for printing
Private Const chnk_pCAL As Long = &H4C414370 'Pixel Calibration
Private Const chnk_sCAL As Long = &H4C414373 'Physical Scale
Private Const chnk_gIFg As Long = &H67464967 'GIF Conversion Info
Private Const chnk_gIFx As Long = &H78464967 'GIF Conversion Info
Private Const chnk_fRAc As Long = &H63415266 'Fractal Parameters

Private cPNGimage As stdPNGdata ' PNG|Bitmap Render class

Public Function LoadFile(ByVal Filename As String, _
            Optional ByVal ValidateAllData As Boolean = True, _
            Optional ByVal Make32Bit As Boolean = False, _
            Optional ByVal AlphaBlendValue As Byte = 255) As Boolean

    ' FileName is the full path & filename of the PNG to load/display
    ' ValidateAllData is True to validate all PNG data, else only critical data is validated
    ' Make32Bit if True will force PNG to 32bpp BMP. Needed for alphablending if desired
    ' AlphaBlendValue is from 0-255 and will blend the final result into a DC
    '   when progressively displaying PNG.

    If c_ClassValidated = False Then
        If ValidateDLLExists = False Then Exit Function
        c_ClassValidated = True
    End If

    If c_Overrides > pdBUSY Then Exit Function ' prevent recursion from the Progress event

    Const png_Signature1 As Long = 1196314761
    Const png_Signature2 As Long = 169478669
    '^^ Complete signature is 8 bytes: 137 80 78 71 13 10 26 10

    Dim ptrLoc As Long          ' used to ensure parsing doesn't go past EOF of corrupted file
    Dim Filenumber As Long      ' the file handle
    Dim gpLong As Long          ' general purpose long value
    Dim lenIDAT As Long         ' running total of the png data size (compressed)
    
    Dim ChunkName As Long       ' name of the chunk
    Dim ChunkLen As Long        ' length of the chunk
    Dim CRC32value As Long      ' CRC value of the chunk
    
    Dim ChunkBytes() As Byte    ' current chunk's bytes
    Dim RawPNGdata() As Byte    ' uncompressed png data
    Dim IDATdata() As Byte      ' compressed png data
    
    Dim cName(3) As Byte        ' for debugging purposes only

    Dim uncmprssSize As Long    ' calculated size of uncompressed PNG data
    Dim strErrorMsg As String   ' crtical error message
    Dim doValidate As Long      ' optional value to CRC check all data
    
    ' reset some stuff
    ClearMemoryVariables
    
    ' attempt to open the file with read access
    On Error Resume Next
    Filenumber = FreeFile
    Open Filename For Binary Access Read As Filenumber
    If Err Then
        MsgBox Err.Description, vbExclamation + vbOKOnly, "Error Accessing File"
        Close #Filenumber
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    ' validate we are looking at a png file
    If LOF(Filenumber) > 56 Then ' minimal (signature=8;header=13,3 rqd chunks=36 min)
        Get Filenumber, , gpLong '  ^^ 3 rqd? IHDR, IEND, & at least 1 IDAT
        If gpLong = png_Signature1 Then
            Get Filenumber, , gpLong
            If gpLong = png_Signature2 Then ptrLoc = 8  ' set pointer to current file position
        End If
    End If
    
    If ptrLoc = 0 Then ' we have an invalid PNG file
        Close #Filenumber
        MsgBox "This is not a valid PNG file", vbExclamation + vbOKOnly, "Invalid Format"
        Exit Function
    End If

    doValidate = Abs(ValidateAllData)
    ReDim IDATdata(lenIDAT)
    
    Do ' read & pre-process the png file
    
        Get Filenumber, , gpLong ' number of bytes for the chunk
        ChunkLen = ntohl(gpLong) ' longs are big endian, need little endian for Windows
        
        ' track position of pointer in the file
        ptrLoc = ptrLoc + ChunkLen + 12 ' 12 = 4byte name + 4byte CRC + 4byte chunk count
        If ptrLoc > LOF(Filenumber) Then
            ' corrupted file; abort
            Close #Filenumber
            MsgBox "This PNG file is corrupted", vbExclamation + vbOKOnly, "Corrupted File"
            Exit Function
        End If
        
        ' read chunk name & chunk data, read CRC separately
        ReDim ChunkBytes(0 To ChunkLen + 3)
        Get Filenumber, , ChunkBytes
        CopyMemory ChunkName, ChunkBytes(0), &H4 ' extract the chunk name
        Get Filenumber, , CRC32value             ' read the CRC value (big endian)
        
        If ChunkLen < 1 Then
            ' should never be <0; however can be zero at anytime
            If ChunkName = chnk_IEND Then Exit Do
            ' TODO: get this out of its own IF statement...Make routines handle it
            Debug.Print "zero length chunk: &H"; Hex(ChunkName) ' debugging
        Else
            
            ' each of the chunk parsing routines will be in a separate function.
            ' This is so that they can easily be modified without affecting any
            ' of the other code. Additionally, it is possible that chunk types will
            ' increase as PNG continues to evolve. Except IDAT, all chunks
            ' are in their own routines; even sRGB which is only a 1 byte chunk.
            Select Case ChunkName
            
            Case chnk_IDAT ' UCase chunk names are critical - CRC check
                ' compressed, filtered image data
                On Error Resume Next
                ' error? what error? all precautions taken in zChunk_IHDR function
                ' However; no predicting "Out of Memory" errors
                If ntohl(CRC32value) = Zcrc32(0, ChunkBytes(0), ChunkLen + 4) Then
                    ReDim Preserve IDATdata(lenIDAT + ChunkLen - 1)
                    CopyMemory IDATdata(lenIDAT), ChunkBytes(4), ChunkLen
                    lenIDAT = UBound(IDATdata) + 1
                Else ' error; CRC validation failed; corrupt file
                    strErrorMsg = "Corrupted Image Data (bad CRC). Cannot continue."
                    Exit Do
                End If
                If Err Then
                    strErrorMsg = "Failed to extract image data. Cannot continue."
                    Exit Do
                End If
                On Error GoTo 0
            
            Case chnk_PLTE ' UCase chunk names are critical - CRC check
                strErrorMsg = zChunk_PLTE(ChunkBytes(), ChunkLen, CRC32value)
                If Len(strErrorMsg) Then Exit Do
            
            Case chnk_bKGD ' suggested window background color
                strErrorMsg = zChunk_bKGD(ChunkBytes(), ChunkLen, CRC32value * doValidate)
                If Len(strErrorMsg) Then Exit Do
                
            Case chnk_tEXt, chnk_zTXt 'uncompressed/compressed text
                strErrorMsg = zChunk_tEXt(ChunkBytes(), ChunkLen, CRC32value * doValidate)
                If Len(strErrorMsg) Then Exit Do
                
            Case chnk_tIME ' image last modified timestamp
                strErrorMsg = zChunk_tIME(ChunkBytes(), ChunkLen, CRC32value * doValidate)
                If Len(strErrorMsg) Then Exit Do
                
            Case chnk_tRNS ' simple transparency option
                ' CRC checked 'cause if invalid, we could generate an out of bounds
                ' error in one of the other routines that reference this array
                strErrorMsg = zChunk_tRNS(ChunkBytes(), ChunkLen, CRC32value)
                If Len(strErrorMsg) Then Exit Do
                
            Case chnk_IHDR ' UCase chunk names are critical - CRC check
                strErrorMsg = zChunk_IHDR(ChunkBytes(), ChunkLen, CRC32value, uncmprssSize)
                If Len(strErrorMsg) Then Exit Do
                
            Case chnk_IEND ' UCase chunk names are critical - CRC check
                ' should CRC check for corrupted file; but why? we're at end of image
                Exit Do
            
            '//////////////////////////////////////
            '/ unhandled chunks, but valid chunks /
            '//////////////////////////////////////
            ' Following have not been programmed into
            ' PNG>BMP conversion/display if required
            
            ' all chunks below are also non-critical; therefore, if they fail to be
            ' translated/crc checked, no biggie: It may affect image quality, but
            ' won't affect whether or not the image can be displayed.
            Case chnk_pHYs
                Call zChunk_pHYs(ChunkBytes(), ChunkLen, CRC32value * doValidate)
            Case chnk_gAMA
                Call zChunk_gAMA(ChunkBytes(), ChunkLen, CRC32value * doValidate)
            Case chnk_cHRM
                Call zChunk_cHRM(ChunkBytes(), ChunkLen, CRC32value * doValidate)
            Case chnk_oFFs
                Call zChunk_oFFs(ChunkBytes(), ChunkLen, CRC32value * doValidate)
            Case chnk_sRGB
                Call zChunk_sRGB(ChunkBytes(), ChunkLen, CRC32value * doValidate)
                
                
            Case chnk_sBIT, chnk_hIST, chnk_iTXt, chnk_pCAL
                CopyMemory cName(0), ChunkName, &H4
                Debug.Print "unprocessed known chunk "; StrConv(cName, vbUnicode)
            Case chnk_iCCP, chnk_sPLT, chnk_fRAc
                CopyMemory cName(0), ChunkName, &H4
                Debug.Print "unprocessed known chunk "; StrConv(cName, vbUnicode)
            Case chnk_sCAL, chnk_gIFg, chnk_gIFx
                CopyMemory cName(0), ChunkName, &H4
                Debug.Print "unprocessed known chunk "; StrConv(cName, vbUnicode)
            
            Case Else
                ' unknown chunks, custom/user-defined chunks?
                ' ToDo: cache & make available via a property or public event
                CopyMemory cName(0), ChunkName, &H4
                Debug.Print "UNKNOWN CHUNK!!! "; StrConv(cName, vbUnicode)
            End Select
    
        End If
    
    Loop

    Close #Filenumber

    If lenIDAT = 0 Or Len(strErrorMsg) > 0 Then  ' invalid png image
        If Len(strErrorMsg) = 0 Then strErrorMsg = "No image data found. Cannot continue."
        If Err Then Err.Clear
        MsgBox strErrorMsg, vbExclamation + vbOKOnly, "Invalid PNG Format"
        ClearMemoryVariables
        Exit Function
    End If

    c_Overrides = pdBUSY + 1 ' allows reading/writing properties but no functions
    RaiseEvent Progress(0) ' inform user we started the process

    On Error Resume Next
    ' we need to uncompress our PNG file
    ReDim RawPNGdata(0 To uncmprssSize - 1)
    gpLong = Zuncompress(RawPNGdata(0), uncmprssSize, IDATdata(0), UBound(IDATdata) + 1)
    If gpLong <> 0 Then
        ' failed to uncompress & shouldn't happen 'cause if I calculated uncmprssSize
        ' wrong, then other calculations in this routine are wrong too
        ' See: CalcUncompressedWidth
        MsgBox "This PNG file is not compressed correctly.", vbOKOnly + vbExclamation, "Invalid Format"
        ClearMemoryVariables
        Exit Function
    End If

   
    ' here we will inform the user that the PNG file is ready. Since the user may
    ' choose to set optional settings that will have an effect on the type of bitmap
    ' we will use to convert the PNG to, we need that info before continuing...
    
    Dim dBitPP As Byte
    Dim tgtWindow As Long, tgtDC As Long
    Dim dcX As Long, dcY As Long
    
    If c_InterlaceOption < pngNeverProgressive Then
        ' if user didn't prohibit progressive display, send request
        RaiseEvent ProgessiveDisplay(png_Width, png_Height, tgtDC, tgtWindow, _
                png_WinBKG, (png_ColorType > 3 Or c_TransColor > -1), dcX, dcY, (png_Interlacing = 1), CalculateOverrides(0, 0, True))
    End If
    
    ' combine all known options into a single variable. Lock options into place
    CalculateOverrides tgtDC, tgtWindow, False
    If Make32Bit = True Or AlphaBlendValue < 255 Then c_Overrides = c_Overrides Or pd32bit
    
    ' tweak some settings to ensure correct bmp representations are performed
    GetDepthInfo dBitPP, 0, 0, 0 ' determine the bit count needed for our bitmap
    If dBitPP = 4 And png_BitDepth = 2 Then
        If IsArrayEmpty(Not png_Palette) Then
        ' special case here for conversion btwn 2bpp(4 color) to 4bpp(8 color)...
        ' The stdPNGdata class will create an 8-color grayscale palette for 4bpp images.
        ' But when we convert from 2bpp to 4bpp, we don't want an 8-color palette,
        ' therefore, we create the 4-color palette now, if needed:
            ' grayscale will be used; create the palette
            ReDim png_Palette(0 To 11)               ' 1st entry is 0/black
            CopyMemory png_Palette(3), 5592405, &H3  ' rgb(85,85,85)
            CopyMemory png_Palette(6), 11184810, &H3 ' rgb(170,170,170)
            CopyMemory png_Palette(9), vbWhite, &H3  ' rgb(255,255,255)
        End If
    End If
    
    Set cPNGimage = New stdPNGdata

    ' here we create the resources needed to convert the PNG file
    If cPNGimage.CreateBitmapBase(dBitPP, png_Width, png_Height, png_Palette(), tgtDC) = False Then
        MsgBox "Resources may be too low, cannot create necessary drawing objects.", vbExclamation + vbOKOnly
        ClearMemoryVariables
        Exit Function
    Else
        ' see if we need any additional resources to display the bitmap
        ' only needed if progressively displaying an image
        If (c_Overrides And pdNoInterlace) = 0 Or c_AlphaColor > -1 Then
            If (dBitPP = 32 Or png_Interlacing = 1 Or c_AlphaColor > -1) Then ' alpha and/or interlacing
                ' will be drawing to a DC, but other resources only needed in some cases
                If cPNGimage.CreateWorkSpace(tgtDC, dcX, dcY, (png_Interlacing = 1), c_AlphaColor) = False Then c_Overrides = c_Overrides Or pdNoInterlace
                ' if memory resources low & failed to create needed resources to
                ' progressively display PNG, don't progressively display
            End If
        End If
        
        ' call function to begin converting PNG to Bitmap
        If png_Interlacing = 0 Then
            gpLong = UnfilterNI(RawPNGdata(), tgtDC, tgtWindow, dcX, dcY, AlphaBlendValue)
        Else
            gpLong = UnfilterInterlaced(RawPNGdata(), tgtDC, tgtWindow, dcX, dcY, AlphaBlendValue)
        End If
    End If
    
    ' return results
    If gpLong = 0 Then
        MsgBox "Failed to properly read PNG image data.", vbExclamation + vbOKOnly, "Invalid Format"
        ClearMemoryVariables
    Else
        If c_AlphaColor > -1 Then cPNGimage.UseWorkingCopy AlphaBlendValue
        
        If (tgtDC <> 0 And tgtWindow <> 0) Then
            c_Overrides = c_Overrides And Not pdLocked
            ' couldn't create the needed workspace to progressively display PNG
            ' or user chose progressive display for non-interlaced but prevented
            ' non-interlaced images from being progressively displayed...
            ' But user is expecting progressive display. The PNG is
            ' now completely processed, so paint it & then refresh the DC
            If (c_Overrides And pdNoInterlace) = pdNoInterlace Or (c_Overrides And 15) = 0 Then
                cPNGimage.DrawToDC tgtDC, dcX, dcY, png_Width, png_Height, 0, 0, png_Width, png_Height, 0, AlphaBlendValue, tgtWindow
            End If
        End If
        LoadFile = True
        If c_DisposableDC Then cPNGimage.ManageDC = False
    End If

    c_Overrides = 0
    
End Function

Public Sub Paint(ByVal hDC As Long, ByVal destX As Long, ByVal destY As Long, _
            Optional destCx As Long, Optional destCy As Long, _
            Optional PngDC As Long, Optional pngX As Long, Optional pngY As Long, _
            Optional pngCx As Long, Optional pngCy As Long, Optional AlphaBlend As Byte = 255)

    If c_Overrides > pdBUSY Then Exit Sub ' prevent recursion from the Progress event
    ' basically a StretchBlt-type call
    ' The AlphaBlend value is ignored if the converted PNG is < 32bpp
    If Not cPNGimage Is Nothing Then
        cPNGimage.DrawToDC hDC, destX, destY, destCx, destCy, pngX, pngY, pngCx, pngCy, PngDC, AlphaBlend
    End If

End Sub

Public Property Get Handle() As Long
    ' the converted PNG's BMP handle
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        If Not cPNGimage Is Nothing Then Handle = cPNGimage.PngHandle
    End If
End Property
Public Property Get hDC() As Long
    ' the DC managed by this class. If hasOwnDC=False then no DC handle is returned
    If c_Overrides < pdLocked Then
        If Not cPNGimage Is Nothing Then hDC = cPNGimage.PngDC
    End If
End Property

Public Property Get HasOwnDC() As Boolean
    ' Determines whether or not the class will keep a DC in memory
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then HasOwnDC = Not c_DisposableDC
End Property

Public Property Let HasOwnDC(bCreateDC As Boolean)
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        c_DisposableDC = Not bCreateDC
        If Not cPNGimage Is Nothing Then cPNGimage.ManageDC = bCreateDC
    End If
End Property

Public Property Get Height() As Long
    ' the PNG's actual height & will also be the converted BMP's height
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then Height = png_Height
End Property

Public Property Get LastModified() As Date
    ' the Last Modified timestamp if PNG provided one
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        If Not IsArrayEmpty(Not png_LastModified) Then
            On Error Resume Next
            Dim iYear As Integer, lDate As Date, lTime As Date
            CopyMemory iYear, png_LastModified(0), 2
            If iYear Then
                ' don't be tempted to combine lines. Doing so is producing error:
                ' Expression too complex. This is only way I can get it to not error out
                lDate = DateSerial(iYear, png_LastModified(2), png_LastModified(3))
                lTime = TimeSerial(png_LastModified(4), png_LastModified(5), png_LastModified(6))
                LastModified = lDate + lTime
            End If
            If Err Then Err.Clear
        End If
    End If
End Property

Public Property Get TransparentStyle() As enumAlphaBkg
    ' Option to set the background color for transparent pixels
    ' Any option other than alphaTransparent destroys any transparency
    '   information within the converted BMP which will be 24bpp
    
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then TransparentStyle = c_AlphaBkg
End Property

Public Property Let TransparentStyle(BkgStyle As enumAlphaBkg)
    If c_Overrides < pdLocked Then c_AlphaBkg = BkgStyle
    ' prevent recursion from the Progress event
End Property

Public Property Get GammaCorrection() As Double
    ' returns the Gamma value if PNG provided one
    ' Note: use 1/GammaCorrection to display values like 2.2, 1.8, etc
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        On Error Resume Next
        GammaCorrection = png_Gamma / 100000
        If Err Then Err.Clear
    End If
End Property

Public Property Get Chromaticity() As Double()
    ' returns Chromaticity information if PNG provided it. The Doubles are translated as:
    ' a return array with UBound<0 means not provided
    'ArrayItem(0) is White X point
    'ArrayItem(1) is White Y point
    'ArrayItem(2) is Red X point
    'ArrayItem(3) is Red Y point
    'ArrayItem(4) is Green X point
    'ArrayItem(5) is Green Y point
    'ArrayItem(6) is Blue X point
    'ArrayItem(7) is Blue Y point
    Dim dRtn() As Double, I As Integer
    ReDim dRtn(-1 To -1)    ' default return array
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        On Error Resume Next
        If Not IsArrayEmpty(Not png_Chromaticity) Then
            ReDim dRtn(0 To 7)
            For I = 0 To 7
                dRtn(I) = png_Chromaticity(I) / 100000
            Next
        End If
        If Err Then Err.Clear
    End If
    Chromaticity = dRtn
End Property

Public Property Get AspectRatio() As Long()
    ' Returns data used for pixel size if PNG provides it. The Longs are translated as:
    'ArrayItem(0) is the X ratio
    'ArrayItem(1) is the Y ratio
    'ArrayItem(2) is either 0=Unknown Measurement, 1=Meters where 1 inch=0.0254 meters
    Dim lRtn() As Long
    ReDim lRtn(-1 To -1)
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        On Error Resume Next
        If Not IsArrayEmpty(Not png_Physical) Then lRtn() = png_Physical()
        If Err Then Err.Clear
    End If
    AspectRatio = lRtn
End Property

Public Property Get Offsets() As Long()
    ' returns information for Page/Display Offsets if PNG provides it
    ' The Longs are translated as
    ' ArrayItem(0) is the X offset
    ' ArrayItem(1) is the Y offset
    ' ArrayItem(2) is either 0=Pixels, or 1=Microns
    Dim lRtn() As Long
    ReDim lRtn(-1 To -1)
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        On Error Resume Next
        If Not IsArrayEmpty(Not png_Offsets) Then lRtn() = png_Offsets()
        If Err Then Err.Clear
    End If
    Offsets = lRtn
End Property

Public Property Get ScanLineWidth() As Long
    ' the converted BMP's DWord aligned ScanLine Width.
    ' The total byte size for the BMP is ScanLineWidth*Height
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then
       If Not cPNGimage Is Nothing Then ScanLineWidth = cPNGimage.PngDataSize \ png_Height
    End If
End Property

Public Property Get StdRGB() As Byte
    ' Returns whether or not Image should be displayed using Standard RGB (sRGB)
    ' Possible values are
    ' 0 = the information is not provided
    ' 1 = Perceptual
    ' 2 = Relative Colorimetric
    ' 3 = Saturation
    ' 4 = Absolute Colorimetric
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then StdRGB = png_sRGB
End Property

Public Property Get Width() As Long
    ' the PNG's actual width & will also be the converted BMP's width
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then Width = png_Width
End Property

Public Property Get Comments(Optional sNoCommentPrefix As String) As String
    ' Returns embedded comments if PNG provided them
    ' the Optional sNoCommentPrefix will be returned if no comments were provided
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        Comments = ExtractText
        If Comments = "" Then Comments = sNoCommentPrefix
    End If
End Property

Public Property Get Palette() As Byte()
    ' Returns the current palette used by the PNG, if any
    ' The palette is in RGB order with one byte for each R,G, & B
    ' The number of palette entries will be evenly divisible by 3
    ' A UBound return array of -1 indicates no palette is applicable
    
    Dim rtnPalette() As Byte, pIndex As Long, stepVal As Long, nrColors As Long
    ReDim rtnPalette(-1 To -1)
    
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        On Error Resume Next
        If Not IsArrayEmpty(Not png_Palette) Then
            rtnPalette() = png_Palette
        Else
            If png_ColorType = 0 Then ' grayscale with implied palette
                If png_BitDepth = 16 Then
                    nrColors = 255      ' ubound of palette entries
                Else
                    nrColors = pow2x8(png_BitDepth) - 1 ' ubound of palette entries
                End If
                ReDim rtnPalette(0 To (nrColors * 3 + 2))  ' each are R,G,B
                stepVal = 255 \ nrColors
                For pIndex = 1 To nrColors
                    ' palette entry zero is always RGB(0,0,0)
                    rtnPalette(pIndex * 3) = pIndex * stepVal
                    rtnPalette(pIndex * 3 + 1) = pIndex * stepVal
                    rtnPalette(pIndex * 3 + 2) = pIndex * stepVal
                Next
            End If
        End If
        If Err Then Err.Clear
    End If
    Palette = rtnPalette()
End Property
Public Property Let Palette(newPalette() As Byte)
' CAUTION** NOT FULLY TESTED

    ' note that a paletted 16bpp PNG is still 8 bytes per R,G,B element
    ' PNG palettes are always R,G,B triples each having 1 byte per R,G,B element
    
    ' You should check the ColorType property before calling this function....
    ' The following ColorType will have a palette
    '   ColorType of 3, no other color type will have a palette
    ' However, ColorType of 0 has an implied grayscale palette although it
    ' doesn't exist. This doesn't mean you can't override the palette by
    ' providing one here.  Doing so will have the routine tweak the ColorType
    ' to be 3 vs 0, so that the colors will be read from the palette you provided.
    ' You should ensure that the passed palette is arrayed (0 To (2^bitcount)*3-1).
    ' Keep in mind that a 16bpp image is considered 8bpp for palette entries.
    
    ' Providing a null palette will force the routine to use grayscale values
    
    If c_Overrides >= pdLocked Then Exit Property
        
    On Error Resume Next
    Dim nrColors As Long, maxEntries As Long, pIndex As Long
    
    If png_ColorType = 0 Then ' grayscale
    
        If Not IsArrayEmpty(Not newPalette) Then
            If LBound(newPalette) Then Exit Property ' must be zero bound
            
            ' user supplied a palette for a color type that
            ' isn't allowed palettes; convert to paletted PNG
            png_ColorType = 3 ' paletted
            ' to ensure no out of bounds check, palette must be maxed out
            If png_BitDepth > 4 Then
                maxEntries = 256 * 3
            Else
                maxEntries = pow2x8(png_BitDepth) * 3
            End If
            ReDim png_Palette(0 To maxEntries - 1)
            nrColors = UBound(newPalette) + 1
            If nrColors > maxEntries Then nrColors = maxEntries
            CopyMemory png_Palette(0), newPalette(0), nrColors
            
            ' to finish converting to a ColorType 3, we need to check
            ' for simple transparency.
            If Not IsArrayEmpty(Not png_TransSimple) Then
                ' the grayscale is using simple transparency
                c_TransColor = png_TransSimple(1)
                ReDim png_TransSimple(0 To maxEntries \ 3 - 1)
                FillMemory png_TransSimple(0), maxEntries \ 3, 255
                png_TransSimple(c_TransColor) = 0
            End If
        'Else no action needed
        End If
                        
    Else
    
        If png_ColorType <> 3 Then Exit Property 'only other type allowed palletes
        
        ' user is providing a modified palette for a paletted PNG
        If IsArrayEmpty(Not newPalette) Then
            ' if the image is not using transparency, we'll simply
            ' say the PNG is now grayscale, otherwise, we'll have to
            ' supply the grayscale values
            If IsArrayEmpty(Not png_TransSimple) Then
                ' no transparency used
                png_ColorType = 0
                Erase png_Palette()
            Else
                If png_BitDepth > 4 Then
                    maxEntries = 255
                Else
                    maxEntries = pow2x8(png_BitDepth) - 1
                End If
                ReDim png_Palette(0 To (maxEntries + 1) * 3 - 1)
                
                For nrColors = 0 To maxEntries
                    png_Palette(pIndex) = (255 \ maxEntries) * nrColors
                    png_Palette(pIndex + 1) = png_Palette(pIndex)
                    png_Palette(pIndex + 2) = png_Palette(pIndex)
                    pIndex = pIndex + 3
                Next
            
            End If
        Else    ' swapping palettes
            If LBound(newPalette) Then Exit Property ' must be zero bound
            
            If png_BitDepth > 4 Then
                maxEntries = 256 * 3
            Else
                maxEntries = pow2x8(png_BitDepth) * 3
            End If
            nrColors = (UBound(newPalette) + 1)
            If nrColors > maxEntries Then nrColors = maxEntries
            ReDim png_Palette(0 To maxEntries - 1)
            CopyMemory png_Palette(0), newPalette(0), nrColors
            ' no modifications of the simple transparency array is needed
        End If
    End If ' color type checks
End Property

Public Property Get IsInterlaced() As Boolean
    ' Returns whether or not the PNG image is in interlaced format
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then IsInterlaced = (png_Interlacing = 1)
End Property

Public Property Get ColorType() As Byte
    ' Returns the PNG Color Type, possible values are
    ' 0 = Grayscale, no physical palette but palette is implied at 2^bitDepth-1
    ' 2 = True Color, individual R,G,B values
    ' 3 = Paletted, pyhsical palette exists in R,G,B where nr colors in palette = (UBound(PaletteSize)+1)\3
    ' 4 = Grayscale + Alpha channel. Palette is same scenario as Color Type zero
    ' 6 = True Color + Alpha channel, individual R,G,B,A values
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then ColorType = png_ColorType
End Property
Public Property Get BitCount_PNG() As Byte
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then
        'R=red, G=green, B=Blue, P=palette index
        Select Case png_ColorType
        Case 0, 3 ' grayscale & paletted (P either 8,4,2 or 1 per Byte)
            BitCount_PNG = png_BitDepth
        Case 2 ' true color (RGB or RRGGBB)
            BitCount_PNG = png_BitDepth * 3
        Case 4 ' grayscale w/alpha channel (PA or PPAA)
            BitCount_PNG = png_BitDepth * 2
        Case 6 ' true color w/alpha channel (RGBA or RRGGBBAA)
            BitCount_PNG = png_BitDepth * 4
        End Select
    End If
End Property
Public Property Get BitCount_BMP() As Byte
    ' The bitmap created to display the PNG. User options and display options
    ' have a direct effect on the bitcount of that BMP
    
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then
        If Not cPNGimage Is Nothing Then BitCount_BMP = cPNGimage.pngDataBitCount
    End If
End Property

Public Property Let ProgressiveDisplay(DisplayType As enumPgDsply)
    ' When progressive display is used, this is the preferred defaults
    ' pngAuto: Let system determine progressive style:
    '   Interlaced Image with on alpha, then pngPixelated is used
    '   Interlaced Image with alpha, then pngFadeIn is used
    '   Non-Interlaced & W*H > (Screen W*H)\3 then pngScanner is used
    If c_Overrides < pdLocked Then ' prevent recursion from the Progress event
        If DisplayType < 7 Then c_InterlaceOption = CByte(DisplayType)
    End If
End Property
Public Property Get ProgressiveDisplay() As enumPgDsply
    ' prevent recursion from the Progress event
    If c_Overrides < pdLocked Then ProgressiveDisplay = c_InterlaceOption
End Property

Private Function CalculateOverrides(tgtDC As Long, tgtWindow As Long, bPrepDisplay As Boolean) As Boolean

' Routine combines several options into a single variable that can be referenced
' throughout the PNG processing functions
    
    If bPrepDisplay Then
        ' calculating suggestion for user whether or not to use progressive display
        If png_Interlacing = 1 Then ' image is interlaced
            CalculateOverrides = True
        Else    ' not interlaced but if > 1/3 screen, suggest it be progressively displayed
            If ((png_Width \ 100) * (png_Height \ 100)) > _
                (((Screen.Width \ Screen.TwipsPerPixelX) \ 100) * _
                    ((Screen.Height \ Screen.TwipsPerPixelY) \ 100)) \ 3 Then CalculateOverrides = True
        End If
    
    Else
        ' PNG is processed, combine user options into a single variable
        Dim tColor As Long
        c_Overrides = 0
        If tgtDC = 0 Or tgtWindow = 0 Then
            ' progressive display is not wanted for this image
            c_Overrides = pdNoInterlace
        Else ' else, it is wanted
            If png_ColorType > 3 Or c_TransColor > -1 Then ' has transparency
                If png_Interlacing = 0 Then   ' not interlaced, but user wants interlaced
                    c_Overrides = pdScanner ' only option for non-interlaced images
                Else
                    c_Overrides = pdFadeIn  ' only option for alpha interlaced images
                End If
            ElseIf png_Interlacing = 1 Then ' non-alpha, interlaced
                If (c_InterlaceOption And pngFadeIn) = 0 Then ' use default (pixelated)
                    c_Overrides = pdPixelated
                Else
                    c_Overrides = pdFadeIn ' use fadein
                End If
            Else
                c_Overrides = pdScanner ' the only option for non-inerlaced images
            End If
        End If
        
        ' flag to indicate alpha/transparency is not used at all
        If (c_TransColor < 0 And png_ColorType < 4) Then
            c_Overrides = c_Overrides Or pdNoAlpha
        Else
            ' now for alternate transparency color (user-defined)
            Select Case c_AlphaBkg
            Case alphaTransparent: ' no alternate color
            Case alphaNoBkgNoAlpha:
                c_Overrides = c_Overrides Or pdNoAlpha
            Case alphaPNGwindowBkg
                c_AlphaColor = png_WinBKG
                If c_AlphaColor < 0 Then c_AlphaColor = vbWhite
            Case alphaTransColorBkg, alphaPNGwindowBkg
                Select Case png_ColorType
                Case 3 ' paletted
                    If c_TransColor > -1 Then ' use suggest DC backcolor if provided
                        c_Overrides = c_Overrides Or pdNoAlpha
                    End If
                Case 0 ' grayscale; same general principle but a little different :)
                    If (c_TransColor > -1) Then ' use transparent color if possible
                        c_AlphaColor = GetPaletteValue(-1, c_TransColor + 0)
                    Else
                        If png_WinBKG < 0 Then
                            c_Overrides = c_Overrides Or pdNoAlpha
                        Else
                            c_AlphaColor = png_WinBKG
                        End If
                    End If
                Case 2
                    If c_TransColor > -1 Then
                        ' the transcolor is in BGR, convert to RGB
                        c_AlphaColor = ((c_TransColor And &HFF) * 65536) Or _
                            ((c_TransColor \ 256&) And &HFF) * 256& Or _
                            ((c_TransColor \ 65536) And &HFF)
                    Else
                        If png_WinBKG > -1 Then
                            c_AlphaColor = png_WinBKG
                        Else
                            c_Overrides = c_Overrides Or pdNoAlpha
                        End If
                    End If
                Case Else
                    c_Overrides = c_Overrides Or pdNoAlpha
                End Select
                ' if all else fails, use white vs black. Better IMO
                If c_AlphaColor < 0 And ((c_Overrides And pdNoAlpha) = 0) Then
                    If png_WinBKG < 0 Then
                        c_AlphaColor = vbWhite
                    Else
                        c_AlphaColor = png_WinBKG
                    End If
                End If
            Case Else
                c_AlphaColor = ConvertVBSysColor(c_AlphaBkg)
            End Select
        End If
        
        ' prevent user from calling any class methods or properties from the
        ' Progress public event....
        c_Overrides = c_Overrides Or pdLocked

    End If

End Function

Private Function CalcUncompressedWidth() As Long

    Dim uncompressedWidth As Long, iBitPP As Byte
    ' following for interlaced images
    Dim Pass As Long, passWidth As Long, passHeight As Long

    On Error GoTo NoLoad

    InitializeMatrix

    ' get the actual bits per pixel the png is using
    ' (i.e., 16bitdepth png @ ColorType 6 = 64bits per pixel)
    GetDepthInfo 0, 0, iBitPP, 0
    
    If png_Interlacing = 0 Then ' no interlacing
        ' uncompressed width will be byte aligned width + 1 for filter byte
        ' multiplied by the height
        passWidth = GetBytesPerPixel(png_Width, iBitPP)
        uncompressedWidth = passWidth * png_Height + png_Height
    Else
        ' interlaced will also be byte aligned but per scanline width
        ' Each of the 7 passes can have different widths + 1 filter byte per line
        For Pass = 1 To 7
            ' calculate number of pixels per scan line
            passWidth = png_Width \ MatrixColAdd(Pass) - (png_Width Mod MatrixColAdd(Pass) > MatrixCol(Pass))
            ' determine number of bytes needed for each scanline
            passWidth = GetBytesPerPixel(passWidth, iBitPP)
            ' calculate number of rows for this scan's pass
            passHeight = png_Height \ MatrixRowAdd(Pass) - (png_Height Mod MatrixRowAdd(Pass) > MatrixRow(Pass))
            ' now get the total bytes needed for the entire pass,
            ' adding 1 filter byte for each line in the pass:  i.e., + passHeight
            uncompressedWidth = uncompressedWidth + passWidth * passHeight + passHeight
        Next
        
    End If
    
    CalcUncompressedWidth = uncompressedWidth

NoLoad:
End Function

Private Sub InitializeMatrix()
        
    ' a quick look up table for bit shifting & palette calculations
    ReDim pow2x8(0 To 7)
    ' using copymemory vs doing the 2^n calculations 8 times
    CopyMemory pow2x8(0), 134480385, &H4
    CopyMemory pow2x8(4), -2143281136, &H4 ' Array(1,2,4,8,16,32,64,128)
    
    ' initialize interlacing matrix, used in the ConvertPNGtoBMP routine and
    ' also used to calculate the uncompressed size of the compressed PNG data
    
    ' Non-interlaced images are considered Pass#8, where interlaced images always
    ' contain 7 passes (1 thru 7).
    
    ' determines what row, in the interlaced image, the current pass begins at
    ReDim MatrixRow(1 To 8)
    CopyMemory MatrixRow(1), 262144, &H4
    CopyMemory MatrixRow(5), 65538, &H4 'Array(0, 0, 4, 0, 2, 0, 1, 0)
    ' determines what column, in the interlaced image, the current pass begins at
    ReDim MatrixCol(1 To 8)
    CopyMemory MatrixCol(1), 33555456, &H4
    CopyMemory MatrixCol(5), 256&, &H4 'Array(0, 4, 0, 2, 0, 1, 0, 0)
    ' determines the row interval of the current pass
    ReDim MatrixRowAdd(1 To 8)
    CopyMemory MatrixRowAdd(1), 67635208, &H4
    CopyMemory MatrixRowAdd(5), 16908804, &H4 'Array(8, 8, 8, 4, 4, 2, 2, 1)
    ' determines the column interval of the current pass
    ReDim MatrixColAdd(1 To 8)
    CopyMemory MatrixColAdd(1), 67373064, &H4
    CopyMemory MatrixColAdd(5), 16843266, &H4 'Array(8, 8, 4, 4, 2, 2, 1, 1)
    
    ' 1st 7 elements of next 2 arrays used for pixellated interlaced images
    
    ' determines the width of each pixellated pixel for the current pass
    ReDim MatrixCX(1 To 8)
    CopyMemory MatrixCX(1), 33817608, &H4
    CopyMemory MatrixCX(5), 16843010, &H4 'Array(8, 4, 4, 2, 2, 1, 1, 1)
    ' determines the height of each pixellated pixel for the current pass
    MatrixCY() = MatrixColAdd() 'Array(8, 8, 4, 4, 2, 2, 1, 1)

End Sub


Private Sub Class_Terminate()
    ClearMemoryVariables    ' clean up
    Set cPNGimage = Nothing
End Sub

Private Function ClearMemoryVariables()
' // LaVolpe, Dec 15
' Function simply zero's out some key variables for reuse of class & loading other files

    If Not cPNGimage Is Nothing Then
        ' no need to exectute lines if no png was previously loaded
    
        Erase png_Comments()          ' compressed text
        Erase png_Palette()           ' palette data
        Erase png_TransSimple()       ' simple transparency data
        Erase png_LastModified()      ' modified date in byte format
        Erase png_Chromaticity()        ' Chromaticity information
        Erase png_Physical()          ' aspect ratio/physical pixel size information
        Erase png_Offsets()           ' image offset information
        png_sRGB = 0                  ' use of standard RGB color space
        
        Erase pow2x8()              ' bit shifting look up table
        Erase MatrixCol()
        Erase MatrixColAdd()
        Erase MatrixCX()
        Erase MatrixCY()
        Erase MatrixRow()
        Erase MatrixRowAdd()
       
        cPNGimage.CleanUp_ALL       ' destroy any resources used/created
        
    End If
    
    c_TransColor = -1   ' cached transparent color
    png_WinBKG = -1     ' cached pNG suggested window bkg color
    c_Overrides = 0     ' class wide options
    c_AlphaColor = -1   ' cached alpha color replacement value

End Function

Private Function ConvertVBSysColor(inColor As Long) As Long
    On Error GoTo ExitRoutine
    If inColor < 0 Then
        ' vb System colors (i.e, vbButton face) are negative & contain the &HFF& value
        ConvertVBSysColor = GetSysColor(inColor And &HFF&)
    Else
        ConvertVBSysColor = inColor
    End If
    Exit Function
    
ExitRoutine:
    ConvertVBSysColor = inColor
End Function

Private Function ConvertPNGtoBMP(rawBytes() As Byte, ByVal scanPass As Byte, _
                    scanCY As Long, rBPRow As Long, startOffset As Long)
                    
' // LaVolpe, Dec 1 thru 10 - Added from Scratch
' 19 Dec -- rewrote to simplify. More & more options made the routine too disorganized
'       added the bNeedAlpha, bNeedBGR, bPackByte flags to categorize needed actions

' This routine (quite big due to various conversion issues) will convert
' PNG bits/bytes to BMP bits/bytes. Common issues to deal with include the
' byte-aligned PNG data may not be the same width as DWord-aligned BMP data. Also,
' BMPs do not support 48bpp nor 2bpp bitmaps. And the biggie: PNGs are RGB
' while DIBs are BGR; otherwise, displaying PNGs would be super quick.
                    
' rawBytes() are the png scanline bytes
' scanPass() is a valid number btwn 1-8 (8 indicates non-interlaced)
' scanCY is number of scanlines (btwn 1 and image height)
' rBPRow is the number of bytes per scanline of the rayBytes array (byte aligned)
' startOffset() is used for scanner-type progression & indicates destination row to begin with

' When no interlacing options are used, this is called once per PNG
' When non-scanner type interlacing is used, called 7 times
' When scanner type interlacing is used, called Height\16+1 times

Dim rRow As Long, rColumn As Long ' current row/column of the png image
Dim dRow As Long, dColumn As Long ' current row/column of destination bitmap
Dim dIndex As Long ' position of dRow in the destBytes() array

Dim dBitPP As Byte, dBytePP As Byte ' nr of bits/bytes per pixel of bitmap
Dim rBytePP As Byte ' nr of bytes per pixel in png image

Dim sub8ptr As Byte ' used to count bits within a byte, for 1,2,4 bpp pngs
Dim tColor As Long  ' color value when copying 3 or 4 bytes to a 3 or 4 byte array
Dim bBlockTransfer As Boolean ' special quick copy flag
Dim blockCX As Long, blockCY As Long, blkIndex As Long
'^^ counter/variables for pixelated interlacing
Dim alphaModifyer As Long, pIndex As Long ' alpha related variables

Dim destBytes() As Byte, workBytes() As Byte ' placeholders for DIB bits (DMA)
Dim TSA As SafeArray1D, TSA_w As SafeArray1D ' array overlays for DIB bits (DMA)
Dim wRowWidth As Long, dRowWidth As Long ' DWord aligned width of bitmap scanline/row

Dim bNeedBGR As Boolean, bPackByte As Boolean, bNeedAlpha As Boolean
'^^ categorized pixel actions

On Error GoTo err_h

' use direct memory access (DMA) to reference the DIB pixel data
With TSA
    .cDims = 1                          ' Number of dimensions
    .cbElements = 1                     ' Size of data elements
    .pvData = cPNGimage.PngDataPtr             ' Data address
    .rgSABound.lLbound = 0              ' Low bound of array
    .rgSABound.cElements = cPNGimage.PngDataSize ' Nr of Elements
End With
CopyMemory ByVal VarPtrArray(destBytes), VarPtr(TSA), 4

' determine the bits/bytes of the png and bitmap images
GetDepthInfo dBitPP, dBytePP, 0, rBytePP

' get location of BMP scanline we are processing from PNG scanline
If startOffset Then dRow = startOffset Else dRow = MatrixRow(scanPass)
dRowWidth = (UBound(destBytes) + 1) \ png_Height

If dBitPP < 24 Then
    ' special case when processing 1, 4 and 8 bpp, non-alpha PNGs
    '       doesn't apply when converting RGB to BGR is in effect (bpp>8) and
    '       only applies when png & bitmap bits-per-pixel are same
    ' Can apply to interlaced images too if on the last pass cause every adjacent
    ' pixel in image's scanline is provided in that pass
    bBlockTransfer = (dBitPP = png_BitDepth And scanPass > 6)
    ' ^^ note: when scanPass=8 then PNG is non-interlaced
End If

If Not bBlockTransfer Then
        
    ' modified 19 Dec to organize the pixel color transfer process for
    ' all possible user-selected transformations
        bNeedAlpha = (dBitPP = 32)
        ' ^^ we will need to know the alpha value
        bNeedBGR = (dBitPP > 8 Or ((c_Overrides And pdPixelated) = pdPixelated))
        ' ^^ we will need to know the RGB in BGR format for pixelated progressive or
        '    if the final BMP will be > 8bpp, with the following exception:
        If bNeedAlpha And png_ColorType > 3 And ((c_Overrides And pdNoAlpha) = 0) Then bNeedBGR = False
        ' ^^ Exception: When png has alpha in it, then we don't always need the BGR
        '    value 'cause it will be calculated separately against an alpha multiplyer
        bPackByte = (dBitPP < 8)
        ' ^^ we will need to modify a byte by bit shifting into it
    ' Note: the only exception not caught here is for non-interlaced, non-alpha 8 bit
    ' And that is handled just before the pixel is transfered from PNG to BMP
    
    If (c_Overrides And pdPixelated) = pdPixelated Then
        ' interlaced; non-alpha channel png using Pixelated vs FadeIn
        If dBitPP > 8 Then  ' 24bit & 32bit
            ' neat...we can reference the same memory space using different arrays.
            ' 24/32bpp interlaced images are processed with no extra overhead, but our
            ' code below is referencing the workBytes array which is normally another DIB.
            ' Instead of adding double the code for two arrays, we hack it here: Both
            ' the workBytes() and destByes() array will point to the same address space.
            TSA_w = TSA
            CopyMemory ByVal VarPtrArray(workBytes), VarPtr(TSA_w), 4
        Else
            ' non-24bit interlacing are converted to 24bit -- another DIB; so ref it
            With TSA_w
                .cDims = 1                            ' Number of dimensions
                .cbElements = 1                       ' Size of data elements
                .pvData = cPNGimage.PngDataPtr_Working       ' Data address
                .rgSABound.lLbound = 0                ' Low bound of array
                .rgSABound.cElements = cPNGimage.PngDataSize_Working ' Nr of Elements
            End With
            CopyMemory ByVal VarPtrArray(workBytes), VarPtr(TSA_w), 4
        End If
        wRowWidth = (UBound(workBytes) + 1) \ png_Height ' bytes per scanline
    End If
End If

For rRow = startOffset To scanCY - 1 ' < in relation to rawBytes() array
    
    dIndex = dRow * dRowWidth ' < destBytes array pointer for 1st pixel in scanline
    rColumn = rRow * rBPRow ' < rawBytes array pointer for 1st pixel in scanline
    
    If bBlockTransfer Then
        ' in some cases, we can just copy straight thru but still must do it
        ' row by row cause PNG byte-aligned width may not = BMP DWord aligned width
        CopyMemory destBytes(dIndex), rawBytes(rColumn), rBPRow
    
    Else    ' interlaced image
    
        dColumn = MatrixCol(scanPass) ' column in relation to the destBytes() array
        sub8ptr = 0
        Do While dColumn < png_Width
        
            If bNeedBGR Then
                ' we need color vs palette index for transfer to 24/32 bit BMP
                Select Case png_ColorType
                Case 2, 6 ' true color without(2)/with(6) alpha
                    If png_BitDepth > 8 Then ' 24,48bpp image
                        tColor = rawBytes(rColumn + 4) Or rawBytes(rColumn + 2) * 256& Or rawBytes(rColumn + 0) * 65536
                    Else
                        tColor = rawBytes(rColumn + 2) Or rawBytes(rColumn + 1) * 256& Or rawBytes(rColumn) * 65536
                    End If
                Case Else ' gray scales(0), gray scale + alpha(4) & paletted(3)
                    tColor = GetPaletteValue(sub8ptr, rawBytes(rColumn), pIndex, True)
                End Select
            End If
            
            If bNeedAlpha Then
                ' we need the alpha value for transfer to 32 bit BMP
                If (c_Overrides And pdNoAlpha) = pdNoAlpha Then ' no alpha to retrieve
                    alphaModifyer = 255 ' ignore alpha values/user-directed
                Else
                    Select Case png_ColorType
                    Case 2 ' true color + simple transparency
                        If png_BitDepth = 16 Then
                            alphaModifyer = 255
                            If tColor = c_TransColor Then ' c_TransColor is BGR vs RGB
                                ' 16bpp: all 48bits must match to be considered transparent
                                If rawBytes(rColumn + 5) Or rawBytes(rColumn + 3) * 256& Or rawBytes(rColumn + 1) * 65536 = _
                                    (png_TransSimple(4) Or (png_TransSimple(2) * 256&) Or (png_TransSimple(0) * 65536)) Then alphaModifyer = 0
                            End If
                        Else
                            If tColor <> c_TransColor Then alphaModifyer = 255
                        End If
                    Case 3 ' paletted + simple transparency in its own palette-alpha table
                        If png_TransSimple(pIndex) <> 0 Then
                            alphaModifyer = png_TransSimple(pIndex)
                            pIndex = pIndex * 3
                            tColor = ((alphaModifyer * png_Palette(pIndex + 2)) \ 255) Or ((alphaModifyer * png_Palette(pIndex + 1)) \ 255) * 256& Or ((alphaModifyer * png_Palette(pIndex)) \ 255) * 65536
                        End If
                    Case 0 ' grayscale with simple transparency
                        If png_BitDepth = 16 Then
                            alphaModifyer = 255
                            If pIndex = c_TransColor Then
                                ' 16bpp: all 16bits must match to be considered transparent
                                If (rawBytes(rColumn + 1) = png_TransSimple(0)) Then alphaModifyer = 0
                            End If
                        Else
                            If pIndex <> c_TransColor Then alphaModifyer = 255
                        End If
                    Case 4  ' grayscale with alpha
                        If png_BitDepth = 16 Then alphaModifyer = rawBytes(rColumn + 2) Else alphaModifyer = rawBytes(rColumn + 1)
                        If alphaModifyer Then _
                            tColor = ((alphaModifyer * rawBytes(rColumn)) \ 255) Or ((alphaModifyer * rawBytes(rColumn)) \ 255) * 256& Or ((alphaModifyer * rawBytes(rColumn)) \ 255) * 65536
                    Case 6 ' true color with alpha
                        If png_BitDepth = 16 Then
                            alphaModifyer = rawBytes(rColumn + 6)
                            If alphaModifyer Then _
                                tColor = ((alphaModifyer * rawBytes(rColumn + 4)) \ 255) Or ((alphaModifyer * rawBytes(rColumn + 2)) \ 255) * 256& Or ((alphaModifyer * rawBytes(rColumn)) \ 255) * 65536
                        Else
                            alphaModifyer = rawBytes(rColumn + 3)
                            If alphaModifyer Then _
                                tColor = ((alphaModifyer * rawBytes(rColumn + 2)) \ 255) Or ((alphaModifyer * rawBytes(rColumn + 1)) \ 255) * 256& Or ((alphaModifyer * rawBytes(rColumn)) \ 255) * 65536
                        End If
                    End Select
                End If
            End If
            
            ' write the result to the BMP
            If bPackByte Then
                ' we will be writing to a 1 or 4 bpp BMP
                If png_BitDepth = 1 Then
                    PackByte rawBytes(rColumn), destBytes(dIndex + dColumn \ 8), sub8ptr, dColumn Mod 8
                Else ' 2,4 bpp :: 2bpp pngs converted to 4bpp bitmaps
                    PackByte rawBytes(rColumn), destBytes(dIndex + dColumn \ 2), sub8ptr, dColumn Mod 2
                End If
            ElseIf bNeedAlpha Then
                ' we will be writing to a 32 bpp BMP
                If alphaModifyer Then ' only modify DIB bits if needed
                    destBytes(dIndex + dColumn * 4 + 3) = alphaModifyer
                    CopyMemory destBytes(dIndex + dColumn * 4), tColor, &H3
                    alphaModifyer = 0
                End If
            Else
                ' we will be writing to a 8 or 24 bit BMP
                If bNeedBGR And dBitPP > 8 Then
                    ' 24 bit BMP, dBytePP should always be 3
                    CopyMemory destBytes(dIndex + dColumn * dBytePP), tColor, dBytePP
                Else                ' 8 bit bmp
                    ' note that in this case, the variable tColor is never calculated above
                    destBytes(dIndex + dColumn * dBytePP) = rawBytes(rColumn)
                End If
            End If
            
            ' ensure our source byte pointer is moved along appropriately
            If png_BitDepth < 8 Then
                sub8ptr = sub8ptr + 1            ' keep track of how many bits processed
                If sub8ptr * png_BitDepth = 8 Then ' when the byte has been processed
                    rColumn = rColumn + 1        ' increment the PNG pointer
                    sub8ptr = 0                  ' reset our counter
                End If
            Else
                rColumn = rColumn + rBytePP      ' else increment by source byte pp
            End If
            
            ' this is a check to see if image is interlaced w/Pixelated option
            If (c_Overrides And pdPixelated) = pdPixelated Then
                If scanPass < 7 Then
                    ' when interlacing & non-alpha, we use an offscreen 24/32bit DIB
                    ' so we can quickly modify pixelated blocks. Packing bytes
                    ' for 1,2,4 bit DIBs is painful. Pixelated blocks are not 1x1,
                    ' get the width & height of this scanline's pixelated data
                    blockCX = Min(dColumn + MatrixCX(scanPass), png_Width)
                    blockCY = Min(dRow + MatrixCY(scanPass), png_Height)
                    
                    If dBytePP = 4 Then
                        pIndex = 4
                        ' when alphablending option is used, then dBytePP will be 4.
                        ' Pixelated is also only used for non-alpha images, therefore,
                        ' pixel alphaValue will always be 255. OR that value to the color.
                        tColor = tColor Or &HFF000000
                    Else
                        pIndex = 3 ' if not 32bpp then it will be 24bpp
                    End If
                    
                    ' loop horizontally btwn 1-8 pixels and fill with same color
                    For blkIndex = (wRowWidth * dRow) + dColumn * pIndex To (wRowWidth * dRow) + blockCX * pIndex - 1 Step pIndex
                        CopyMemory workBytes(blkIndex), tColor, pIndex
                    Next
                    ' now loop vertically, 1-7 rows & fill with the previous row
                    For blkIndex = dRow + 1 To blockCY - 1
                        CopyMemory workBytes((blkIndex * wRowWidth) + (dColumn * pIndex)), workBytes((wRowWidth * dRow) + (dColumn * pIndex)), (blockCX - dColumn) * pIndex
                    Next
                End If
            End If
            
            dColumn = dColumn + MatrixColAdd(scanPass)
        Loop
    End If  ' check for quick copy
    
    dRow = dRow + MatrixRowAdd(scanPass)
Next

' clean up & return result
CopyMemory ByVal VarPtrArray(destBytes), 0&, 4
If TSA_w.cDims Then CopyMemory ByVal VarPtrArray(workBytes), 0&, 4
ConvertPNGtoBMP = True

Exit Function

err_h:
Err.Clear
'Stop       ' un rem when debugging improper display
'Resume
If TSA.cDims Then CopyMemory ByVal VarPtrArray(destBytes), 0&, 4
If TSA_w.cDims Then CopyMemory ByVal VarPtrArray(workBytes), 0&, 4
End Function

Private Function ExtractText() As String
' // LaVolpe, Dec 1 thru 10 - Added; consolidated into one function
' Function extracts all the text chunks found within the PNG

'http://www.w3.org/TR/PNG/#11tEXt
'http://www.w3.org/TR/PNG/#11zTXt
'http://www.w3.org/TR/PNG/#11iTXt

    ' see if any text items were found?
    If IsArrayEmpty(Not png_Comments) Then Exit Function
        
    Dim Ptr As Long, ChunkName As Long, chunkSize As Long, textLen As Long
    Dim sText As String, sKeyWord As String
    Dim sBytes() As Byte
    Dim deflateSize As Long, lRtn As Long
    
    On Error GoTo EndTextData ' in case of garbled or corrupted data
    
    Do
    
        ' extract the text chunk size & chunk name
        CopyMemory chunkSize, png_Comments(Ptr), &H4
        CopyMemory ChunkName, png_Comments(Ptr + 4), &H4
        Ptr = Ptr + 8   ' adjust pointer
    
        Select Case ChunkName
        
        Case chnk_tEXt  ' uncompressed text format
            'Keyword:        1-79 bytes (character string)
            'Null separator: 1 byte
            'Text:           n bytes (character string)
            
            sKeyWord = GetKeyWord(Ptr)  ' get the keyword from the chunk
            If Len(sKeyWord) Then
            
                If Len(sText) Then sText = sText & vbCrLf

                ' determine where in the chunk the actual text begins skipping the null separator
                textLen = chunkSize - Len(sKeyWord) - 1 ' cache calculation
                
                ' build byte array to use strConv on
                ReDim sBytes(0 To textLen - 1)
                CopyMemory sBytes(0), png_Comments(Ptr + Len(sKeyWord) + 1), textLen
                ' per png specs; don't allow ESC to be returned
                sText = sText & sKeyWord & ": " & Replace$(StrConv(sBytes(), vbUnicode), vbKeyEscape, " ")
                
            End If
            
        Case chnk_zTXt ' compressed text format
            'Keyword:            1-79 bytes (character string)
            'Null separator:     1 byte
            'Compression method: 1 byte
            'Compressed Text:    n bytes
        
            sKeyWord = GetKeyWord(Ptr)  ' get the keyword from the chunk
            If Len(sKeyWord) Then
            
                textLen = Len(sKeyWord) + Ptr + 2 ' position in array where text starts
                ' if the compression byte isn't zero, we will abort
                If png_Comments(textLen - 1) = 0 Then
                
                    On Error Resume Next
                    If Len(sText) Then sText = sText & vbCrLf
                
                    ' no way of knowing what the uncompressed size really is. Guesstimate
                    deflateSize = chunkSize + chunkSize
                    Do
                        ' oversize array & try to uncompress
                        ReDim sBytes(0 To deflateSize)
                        lRtn = Zuncompress(sBytes(0), deflateSize, png_Comments(textLen), chunkSize)
                        
                        If lRtn = 0 Then
                            ' done, the deflateSize is the actual decompressed length
                            ReDim Preserve sBytes(0 To deflateSize - 1)
                            ' per png specs; don't allow ESC to be returned
                            sText = sText & sKeyWord & ": " & Replace$(StrConv(sBytes, vbUnicode), vbKeyEscape, " ")
                            Exit Do
                        
                        Else ' didn't uncompress; try again?
                            
                            ' didn't supply a big enough buffer; increment it & try again
                            ' Note: theoretically, this could go to 200+MB.
                            ' Use commonsense assuming compression will always be < 500%
                            deflateSize = deflateSize + chunkSize * 2
                            If deflateSize > chunkSize * 50 Then Exit Do ' is this enough?
                            If Err Then
                                Err.Clear
                                Exit Do
                            End If
                        End If
                    Loop
                End If  ' valid compression byte check
            End If ' valid keyword check
            
        
        Case chnk_iTXt  ' to do; unicode
        
        End Select
        
        Ptr = Ptr + chunkSize
        Erase sBytes
        
        ' ptr will be a position within the png_Comments array
    Loop Until Ptr > UBound(png_Comments)
        
        
EndTextData:
    ' do some simple formatting for VB purposes:
    sText = Replace$(sText, vbCrLf, vbLf)
    sText = Replace$(sText, vbLf, vbCrLf)
    ExtractText = sText
    
End Function

Private Function GetBytesPerPixel(totalWidth As Long, btsPerPixel As Byte) As Long
' // LaVolpe, Dec 1 thru 10
' returns number of bytes required to display n pixels at p color depth (byte aligned)
    GetBytesPerPixel = (totalWidth * btsPerPixel + 7) \ 8

End Function

Private Sub GetDepthInfo(destBitPP As Byte, destBytePP As Byte, _
                rawBitsPP As Byte, rawBytesPP As Byte)
    
' returns the bits per pixel & bytes per pixel for the destination bitmap
' and also the respective values for the png image
    
' PNG > DIB bmp (per pixel) conversion chart I use throughout the routines:
'Color Type     bit depth   PNG bits/bytes per pixel  BMP bits/bytes pp (ignore alpha)
'----------     ---------   ------------------------- --------------------------------
'0 gray scale   1                   1   1                   1   1   (1  1)
'               2                   2   1                   4   1   (4  1)
'               4                   4   1                   4   1   (4  1)
'               8                   8   1                   8   1   (8  1)
'               16                  16  2                   8   1   (8  1)

'2 true color   8                   24  3                   24  3   (24 3)
'               16                  48  6                   24  3   (24 3)

'3 palette      1                   1   1                   1   1   (1  1)
'               2                   2   1                   4   1   (4  1)
'               4                   4   1                   4   1   (4  1)
'               8                   8   1                   8   1   (8  1)
'               16                  16  2                   8   1   (8  1)

'4 gray+alpha   8                   16  2                   32  4   (24 3)
'               16                  32  4                   32  4   (24 3)

'6 true+alpha   8                   32  4                   32  4   (24 3)
'               16                  64  8                   32  4   (24 3)

'any bit depth that uses simple transparency (trns chunk)   32  4   (n/a)
'any bit depth where AlphaBlend option is used              32  4   (n/a)

' Special cases. When user wants to overwrite transparency with a
'   non-transparent color (setting the TransparentStyle to anything other than
'   alphaTransparent), then the final bitmap is converted from 32bpp Alpha
'   to 24bpp non-alpha and transparency is lost. If transparency is desired,
'   and also a different bkg color is desired, you should load PNG and then
'   Paint it to an offscreen DC which contains a bitmap of the color wanted.
'--------------------------------------------------------------------------

    Select Case png_ColorType
    Case 6 ' true color w/alpha (only 8,16 bpp pngs)
        If (c_Overrides And pdNoAlpha) = pdNoAlpha Then
            destBitPP = 24: destBytePP = 3 ' scale down if no alpha is being used
        Else
            destBitPP = 32: destBytePP = 4
        End If
        rawBytesPP = 4 * (png_BitDepth \ 8): rawBitsPP = png_BitDepth * 4
    Case 4: ' grayscale w/alpha (only 8,16 bpp pngs)
        If (c_Overrides And pdNoAlpha) = pdNoAlpha Then
            destBitPP = 24: destBytePP = 3 ' scale down if no alpha is being used
        Else
            destBitPP = 32: destBytePP = 4
        End If
        rawBytesPP = 2 * (png_BitDepth \ 8): rawBitsPP = png_BitDepth * 2
    Case 2: ' true color (rgb triples) (8,16 bpp pngs)
        destBitPP = 24: destBytePP = 3
        rawBytesPP = 3 * (png_BitDepth \ 8): rawBitsPP = png_BitDepth * 3
    Case 0 ' grayscale images (all bit depths)
        If png_BitDepth = 2 Then ' special case as MS bitmaps don't do 2bpp
            destBitPP = 4: destBytePP = 1: rawBytesPP = 1: rawBitsPP = 2
        ElseIf png_BitDepth > 4 Then ' (8,16 bpp pngs)
            destBitPP = 8: destBytePP = 1
            rawBytesPP = png_BitDepth \ 8: rawBitsPP = png_BitDepth
        Else ' (1,4 bpp pngs)
            destBitPP = png_BitDepth: destBytePP = 1
            rawBytesPP = 1: rawBitsPP = png_BitDepth
        End If
    Case 3: ' palette entries (1,2,4,8 bpp pngs)
        If png_BitDepth = 2 Then destBitPP = 4 Else destBitPP = png_BitDepth
        destBytePP = 1: rawBytesPP = 1: rawBitsPP = png_BitDepth
    End Select
    
    ' tweak for transparency which doesn't affect the actual png bitcounts:
    If c_TransColor > -1 And (c_Overrides And pdNoAlpha) = 0 Then
        ' If any transparency is used unless alpha is being prohibited
        destBitPP = 32: destBytePP = 4
    ElseIf ((c_Overrides And pd32bit) = pd32bit) Then
        destBitPP = 32: destBytePP = 4 ' force if user wants it forced
    ElseIf ((c_Overrides And pdFadeIn) = pdFadeIn) And png_Interlacing = 1 Then
        ' when fade-in interlaced, use transparency
        destBitPP = 32: destBytePP = 4
    End If
    
End Sub

Private Function GetKeyWord(fromPtr As Long) As String
    
' // LaVolpe, Dec 1 thru 10 - Added, returns keyword from tEXt or zTXt
    
Dim zTxtStart As Long
Dim txtData() As Byte
    
For zTxtStart = fromPtr To 80 + fromPtr
    ' per specs, the text name must not have chr$(0) in it
    ' look for the chr$(0) which would be the end of the text keyword
    If png_Comments(zTxtStart) = 0 Then ' null character
        ' copy the keyword into tmp array & convert from binary to ascii
        ReDim txtData(0 To zTxtStart - fromPtr - 1)
        CopyMemory txtData(0), png_Comments(fromPtr), zTxtStart - fromPtr
        GetKeyWord = StrConv(txtData, vbUnicode)
        Exit For
    End If
Next
    
End Function

Private Function GetPaletteValue(ByVal PixelPos As Long, PixelValue As Byte, _
                Optional RtnIndex As Long, Optional SwapRGB As Boolean) As Long
' // LaVolpe, Dec 1 thru 10 - added from scratch
' with progressive display and alphablending is it is necessary to be able to read
' any pixel regardless of the format/color depth. I am not 100% sure that these
' formulas are exact. Feel free to correct any mistakes I made

' Note: the bytes when received here will be in left to right order
' In other words, position 0 will be bit 7
' A lot of bit shifting going on here

    Dim Offset As Byte
    Dim rgbDat(0 To 2) As Byte
    Dim Color As Long

    Select Case png_BitDepth
    Case 1: ' 2 color palette (each bit represents entry in the palette or 8 pixels per byte)
        ' passed PixelPos is the 0-based position in the byte passed
        ' passed PixelValue is the value returned from the png stream
        
        ' convert the PixelPosition to a bit which would be the palette index
        If PixelPos < 0 Then    ' pixelvalue is the index
            RtnIndex = PixelValue
        Else
            Offset = (PixelPos + 1) Mod 8
            If Offset Then Offset = 8 - Offset
            RtnIndex = (PixelValue \ pow2x8(Offset)) And &H1
        End If
        If png_ColorType = 0 Then  ' grayscale
            Color = vbWhite * RtnIndex  ' b&w: 0=black, 1=white
            ' swapping colors if requested not necessary as Black/White reversed is still Black/White
        Else
            ' 2 color palette: 0=1st color, 1=2nd color
            Color = RtnIndex * 3
            If SwapRGB Then
                Color = png_Palette(Color + 2) Or (png_Palette(Color + 1) * 256&) Or (png_Palette(Color) * 65536)
            Else
                CopyMemory Color, png_Palette(Color), &H3
            End If
        End If
        
    Case 2: ' 4 pixels per byte, 4 color palette
        ' same logic as above, except 2 bits = 1 pixel
        ' Possible values will be 0-3
        If PixelPos < 0 Then ' pixelvalue is the index
            RtnIndex = PixelValue
        Else
            Offset = (PixelPos + 1) Mod 4
            If Offset Then Offset = 4 - Offset
            RtnIndex = (PixelValue \ pow2x8(Offset * 2)) And &H3
        End If
        If png_ColorType = 3 Then ' paletted
            Color = RtnIndex * 3
            If SwapRGB Then
                Color = png_Palette(Color + 2) Or (png_Palette(Color + 1) * 256&) Or (png_Palette(Color) * 65536)
            Else
                CopyMemory Color, png_Palette(Color), &H3
            End If
        Else ' grayscale; no palette provided
            ' swapping grayscale if requested is unnecessary
            RtnIndex = RtnIndex * 85
            Color = RtnIndex + RtnIndex * 256& + RtnIndex * 65536
        End If
        
    Case 4: ' 2 pixels per byte, 16 color palette
        If PixelPos < 0 Then ' pixelvalue is the index
            RtnIndex = PixelValue
        Else
            If PixelPos = 0 Then RtnIndex = (PixelValue And 240) \ 16 Else RtnIndex = PixelValue And 15
        End If
        If png_ColorType = 3 Then ' paletted
            Color = RtnIndex * 3
            If SwapRGB Then
                Color = png_Palette(Color + 2) Or (png_Palette(Color + 1) * 256&) Or (png_Palette(Color) * 65536)
            Else
                CopyMemory Color, png_Palette(Color), &H3
            End If
        Else ' grayscale; no palette provided
            ' swapping grayscale if requested is unnecessary
            Color = RtnIndex * 17
            Color = Color Or Color * 256& Or Color * 65536
        End If
        
    Case 8, 16: ' 1 pixel per byte, 256 color palette
        RtnIndex = PixelValue
        If png_ColorType = 3 Then 'paletted
            Color = RtnIndex * 3
            If SwapRGB Then
                Color = png_Palette(Color + 2) Or (png_Palette(Color + 1) * 256&) Or (png_Palette(Color) * 65536)
            Else
                Color = png_Palette(Color) + png_Palette(Color + 1) * 256& + png_Palette(Color + 2) * 65536
            End If
        Else 'grayscale; no palette provided
            ' swapping grayscale if requested is unnecessary
            Color = PixelValue + PixelValue * 256& + PixelValue * 65536
        End If
    
    End Select
    
    GetPaletteValue = Color
    
End Function

Private Function IsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
' // LaVolpe, Dec 1 thru 10
  ' test to see if an array has been initialized
  IsArrayEmpty = (lArrayPointer = -1)
End Function

Private Function MakeAlphaPixel(Color As Long, X As Long, Y As Long, alphaByte As Long) As Long

' // LaVolpe, Dec 1 thru 10
' Under construction::: This may be a helper function to use for Win95
' Still need to test stretching an image and getting appropriate alphablends
'       without using the AlphaBlend API

' This routine is only used when specifically selected or the operating system
' cannot use GDIplus or AlphaBlend. Slowest method & primarily for (Win95/NT4<)
' http://en.wikipedia.org/wiki/History_of_Microsoft_Windows

Dim alphaColor As Long
Dim srcColor(0 To 2) As Byte
Dim pngColor(0 To 2) As Byte

Select Case alphaByte
Case 255: alphaColor = Color   ' fully opaque
'Case 0: alphaColor = GetPixel(m_PicBox.hDC, X, Y) ' fully transparent
Case Else
    'CopyMemory srcColor(0), GetPixel(m_PicBox.hDC, X, Y), &H3
    CopyMemory pngColor(0), Color, &H3
    srcColor(0) = (alphaByte * (0 + pngColor(0) - srcColor(0))) \ 255 + srcColor(0)
    srcColor(1) = (alphaByte * (0 + pngColor(1) - srcColor(1))) \ 255 + srcColor(1)
    srcColor(2) = (alphaByte * (0 + pngColor(2) - srcColor(2))) \ 255 + srcColor(2)
    CopyMemory alphaColor, srcColor(0), &H3
End Select

MakeAlphaPixel = alphaColor
End Function

Private Function Min(valA As Long, ValB As Long) As Long
' // LaVolpe, Dec 1 thru 10 - helper function for progressive display
    If valA < ValB Then
        Min = valA
    Else
        Min = ValB
    End If
End Function

Private Sub PackByte(srcByte As Byte, DestByte As Byte, ByVal byteOffsetSrc As Long, ByVal destOffset As Byte)

' // LaVolpe, Dec 1 thru 10 - added from scratch
' Function modifies a 1, 2 or 4 bit per pixel byte

Dim bitPtr As Byte  ' value of the bit or bits at the specified position

Select Case png_BitDepth
Case 1 ' 8 bits per pixel, 8 pixels per byte
    ' we need to value of the bit (either 0 or 1) from the passed source byte
    ' the source positions will be inverted
    byteOffsetSrc = (byteOffsetSrc + 1) Mod 8
    If byteOffsetSrc Then byteOffsetSrc = 8 - byteOffsetSrc
    bitPtr = (srcByte \ pow2x8(byteOffsetSrc)) And &H1
    ' get the position in the destination byte where the bit will be placed
    destOffset = (destOffset + 1) Mod 8
    ' the bitmap's rgb values are also inverted BGR vs RGB
    If destOffset Then destOffset = 8 - destOffset
    ' place the bit at that position
    DestByte = DestByte Or (bitPtr * pow2x8(destOffset))
    
Case 4 ' 4 bits per pixel, 2 pixels per byte
    ' same logic as above except the value we want is contained in 4 bits
    If byteOffsetSrc Then
        bitPtr = srcByte And 15
    Else
        bitPtr = (srcByte And 240) \ 16
    End If
    destOffset = destOffset Xor 1
    DestByte = DestByte Or (bitPtr * pow2x8(destOffset * 4))

Case 2 ' the BITMAPINFO needed for CreateDIBitmap or CreateDibSection does
       ' appears to not support 2 bits per pixel formats. Therefore, we will
       ' will convert a 2bpp png image to a 4bpp image. Not too difficult :)
    byteOffsetSrc = (byteOffsetSrc + 1) Mod 4
    If byteOffsetSrc Then byteOffsetSrc = 4 - byteOffsetSrc
    bitPtr = ((srcByte \ (pow2x8(byteOffsetSrc * 2))) And &H3)
    destOffset = destOffset Xor 1
    DestByte = DestByte Or (bitPtr * pow2x8(destOffset * 4))
    


Case Else
    

End Select

End Sub

Private Function PaethPredictor(Left As Integer, Above As Integer, UpperLeft As Integer) As Integer

' // LaVolpe, Dec 1 thru 10 - rewrote for understanding & commented/linked

' http://www.w3.org/TR/PNG/#9-table91
' algorithm is used for both encoding & decoding the png image's filter
' based off of the formula created by Alan W. Paeth & provided fully in url above


    Dim pa As Integer, pB As Integer, pC As Integer, p As Integer
    p = Left + Above - UpperLeft
    pa = Abs(p - Left)
    pB = Abs(p - Above)
    pC = Abs(p - UpperLeft)
    
    ' tie breaker
    ' The order in which the comparisons are performed is critical and shall not be altered
    If (pa <= pB) And (pa <= pC) Then
     PaethPredictor = Left
    ElseIf pB <= pC Then
     PaethPredictor = Above
    Else
     PaethPredictor = UpperLeft
    End If

End Function

Private Sub DecodeFilter_Avg(Filtered() As Byte, UnFiltered() As Byte, RowNr As Long, ScanLine As Long, stepVal As Byte, Optional iOffset As Long)

' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
'http://www.w3.org/TR/PNG/#9-table91
'Filters may use the original values of the following bytes to generate the new byte value:
'
'x  the byte being filtered;
'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
'b  the byte corresponding to x in the previous scanline;
'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).

' algorithm: Recon(x) = Filt(x) + floor((Recon(a) + Recon(b)) / 2)
' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes

    Dim ppLeft As Integer, ppTop As Integer
    Dim X As Long, startByte As Long
    
    startByte = RowNr * ScanLine - RowNr - 1
    
    For X = startByte + 1 To startByte + ScanLine - 1
    
        If RowNr Then ppTop = UnFiltered(X - ScanLine + 1)
        If X - stepVal > startByte Then ppLeft = UnFiltered(X - stepVal)
        
        UnFiltered(X) = (0 + Filtered(X + RowNr + 1 + iOffset) + (ppLeft + ppTop) \ 2) Mod 256
        
    Next

End Sub

Private Sub DecodeFilter_Paeth(Filtered() As Byte, UnFiltered() As Byte, RowNr As Long, ScanLine As Long, stepVal As Byte, Optional iOffset As Long)

' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
'http://www.w3.org/TR/PNG/#9-table91

'Filters may use the original values of the following bytes to generate the new byte value:
'
'x  the byte being filtered;
'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
'b  the byte corresponding to x in the previous scanline;
'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).

' algorithm: Recon(x) = Filt(x) + PaethPredictor(Recon(a), Recon(b), Recon(c))
' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes

    Dim ppLeft As Integer, ppTop As Integer, ppTopLeft As Integer
    Dim X As Long, startByte As Long
    
    startByte = RowNr * ScanLine - RowNr - 1
    
    For X = startByte + 1 To startByte + ScanLine - 1
        
        If X - stepVal > startByte Then ' we are not on the 1st pixel
            ppLeft = UnFiltered(X - stepVal)
            If RowNr Then
                ppTop = UnFiltered(X - ScanLine + 1)
                ppTopLeft = UnFiltered(X - ScanLine + 1 - stepVal)
            End If
        Else
            If RowNr Then ppTop = UnFiltered(X - ScanLine + 1)
        End If
        
        UnFiltered(X) = (Filtered(X + RowNr + 1 + iOffset) + PaethPredictor(ppLeft, ppTop, ppTopLeft)) Mod 256

    Next

End Sub

Private Sub DecodeFilter_Sub(Filtered() As Byte, UnFiltered() As Byte, RowNr As Long, ScanLine As Long, stepVal As Byte, Optional iOffset As Long)

' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
'http://www.w3.org/TR/PNG/#9-table91

'Filters may use the original values of the following bytes to generate the new byte value:
'
'x  the byte being filtered;
'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
'b  the byte corresponding to x in the previous scanline;
'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).

' algorithm: Recon(x) = Filt(x) + Recon(a)
' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes

    Dim startByte As Long
    Dim n As Long, X As Long
    
    startByte = RowNr * ScanLine - RowNr
    ' 1st n bytes for 1st pixel are unfiltered
    CopyMemory UnFiltered(startByte), Filtered(startByte + RowNr + 1 + iOffset), stepVal
    
    For n = startByte + stepVal To startByte + ScanLine - 2 Step stepVal
        For X = n To n + stepVal - 1
            UnFiltered(X) = (0 + UnFiltered(X - stepVal) + Filtered(X + RowNr + 1 + iOffset)) Mod 256
        Next
    Next
    
End Sub

Private Sub DecodeFilter_Up(Filtered() As Byte, UnFiltered() As Byte, RowNr As Long, ScanLine As Long, Optional iOffset As Long)

' // LaVolpe, Dec 1 thru 10 - rewrote for understanding; nothing wrong with original
'http://www.w3.org/TR/PNG/#9-table91

'Filters may use the original values of the following bytes to generate the new byte value:
'
'x  the byte being filtered;
'a  the byte corresponding to x in the pixel immediately before the pixel containing x (or the byte immediately before x, when the bit depth is less than 8);
'b  the byte corresponding to x in the previous scanline;
'c  the byte corresponding to b in the pixel immediately before the pixel containing b (or the byte immediately before b, when the bit depth is less than 8).

' algorithm:  Recon(x) = Filt(x) + Recon(b)
' Unsigned arithmetic modulo 256 is used, so that both the inputs and outputs fit into bytes
    
    Dim startByte As Long
    Dim X As Long, ppTop As Integer
    
    startByte = RowNr * ScanLine - RowNr
    
    For X = startByte To startByte + ScanLine - 2
        If RowNr Then ppTop = UnFiltered(X - ScanLine + 1)
        UnFiltered(X) = (ppTop + Filtered(X + RowNr + 1 + iOffset)) Mod 256
    Next

End Sub

Private Function UnfilterInterlaced(Filtered() As Byte, _
            tgtDC As Long, tgtWindow As Long, dcX As Long, dcY As Long, AlphaBlend As Byte) As Boolean

' // LaVolpe, Dec 1 thru 10 - built from scratch

' ToDo: this routine assumes pixels are 1x1 squares. With a 6:9 aspect ratio,
'   the routine will not scale. This should be adjusted when the routines are
'   fully capable of supporting aspect ratios of other than 1:1 or should it be stretched?

' http://www.libpng.org/pub/png/spec/1.2/PNG-DataRep.html#DR.Interlaced-data-order

' Progressive display/scan order per 8 pixel blocks (64 total pixels)
'   1 6 4 6 2 6 4 6     ' 1st scan: 1 pixel (@col 0), row 0 [1/64 of total image]
'   7 7 7 7 7 7 7 7     ' 2nd scan: 1 pixel (@col 4), row 0 [1/32 of image shown]
'   5 6 5 6 5 6 5 6     ' 3rd scan: 2 pixels (@cols 0:4), row 4 [1/16 of image]
'   7 7 7 7 7 7 7 7     ' 4th scan: 4 pixels (@cols 2:6), rows 0:4 [1/8]
'   3 6 4 6 3 6 4 6     ' 5th scan: 8 pixels (@cols 0:2:4:6), rows 2:6 [1/4]
'   7 7 7 7 7 7 7 7     ' 6th scan: 16 pixels (@cols 1:3:5:7), rows 0:2:4:6 [1/2]
'   5 6 5 6 5 6 5 6     ' 7th scan: 32 pixels (@cols all), rows 1:3:5:7 [100%]
'   7 7 7 7 7 7 7 7                 64 pixels, 15 scanlines over 7 passes

' Note : all logic in this routine is based off of the above grid.
' Scanline widths are only guaranteed to be same for each scanline in the same pass.
' Scanlines can be padded both horizontally & vertically if the image doesn't fit into
'   a nice 8x8 grid evenly.
' Each scanline in interlaced image is also filtered, but they are filtered in relation
' to only the other scanlines in the same pass, different than non-interlaced images.
' Think of non-interlaced images as single-pass interlaced images.

    ' counter variables
    Dim Pass As Byte, srcRow As Long
    ' sizing/bit alignment variables
    Dim nr8wide As Long, nr8high As Long
    Dim nrBytes As Long, passPtr As Long
    ' our temporary arrays
    Dim UnFiltered() As Byte     ' the unfiltered, complete image
    Dim InterlacePass() As Byte  ' unfiltered progressive display (used 7x for 7 passes)
    ' bytes and bits per pixel values
    Dim bytesPP As Byte, BPRow As Long, bitPP As Byte, bmpBPP As Byte
    ' allows direct memory access (DMA) to the DIB bits
    Dim TSA As SafeArray1D
    
    ' need bit & byte information
    GetDepthInfo bmpBPP, 0, bitPP, bytesPP
    ' create safearray structure
    With TSA
        .cDims = 1                          ' Number of dimensions
        .cbElements = 1                     ' Size of data elements
        .pvData = cPNGimage.PngDataPtr             ' Data address
        .rgSABound.lLbound = 0              ' Low bound of array
        .rgSABound.cElements = cPNGimage.PngDataSize ' Nr of elements in array
    End With

    ' overlay the safearray over our uninitialized array
    CopyMemory ByVal VarPtrArray(UnFiltered), VarPtr(TSA), 4

    ' interlaced images always come in 7 passes; although not all passes may be used
    For Pass = 1 To 7
    
        ' ensure bounds are valid. If image is smaller than 8x8 not all passes are valid/used
        ' Tested with images as small as 1x1
    
        ' calculate nr of pixels for this pass that will fit in width of image
        nr8wide = png_Width \ MatrixColAdd(Pass) - (png_Width Mod MatrixColAdd(Pass) > MatrixCol(Pass))
        If nr8wide > 0 Then
            
            ' calcuate nr of rows for this pass that will fit in height of image
            nr8high = png_Height \ MatrixRowAdd(Pass) - (png_Height Mod MatrixRowAdd(Pass) > MatrixRow(Pass))
            If nr8high > 0 Then
    
                ' calculate row bytes for the interlaced image, byte aligned
                BPRow = GetBytesPerPixel(nr8wide, bitPP)
                ' how many bytes are needed for the complete pass, less filter byte?
                nrBytes = BPRow * nr8high
                BPRow = BPRow + 1
                '^^ the filter routines expect the filter byte to be in its parameters, so add it
                
                ' size the temp array to hold the unfiltered scanlines. This will be transfered over
                ' to the unfiltered array after unfiltering has been accomplished
                If Pass = 1 Then
                    ReDim InterlacePass(0 To nrBytes)
                ElseIf nrBytes > UBound(InterlacePass) Then
                    '^^ array could be large; don't resize if unnecessary
                    ReDim InterlacePass(0 To nrBytes - 1)
                End If
    
                ' unfilter the scanlines
                For srcRow = 0 To nr8high - 1
                    Select Case Filtered(BPRow * srcRow + passPtr)
                    Case 0: ' no filtering, straight copy to our temp array
                        CopyMemory InterlacePass(srcRow * BPRow - srcRow), Filtered(BPRow * srcRow + passPtr + 1), BPRow - 1
                    Case 1: ' sub filter
                        DecodeFilter_Sub Filtered, InterlacePass, srcRow, BPRow, bytesPP, passPtr
                    Case 2: ' up filter
                        DecodeFilter_Up Filtered, InterlacePass, srcRow, BPRow, passPtr
                    Case 3: ' average filter
                        DecodeFilter_Avg Filtered, InterlacePass, srcRow, BPRow, bytesPP, passPtr
                    Case 4: ' paeth filter
                        DecodeFilter_Paeth Filtered, InterlacePass, srcRow, BPRow, bytesPP, passPtr
                    Case Else
                        ' If we got here, there is a different filtering mechanism at large
                        CopyMemory ByVal VarPtrArray(UnFiltered), 0&, 4
                        Exit Function
                    End Select
                Next
        
                ' offset the filtered array pointer to account for the 1byte filter flag per scanline
                ' This will point to the next pass's X,Y position in the Unfiltered() array
                passPtr = passPtr + nrBytes + nr8high
            
                ' send unfiltered array to be transfered to the DIB(s)
                If ConvertPNGtoBMP(InterlacePass(), Pass, nr8high, BPRow - 1, 0) = False Then Exit For
            
                ' display the interlaced pass if appropriate
                RaiseEvent Progress(Pass * 100 \ 7)
                If (c_Overrides And pdNoInterlace) = 0 Then
                    If png_Interlacing = 1 Or bmpBPP = 32 Then _
                        cPNGimage.DisplayInterlaced tgtDC, tgtWindow, dcX, dcY, MatrixRow(Pass) + 0, MatrixRowAdd(Pass) + 0&, (Pass = 7), False, c_AlphaColor, AlphaBlend
                End If
        
            End If ' check for nr8high < 1
        
        End If ' check for nr8wide < 1
    
    Next Pass

    ' remove the safearray overlay
    CopyMemory ByVal VarPtrArray(UnFiltered), 0&, 4
    
    If Pass < 8 Then Exit Function

    UnfilterInterlaced = True
    
End Function

Private Function UnfilterNI(FilteredData() As Byte, _
            tgtDC As Long, tgtWindow As Long, dcX As Long, dcY As Long, AlphaBlend As Byte) As Boolean

' // LaVolpe, Dec 1 thru 10 - completely rewritten to remove excess large array usage
' http://www.w3.org/TR/PNG/#9-table91

    Dim UnFiltered() As Byte, Row As Long
    Dim BPRow As Long, lBpp As Byte, dBpp As Byte, stepVal As Byte
    Dim lastLineBlt As Long
    
    Const fauxInterlace As Long = 16
    '^^ tried other values ranging from 4 to 32 and 16 is best performance
    
    GetDepthInfo dBpp, 0, lBpp, stepVal
    BPRow = GetBytesPerPixel(png_Width, lBpp) + 1
    ReDim UnFiltered(0 To (BPRow - 1) * png_Height - 1)
    
    For Row = 0 To png_Height - 1

        Select Case FilteredData(BPRow * Row)
        Case 0 'None
            CopyMemory UnFiltered(BPRow * Row - Row), FilteredData(BPRow * Row + 1), BPRow - 1
        Case 1 'Sub
            DecodeFilter_Sub FilteredData, UnFiltered, Row, BPRow, stepVal
        Case 2 'Up
            DecodeFilter_Up FilteredData, UnFiltered, Row, BPRow
        Case 3 'Average
            DecodeFilter_Avg FilteredData, UnFiltered, Row, BPRow, stepVal
        Case 4 'Paeth
            DecodeFilter_Paeth FilteredData, UnFiltered, Row, BPRow, stepVal
        Case Else ' invalid filter type
            Exit Function
        End Select
        
        If Row Mod fauxInterlace = 0 Then
            If Row Then
                RaiseEvent Progress(Row * 100 \ png_Height)
                lastLineBlt = Row
                If ConvertPNGtoBMP(UnFiltered(), 8, Row, BPRow - 1, Row - fauxInterlace) = False Then Exit Function
                If (c_Overrides And pdScanner) = pdScanner Then _
                    cPNGimage.DisplayInterlaced tgtDC, tgtWindow, dcX, dcY, Row - fauxInterlace, fauxInterlace, (Row = png_Height), True, c_AlphaColor, AlphaBlend
            End If
        End If
        
    Next Row
    
    RaiseEvent Progress(100)
    If lastLineBlt < fauxInterlace Then
        If ConvertPNGtoBMP(UnFiltered(), 8, png_Height, BPRow - 1, 0) = False Then Exit Function
        If (c_Overrides And pdScanner) = pdScanner Then _
            cPNGimage.DisplayInterlaced tgtDC, tgtWindow, dcX, dcY, 0, png_Height, True, True, c_AlphaColor, AlphaBlend
    Else
        Row = png_Height - lastLineBlt
        If Row > 1 Then
            If ConvertPNGtoBMP(UnFiltered(), 8, png_Height, BPRow - 1, png_Height - Row) = False Then Exit Function
        End If
        If (c_Overrides And pdScanner) = pdScanner Then _
            cPNGimage.DisplayInterlaced tgtDC, tgtWindow, dcX, dcY, png_Height - Row, Row, True, True, c_AlphaColor, AlphaBlend
    End If
    UnfilterNI = True
    
End Function

Private Function ValidateDLLExists() As Boolean
    ' function ensures DLLs of interest exist in user's system path
    ' side note: DLL filename is not case-sensitive, call names are!
    
    Dim lb As Long, pa As Long, sRtn As String
    Dim I As Integer, sDLL As String, sCall As String
    
    For I = 1 To 2
        Select Case I
        Case 1
            sDLL = "zlibvb.dll": sCall = "uncompress"
        Case 2
            sDLL = "ws2_32.dll": sCall = "htonl"
        End Select
        'attempt to open the DLL to be checked
        lb = LoadLibrary(sDLL)
        'if so, retrieve the address of one of the function calls
        If lb Then pa = GetProcAddress(lb, sCall)
        
        If pa = 0 Then  ' nope, don't got it; can't run project
            If Len(sRtn) Then sRtn = sRtn & vbCrLf
            sRtn = sRtn & sDLL
        Else            ' got it, free references
            FreeLibrary lb
            lb = 0
            pa = 0
        End If
    Next
    If Len(sRtn) Then
        sRtn = "Following DLL(s) not found in your system path:" & vbCrLf & sRtn
        MsgBox sRtn, vbExclamation + vbOKOnly, "Missing Library"
    Else
        ValidateDLLExists = True
    End If
    
End Function

Private Function zChunk_bKGD(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
'http://www.w3.org/TR/PNG/#11bKGD
    
    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    
    On Error Resume Next
    Select Case png_ColorType
        Case 2, 6 ' true color with/without transparency
            png_WinBKG = Buffer(5) Or Buffer(7) * 256& Or Buffer(9) * 65536
        Case 3 ' grayscale or paletted
            png_WinBKG = GetPaletteValue(0, Buffer(4))
        Case 0, 4 ' grayscale with/without transparency
            png_WinBKG = GetPaletteValue(0, Buffer(5))
    End Select
    If Err Then
        png_WinBKG = -1 ' no suggested window background color
        Err.Clear
    End If

End Function

Private Function zChunk_IHDR(Buffer() As Byte, bufLen As Long, crcValue As Long, cmprSize As Long) As String
'http://www.w3.org/TR/PNG/#11IHDR
    ' IHDR structure
    '    Width As Long              << cannot be negative
    '    Height As Long             << cannot be negative
    '    BitDepth As Byte           << must be 1,2,4,8,16
    '    ColorType As Byte          << must be 0,2,3,4,6
    '    Compression As Byte        << must be zero
    '    Filter As Byte             << must be zero
    '    Interlacing As Byte        << must be zero or one
    
    On Error Resume Next
    If ntohl(crcValue) = Zcrc32(0, Buffer(0), bufLen + 4) Then
    
        Dim lRtn As Long, sReturn As String
        
        CopyMemory png_Width, Buffer(4), 4
        png_Width = ntohl(png_Width)
        CopyMemory png_Height, Buffer(8), 4
        png_Height = ntohl(png_Height)
        
        If png_Width < 1 Or png_Height < 1 Then
            sReturn = "Corrupted Image Header. Cannot continue."
        
        Else
            
            If Buffer(14) > 0 Then
                sReturn = "Invalid Compression Flag in Header. Cannot continue."
            Else
                If Buffer(15) > 0 Then
                    sReturn = "Invalid Filter Flag in Header. Cannot continue."
                Else
                    
                    png_BitDepth = Buffer(12)
                    Select Case png_BitDepth
                    Case 1, 2, 4, 8, 16
                        ' it is a valid bit depth
                        png_ColorType = Buffer(13)
                        Select Case png_ColorType
                        Case 0, 2, 3, 4, 6
                            ' it is a valid color type
                            png_Interlacing = Buffer(16)
                            If png_Interlacing > 1 Then
                                sReturn = "Invalid Interlacing Flag in Header. Cannot continue."
                            End If
                        Case Else
                            sReturn = "Invalid Color Type Flag in Header. Cannot continue."
                        End Select
                    Case Else
                        sReturn = "Invalid Bit Depth Flag in Header. Cannot continue."
                    End Select
                    
                End If  ' Filter flag
            End If  ' Compression flag
        End If  ' Dimensions
        
        If sReturn = "" Then
            ' check for png sizes that would cause overflow errors in other calculations...
            ' This has 2 basic checks
            ' check DWord width alignment * height first are within bounds
            lRtn = ((((png_Width * 32) + &H1F) And Not &H1F&) \ &H8) * png_Height
            ' see if uncompress png data is too long
            If Not Err Then cmprSize = CalcUncompressedWidth()
            If cmprSize = 0 Then
                If Err Then Err.Clear
                ' either the dWord aligned bytes required for the bitmap
                ' are too large (Long value) or the size of the total
                ' uncompressed pixel array is too large (Long value).
                ' Image is way too big to process anyway! would require GBs of memory
                sReturn = "Image is too large to process. Cannot continue."
            End If
        End If
    Else ' error; CRC validation failed; corrupt file
        sReturn = "Corrupted Image Header (bad CRC). Cannot continue."
    End If

    zChunk_IHDR = sReturn

End Function

Private Function zChunk_PLTE(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
' http://www.w3.org/TR/PNG/#11PLTE
    
    If png_ColorType = 0 Or png_ColorType = 4 Then Exit Function
    '^^ per specs, palettes shall not appear for those color types
    '   Since we can ignore the palette, we won't trigger a critcal error
    
    If ntohl(crcValue) = Zcrc32(0, Buffer(0), bufLen + 4) Then
        
        ' per png specs, palette must be divisible by 3
        If (bufLen \ 3) * 3 = bufLen Then
            ReDim png_Palette(bufLen - 1)
            CopyMemory png_Palette(0), Buffer(4), bufLen
        Else ' error
            zChunk_PLTE = "Invalid Palette. Cannot continue."
        End If
    Else ' error; CRC validation failed; corrupt file
        zChunk_PLTE = "Invalid Palette (bad CRC). Cannot continue."
    End If

End Function

Private Function zChunk_tIME(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
'http://www.w3.org/TR/PNG/#11tIME

    On Error GoTo ExitMe
    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    
    Dim iYear As Integer
    If bufLen > 6 Then ' format...7bytes
        'Year:   2 bytes (4 digits)
        'Month:  1 byte (1-12)
        'Day:    1 byte (1-31)
        'Hour:   1 byte (0-23)
        'Minute: 1 byte (0-59)
        'Second: 1 byte (0-60)
        ReDim png_LastModified(0 To 7)
        CopyMemory iYear, Buffer(4), &H2 ' year
        ' 4-digit year; convert to little endian
        iYear = ntohs(iYear)
        CopyMemory png_LastModified(0), iYear, 2      ' put year in array
        CopyMemory png_LastModified(2), Buffer(6), 5  ' put rest of date in array
    ' else not a valid chunk; however; not a critical chunk either
    End If

ExitMe:
End Function

Private Function zChunk_tRNS(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
'http://www.w3.org/TR/PNG/#11tRNS
    
    If png_ColorType > 3 Then Exit Function
    ' Per specs, the tRNS chunk shall not be used for Color Types 4 and 6

    On Error GoTo ExitMe
    If ntohl(crcValue) = Zcrc32(0, Buffer(0), bufLen + 4) Then
    
        Dim UB As Long
    
        ReDim png_TransSimple(0 To bufLen - 1)
        CopyMemory png_TransSimple(0), Buffer(4), bufLen
    
        If png_ColorType = 0 Then ' grayscale with simple transparency
            c_TransColor = png_TransSimple(1) ' color-index value
        ElseIf png_ColorType = 2 Then ' rgb triple + grayscale
            ' save as BGR to be compared against PNG samples
            c_TransColor = png_TransSimple(5) Or png_TransSimple(3) * 256& Or png_TransSimple(1) * 65536
        ElseIf png_ColorType = 3 Then ' PNGhdr.TransSimple() is an array
            ' to prevent out of bounds errors, ensure array is 255
            If UBound(png_TransSimple) < 255 Then ' pngs are not required to provide all
                UB = UBound(png_TransSimple)      ' alpha values for every color in
                ReDim Preserve png_TransSimple(0 To 255) ' the palette.
                FillMemory png_TransSimple(UB + 1), 255 - UB, 255
            End If
            c_TransColor = 0 ' simply a flag > -1, has no other meaning
                             ' alpha values are in the png_TransSimple() array
        End If
        If Err Then
            Err.Clear   ' an error regarding the PNGhdr.TransSimple() array. Ignore simple transparency
            c_TransColor = -1 ' prevent transparency
        End If
    
    'Else error, but not a critical error - pic may look crappy
    End If
    
    
ExitMe:
End Function

Private Function zChunk_tEXt(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
'http://www.w3.org/TR/PNG/#11tEXt
'http://www.w3.org/TR/PNG/#11zTXt

    On Error GoTo ExitMe
    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    
    ' not needed to display image; cache & extract via property if needed
    If IsArrayEmpty(Not png_Comments) Then
        ReDim png_Comments(0 To bufLen + 7)
    Else
        ReDim Preserve png_Comments(0 To UBound(png_Comments) + bufLen + 8)
    End If
    ' save the chunk name & bytes
    CopyMemory png_Comments(UBound(png_Comments) - bufLen - 3), Buffer(0), bufLen + 4
    ' also save the chunk size already converted to proper endian
    CopyMemory png_Comments(UBound(png_Comments) - bufLen - 7), bufLen, 4

ExitMe:
End Function

Private Function zChunk_gAMA(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
' http://www.w3.org/TR/PNG/#11gAMA

' Gamma correction; possibly applicable to BitmapInfoHeader v5
    On Error GoTo ExitMe
    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    
    If bufLen = 4 Then
        CopyMemory png_Gamma, Buffer(4), 4
        png_Gamma = ntohl(png_Gamma)
    End If
    
ExitMe:
End Function

Private Function zChunk_cHRM(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
' http://www.w3.org/TR/PNG/#11cHRM
    
' Chromatic information; possibly applicable to BitmapInfoHeader v5
    On Error GoTo ExitMe
    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    
    If bufLen > 31 Then
        Dim I As Integer, lItem As Long
        ReDim png_Chromaticity(0 To 7)
        For I = 0 To 7
            CopyMemory lItem, Buffer(I * 4 + 4), 4
            png_Chromaticity(I) = ntohl(lItem)
        Next
    End If
    
ExitMe:
End Function

Private Function zChunk_pHYs(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
'http://www.libpng.org/pub/png/spec/1.2/PNG-Chunks.html#C.pHYs

On Error GoTo ExitMe

    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    
    If bufLen > 8 Then
        'Pixels per unit, X axis: 4 bytes (unsigned integer)
        'Pixels per unit, Y axis: 4 bytes (unsigned integer)
        'Unit specifier:          1 byte (1=meter with 1 inch=.0254 meters)
        Dim lItem As Long
        ReDim png_Physical(0 To 2)
        CopyMemory lItem, Buffer(4), &H4
        png_Physical(0) = ntohl(lItem)
        CopyMemory lItem, Buffer(8), &H4
        png_Physical(1) = ntohl(lItem)
        png_Physical(2) = Buffer(12)
        If Buffer(12) = 1 Then ' units is meter
            ' note to self: this appears to be a LogPixelsX/Y setting when
            ' multiplying the X,Y ppu by .0254 (meters per inches)
            '   Example: if pixels per unit = 2834 or 3780
            '   CLng(2834 * .0254) = 72, CLng(3780 * .0254) = 96
            ' therefore, these values could be used with API SetViewportExtEx ?
        Else
            ' values are an aspect ratio; possibly for windows media centers?
            ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmencode/htm/preservingthepixelaspectratio.asp
            ' But for hDCs, maybe stretchblt to that ratio? or from that ratio?
        End If
    ' else invalid chunk
    End If
    
ExitMe:
End Function

Private Function zChunk_oFFs(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
' http://www.libpng.org/pub/png/book/chapter11.html#png.ch11.div.10

' physical offset of image (primarily used for printing, but could be for display too.
    On Error GoTo ExitMe
    
    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    
    If bufLen > 8 Then
        Dim lItem As Long
        ReDim png_Offsets(0 To 2)
        CopyMemory lItem, Buffer(4), &H4
        png_Offsets(0) = ntohl(lItem)         ' left offset
        CopyMemory lItem, Buffer(8), &H4
        png_Offsets(1) = ntohl(lItem)         ' top offset
        png_Offsets(2) = Buffer(9)            ' unit: either pixel(0) or micron(1)
    End If

ExitMe:
End Function

Private Function zChunk_sRGB(Buffer() As Byte, bufLen As Long, crcValue As Long) As String
' http://www.libpng.org/pub/png/spec/1.2/PNG-Chunks.html#C.sRGB
    
    If crcValue Then
        If ntohl(crcValue) <> Zcrc32(0, Buffer(0), bufLen + 4) Then Exit Function
    End If
    ' real values will be btwn 0 and 3, but tweaked so 0 means no sRGB chunk rcvd
    If Buffer(4) < 4 Then png_sRGB = Buffer(4) + 1
    '0: Perceptual
    '1: Relative colorimetric
    '2: Saturation
    '3: Absolute colorimetric

End Function


