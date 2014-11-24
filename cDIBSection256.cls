VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


' ==================================================================================
' cDIBSection256.cls
' Copyright © 1999 Steve McMahon
' Visit vbAccelerator at http://vbaccelerator.com
'
' Creates and manages a 256 colour GDI DibSection.  This is DIB
' in which the bitmap bits are stored in windows memory so can
' be modified.  Also, there are only 256 colours (1 byte/pixel)
' and the colour palette can be modified using GetDIBColorTable
' and SetDIBColorTable.  This means fades etc can be achieved
' by simply manipulating the DIB Color Table, rather than
' modifying the bitmap bits.  By doing this, a fade on a
' 512x512 fade can run much quicker than the equivalent for a
' True Colour DIB.
' The speed you run at depends on how your gfx driver implements
' DIB colour tables.  On a 8Mb ATI Xpert@Work, Win95, this code
' runs at > 300 fps for a 256x256 DIB!  However, on a 4Mb Matrox
' Millenium, NT it runs at less speed.
'
' Note: for best performance, when compiling an executable check
' all the boxes on the Properties-Compile tab Advanced Optimisations
' button, particularly Remove Array Bounds checks.
' ==================================================================================


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                               lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long




Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
                                          (ByVal hdc As Long, _
                                           pBitmapInfo As BITMAPINFO256, _
                                           ByVal un As Long, _
                                           lplpVoid As Long, _
                                           ByVal Handle As Long, _
                                           ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0        '  color table in RGBs
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As Any) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Private Declare Function GetDIBits256 Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long

Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Start of structure:
Private Const BITMAPTYPE As Integer = &H4D42
Private Type BITMAPFILEHEADER
    bfType As Integer        '- type  ="BM" i.e &H4D42 - 2
    bfSize As Long        ' - size in bytes of file - 6
    bfReserved1 As Integer        ' - reserved, must be 0 - 8
    bfReserved2 As Integer        ' - reserved, must be 0 - 10
    bfOffBits As Long        ' offset from this structure to the bitmap bits - 14
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Long, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO256, ByVal wUsage As Long) As Long



Private m_hDIb As Long
Private m_hBmpOld As Long
Private m_hDC As Long
Private m_lPtr As Long
Private m_tBI As BITMAPINFO256
' for speed - declare RGB array as global
Private tRGB(0 To 256) As RGBQUAD

Public Function CreateDIB( _
       ByVal lHdc As Long, _
       ByVal lWidth As Long, _
       ByVal lHeight As Long, _
       ByRef hDib As Long _
       ) As Boolean
        '<EhHeader>
        On Error GoTo CreateDIB_Err
        '</EhHeader>
    Dim i As Long
    Dim j As Long
100 With m_tBI.bmiHeader
102     .biSize = Len(m_tBI.bmiHeader)
104     .biWidth = lWidth
106     .biHeight = lHeight
108     .biPlanes = 1
110     .biBitCount = 8
112     .biCompression = BI_RGB
114     .biSizeImage = BytesPerScanLine * .biHeight
    End With
    ' Create colorful palette as default:
116 For i = 0 To 15
118     With m_tBI.bmiColors(i)
120         j = i
122         .rgbBlue = (j * 255) \ 16
124         .rgbGreen = (j * 128) \ 16
126         .rgbRed = (j * 96) \ 16
        End With
128 Next i


130 hDib = CreateDIBSection( _
           lHdc, _
           m_tBI, _
           DIB_RGB_COLORS, _
           m_lPtr, _
           0, 0)
132 CreateDIB = (hDib <> 0)
        '<EhFooter>
        Exit Function

CreateDIB_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.CreateDIB", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Function
Public Function Create( _
       ByVal lWidth As Long, _
       ByVal lHeight As Long _
       ) As Boolean
        '<EhHeader>
        On Error GoTo Create_Err
        '</EhHeader>
    Dim lHDCDesk As Long
100 ClearUp
102 lHDCDesk = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
104 m_hDC = CreateCompatibleDC(lHDCDesk)
106 DeleteDC lHDCDesk
108 If (m_hDC <> 0) Then
110     If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
112         m_hBmpOld = SelectObject(m_hDC, m_hDIb)
114         Create = True
        Else
116         DeleteObject m_hDC
118         m_hDC = 0
        End If
    End If
        '<EhFooter>
        Exit Function

Create_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.Create", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Function
Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
        '<EhHeader>
        On Error GoTo BytesPerScanLine_Err
        '</EhHeader>
100 BytesPerScanLine = (m_tBI.bmiHeader.biWidth + 3) And &HFFFFFFFC
        '<EhFooter>
        Exit Property

BytesPerScanLine_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.BytesPerScanLine", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property

Public Property Get Width() As Long
        '<EhHeader>
        On Error GoTo Width_Err
        '</EhHeader>
100 Width = m_tBI.bmiHeader.biWidth
        '<EhFooter>
        Exit Property

Width_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.Width", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property
Public Property Get Height() As Long
        '<EhHeader>
        On Error GoTo Height_Err
        '</EhHeader>
100 Height = m_tBI.bmiHeader.biHeight
        '<EhFooter>
        Exit Property

Height_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.Height", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property

Public Sub LoadPictureBlt( _
       ByVal lHdc As Long, _
       Optional ByVal lSrcLeft As Long = 0, _
       Optional ByVal lSrcTop As Long = 0, _
       Optional ByVal lSrcWidth As Long = -1, _
       Optional ByVal lSrcHeight As Long = -1, _
       Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
       )
        '<EhHeader>
        On Error GoTo LoadPictureBlt_Err
        '</EhHeader>
    Dim lC As Long
100 lC = GetDIBColorTable(lHdc, 0, 256, tRGB(0))
102 Debug.Assert (lC = 256)
104 If (lC > 0) Then
106     SetDIBColorTable m_hDC, 0, lC, tRGB(0)
    End If
108 If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
110 If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
112 BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lHdc, lSrcLeft, lSrcTop, eRop

        '<EhFooter>
        Exit Sub

LoadPictureBlt_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.LoadPictureBlt", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Sub


Public Sub PaintPicture( _
       ByVal lHdc As Long, _
       Optional ByVal lDestLeft As Long = 0, _
       Optional ByVal lDestTop As Long = 0, _
       Optional ByVal lDestWidth As Long = -1, _
       Optional ByVal lDestHeight As Long = -1, _
       Optional ByVal lSrcLeft As Long = 0, _
       Optional ByVal lSrcTop As Long = 0, _
       Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
       )
        '<EhHeader>
        On Error GoTo PaintPicture_Err
        '</EhHeader>
100 If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
102 If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
104 BitBlt lHdc, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
        '<EhFooter>
        Exit Sub

PaintPicture_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.PaintPicture", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Sub

Public Property Get hdc() As Long
        '<EhHeader>
        On Error GoTo hdc_Err
        '</EhHeader>
100 hdc = m_hDC
        '<EhFooter>
        Exit Property

hdc_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.hdc", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property
Public Property Get hDib() As Long
        '<EhHeader>
        On Error GoTo hDib_Err
        '</EhHeader>
100 hDib = m_hDIb
        '<EhFooter>
        Exit Property

hDib_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.hDib", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property
Public Property Get DIBSectionBitsPtr() As Long
        '<EhHeader>
        On Error GoTo DIBSectionBitsPtr_Err
        '</EhHeader>
100 DIBSectionBitsPtr = m_lPtr
        '<EhFooter>
        Exit Property

DIBSectionBitsPtr_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.DIBSectionBitsPtr", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property

Public Sub ClearUp()
        '<EhHeader>
        On Error GoTo ClearUp_Err
        '</EhHeader>
100 If (m_hDC <> 0) Then
102     If (m_hDIb <> 0) Then
104         SelectObject m_hDC, m_hBmpOld
106         DeleteObject m_hDIb
        End If
108     DeleteObject m_hDC
    End If
110 m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0
        '<EhFooter>
        Exit Sub

ClearUp_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.ClearUp", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Sub

Public Property Get Color(ByVal nIndex As Long) As Long
        '<EhHeader>
        On Error GoTo Color_Err
        '</EhHeader>
    Dim lC As Long
    Dim tRGBItem As RGBQUAD
100 lC = GetDIBColorTable(m_hDC, nIndex, 1, tRGBItem)
102 Debug.Assert (lC = 1)
104 If (lC = 1) Then
106     Color = tRGBItem.rgbRed Or (tRGBItem.rgbGreen * &H100) Or (tRGBItem.rgbBlue * &H10000)
    Else
108     Color = -1        ' CLR_INVALID
    End If
        '<EhFooter>
        Exit Property

Color_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.Color", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property




Public Property Get BitmapInfo() As BITMAPINFO256
        '<EhHeader>
        On Error GoTo BitmapInfo_Err
        '</EhHeader>
100     BitmapInfo = m_tBI
        '<EhFooter>
        Exit Property

BitmapInfo_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.BitmapInfo", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property








Public Property Let Color(ByVal nIndex As Long, ByVal lColor As Long)
        '<EhHeader>
        On Error GoTo Color_Err
        '</EhHeader>
    Dim lC As Long
    Dim tRGBItem As RGBQUAD
    Dim lB As Long, lG As Long, lR As Long
100 lB = (lColor And &HFF0000) \ &H10000
102 lG = (lColor And &HFF00&) \ &H100
104 lR = (lColor And &HFF)
106 tRGBItem.rgbBlue = lB
108 tRGBItem.rgbGreen = lG
110 tRGBItem.rgbRed = lR
112 lC = SetDIBColorTable(m_hDC, nIndex, 1, tRGBItem)

114 m_tBI.bmiColors(nIndex).rgbRed = lR
116 m_tBI.bmiColors(nIndex).rgbGreen = lG
118 m_tBI.bmiColors(nIndex).rgbBlue = lB

        '<EhFooter>
        Exit Property

Color_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.Color", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Property
Public Sub SetPalette( _
       lColor() As Long _
       )
        '<EhHeader>
        On Error GoTo SetPalette_Err
        '</EhHeader>
    Dim tRGBOut(0 To 256) As RGBQUAD
    Dim lC As Long
    Dim i As Long

100 lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
102 For i = LBound(lColor) To UBound(lColor)
104     tRGB(i).rgbBlue = (lColor(i) And &HFF0000) \ &H10000
106     tRGB(i).rgbGreen = (lColor(i) And &HFF00&) \ &H100&
108     tRGB(i).rgbRed = (lColor(i) And &HFF&)
110 Next i
112 lC = SetDIBColorTable(m_hDC, 0, 256, tRGB(0))

        '<EhFooter>
        Exit Sub

SetPalette_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.SetPalette", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Sub

Public Sub CopyPalette(ByRef cDIB As cDIBSection256)
        '<EhHeader>
        On Error GoTo CopyPalette_Err
        '</EhHeader>
    Dim lC As Long
100 lC = GetDIBColorTable(cDIB.hdc, 0, 256, tRGB(0))
102 Debug.Assert (lC = 256)
104 If (lC > 0) Then
106     lC = SetDIBColorTable(m_hDC, 0, 256, tRGB(0))
108     Debug.Assert (lC = 256)
    End If
        '<EhFooter>
        Exit Sub

CopyPalette_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.CopyPalette", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Sub
Public Sub FixPalette()
        '<EhHeader>
        On Error GoTo FixPalette_Err
        '</EhHeader>
    Dim lC As Long, i As Long

    ' Fix up the palette to match the current DIB colour table
100 lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
102 Debug.Assert (lC = 256)
104 For i = 0 To lC - 1
106     LSet m_tBI.bmiColors(i) = tRGB(i)
    Next

        '<EhFooter>
        Exit Sub

FixPalette_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.FixPalette", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Sub


Public Function SavePicture(ByVal sFilename As String) As Boolean
        '<EhHeader>
        On Error GoTo SavePicture_Err
        '</EhHeader>
    Dim lC As Long, i As Long

    ' Fix up the palette to match the current DIB colour table
100 lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
102 Debug.Assert (lC = 256)
104 For i = 0 To lC - 1
106     LSet m_tBI.bmiColors(i) = tRGB(i)
    Next
    ' Save to BMP with 256 colour palette:
108 SavePicture = SaveToBitmap(m_tBI, m_lPtr, sFilename)

        '<EhFooter>
        Exit Function

SavePicture_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.SavePicture", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Function
Private Function SaveToBitmap(ByRef tBI As BITMAPINFO256, ByVal lPtrBits As Long, ByVal sFilename As String)
        '<EhHeader>
        On Error GoTo SaveToBitmap_Err
        '</EhHeader>
    Dim tBH As BITMAPFILEHEADER
    Dim tRGBQ As RGBQUAD
    Dim hFile As Long
    Dim lBytesWritten As Long
    Dim lSize As Long
    Dim lR As Long
    Dim bErr As Boolean
    Dim hMem As Long, lPtr As Long
    Dim lErr As Long

    ' Prepare the BITMAPFILEHEADER
100 With tBH
102     .bfType = BITMAPTYPE
104     .bfOffBits = 14 + Len(tBI)
106     .bfSize = .bfOffBits + tBI.bmiHeader.biSizeImage
    End With
108 hFile = CreateFile(sFilename, _
                       GENERIC_READ Or GENERIC_WRITE, _
                       ByVal 0&, _
                       ByVal 0&, _
                       CREATE_ALWAYS, _
                       FILE_ATTRIBUTE_NORMAL, _
                       0)
110 lErr = Err.LastDllError
112 If (hFile = INVALID_HANDLE_VALUE) Then
        ' error
114     Err.Raise 17, App.EXEName & ".cDIBSection256", ApiError(lErr)
    Else

        ' Writing the BITMAPFILEINFOHEADER is somewhat painful
        ' due to non-byte alignment of structure...
116     hMem = GlobalAlloc(GPTR, 14)
118     lPtr = GlobalLock(hMem)
120     CopyMemory ByVal lPtr, tBH.bfType, 2
122     CopyMemory ByVal lPtr + 2, tBH.bfSize, 4
124     CopyMemory ByVal lPtr + 6, 0&, 4
126     CopyMemory ByVal lPtr + 10, tBH.bfOffBits, 4
128     lSize = 14
130     lR = WriteFile(hFile, ByVal lPtr, lSize, lBytesWritten, ByVal 0&)
132     GlobalUnlock hMem
134     GlobalFree hMem

        ' Add the BITMAPINFOHEADER and colour palette:
136     bErr = FileErrHandler(lR, lSize, lBytesWritten)
138     If Not bErr Then
140         lSize = Len(tBI)
142         lR = WriteFile(hFile, tBI, lSize, lBytesWritten, ByVal 0&)
144         bErr = FileErrHandler(lR, lSize, lBytesWritten)
        End If

146     If Not bErr Then
            ' Its easy to write the bitmap data, though...
148         lSize = tBI.bmiHeader.biSizeImage
150         lR = WriteFile(hFile, ByVal lPtrBits, lSize, lBytesWritten, ByVal 0&)
152         bErr = FileErrHandler(lR, lSize, lBytesWritten)
        End If


154     CloseHandle hFile
        'SavePicture = Not (bErr)
    End If

        '<EhFooter>
        Exit Function

SaveToBitmap_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.SaveToBitmap", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Function
Private Function ApiError(ByVal e As Long) As String
        '<EhHeader>
        On Error GoTo ApiError_Err
        '</EhHeader>
    Dim s As String, c As Long
100 s = String(256, 0)
102 c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, e, 0&, s, Len(s), ByVal 0)
104 If c Then ApiError = Left$(s, c)
        '<EhFooter>
        Exit Function

ApiError_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.ApiError", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Function

Private Function FileErrHandler(ByVal lR As Long, ByVal lSize As Long, ByVal lBytes As Long) As Boolean
        '<EhHeader>
        On Error GoTo FileErrHandler_Err
        '</EhHeader>
100 If (lR = 0) Or Not (lSize = lBytes) Then
        'Err.Raise
102     FileErrHandler = True
    End If
        '<EhFooter>
        Exit Function

FileErrHandler_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cDIBSection256.FileErrHandler", _
                  "cDIBSection256 component failure"
        '</EhFooter>
End Function



Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
ClearUp
End Sub




