VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' GDI+ Wrapper written by Paul
' http://www.syix.com/wpsjr1/index.html
'
' Wrapper will Dispose on setting equal to Nothing (terminate event)
' And you are responsible for Disposing before any new bitmap is created
' Failure to Dispose properly will lead to various problems/crashes.
'
' Currently, functions that make a new bitmap are:
' CreateFromScan0, CreateFromHBITMAP and LoadFromFile
'
' 10/04/2004 EnumEncoders and associated functions were moved to cGDIP
' This class is made more efficient without the Enum, which only needs to be performed once anyway.

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICDESC, riid As CLSID, ByVal fOwn As Long, ByRef lplpvObj As Any)

Public gdip As cGDIP ' Reference for Error Handling

Private Bitmap          As Long
Private lCurErr         As Long
Private lPixelFormat    As Long
Private m_lFrameCount   As Long
Private m_lWidth        As Long
Private m_lHeight       As Long
Private rc              As RECT
Private Stream          As IUnknown
Private bProperties()   As Byte
Private pi              As PROPERTYITEM
Private m_lPropSize     As Long
Private b()             As Byte

Private Function BitmapFromStream(ByVal IStream As IUnknown) As Long
  If Not (IStream Is Nothing) Then
    BitmapFromStream = Execute(GdipLoadImageFromStream(IStream, Bitmap))
    If BitmapFromStream = Status.Ok Then
      GetPixelFormat
      GetDimension
    End If
  End If
End Function

Public Function CreateFromScan0(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal ePixelFormats As PixelFormats, ByRef lDataPtr As Long)
  CreateFromScan0 = Execute(GdipCreateBitmapFromScan0(lWidth, lHeight, lStride, ePixelFormats, ByVal lDataPtr, Bitmap))
  GetPixelFormat
  GetDimension
End Function

' Assumes a zero-based byte array containing an entire image is passed
Public Function CreateFromStream(ByRef bIn() As Byte) As Long
  CreateStreamOnHGlobal bIn(0), False, Stream
  CreateFromStream = BitmapFromStream(Stream)
End Function

Public Function CreateFromHBitmap(ByVal hBitmap As Long) As Long
  CreateFromHBitmap = Execute(GdipCreateBitmapFromHBITMAP(hBitmap, 0, Bitmap))
  GetPixelFormat
  GetDimension
End Function

Public Function CreateFromGraphics(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lGraphicHandle As Long) As Long
  CreateFromGraphics = Execute(GdipCreateBitmapFromGraphics(lWidth, lHeight, lGraphicHandle, Bitmap))
  GetPixelFormat
  GetDimension
End Function

Public Function CreateFromClone(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lBitmapHandle As Long, ByVal ePixelFormat As PixelFormats) As Long
  CreateFromClone = Execute(GdipCloneBitmapAreaI(0, 0, lWidth, lHeight, ePixelFormat, lBitmapHandle, Bitmap))
End Function

' I purposely did not use GdipCreateBitmapFromResource for now,
' since it doesn't work on anything but bitmap resources.
' This should work on all GDI+ supported image types
' Store images as "CUSTOM" resources with the resource editor addin
Public Function CreateFromResource(ByVal lResID As Long) As Long
  'Dim b() As Byte ' Changed to module wide scope, 02/14/2004
  b = LoadResData(lResID, "CUSTOM")
  CreateStreamOnHGlobal b(0), False, Stream
  CreateFromResource = BitmapFromStream(Stream)
End Function

Public Function CreateThumbFromImage(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lBitmapHandle As Long) As Long
  CreateThumbFromImage = Execute(GdipGetImageThumbnail(lBitmapHandle, lWidth, lHeight, Bitmap))
  GetPixelFormat
  GetDimension
End Function

Public Function ConvertFormat(ByVal eFormat As PixelFormats, ByVal eDitherType As DitherType, ByVal ePaletteType As PaletteType, ByVal lPPalette As Long, ByVal sngAlphaThresholdPercent As Single) As Long
  ConvertFormat = Execute(GdipBitmapConvertFormat(Bitmap, eFormat, eDitherType, ePaletteType, ByVal lPPalette, sngAlphaThresholdPercent))
End Function

Public Property Get CurErr() As Long
  CurErr = lCurErr
End Property

Public Sub Dispose()
  If Bitmap Then
    GdipDisposeImage Bitmap
    Bitmap = 0
  End If
  Set Stream = Nothing
End Sub

Private Function Execute(ByVal lReturn As Long) As Long
  If lReturn = 0 Then
    lCurErr = 0
  Else
    lCurErr = lReturn
    If Not gdip Is Nothing Then gdip.Error lReturn, "cBitmap"
  End If
  Execute = lCurErr
End Function

' number of frames in a multi-page tiff or animated gif
Public Property Get FrameCount() As Long
  FrameCount = m_lFrameCount
End Property

Public Function LoadFromFile(ByVal sFilename As String) As Long
  LoadFromFile = Execute(GdipLoadImageFromFile(sFilename, Bitmap))
  GetPixelFormat
  GetDimension
End Function

' Returns the number of bytes from the stream in the array
' 0 if failure
Friend Function GetBitmapStream(ByRef bBmpBits() As Byte, ByRef tEncoder As CLSID, ByVal lEncParamPointer As Long) As Long
  Dim Stream As IStream
  Dim lBytesRead As Long
  Dim curMax As Currency
  
  ' True allows the global memory to be freed automatically,
  ' after the destruction of the IStream
  If CreateStreamOnHGlobal(ByVal 0, True, Stream) = 0 Then
    If Execute(GdipSaveImageToStream(Bitmap, Stream, tEncoder, ByVal lEncParamPointer)) = 0 Then
      
      ' Here is an alternate method of getting information from the stream
      ' It's not nearly as versitile, and liable to be inaccurate,
      ' as GlobalSize might be rounded.
    
      ' I thought I'd include it, as the IStream interface need not be
      ' defined to use this code, simply; Dim stream as IUnknown
    
      ' Dim p As Long
      ' Dim hGlob As Long
      ' Dim lSizeGlob As Long
    
      ' GetHGlobalFromStream stream, p
      ' RtlMoveMemory hGlob, ByVal p, 4 ' could also GlobalLock here to deref and lock, remember to unlock later
      ' lSizeGlob = GlobalSize(hGlob)
      ' If lSizeGlob Then
      '   ReDim bBmpBits(lSizeGlob - 1)
      '   RtlMoveMemory bBmpBits(0), ByVal hGlob, lSizeGlob
      ' End If

      ' Note: you could seek anywhere in the stream and get any portion
      ' of the bitmap stream.
      Stream.Seek 0@, STREAM_SEEK_END, curMax ' get the size of the stream
      curMax = curMax * 10000 ' adjust for int64

      Stream.Seek 0@, STREAM_SEEK_SET, 0@ ' return to the beginning

      If curMax > 0 Then
        ReDim bBmpBits(curMax - 1)
        Stream.Read bBmpBits(0), curMax, lBytesRead ' note that lBytesRead could overflow, please fix this if you're reading more than 2 Gigs :P
        GetBitmapStream = lBytesRead
      End If
    End If
  End If

  Set Stream = Nothing
End Function

Private Sub GetDimension()
  Dim sngWidth  As Single
  Dim sngHeight As Single
  
  Execute GdipGetImageDimension(Bitmap, sngWidth, sngHeight)
  m_lWidth = sngWidth
  m_lHeight = sngHeight
  rc.Right = m_lWidth
  rc.Bottom = m_lHeight
End Sub

Friend Sub GetPalette(ByVal lPalPointer As Long, ByVal lNumColors As Long)
  Dim lSize As Long
  
  lSize = (lNumColors * 4) + 8
  Execute GdipGetImagePalette(Bitmap, ByVal lPalPointer, lSize)
End Sub

Private Sub GetPixelFormat()
  Execute GdipGetImagePixelFormat(Bitmap, lPixelFormat)
End Sub

Public Property Get Handle() As Long
  Handle = Bitmap
End Property

' So far this should only be changed from cGraphics.MakeTransparent
Public Property Let Handle(ByVal lHandle As Long)
  Bitmap = lHandle
End Property

Public Property Get Height() As Long
  Height = m_lHeight
End Property

' Useful for Animated GIF, gets the delay for a given frame (0 based)
Public Function GetDelay(ByVal lFrame As Long) As Long
  Dim vDelay As Variant
  
  GetPropertyItem PropertyTagFrameDelay, vDelay, lFrame
  GetDelay = vDelay
End Function

Public Function GetLoopCount() As Long
  Dim vLoopCount As Variant
  
  If GetPropertyItem(PropertyTagLoopCount, vLoopCount, 0) = Status.Ok Then
    GetLoopCount = vLoopCount
  End If
End Function

Public Function GetPropertyItem(ByVal eTag As PropertyTags, ByRef vValue As Variant, ByVal lPropertyIndex As Long) As Long
  Dim lSize As Long
  
  GetPropertyItem = Execute(GdipGetPropertyItemSize(Bitmap, eTag, lSize))
  If GetPropertyItem = Status.Ok Then
    ReDim bProperties(lSize - 1)
    GetPropertyItem = Execute(GdipGetPropertyItem(Bitmap, eTag, lSize, bProperties(0)))
    If GetPropertyItem = Status.Ok Then
      RtlMoveMemory pi, bProperties(0), LenB(pi) ' 2 bytes padding in memory
      If pi.Value And pi.length > 0 Then
        ReDim bProperties(pi.length - 1)
        RtlMoveMemory bProperties(0), ByVal pi.Value, pi.length
        m_lPropSize = pi.length
        If pi.type = PropertyTagTypeByte Or pi.type = PropertyTagTypeShort Or pi.type = PropertyTagTypeLong Then
          vValue = GetPropertyNumeric(pi.type, lPropertyIndex)
        ElseIf pi.type = PropertyTagTypeASCII Or pi.type = PropertyTagTypeRational Then
          'vValue = GetPropertyString(pi.Type, lPropertyIndex)
        Else
          Debug.Print "Error - Unhandled PropertyTagType"
        End If
      End If
    End If
  End If
End Function

Private Function GetPropertyNumeric(ByVal eTagtype As PropertyTagTypes, ByVal lPropIndex As Long) As Long
  Dim lSize As Long
  Dim lIndex As Long
  
    Select Case eTagtype
      Case PropertyTagTypeByte
        lSize = 1
      
      Case PropertyTagTypeShort
        lSize = 2
      
      Case PropertyTagTypeLong
        lSize = 4
    End Select
    
    lIndex = lPropIndex * lSize
    If lIndex >= 0 And lIndex < m_lPropSize Then
      RtlMoveMemory GetPropertyNumeric, bProperties(lIndex), lSize
    End If
End Function

'Private Function GetPropertyString(ByVal eTagtype As PropertyTagTypes, ByVal lPropIndex As Long) As String
'  If lPropIndex >= 0 And lPropIndex < m_lPropCount Then
'    ' todo
'  End If
'End Function

Friend Function GetPixelARGB(ByVal X As Long, ByVal Y As Long) As ARGB
  Execute GdipBitmapGetPixel(Bitmap, X, Y, GetPixelARGB)
End Function

Public Function GetPixelLong(ByVal X As Long, ByVal Y As Long) As Long
  Execute GdipBitmapGetPixel(Bitmap, X, Y, GetPixelLong)
End Function

Private Function IPictureFromBitmap(ByVal lBitmap As Long) As IPicture
  Dim lError  As Long
  Dim hBitmap As Long
  Dim picdes  As PICDESC

  If Bitmap = 0 Then Exit Function
    
  lError = GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0)
    
  If hBitmap <> 0 Then
    picdes.cbSizeOfStruct = Len(picdes)
    picdes.picType = vbPicTypeBitmap
    picdes.hGdiObj = hBitmap
    OleCreatePictureIndirect picdes, StringToGuid(IPictureCLSID), True, IPictureFromBitmap
  Else
    Error lError
  End If
  ' dispose of bitmap in caller
End Function

Public Function InitializePalette(ByVal lPPalette As Long, ByVal ePalType As PaletteType, ByVal bUseTransparent As Boolean, Optional ByVal lOptimalColors As Long = 0) As Long
  If lOptimalColors > 0 Then
    InitializePalette = Execute(GdipInitializePalette(ByVal lPPalette, PaletteTypeOptimal, lOptimalColors, CLng(Abs(bUseTransparent)), Bitmap))
  Else
    InitializePalette = Execute(GdipInitializePalette(ByVal lPPalette, ePalType, 0, CLng(Abs(bUseTransparent)), 0))
  End If
End Function

Friend Function LockBits(ByRef rc As RECT, ByVal eMode As ImageLockMode, ByVal eFormat As PixelFormats, ByRef bd As BitmapData) As Long
  LockBits = Execute(GdipBitmapLockBits(Bitmap, rc, eMode, eFormat, bd))
End Function

Public Function MakeTransparent(ByVal lColor As Long) As Long
  Dim bmp As cBitmap
  Dim gfx As cGraphics
  Set bmp = New cBitmap
  Set gfx = New cGraphics
  Dim rc As RECT
  Dim lAttrib As Long
  Dim lTmpHandle As Long

  If bmp.CreateFromScan0(m_lWidth, m_lHeight, 0, PixelFormat32bppARGB, 0) = Ok Then
    If gfx.GetGraphicsContext(bmp.Handle) = Ok Then
      If gfx.Clear = Ok Then
        rc.Right = m_lWidth
        rc.Bottom = m_lHeight
        
        If GdipCreateImageAttributes(lAttrib) = Ok Then
          If GdipSetImageAttributesColorKeys(lAttrib, ColorAdjustTypeDefault, 1, lColor, lColor) = Ok Then
            If gfx.DrawImageRectRectIAttr(Bitmap, 0, 0, m_lWidth, m_lHeight, 0, 0, m_lWidth, m_lHeight, lAttrib) = Ok Then
              GdipDisposeImageAttributes lAttrib
              lTmpHandle = Bitmap
              Bitmap = bmp.Handle
              bmp.Handle = lTmpHandle
            End If
          End If
        End If
      End If
    End If
  End If
  
  Set gfx = Nothing
  Set bmp = Nothing
End Function

Public Property Get Picture() As StdPicture
  Set Picture = IPictureFromBitmap(Bitmap)
End Property

Public Property Get PixelFormat() As PixelFormats
  PixelFormat = lPixelFormat
End Property

Friend Property Get Rectangle() As RECT
  Rectangle = rc
End Property

Public Function RotateFlip(ByVal eFlipType As RotateFlipType) As Long
  RotateFlip = Execute(GdipImageRotateFlip(Bitmap, eFlipType))
End Function

' At the time I wrote this class there were only two multi-page formats supported
' Consider adding an enum if more formats arise
Public Function SelectActiveFrame(ByVal lFrame As Long, Optional ByVal lIsGIF As Boolean = True) As Long
  Dim cid As CLSID
  If lIsGIF Then
    cid = StringToGuid(FrameDimensionTime)
  Else ' TIFF
    cid = StringToGuid(FrameDimensionPage)
  End If
  SelectActiveFrame = Execute(GdipImageSelectActiveFrame(Bitmap, cid, lFrame))
  If m_lFrameCount = 0 Then GdipImageGetFrameCount Bitmap, cid, m_lFrameCount
  GetDimension ' frames can be different sizes
End Function

Public Sub SetTransColor(ByVal lIndex As Byte)
  Dim gpal As ColorPalette256
  
  If Bitmap Then
    GetPalette VarPtr(gpal), 256
    gpal.Entries(lIndex).Alpha = 0
    SetPalette VarPtr(gpal)
  End If
End Sub

Friend Function SaveAdd(ByRef enc As EncoderParameters) As Long
  SaveAdd = Execute(GdipSaveAdd(Bitmap, enc))
End Function

Friend Function SaveAddImage(ByVal sFilename As String, ByVal lBitmapHandle As Long, ByRef enc As EncoderParameters) As Long
  SaveAddImage = Execute(GdipSaveAddImage(Bitmap, lBitmapHandle, enc))
End Function

Friend Function SaveToFile(ByVal sFilename As String, ByRef tEncoder As CLSID, ByVal lEncParamPointer As Long) As Long
  SaveToFile = Execute(GdipSaveImageToFile(Bitmap, sFilename, tEncoder, ByVal lEncParamPointer))
End Function

Public Function SetPalette(ByVal lPalPointer As Long) As Long
  SetPalette = Execute(GdipSetImagePalette(Bitmap, ByVal lPalPointer))
End Function

Friend Function StringToGuid(ByRef sGuid As String) As CLSID
  CLSIDFromString sGuid, StringToGuid
End Function

Friend Function UnlockBits(ByRef bd As BitmapData) As Long
  UnlockBits = Execute(GdipBitmapUnlockBits(Bitmap, bd))
End Function

Public Property Get Width() As Long
  Width = m_lWidth
End Property

' Private Sub Class_Initialize()
  ' EnumEncoders moved to cGDIP 10/04/2004 (along with associated functions)
' End Sub

Private Sub Class_Terminate()
  Dispose
End Sub
