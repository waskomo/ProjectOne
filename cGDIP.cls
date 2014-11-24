VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' GDI+ Wrapper written by Paul
' http://www.syix.com/wpsjr1/index.html

' Initialization class
' Seperated off from cBitmap class 02/14/2004

Event Error(ByVal lGdiError As Long, ByVal sErrorDesc As String, ByVal sClassName As String)

Private lToken          As Long
Private m_lNumCodecs    As Long
Private tGuids()        As CLSID        ' Array of GUIDs for codecs
Private colCodecs       As Collection   ' Search into codec index on file extension

Private Sub Class_Initialize()
  Dim gsi As GdiplusStartupInput
  
  gsi.GdiplusVersion = 1
  
  ' next line used when debugging in VC, a la DebugBreak API
  ' gsi.DebugEventCallback = PassAddress(AddressOf DebugEventProc)
  If GdiplusStartup(lToken, gsi, ByVal 0) <> Status.Ok Then
    Debug.Print "GdiplusStartup Error!"
  Else
    EnumEncoders
  End If
End Sub

Public Function Error(ByVal lGdiError As Long, ByVal sClassName As String)
  RaiseEvent Error(lGdiError, GdiErrorString(lGdiError), sClassName)
End Function

Private Function GdiErrorString(ByVal lError As Status) As String
  Dim s As String
  
  Select Case lError
    Case GenericError:              s = "Generic Error"
    Case InvalidParameter:          s = "Invalid Parameter"
    Case OutOfMemory:               s = "Out Of Memory"
    Case ObjectBusy:                s = "Object Busy"
    Case InsufficientBuffer:        s = "Insufficient Buffer"
    Case NotImplemented:            s = "Not Implemented"
    Case Win32Error:                s = "Win32 Error"
    Case WrongState:                s = "Wrong State"
    Case Aborted:                   s = "Aborted"
    Case FileNotFound:              s = "File Not Found"
    Case ValueOverflow:             s = "Value Overflow"
    Case AccessDenied:              s = "Access Denied"
    Case UnknownImageFormat:        s = "Unknown Image Format"
    Case FontFamilyNotFound:        s = "FontFamily Not Found"
    Case FontStyleNotFound:         s = "FontStyle Not Found"
    Case NotTrueTypeFont:           s = "Not TrueType Font"
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version"
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized"
    Case PropertyNotFound:          s = "Property Not Found"
    Case PropertyNotSupported:      s = "Property Not Supported"
    Case Else:                      s = "Unknown GDI+ Error"
  End Select
  
  GdiErrorString = s
End Function

Private Function EnumEncoders() As Long
  Dim lNumEncoders  As Long
  Dim lEncoderSize  As Long
  Dim lError        As Long
  Dim b()           As Byte
  Dim i             As Long
  Dim codecs()      As ImageCodecInfo
  
  lError = GdipGetImageEncodersSize(lNumEncoders, lEncoderSize)
  If lError = 0 Then
    ReDim codecs(lNumEncoders - 1)
    ReDim b(lEncoderSize - 1)
    
    lError = GdipGetImageEncoders(lNumEncoders, lEncoderSize, b(0))
    If lError = 0 Then
      RtlMoveMemory codecs(0), b(0), lNumEncoders * LenB(codecs(0))
      ReDim tGuids(lNumEncoders - 1)
      m_lNumCodecs = lNumEncoders
      Set colCodecs = Nothing
      Set colCodecs = New Collection
      
      Do While lNumEncoders
        lNumEncoders = lNumEncoders - 1
        tGuids(lNumEncoders) = codecs(lNumEncoders).CLSID
        ParseOnChar StringFromPointerW(codecs(lNumEncoders).pwszFilenameExtension), ";", lNumEncoders
      Loop
    Else
      Error lError
    End If
  Else
    Error lError
  End If
End Function

Friend Property Get EncoderGuid(ByVal lIndex As Long) As CLSID
  EncoderGuid = tGuids(lIndex)
End Property

' do not compare this to a boolean
'  returns -1 for not found, 0-positive GUID index for found
Public Function ExtensionExists(ByRef sKey As String) As Long
  On Error GoTo ErrorHandler
  ExtensionExists = True ' invalid index
  
  If Not colCodecs Is Nothing Then
    ExtensionExists = colCodecs.Item(sKey)
  End If
  
  Exit Function
ErrorHandler:
  ' exit silently
End Function

' included instead of using Split() for the VB5 set :)
Private Sub ParseOnChar(ByRef sIn As String, ByRef sChar As String, ByVal lGuidIndex As Long)
  Dim lStartPosition As Long
  Dim lFoundPosition As Long
  Dim sItem          As String
  
  lFoundPosition = InStr(sIn, sChar)
  lStartPosition = 1
  
  Do While lFoundPosition
    sItem = Mid$(sIn, lStartPosition, lFoundPosition - lStartPosition)
    colCodecs.Add lGuidIndex, sItem
    lStartPosition = lFoundPosition + 1
    lFoundPosition = InStr(lStartPosition, sIn, sChar)
  Loop
  
  sItem = Trim$(Mid$(sIn, lStartPosition))
  If LenB(sItem) Then colCodecs.Add lGuidIndex, sItem
End Sub

Friend Function StringToGuid(ByRef sGuid As String) As CLSID
  CLSIDFromString sGuid, StringToGuid
End Function

Private Function StringFromPointerW(ByVal lPointer As Long) As String
  Dim lLength As Long
  
  If lPointer Then
    lLength = lstrlenW(lPointer)
    StringFromPointerW = Space$(lLength)
    RtlMoveMemory ByVal StrPtr(StringFromPointerW), ByVal lPointer, lLength * 2
  End If
End Function



Private Sub Class_Terminate()
  If lToken Then GdiplusShutdown lToken
End Sub
