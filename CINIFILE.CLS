VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInifile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' =========================================================
' Class:    cIniFile
' Author:   Steve McMahon
' Date  :   21 Feb 1997
'
' A nice class wrapper around the INIFile functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
' Updated 10 May 1998 for VB5.
'   * Added EnumerateAllSections method
'   * Added Load and Save form position methods
' =========================================================

Private m_sPath As String
Private m_sKey As String
Private m_sSection As String
Private m_sDefault As String
Private m_lLastReturnCode As Long

#If Win32 Then
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    ' Profile String functions:
    Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If

Property Get LastReturnCode() As Long
        '<EhHeader>
        On Error GoTo LastReturnCode_Err
        '</EhHeader>
100     LastReturnCode = m_lLastReturnCode
        '<EhFooter>
        Exit Property

LastReturnCode_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.LastReturnCode", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Get Success() As Boolean
        '<EhHeader>
        On Error GoTo Success_Err
        '</EhHeader>
100     Success = (m_lLastReturnCode <> 0)
        '<EhFooter>
        Exit Property

Success_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Success", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Let Default(sDefault As String)
        '<EhHeader>
        On Error GoTo Default_Err
        '</EhHeader>
100     m_sDefault = sDefault
        '<EhFooter>
        Exit Property

Default_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Default", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Get Default() As String
        '<EhHeader>
        On Error GoTo Default_Err
        '</EhHeader>
100     Default = m_sDefault
        '<EhFooter>
        Exit Property

Default_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Default", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Let Path(sPath As String)
        '<EhHeader>
        On Error GoTo Path_Err
        '</EhHeader>
100     m_sPath = sPath
        '<EhFooter>
        Exit Property

Path_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Path", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Get Path() As String
        '<EhHeader>
        On Error GoTo Path_Err
        '</EhHeader>
100     Path = m_sPath
        '<EhFooter>
        Exit Property

Path_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Path", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Let Key(sKey As String)
        '<EhHeader>
        On Error GoTo Key_Err
        '</EhHeader>
100     m_sKey = sKey
        '<EhFooter>
        Exit Property

Key_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Key", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Get Key() As String
        '<EhHeader>
        On Error GoTo Key_Err
        '</EhHeader>
100     Key = m_sKey
        '<EhFooter>
        Exit Property

Key_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Key", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Let Section(sSection As String)
        '<EhHeader>
        On Error GoTo Section_Err
        '</EhHeader>
100     m_sSection = sSection
        '<EhFooter>
        Exit Property

Section_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Section", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Get Section() As String
        '<EhHeader>
        On Error GoTo Section_Err
        '</EhHeader>
100     Section = m_sSection
        '<EhFooter>
        Exit Property

Section_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Section", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Get Value() As String
        '<EhHeader>
        On Error GoTo Value_Err
        '</EhHeader>
    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer

100     sBuf = Space$(255)
102     iSize = Len(sBuf)
104     iRetCode = GetPrivateProfileString(m_sSection, m_sKey, m_sDefault, sBuf, iSize, m_sPath)
106     If (iSize > 0) Then
108         Value = Left$(sBuf, iRetCode)
        Else
110         Value = ""
        End If

        '<EhFooter>
        Exit Property

Value_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Value", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Let Value(sValue As String)
        '<EhHeader>
        On Error GoTo Value_Err
        '</EhHeader>
    Dim iPos As Integer
         'Strip Chr$(0):
100     iPos = InStr(sValue, Chr$(0))
102     Do While iPos <> 0
104         sValue = Left$(sValue, (iPos - 1)) & Mid$(sValue, (iPos + 1))
106         iPos = InStr(sValue, Chr$(0))
        Loop
108     m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, sValue, m_sPath)
        '<EhFooter>
        Exit Property

Value_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Value", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Public Sub DeleteKey()
        '<EhHeader>
        On Error GoTo DeleteKey_Err
        '</EhHeader>
100     m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, 0&, m_sPath)
        '<EhFooter>
        Exit Sub

DeleteKey_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.DeleteKey", _
                  "cInifile component failure"
        '</EhFooter>
End Sub
Public Sub DeleteSection()
        '<EhHeader>
        On Error GoTo DeleteSection_Err
        '</EhHeader>
100     m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, 0&, m_sPath)
        '<EhFooter>
        Exit Sub

DeleteSection_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.DeleteSection", _
                  "cInifile component failure"
        '</EhFooter>
End Sub
Property Get INISection() As String
        '<EhHeader>
        On Error GoTo INISection_Err
        '</EhHeader>
    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer

100     sBuf = Space$(8192)
102     iSize = Len(sBuf)
104     iRetCode = GetPrivateProfileString(m_sSection, 0&, m_sDefault, sBuf, iSize, m_sPath)
106     If (iSize > 0) Then
108         INISection = Left$(sBuf, iRetCode)
        Else
110         INISection = ""
        End If

        '<EhFooter>
        Exit Property

INISection_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.INISection", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Let INISection(sSection As String)
        '<EhHeader>
        On Error GoTo INISection_Err
        '</EhHeader>
100     m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, sSection, m_sPath)
        '<EhFooter>
        Exit Property

INISection_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.INISection", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Property Get Sections() As String
        '<EhHeader>
        On Error GoTo Sections_Err
        '</EhHeader>
    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer

100     sBuf = Space$(8192)
102     iSize = Len(sBuf)
104     iRetCode = GetPrivateProfileString(0&, 0&, m_sDefault, sBuf, iSize, m_sPath)
106     If (iSize > 0) Then
108         Sections = Left$(sBuf, iRetCode)
        Else
110         Sections = ""
        End If

        '<EhFooter>
        Exit Property

Sections_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.Sections", _
                  "cInifile component failure"
        '</EhFooter>
End Property
Public Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef iCount As Long)
        '<EhHeader>
        On Error GoTo EnumerateCurrentSection_Err
        '</EhHeader>
    Dim sSection As String
    Dim iPos As Long
    Dim iNextPos As Long
    Dim sCur As String
    
100     iCount = 0
102     Erase sKey
104     sSection = INISection
106     If (Len(sSection) > 0) Then
108         iPos = 1
110         iNextPos = InStr(iPos, sSection, Chr$(0))
112         Do While iNextPos <> 0
114             sCur = Mid$(sSection, iPos, (iNextPos - iPos))
116             If (sCur <> Chr$(0)) Then
118                 iCount = iCount + 1
120                 ReDim Preserve sKey(1 To iCount) As String
122                 sKey(iCount) = Mid$(sSection, iPos, (iNextPos - iPos))
124                 iPos = iNextPos + 1
126                 iNextPos = InStr(iPos, sSection, Chr$(0))
                End If
            Loop
        End If
        '<EhFooter>
        Exit Sub

EnumerateCurrentSection_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.EnumerateCurrentSection", _
                  "cInifile component failure"
        '</EhFooter>
End Sub
Public Sub EnumerateAllSections(ByRef sSections() As String, ByRef iCount As Long)
        '<EhHeader>
        On Error GoTo EnumerateAllSections_Err
        '</EhHeader>
    Dim sIniFile As String
    Dim iPos As Long
    Dim iNextPos As Long
    Dim sCur As String
    
100     iCount = 0
102     Erase sSections
104     sIniFile = Sections
106     If (Len(sIniFile) > 0) Then
108         iPos = 1
110         iNextPos = InStr(iPos, sIniFile, Chr$(0))
112         Do While iNextPos <> 0
114             If (iNextPos <> iPos) Then
116                 sCur = Mid$(sIniFile, iPos, (iNextPos - iPos))
118                 iCount = iCount + 1
120                 ReDim Preserve sSections(1 To iCount) As String
122                 sSections(iCount) = sCur
                End If
124             iPos = iNextPos + 1
126             iNextPos = InStr(iPos, sIniFile, Chr$(0))
            Loop
        End If

        '<EhFooter>
        Exit Sub

EnumerateAllSections_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.cInifile.EnumerateAllSections", _
                  "cInifile component failure"
        '</EhFooter>
End Sub
Public Sub SaveFormPosition(ByRef frmThis As Object)
Dim sSaveKey As String
Dim sSaveDefault As String
On Error GoTo SaveError
    sSaveKey = Key
    If Not (frmThis.WindowState = vbMinimized) Then
        Key = "Maximised"
        Value = (frmThis.WindowState = vbMaximized) * -1
        If (frmThis.WindowState <> vbMaximized) Then
            Key = "Left"
            Value = frmThis.Left
            Key = "Top"
            Value = frmThis.Top
            Key = "Width"
            Value = frmThis.Width
            Key = "Height"
            Value = frmThis.Height
        End If
    End If
    Key = sSaveKey
    Exit Sub
SaveError:
    Key = sSaveKey
    m_lLastReturnCode = 0
    Exit Sub
End Sub
Public Sub LoadFormPosition(ByRef frmThis As Object, Optional ByRef lMinWidth = 0, Optional ByRef lMinHeight = 0)
Dim sSaveKey As String
Dim sSaveDefault As String
Dim lLeft As Long
Dim lTop As Long
Dim lWidth As Long
Dim lHeight As Long
On Error GoTo LoadError
    sSaveKey = Key
    sSaveDefault = Default
    Default = "FAIL"
    Key = "Left"
    lLeft = CLngDefault(Value, frmThis.Left)
    Key = "Top"
    lTop = CLngDefault(Value, frmThis.Top)
    Key = "Width"
    lWidth = CLngDefault(Value, frmThis.Width)
    'MsgBox Str(lWidth)
    If (lWidth < lMinWidth) Then lWidth = lMinWidth
    Key = "Height"
    lHeight = CLngDefault(Value, frmThis.Height)
    If (lHeight < lMinHeight) Then lHeight = lMinHeight
    If (lLeft < 0 * Screen.TwipsPerPixelX) Then lLeft = 0 * Screen.TwipsPerPixelX
    If (lTop < 0 * Screen.TwipsPerPixelY) Then lTop = 0 * Screen.TwipsPerPixelY
    If (lLeft + lWidth > Screen.Width - 4 * Screen.TwipsPerPixelX) Then
        lLeft = Screen.Width - 4 * Screen.TwipsPerPixelX - lWidth
        If (lLeft < 0 * Screen.TwipsPerPixelX) Then lLeft = 0 * Screen.TwipsPerPixelX
        If (lLeft + lWidth > Screen.Width - 4 * Screen.TwipsPerPixelX) Then
            lWidth = Screen.Width - lLeft - 4 * Screen.TwipsPerPixelX
        End If
    End If
    If (lTop + lHeight > Screen.Height - 4 * Screen.TwipsPerPixelY) Then
        lTop = Screen.Height - 4 * Screen.TwipsPerPixelY - lHeight
        If (lTop < 4 * Screen.TwipsPerPixelY) Then lTop = 4 * Screen.TwipsPerPixelY
        If (lTop + lHeight > Screen.Height - 4 * Screen.TwipsPerPixelY) Then
            lHeight = Screen.Height - lTop - 4 * Screen.TwipsPerPixelY
        End If
    End If
    If (lWidth >= lMinWidth) And (lHeight >= lMinHeight) Then
        frmThis.Move lLeft, lTop, lWidth, lHeight
    End If
    Key = "Maximised"
    If (CLngDefault(Value, 0) <> 0) Then
        frmThis.WindowState = vbMaximized
    End If
    Key = sSaveKey
    Default = sSaveDefault
    Exit Sub
LoadError:
    Key = sSaveKey
    Default = sSaveDefault
    m_lLastReturnCode = 0
    Exit Sub
End Sub
Public Function CLngDefault(ByVal sString As String, Optional ByVal lDefault As Long = 0) As Long
Dim lR As Long
On Error Resume Next
    lR = CLng(sString)
    If (Err.Number <> 0) Then
        CLngDefault = lDefault
    Else
        CLngDefault = lR
    End If
End Function
