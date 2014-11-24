VERSION 5.00
Begin VB.UserControl MyPalette 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   PropertyPages   =   "MyPalette.ctx":0000
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   Begin VB.HScrollBar HScroll 
      Height          =   247
      LargeChange     =   50
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2574
      Width           =   3640
   End
   Begin VB.VScrollBar VScroll 
      Height          =   2938
      LargeChange     =   50
      Left            =   3861
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   247
   End
   Begin VB.PictureBox Surface 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2119
      Left            =   234
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   0
      Top             =   234
      Width           =   3172
      Begin VB.Label Label 
         Height          =   481
         Left            =   1521
         TabIndex        =   3
         Top             =   1053
         Width           =   1300
      End
   End
   Begin VB.Image SurfaceBuffer 
      Height          =   364
      Left            =   3744
      Top             =   3159
      Visible         =   0   'False
      Width           =   715
   End
End
Attribute VB_Name = "MyPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Event BoxClicked(Button As Integer, Index As Byte)

Public Enum BorderConstants
    flat = 0
    Deep = 1
    Bump = 2
End Enum

Private c_BoxBorderStyle As BorderConstants

Private c_BoxCount As Byte
Private c_PaletteRGB(255) As Long

Private c_BoxSize As Long          'size of color Box

Private c_BoxSpacing As Long       'spacing between color Box

Dim c_BoxRight As Long
Dim c_BoxBottom As Long

Private c_BoxTop As Long             'topleft corner of color Box on picbox
Private c_BoxLeft As Long

Private c_BoxPerColumn As Single
Private c_BoxPerLine As Long    'how many color box per line
Private c_BackgrColor As Long     'background of palette display
Private c_BoxBorderColor As Long  'border color of color Box
Private c_BoxHasBorders As Boolean

Private c_BoxTextColor As Long
Private c_BoxHasText As Long

Private c_MouseoverBox As Byte

Private c_x As Long
Private c_y As Long

Private c_BoxStep As Long

Private c_BorderWidth

Private c_Enabled

'inner dimensions of a window
Private H As Long
Private W As Long


Public Property Get UCWidth() As Long
        '<EhHeader>
        On Error GoTo UCWidth_Err
        '</EhHeader>
100     UCWidth = UserControl.Width
        '<EhFooter>
        Exit Property

UCWidth_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.UCWidth", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Get SurfaceWidth() As Long
        '<EhHeader>
        On Error GoTo SurfaceWidth_Err
        '</EhHeader>
100     SurfaceWidth = Surface.Width
        '<EhFooter>
        Exit Property

SurfaceWidth_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.SurfaceWidth", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get SurfaceHeight() As Long
        '<EhHeader>
        On Error GoTo SurfaceHeight_Err
        '</EhHeader>
100     SurfaceHeight = Surface.Height
        '<EhFooter>
        Exit Property

SurfaceHeight_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.SurfaceHeight", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let Enabled(ByVal Number As Boolean)
        '<EhHeader>
        On Error GoTo Enabled_Err
        '</EhHeader>
100     c_Enabled = Number
102     UserControl.Enabled = c_Enabled
        '<EhFooter>
        Exit Property

Enabled_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.Enabled", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Get Enabled() As Boolean
        '<EhHeader>
        On Error GoTo Enabled_Err
        '</EhHeader>
100     Enabled = c_Enabled
        '<EhFooter>
        Exit Property

Enabled_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.Enabled", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxBorderStyle(ByVal Number As BorderConstants)
        '<EhHeader>
        On Error GoTo BoxBorderStyle_Err
        '</EhHeader>
100     c_BoxBorderStyle = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxBorderStyle_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxBorderStyle", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Get BoxBorderStyle() As BorderConstants
        '<EhHeader>
        On Error GoTo BoxBorderStyle_Err
        '</EhHeader>
100     BoxBorderStyle = c_BoxBorderStyle
        '<EhFooter>
        Exit Property

BoxBorderStyle_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxBorderStyle", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxCount(ByVal Number As Byte)
        '<EhHeader>
        On Error GoTo BoxCount_Err
        '</EhHeader>
100     c_BoxCount = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxCount_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxCount", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxCount() As Byte
        '<EhHeader>
        On Error GoTo BoxCount_Err
        '</EhHeader>
100     BoxCount = c_BoxCount
        '<EhFooter>
        Exit Property

BoxCount_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxCount", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let PaletteRGB(ByVal EntryNr As Byte, ByVal ColorRGB As Long)
        '<EhHeader>
        On Error GoTo PaletteRGB_Err
        '</EhHeader>
100     c_PaletteRGB(EntryNr) = ColorRGB
        'InitSurface
        '<EhFooter>
        Exit Property

PaletteRGB_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.PaletteRGB", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Get PaletteRGB(ByVal EntryNr As Byte) As Long
        '<EhHeader>
        On Error GoTo PaletteRGB_Err
        '</EhHeader>
100     PaletteRGB = c_PaletteRGB(EntryNr)
        '<EhFooter>
        Exit Property

PaletteRGB_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.PaletteRGB", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxSize(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BoxSize_Err
        '</EhHeader>
100     c_BoxSize = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxSize_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxSize", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxSize() As Long
        '<EhHeader>
        On Error GoTo BoxSize_Err
        '</EhHeader>
100     BoxSize = c_BoxSize
        '<EhFooter>
        Exit Property

BoxSize_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxSize", _
                  "MyPalette component failure"
        '</EhFooter>
End Property



Public Property Let BoxSpacing(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BoxSpacing_Err
        '</EhHeader>
100     c_BoxSpacing = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxSpacing_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxSpacing", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxSpacing() As Long
        '<EhHeader>
        On Error GoTo BoxSpacing_Err
        '</EhHeader>
100     BoxSpacing = c_BoxSpacing
        '<EhFooter>
        Exit Property

BoxSpacing_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxSpacing", _
                  "MyPalette component failure"
        '</EhFooter>
End Property



Public Property Let BoxTop(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BoxTop_Err
        '</EhHeader>
100     c_BoxTop = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxTop_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxTop", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxTop() As Long
        '<EhHeader>
        On Error GoTo BoxTop_Err
        '</EhHeader>
100     BoxTop = c_BoxTop
        '<EhFooter>
        Exit Property

BoxTop_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxTop", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxLeft(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BoxLeft_Err
        '</EhHeader>
100     c_BoxLeft = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxLeft_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxLeft", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxLeft() As Long
        '<EhHeader>
        On Error GoTo BoxLeft_Err
        '</EhHeader>
100     BoxLeft = c_BoxLeft
        '<EhFooter>
        Exit Property

BoxLeft_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxLeft", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxPerColumn(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BoxPerColumn_Err
        '</EhHeader>
100     c_BoxPerColumn = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxPerColumn_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxPerColumn", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxPerColumn() As Long
        '<EhHeader>
        On Error GoTo BoxPerColumn_Err
        '</EhHeader>
100     BoxPerColumn = c_BoxPerColumn
        '<EhFooter>
        Exit Property

BoxPerColumn_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxPerColumn", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxPerLine(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BoxPerLine_Err
        '</EhHeader>
100     c_BoxPerLine = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxPerLine_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxPerLine", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxPerLine() As Long
        '<EhHeader>
        On Error GoTo BoxPerLine_Err
        '</EhHeader>
100     BoxPerLine = c_BoxPerLine
        '<EhFooter>
        Exit Property

BoxPerLine_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxPerLine", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BackgrColor(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BackgrColor_Err
        '</EhHeader>
100     c_BackgrColor = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BackgrColor_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BackgrColor", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BackgrColor() As Long
Attribute BackgrColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
        '<EhHeader>
        On Error GoTo BackgrColor_Err
        '</EhHeader>
100     BackgrColor = c_BackgrColor
        '<EhFooter>
        Exit Property

BackgrColor_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BackgrColor", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxBorderColor(ByVal Number As Long)
        '<EhHeader>
        On Error GoTo BoxBorderColor_Err
        '</EhHeader>
100     c_BoxBorderColor = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxBorderColor_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxBorderColor", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxBorderColor() As Long
Attribute BoxBorderColor.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
        '<EhHeader>
        On Error GoTo BoxBorderColor_Err
        '</EhHeader>
100     BoxBorderColor = c_BoxBorderColor
        '<EhFooter>
        Exit Property

BoxBorderColor_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxBorderColor", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxHasBorders(ByVal Number As Boolean)
        '<EhHeader>
        On Error GoTo BoxHasBorders_Err
        '</EhHeader>
100     c_BoxHasBorders = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxHasBorders_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxHasBorders", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxHasBorders() As Boolean
        '<EhHeader>
        On Error GoTo BoxHasBorders_Err
        '</EhHeader>
100     BoxHasBorders = c_BoxHasBorders
        '<EhFooter>
        Exit Property

BoxHasBorders_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxHasBorders", _
                  "MyPalette component failure"
        '</EhFooter>
End Property

Public Property Let BoxHasText(ByVal Number As Boolean)
        '<EhHeader>
        On Error GoTo BoxHasText_Err
        '</EhHeader>
100     c_BoxHasText = Number
102     InitSurface
        '<EhFooter>
        Exit Property

BoxHasText_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxHasText", _
                  "MyPalette component failure"
        '</EhFooter>
End Property
Public Property Get BoxHasText() As Boolean
        '<EhHeader>
        On Error GoTo BoxHasText_Err
        '</EhHeader>
100     BoxHasText = c_BoxHasText
        '<EhFooter>
        Exit Property

BoxHasText_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.BoxHasText", _
                  "MyPalette component failure"
        '</EhFooter>
End Property


Private Sub Surface_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Surface_MouseMove_Err
        '</EhHeader>
    Dim Row As Long
    Dim Column As Long

100 c_x = X - c_BoxLeft + BoxSpacing / 2
102 c_y = Y - c_BoxTop + BoxSpacing / 2

104 Column = Int(c_x / c_BoxStep)
106 Row = Int(c_y / c_BoxStep)

108 c_MouseoverBox = (Column + Row * c_BoxPerLine) And 255

110 c_y = c_BoxTop + c_BoxStep * Row
112 c_x = c_BoxLeft + c_BoxStep * Column
       

114 If c_MouseoverBox <= c_BoxCount And X < c_BoxRight And X > c_BoxLeft And Y < c_BoxBottom And Y > c_BoxTop Then
116     Surface.Picture = SurfaceBuffer.Picture
118     Surface.DrawMode = 13 '7
120     Surface.Line (c_x, c_y)-(c_x + c_BoxSize, c_y + c_BoxSize), RGB(255, 255, 255), B
122     Surface.DrawMode = 13
    Else
124     Surface.Picture = SurfaceBuffer.Picture
    End If

        '<EhFooter>
        Exit Sub

Surface_MouseMove_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.Surface_MouseMove", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub

Private Sub Surface_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Surface_MouseUp_Err
        '</EhHeader>

100 If c_MouseoverBox <= c_BoxCount Then
102     RaiseEvent BoxClicked(Button, c_MouseoverBox)
    End If

        '<EhFooter>
        Exit Sub

Surface_MouseUp_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.Surface_MouseUp", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub





Private Sub UserControl_Initialize()
        '<EhHeader>
        On Error GoTo UserControl_Initialize_Err
        '</EhHeader>

100 Surface.AutoRedraw = True
102 c_BorderWidth = 4

        '<EhFooter>
        Exit Sub

UserControl_Initialize_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.UserControl_Initialize", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub
Public Function InitSurface()
        '<EhHeader>
        On Error GoTo InitSurface_Err
        '</EhHeader>
        Dim c_Entry As Byte


100     c_BoxStep = c_BoxSize + c_BoxSpacing

102     c_BoxRight = (c_BoxStep * c_BoxPerLine)
104     c_BoxPerColumn = Int(c_BoxCount / c_BoxPerLine) + 1

106     If c_BoxPerColumn <= c_BoxCount / c_BoxPerLine Then c_BoxPerColumn = c_BoxPerColumn + 1
108     c_BoxBottom = c_BoxTop + c_BoxStep * c_BoxPerColumn

110     Surface.Width = c_BoxLeft * 2 + c_BoxRight - c_BoxSpacing
112     Surface.Height = c_BoxTop * 2 + c_BoxPerColumn * c_BoxStep - c_BoxSpacing
114     Surface.BackColor = c_BackgrColor
116     Surface.Line (0, 0)-(Surface.Width, Surface.Height), c_BackgrColor, BF
    
118     UserControl.BackColor = c_BackgrColor

120     Surface.Move 0, 0


122     c_Entry = 0

124     Surface.Forecolor = RGB(255, 255, 255)

126     For c_y = c_BoxTop To c_BoxBottom Step c_BoxStep
128         For c_x = c_BoxLeft To c_BoxRight Step c_BoxStep

130             If c_Entry <= c_BoxCount Then

132                 Surface.Line (c_x, c_y)-(c_x + c_BoxSize, c_y + c_BoxSize), c_PaletteRGB(c_Entry), BF

134                 If c_BoxHasBorders = True Then

136                     Select Case c_BoxBorderStyle
                            Case flat
138                             Surface.Line (c_x, c_y)-(c_x + c_BoxSize, c_y + c_BoxSize), c_BoxBorderColor, B
140                         Case Deep
142                             Surface.Line (c_x, c_y)-(c_x, c_y + c_BoxSize), Darken(c_PaletteRGB(c_Entry))
144                             Surface.Line (c_x, c_y)-(c_x + c_BoxSize, c_y), Darken(c_PaletteRGB(c_Entry))
146                             Surface.Line (c_x + c_BoxSize, c_y + c_BoxSize)-(c_x, c_y + c_BoxSize), Lighten(c_PaletteRGB(c_Entry))
148                             Surface.Line (c_x + c_BoxSize, c_y + c_BoxSize)-(c_x, c_y + c_BoxSize), Lighten(c_PaletteRGB(c_Entry))
150                         Case Bump
152                             Surface.Line (c_x, c_y)-(c_x, c_y + c_BoxSize), Lighten(c_PaletteRGB(c_Entry))
154                             Surface.Line (c_x, c_y)-(c_x + c_BoxSize, c_y), Lighten(c_PaletteRGB(c_Entry))
156                             Surface.Line (c_x + c_BoxSize, c_y + c_BoxSize)-(c_x, c_y + c_BoxSize), Darken(c_PaletteRGB(c_Entry))
158                             Surface.Line (c_x + c_BoxSize, c_y + c_BoxSize)-(c_x, c_y + c_BoxSize), Darken(c_PaletteRGB(c_Entry))
                        End Select

                    End If

160                 If c_BoxHasText Then
162                     Surface.Forecolor = (Not c_PaletteRGB(c_Entry)) And (2 ^ 24 - 1)
164                     Surface.CurrentX = c_x + 2
166                     Surface.CurrentY = c_y + 2
168                     Surface.Print Hex$(c_Entry)
                    End If

                End If

170             c_Entry = c_Entry + 1

172         Next c_x
174     Next c_y

176     SurfaceBuffer.Picture = Surface.Image
    
178     Call GuiResize

        '<EhFooter>
        Exit Function

InitSurface_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.InitSurface", _
                  "MyPalette component failure"
        '</EhFooter>
End Function
Private Function Darken(Color As Long) As Long
        '<EhHeader>
        On Error GoTo Darken_Err
        '</EhHeader>
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte

100    R = Color And 255
102    G = Int(Color / 256) And 255
104    B = Int(Color / 65536) And 255
   
106    R = R * 0.5
108    G = G * 0.5
110    B = B * 0.5
   
112    Darken = RGB(R, G, B)
        '<EhFooter>
        Exit Function

Darken_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.Darken", _
                  "MyPalette component failure"
        '</EhFooter>
End Function

Private Function Lighten(Color As Long) As Long
        '<EhHeader>
        On Error GoTo Lighten_Err
        '</EhHeader>
    Dim R As Long
    Dim G As Long
    Dim B As Long

100    R = Color And 255
102    G = Int(Color / 256) And 255
104    B = Int(Color / 65536) And 255
   
106    R = R * 2
108    G = G * 2
110    B = B * 2
   
112    If R > 255 Then R = 255
114    If G > 255 Then G = 255
116    If B > 255 Then B = 255
   
118    Lighten = RGB(R, G, B)
        '<EhFooter>
        Exit Function

Lighten_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.Lighten", _
                  "MyPalette component failure"
        '</EhFooter>
End Function

Private Sub UserControl_InitProperties()
        '<EhHeader>
        On Error GoTo UserControl_InitProperties_Err
        '</EhHeader>
    Dim X As Long

100 c_BoxCount = 15

102 For X = 0 To 255
        'c_PaletteRGB(X) = 0
104 Next X

106 c_BoxSize = 16    'size of color Box

108 c_BoxSpacing = 4  'spacing between color Box

110 c_BoxTop = 4             'topleft corner of color Box on picbox
112 c_BoxLeft = 4

114 c_BoxPerColumn = 0
116 c_BoxPerLine = 8   'how many color box per line
118 c_BackgrColor = RGB(0, 0, 255)  'background of palette display
120 c_BoxBorderColor = RGB(255, 255, 255) 'border color of color Box
122 c_BoxHasBorders = True
124 c_BoxTextColor = RGB(255, 255, 255)
126 c_BoxHasText = True
128 c_BoxBorderStyle = flat

130 Call InitSurface

        '<EhFooter>
        Exit Sub

UserControl_InitProperties_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.UserControl_InitProperties", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_ReadProperties_Err
        '</EhHeader>

100 c_BoxCount = PropBag.ReadProperty("BoxCount")

102 c_BoxSize = PropBag.ReadProperty("BoxWidth")  'size of color Box

104 c_BoxSpacing = PropBag.ReadProperty("BoxHSpacing")  'spacing between color Box

106 c_BoxTop = PropBag.ReadProperty("BoxTop")             'topleft corner of color Box on picbox
108 c_BoxLeft = PropBag.ReadProperty("BoxLeft")

110 c_BoxPerLine = PropBag.ReadProperty("BoxPerLine")   'how many color box per line
112 c_BackgrColor = PropBag.ReadProperty("BackgrColor")  'background of palette display
114 c_BoxBorderColor = PropBag.ReadProperty("BoxBorderColor") 'border color of color Box
116 c_BoxHasBorders = PropBag.ReadProperty("BoxHasBorders")
118 c_BoxTextColor = PropBag.ReadProperty("BoxTextColor")
120 c_BoxHasText = PropBag.ReadProperty("BoxHasText")
122 c_BoxBorderStyle = PropBag.ReadProperty("BoxBorderStyle")

124 Call InitSurface

        '<EhFooter>
        Exit Sub

UserControl_ReadProperties_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.UserControl_ReadProperties", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_WriteProperties_Err
        '</EhHeader>


100 PropBag.WriteProperty "BoxCount", c_BoxCount

102 PropBag.WriteProperty "BoxWidth", c_BoxSize

104 PropBag.WriteProperty "BoxHSpacing", c_BoxSpacing

106 PropBag.WriteProperty "BoxTop", c_BoxTop
108 PropBag.WriteProperty "BoxLeft", c_BoxLeft

110 PropBag.WriteProperty "BoxPerLine", c_BoxPerLine
112 PropBag.WriteProperty "BackgrColor", c_BackgrColor

114 PropBag.WriteProperty "BoxborderColor", c_BoxBorderColor
116 PropBag.WriteProperty "BoxHasBorders", c_BoxHasBorders
118 PropBag.WriteProperty "BoxTextColor", c_BoxTextColor
120 PropBag.WriteProperty "BoxHasText", c_BoxHasText
122 PropBag.WriteProperty "BoxBorderStyle", c_BoxBorderStyle
        '<EhFooter>
        Exit Sub

UserControl_WriteProperties_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.UserControl_WriteProperties", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

Call GuiResize

End Sub

Private Sub GuiResize()
    Dim Dummy
    H = UserControl.ScaleHeight
    W = UserControl.ScaleWidth

    'scrollbar visibility
    If ((H < Surface.Height) Or (W < Surface.Width)) = True Then
        VScroll.Visible = True
        HScroll.Visible = True
    Else
        VScroll.Visible = False
        HScroll.Visible = False
    End If


    'scrollbar max value
    HScroll.Enabled = W < Surface.Width
    VScroll.Enabled = H < Surface.Height
    If HScroll.Enabled Then HScroll.Max = Surface.Width - W + VScroll.Width
    If VScroll.Enabled Then VScroll.Max = Surface.Height - H + HScroll.Height

    If VScroll.Visible Then Dummy = 1 Else Dummy = 0
    
    On Error Resume Next
    HScroll.Move 0, H - HScroll.Height, W - VScroll.Width * Dummy, HScroll.Height
    On Error GoTo 0

    If HScroll.Visible Then Dummy = 1 Else Dummy = 0
    If H - HScroll.Height * Dummy > 0 Then
        VScroll.Move W - VScroll.Width, 0, VScroll.Width, H - HScroll.Height * Dummy
    End If

    Call LabelPos

End Sub

Private Sub LabelPos()
        '<EhHeader>
        On Error GoTo LabelPos_Err
        '</EhHeader>
100 If VScroll.Visible Or HScroll.Visible Then
102     Label.Visible = True
        'Label.Move VScroll.Left - Surface.Left, HScroll.Top - Surface.Top
104     Label.Move VScroll.Left, HScroll.Top
    Else
106     Label.Visible = False
    End If
        '<EhFooter>
        Exit Sub

LabelPos_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.LabelPos", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub

Private Sub VScroll_Change()
        '<EhHeader>
        On Error GoTo VScroll_Change_Err
        '</EhHeader>
100 Surface.Top = -VScroll.Value
102 Call LabelPos
        '<EhFooter>
        Exit Sub

VScroll_Change_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.VScroll_Change", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub
Private Sub VScroll_scroll()
        '<EhHeader>
        On Error GoTo VScroll_scroll_Err
        '</EhHeader>
100 Surface.Top = -VScroll.Value
102 Call LabelPos
        '<EhFooter>
        Exit Sub

VScroll_scroll_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.VScroll_scroll", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub
Private Sub HScroll_Change()
        '<EhHeader>
        On Error GoTo HScroll_Change_Err
        '</EhHeader>
100 Surface.Left = -HScroll.Value
102 Call LabelPos
        '<EhFooter>
        Exit Sub

HScroll_Change_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.HScroll_Change", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub
Private Sub HScroll_scroll()
        '<EhHeader>
        On Error GoTo HScroll_scroll_Err
        '</EhHeader>
100 Surface.Left = -HScroll.Value
102 Call LabelPos
        '<EhFooter>
        Exit Sub

HScroll_scroll_Err:
        VBA.Err.Raise vbObjectError + 100, _
                  "Project_One.MyPalette.HScroll_scroll", _
                  "MyPalette component failure"
        '</EhFooter>
End Sub



