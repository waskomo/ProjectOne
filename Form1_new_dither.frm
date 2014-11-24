VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DDA53BD0-2CD0-11D4-8ED4-00E07D815373}#1.0#0"; "MBMouse.ocx"
Begin VB.Form Form1 
   Caption         =   "Zoom Window"
   ClientHeight    =   4425
   ClientLeft      =   510
   ClientTop       =   3810
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MouseIcon       =   "Form1_new_dither.frx":0000
   ScaleHeight     =   4425
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6240
      Top             =   2400
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   16
      Left            =   10200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Buffer 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   1575
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   0
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3975
      LargeChange     =   32
      Left            =   9720
      Max             =   200
      SmallChange     =   4
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Value           =   20
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Text            =   "Form1_new_dither.frx":1CCA
      Top             =   0
      Width           =   3735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   80
      Left            =   0
      Max             =   320
      SmallChange     =   4
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3960
      Width           =   9735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.bmp"
      Filter          =   "bmp,gif"
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   15
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   3360
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      HasDC           =   0   'False
      Height          =   255
      Index           =   14
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   3360
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   13
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   12
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   11
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   10
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   9
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   8
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   7
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   6
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   5
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   4
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   3
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   2
      Left            =   10080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Index           =   1
      Left            =   10440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Display 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   2055
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox ZoomPic 
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   2760
      MouseIcon       =   "Form1_new_dither.frx":1CD2
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   0
      Top             =   840
      Width           =   2820
   End
   Begin VB.Frame Frame1 
      Caption         =   "colors"
      Height          =   3135
      Left            =   9960
      TabIndex        =   17
      Top             =   600
      Width           =   855
   End
   Begin MBMouseHelper.MouseHelper MouseHelper1 
      Left            =   4920
      Top             =   1800
      _ExtentX        =   900
      _ExtentY        =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'todo:

'- 4th color: replace, find closest
'- save/load modes in 64 native format
'- reconvert after fill shouldnt optimize d021
'- undo for d021
'- why undo is enabled after resetundo ?
'- brush: if whole char covered it should overwrite it without pixelfit
'- screen array: add screen index for closer 64 modell
'- debug delete char
'- copy area
'- cursor shows drawmode
'- fix fill writes undo twice
'- debug save koala

'declares for 8 bit paletted image

'Option Explicit


'notes:
'pixels(319,199): array of pic being drawn
'colmap(319,199): array of colors converted when loading pic

'to prevent close button to appear
Private Const SC_CLOSE As Long = &HF060&
Private Const MF_BYCOMMAND = &H0&
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Dim IncColors(15) As Byte
Dim ChoX As Integer
Dim ChoY As Integer

Const UnUsed = 16

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

Dim tsain As SAFEARRAY2D
Dim Pixels() As Byte
Public PixelsDib As New cDIBSection256

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' to get the inner dimensions of a window
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Dim Rectangle As RECT

' own stuff
Dim ox As Integer
Dim oy As Integer

Dim ScrBanks As Byte
Dim BitBanks As Byte

Dim ScrBank As Byte
Dim BitBank As Byte

Dim TextLine As String ', fil1 As File, ts As TextStream
Dim coX As Integer
Dim coY As Integer
'Dim woX As Integer
'Dim woY As Integer
Dim Color As Long
Dim LeftCol As Long
Dim RightCol As Long
Public LeftColN As Byte
Dim RightColN As Byte
Dim PaintNow As Boolean
Dim PressedKey As String, Filename As String, CopyMode As String
Dim c64cols(15) As Long
Dim ActCol As Long, actcolor As Long
Dim ColNum As Byte
Public BackgrIndex As Byte
Public ZoomWinLeft As Integer, ZoomWinTop As Integer
Dim Ax As Integer, Ay As Integer
Dim OldAx As Integer, OldAy As Integer
Dim cx As Integer, cy As Integer
Dim ZoomPicAbsX As Integer, ZoomPicAbsY As Integer
Dim cols(15) As Long, actcols(2) As Long
Dim U As Integer
Dim cd(15, 15) As Integer, d(3) As Integer, c(3) As Integer
Dim ColMap(319, 199) As Byte ', pixels(319, 199) As Byte
Dim cor(15) As Long, cog(15) As Long, cob(15) As Long
Dim try(15) As Long
Dim err As Double
Dim Draw As Boolean
Dim Pic1Click As Boolean, PaintPic1 As Boolean

Dim OldZoomX As Integer, OldZoomY As Integer
Dim DontUpdate As Boolean
Dim MostFreqCol As Long

Public Reso_PixLeft As Boolean, Reso_PixRight As Boolean, Reso_PixAvg As Boolean
Public optbackgr As Boolean, mostfreqbackgr As Boolean, usrdefbackgr As Boolean
Public UsrBackgr As Byte 'user selected backgr when converting
Public mostfrq3 As Boolean, optchars As Boolean
Public FixWithBackgr As Byte
Public Force_Greys As Boolean
Public Tresh_Grey As Integer
Public Dither As Boolean
Public Uni_Dither As Boolean
Public Avg_Dither As Boolean
Public Tresh_Dith As Integer
Public Uni_Strength As Byte

Dim fso As Variant

Dim DitherMode As Boolean
Dim FastSetPixel As Boolean
Dim pattern(1, 1) As Byte

Dim FillCol As Byte, Area As Integer
Dim FillMode As String
Public PicName As String

Dim PixelGrid As Boolean
Dim CharGrid As Boolean
Dim PixelBox As Boolean

Dim GfxMode As String

Public ResoDiv As Byte
'Dim OldY As Long
'Dim OldX As Long
Dim x As Integer
Dim y As Integer

Dim FliMul As Byte
Public ZoomHeight As Byte
Public ZoomWidth As Byte
Public ZoomScale As Long
Dim Brush(15, 15) As Byte
Dim UsedCols(39, 24, 16) As Byte
Dim d800(39, 24) As Byte

'Private Type colram_type
'Color As Byte
'Count As Byte
'End Type

'Dim ColRam(39, 24) As colram_type
Dim ScrRam(1, 39, 199, 1) As Byte '96000 byte
Dim Bank As Byte
Dim OldBank As Byte


Dim XFliLimit As Byte
Dim DitherFill As Boolean
Dim Tx As Integer

Public StretchPic As Boolean
Public KeepAspect As Boolean

'bitmap 16000
'scrram 32000
'pixels 64000
'd800    1000
'------------
'      113000

Dim UnBitMap(1, 39, 199, 15) As Byte
Dim UnScrRam(1, 39, 199, 1, 15) As Byte
Dim UnPixels(319, 199, 15) As Byte
Dim UnD800(39, 24, 15) As Byte

Dim RedoCount As Byte
Dim UndoPtr As Byte
Dim UndoOffs As Byte

Dim FillChangedPic As Boolean
Dim Sx As Integer
Dim Sy As Integer


' new attribute system

          'bank,x,y
Dim BITMAP(1, 39, 199) As Byte

Dim BitCount00(255) As Byte
Dim BitCount01(255) As Byte
Dim BitCount10(255) As Byte
Dim BitCount11(255) As Byte

Dim BitCount1(255) As Byte
Dim BitCount0(255) As Byte


Dim Bp00 As Integer
Dim Bp01 As Integer
Dim Bp10 As Integer
Dim Bp11 As Integer

Dim Bp1 As Integer
Dim Bp0 As Integer

Dim Shift1(7) As Byte
Dim Mask0(7) As Byte

Dim Shift(3, 3) As Byte
Dim Mask00(3) As Byte

'----------------- toolbar variables

Public PaintMode As String

'----------------- brush stuff
Public Br As Double
Public Bdither As Byte
Public Bsize As Double
Public Bcolor1 As Byte
Public Bcolor2 As Byte
Dim Barray(31, 31) As Byte


Public Sub BrushPreCalc()
Dim x As Double
Dim y As Double
Dim bayer(3, 3) As Byte
Dim d As Double
Dim round As Double
Dim Color As Byte

For x = 0 To 31
 For y = 0 To 31
 Barray(x, y) = 0
 Next y
Next x

round = Int((Bsize / 2) / ResoDiv) * ResoDiv

For x = -round To round Step ResoDiv
 For y = -round To round Step 1

  d = Sqr(x * x + y * y)

  If d < Bsize / 2 Then
   
   Barray(Int(31 / 2) + x, Int(31 / 2) + y) = 1
   
   If ResoDiv = 2 Then _
   Barray(Int(31 / 2) + x + 1, Int(31 / 2) + y) = 1
  
  End If

 Next y
Next x

End Sub
Private Sub PlotBrush(ByVal pbx, ByVal pby)
Dim px As Byte
Dim py As Byte
Dim bayer(3, 3) As Byte
Dim pcol As Byte
Dim fx As Integer
Dim fy As Integer
Dim qx As Integer
Dim qy As Integer

bayer(0, 0) = 1: bayer(0, 1) = 9: bayer(0, 2) = 3: bayer(0, 3) = 11
bayer(1, 0) = 13: bayer(1, 1) = 5: bayer(1, 2) = 15: bayer(1, 3) = 7
bayer(2, 0) = 4: bayer(2, 1) = 12: bayer(2, 2) = 2: bayer(2, 3) = 10
bayer(3, 0) = 16: bayer(3, 1) = 8: bayer(3, 2) = 14: bayer(3, 3) = 6

FastSetPixel = True
'If ResoDiv = 2 Then pbx = pbx And 510

For px = 0 To 31 Step ResoDiv
 For py = 0 To 31
 
  If Barray(px, py) = 1 Then
     
   qx = px + pbx - 16
   qy = py + pby - 16
   
   If qx >= 0 And qx <= 319 And _
   qy >= 0 And qy <= 199 Then
   
   fx = (Int((qx) / ResoDiv) * 1) And 3
   fy = (qy) And 3
   If Bdither < bayer(fx And 3, fy And 3) Then _
   pcol = Bcolor1 Else _
   pcol = Bcolor2
     
   Call Setpixel(qx, qy, pcol)
   'If PixelFits(qx, qy, pcol) Then
   '   Pixels(qx, qy) = pcol
   '   If ResoDiv = 2 Then Pixels(qx + 1, qy) = pcol
   'End If
   
   End If
   
  End If
  
 Next py
Next px

FastSetPixel = False

End Sub
Private Sub ResetUndo()

MDIForm1.mnuRedo.Enabled = False
'MDIForm1.mnuUndo.Enabled = False

UndoPtr = 0
UndoOffs = 0

CopyMemory UnBitMap(0, 0, 0, Ptr), ByVal VarPtr(BITMAP(0, 0, 0)), LenB(BITMAP(0, 0, 0)) * 15999
CopyMemory UnScrRam(0, 0, 0, 0, Ptr), ByVal VarPtr(ScrRam(0, 0, 0, 0)), LenB(ScrRam(0, 0, 0, 0)) * 31999
CopyMemory UnPixels(0, 0, Ptr), ByVal VarPtr(Pixels(0, 0)), LenB(Pixels(0, 0)) * 63999
CopyMemory UnD800(0, 0, Ptr), ByVal VarPtr(d800(0, 0)), LenB(d800(0, 0)) * 9999

Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
End Sub
Private Sub WriteUndo()
Dim Ptr As Byte

RedoCount = 0
MDIForm1.mnuRedo.Enabled = False

If UndoPtr < 15 Then
 UndoPtr = UndoPtr + 1
Else
 UndoOffs = (UndoOffs + 1) And 15
End If

Ptr = (UndoPtr + UndoOffs) And 15

'bitmap 16000
'scrram 32000
'pixels 64000
'd800    1000
'------------
'      113000

CopyMemory UnBitMap(0, 0, 0, Ptr), ByVal VarPtr(BITMAP(0, 0, 0)), LenB(BITMAP(0, 0, 0)) * 16000
CopyMemory UnScrRam(0, 0, 0, 0, Ptr), ByVal VarPtr(ScrRam(0, 0, 0, 0)), LenB(ScrRam(0, 0, 0, 0)) * 32000
CopyMemory UnPixels(0, 0, Ptr), ByVal VarPtr(Pixels(0, 0)), LenB(Pixels(0, 0)) * 64000
CopyMemory UnD800(0, 0, Ptr), ByVal VarPtr(d800(0, 0)), LenB(d800(0, 0)) * 1000

Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
End Sub
Private Sub ReadUndo()
Dim Ptr As Byte

MDIForm1.mnuRedo.Enabled = True

If UndoPtr > 0 Then
 UndoPtr = UndoPtr - 1
 RedoCount = RedoCount + 1
End If


Ptr = (UndoPtr + UndoOffs) And 15

CopyMemory BITMAP(0, 0, 0), ByVal VarPtr(UnBitMap(0, 0, 0, Ptr)), LenB(UnBitMap(0, 0, 0, 0)) * 16000
CopyMemory ScrRam(0, 0, 0, 0), ByVal VarPtr(UnScrRam(0, 0, 0, 0, Ptr)), LenB(UnScrRam(0, 0, 0, 0, 0)) * 32000
CopyMemory Pixels(0, 0), ByVal VarPtr(UnPixels(0, 0, Ptr)), LenB(UnPixels(0, 0, 0)) * 64000
CopyMemory d800(0, 0), ByVal VarPtr(UnD800(0, 0, Ptr)), LenB(UnD800(0, 0, 0)) * 1000

'If UndoPtr > 0 Then UndoPtr = UndoPtr - 1

Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
End Sub

Private Sub Redo()
Dim Ptr As Byte

If RedoCount > 0 Then
 'UndoPtr = UndoPtr - 1
 RedoCount = RedoCount - 1
 UndoPtr = UndoPtr + 1
End If

If RedoCount <> 0 Then
 MDIForm1.mnuRedo.Enabled = True
Else
 MDIForm1.mnuRedo.Enabled = False
End If

Ptr = (UndoPtr + UndoOffs) And 15

CopyMemory BITMAP(0, 0, 0), ByVal VarPtr(UnBitMap(0, 0, 0, Ptr)), LenB(UnBitMap(0, 0, 0, 0)) * 16000
CopyMemory ScrRam(0, 0, 0, 0), ByVal VarPtr(UnScrRam(0, 0, 0, 0, Ptr)), LenB(UnScrRam(0, 0, 0, 0, 0)) * 32000
CopyMemory Pixels(0, 0), ByVal VarPtr(UnPixels(0, 0, Ptr)), LenB(UnPixels(0, 0, 0)) * 64000
CopyMemory d800(0, 0), ByVal VarPtr(UnD800(0, 0, Ptr)), LenB(UnD800(0, 0, 0)) * 1000


Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
End Sub

Public Sub mnuRedo_Click()
Call Redo
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call Form2.ReDrawDisp

End Sub

Public Sub mnuUndo_Click()

Call ReadUndo

Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call Form2.ReDrawDisp

End Sub

Private Sub grid(ZoomX, ZoomY)
Dim XX As Integer
Dim YY As Integer

ZoomPic.Scale (0, 0)-(ZoomWidth, ZoomHeight)

gridx = 8 - (Int(ZoomX) Mod 8)
gridy = 8 - (Int(ZoomY) Mod 8)


'Draw pixel grid
If PixelGrid = True Then
For XX = 0 To ZoomWidth Step ResoDiv
 'If (XX + (8 - gridx)) Mod 8 <> 0 Then ZoomPic.Line (XX, 0)-(XX, ZoomHeight), 0
 ZoomPic.Line (XX, 0)-(XX, ZoomHeight), 0
Next XX

For YY = 0 To ZoomHeight
 'If (YY + (8 - gridy)) Mod 8 <> 0 Then ZoomPic.Line (0, YY)-(ZoomWidth, YY), 0
 ZoomPic.Line (0, YY)-(ZoomWidth, YY), 0
Next YY
End If


'Draw charborders grid
If CharGrid = True Then
'ZoomPic.DrawMode = 7 '7=eor,13=normal
For XX = gridx To ZoomWidth Step 8  '/ ResoDiv
 ZoomPic.Line (XX, 0)-(XX, ZoomHeight), RGB(188, 188, 188)
 'ZoomPic.Line (XX, 0)-(XX, ZoomHeight), RGB(255, 255, 255)
Next XX

For YY = gridy To ZoomHeight Step 8 '8
 ZoomPic.Line (0, YY)-(ZoomWidth, YY), RGB(188, 188, 188)
 'ZoomPic.Line (0, YY)-(ZoomWidth, YY), RGB(255, 255, 255)
Next YY
'ZoomPic.DrawMode = 13 '7=eor,13=normal
End If

End Sub

Private Sub CoordLimit(x, y)
x = Int(x / ResoDiv) * ResoDiv
y = Int(y)
If x >= 320 - ZoomWidth Then x = 320 - ZoomWidth
'If x < XFliLimit Then x = XFliLimit
If x < 0 Then x = 0
If y >= 200 - ZoomHeight Then y = 200 - ZoomHeight
If y < 0 Then y = 0
End Sub
Private Sub ZoomWin(ZoomX, ZoomY)
 Dim zw As Integer, zh As Integer, TempZoomX As Integer
 
 Call CoordLimit(ZoomX, ZoomY)
 
 TempZoomX = Int(ZoomX / ResoDiv) * ResoDiv
 
 zw = ZoomWidth * ZoomScale
 zh = ZoomHeight * ZoomScale
 ZoomPic.Scale (0, 0)-(zw, zh)
 StretchBlt ZoomPic.hdc, 0, 0, zw, zh, PixelsDib.hdc, TempZoomX, ZoomY, ZoomWidth, ZoomHeight, vbSrcCopy
 ZoomPic.Scale (0, 0)-(ZoomWidth, ZoomHeight)
 
 HScroll1.Value = TempZoomX
 VScroll1.Value = ZoomY
 ZoomWinLeft = HScroll1.Value
 ZoomWinTop = VScroll1.Value
 Call grid(ZoomWinLeft, ZoomWinTop)
 Call Form2.ReDrawDisp
 
End Sub




Public Sub Display_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
If x > 319 Then x = 319 - ResoDiv
If y > 199 Then y = 199
If x < 0 Then x = 0
If y < 0 Then y = 0

Ax = Int(x / ResoDiv) * ResoDiv: Ay = Int(y)

'statusbar update
MDIForm1.StatusBar1.Panels(5).Text = "x:" + Str(Ax / ResoDiv)
MDIForm1.StatusBar1.Panels(6).Text = "y:" + Str(Ay)
MDIForm1.StatusBar1.Panels(7).Text = "c:" + Str(Int(Ax / 8))
MDIForm1.StatusBar1.Panels(8).Text = "r:" + Str(Int(Ay / 8))

Call CharCols

If Pic1Click = True Then
 Call ZoomWin(Ax - ZoomWidth / 2, Ay - ZoomHeight / 2)
 ZoomPic.Refresh
End If
 
If PaintPic1 = True Then
 If Sqr((Ax - OldAx) ^ 2 + (Ay - OldAy) ^ 2) < -1 Then
    Call Paint(Ax, Ay, LeftColN)
 Else
    Call DrawLine(OldAx, OldAy, Ax, Ay, LeftColN)
    Call Form2.ReDrawDisp
 End If
End If

Call Form2.ShowCursor(Ax, Ay)
OldAx = Ax: OldAy = Ay

End Sub
Public Sub Display_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 


OldAx = Ax
OldAy = Ay

If PaintMode = "draw" Then

 If Button = 2 Then
  DontUpdate = True
  Pic1Click = True
  Call ZoomWin(Ax - (ZoomWidth / 2), Ay - (ZoomHeight / 2))
 End If
 
 If Button = 1 Then
  PaintPic1 = True
  ColNum = LeftColN
  Call Setpixel(Ax, Ay, LeftColN)
 End If
 
ElseIf PaintMode = "brush" Then

 If Button = 2 Then
  DontUpdate = True
  Pic1Click = True
  Call ZoomWin(Ax - (ZoomWidth / 2), Ay - (ZoomHeight / 2))
 End If
 
 If Button = 1 Then
  PaintPic1 = True
  Call PlotBrush(Ax, Ay)
 End If

ElseIf PaintMode = "fill" Then
   
   If Button = 1 Then
    FillCol = LeftColN
    Call StartFill
   End If
   
  If Button = 2 Then
   DontUpdate = True
   Pic1Click = True
   Call ZoomWin(Ax - (ZoomWidth / 2), Ay - (ZoomHeight / 2))
  End If

End If



End Sub

Public Sub Display_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
 Pic1Click = False   ' Turn off painting.
 DontUpdate = False  ' Let scrollers call zoomwin
 Call Form2.ReDrawDisp
ElseIf Button = 1 Then
 PaintPic1 = False
 Call ZoomWin(ZoomWinLeft, ZoomWinTop)
 Call WriteUndo
End If

End Sub

Public Sub mnuBrush_Click()
BrushDialog.Show vbModal
End Sub

Public Sub mnuDitherfill_Click()
MDIForm1.mnuDitherfill.Checked = Not MDIForm1.mnuDitherfill.Checked
DitherFill = Not DitherFill
End Sub

Public Sub mnuPixelGrid_Click()
MDIForm1.mnuPixelGrid.Checked = Not MDIForm1.mnuPixelGrid.Checked
PixelGrid = MDIForm1.mnuPixelGrid.Checked
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
End Sub

Public Sub mnuCharGrid_Click()
MDIForm1.mnuCharGrid.Checked = Not MDIForm1.mnuCharGrid.Checked
CharGrid = MDIForm1.mnuCharGrid.Checked
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
End Sub
Public Sub mnuPixelBox_Click()
MDIForm1.mnuPixelBox.Checked = Not MDIForm1.mnuPixelBox.Checked
PixelBox = MDIForm1.mnuPixelBox.Checked
Call ZoomWin(ZoomWinLeft, ZoomWinTop)

End Sub

Public Sub mnuCompensatingFill_Click()
MDIForm1.mnuCompensatingFill.Checked = True
MDIForm1.mnuStrictFill.Checked = False
FillMode = "compensating"
End Sub

Private Sub UnCheckModes()
MDIForm1.mnuGfxModeHires.Checked = False
MDIForm1.mnuGfxModeKoala.Checked = False
MDIForm1.mnuGfxModeDrazlace.Checked = False
MDIForm1.mnuGfxModeAfli.Checked = False
MDIForm1.mnuGfxModeIFli.Checked = False
MDIForm1.mnuGfxModeFli.Checked = False
MDIForm1.mnuGfxModeDrazlaceSpec.Checked = False
MDIForm1.mnuGfxModeUfli.Checked = False
Form1.Caption = "Zoom Window 1:" + Str(ZoomScale) + "  Mode: " + GfxMode
End Sub
Private Sub Copy2Convbuff()
For x = 0 To 319
For y = 0 To 199
ColMap(x, y) = Pixels(x, y)
Next y
Next x
End Sub
Public Sub mnuGfxModeHires_Click()

GfxMode = "hires"
ResoDiv = 1
FliMul = 8
ScrBanks = 0
BitBanks = 0
XFliLimit = 0
Call UnCheckModes
MDIForm1.mnuGfxModeHires.Checked = True
Call Copy2Convbuff
Call Hires_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub

Public Sub mnuGfxModeKoala_Click()

GfxMode = "koala"
ResoDiv = 2
FliMul = 8
ScrBanks = 0
BitBanks = 0
XFliLimit = 0
Call UnCheckModes
MDIForm1.mnuGfxModeKoala.Checked = True
Call Copy2Convbuff
Call Multicol_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub

Public Sub mnuGfxModeDrazlace_Click()

GfxMode = "drazlace"
ResoDiv = 1
FliMul = 8
ScrBanks = 0
BitBanks = 1
XFliLimit = 0
Call UnCheckModes
MDIForm1.mnuGfxModeDrazlace.Checked = True
Call Copy2Convbuff
Call Multicol_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub
Public Sub mnuGfxModeAfli_Click()

GfxMode = "afli"
ResoDiv = 1
FliMul = 1
ScrBanks = 0
BitBanks = 0
XFliLimit = 24
Call UnCheckModes
MDIForm1.mnuGfxModeAfli.Checked = True
Call Copy2Convbuff
Call Hires_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub
Public Sub mnuGfxModeIFli_Click()

GfxMode = "ifli"
ResoDiv = 1
FliMul = 1
ScrBanks = 1
BitBanks = 1
XFliLimit = 24
Call UnCheckModes
MDIForm1.mnuGfxModeIFli.Checked = True
Call Copy2Convbuff
Call Multicol_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub

Public Sub mnuGfxModeFli_Click()

GfxMode = "fli"
ResoDiv = 2
FliMul = 1
ScrBanks = 0
BitBanks = 0
XFliLimit = 24
Call UnCheckModes
MDIForm1.mnuGfxModeFli.Checked = True
Call Copy2Convbuff
Call Multicol_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub

Public Sub mnuGfxModeDrazlaceSpec_Click()

GfxMode = "drazlacespec"
ResoDiv = 1
FliMul = 8
ScrBanks = 1
BitBanks = 1
XFliLimit = 0
Call UnCheckModes
MDIForm1.mnuGfxModeDrazlaceSpec.Checked = True
Call Copy2Convbuff
Call Multicol_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub
Public Sub mnuGfxModeUfli_Click()

GfxMode = "ufli"
ResoDiv = 1
FliMul = 2
ScrBanks = 0
BitBanks = 0
XFliLimit = 24
Call UnCheckModes
MDIForm1.mnuGfxModeUfli.Checked = True
Call Copy2Convbuff
Call Ufli_Attrib
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
Call ResetUndo

End Sub

Public Sub mnuStrictFill_Click()
mnuStrictFill.Checked = True
mnuCompensatingFill.Checked = False
FillMode = "strict"
End Sub

Public Sub mnuCopy_Click()

Clipboard.Clear
Clipboard.SetData Form2.Picture2.Image
End Sub


Public Sub mnuConvertOptions_Click()

ConvertDialog.Show vbModal

End Sub

Public Sub mnuLoadPicture_Click()
Call Load_Click
End Sub

Public Sub mnuZoomScale_Click()
ZoomScaleDialog.Show vbModal
End Sub




Private Sub Timer1_Timer()

For z = 0 To 15

 If MDIForm1.ColorPicker1.Color = c64cols(z) Then LeftColN = z
 If MDIForm1.ColorPicker2.Color = c64cols(z) Then RightColN = z

Next z

 MDIForm1.StatusBar1.Panels(9).Picture = Picture3(LeftColN).Image
 MDIForm1.StatusBar1.Panels(10).Picture = Picture3(RightColN).Image

End Sub

Private Sub VScroll1_scroll()
DontUpdate = True
ZoomWinTop = VScroll1.Value: Call ZoomWin(ZoomWinLeft, VScroll1.Value)
ZoomPic.Refresh
DontUpdate = False
End Sub
Private Sub HScroll1_scroll()
DontUpdate = True
ZoomWinLeft = HScroll1.Value: Call ZoomWin(ZoomWinLeft, ZoomWinTop)
ZoomPic.Refresh
DontUpdate = False
End Sub
Private Sub VScroll1_Change()
If DontUpdate = False Then
ZoomWinTop = VScroll1.Value
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
ZoomPic.Refresh
End If
End Sub
Private Sub HScroll1_Change()
If DontUpdate = False Then
ZoomWinLeft = HScroll1.Value
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
ZoomPic.Refresh
End If
End Sub

Private Sub CharCols()

cx = Int(Ax / 8)
cy = Int(Ay / FliMul)
dy = Int(Ay / 8)

If ScrBanks = 1 Then Bank = Ax And 1 Else Bank = 0

If cx <> ChoX Or cy <> ChoY Or Bank <> OldBank Then
 
 MDIForm1.StatusBar1.Panels(1).Picture = Picture3(ScrRam(Bank, cx, cy, 0)).Image
 MDIForm1.StatusBar1.Panels(2).Picture = Picture3(ScrRam(Bank, cx, cy, 1)).Image
 MDIForm1.StatusBar1.Panels(3).Picture = Picture3(d800(cx, dy)).Image
 MDIForm1.StatusBar1.Panels(4).Picture = Picture3(BackgrIndex).Image
 
End If


'If BankNum = 1 Then Bank = Ax And 1 Else Bank = 0

'Text1.Text = Str(ScrRam(Bank, cx, cy, 2).Color) + " " + Str(ScrRam(Bank, cx, cy, 0).Color) + " " + Str(ScrRam(Bank, cx, cy, 1).Color) + ":" + Str(amount) + " " + Str(ScrRam(Bank, cx, cy, 0).Count) + " " + Str(ScrRam(Bank, cx, cy, 1).Count) + " X" + Str(Ax And 1)
'Text1.Text = Str(BITMAP(0, cx, Ay)) + " " + Str(BitCount01(BITMAP(0, cx, Ay)))

ChoX = cx
ChoY = cy
OldBank = Bank

End Sub

Private Sub Paint(drX As Integer, drY As Integer, ByVal drColor As Byte)

If PaintMode = "draw" Then
    Call Setpixel(drX, drY, drColor)
ElseIf PaintMode = "brush" Then
    Call PlotBrush(drX, drY)
End If

End Sub
Private Sub Setpixel(spX As Integer, spY As Integer, ByVal spColor As Byte)


'spX = Int(spX / ResoDiv) * ResoDiv

If spX >= 0 And spX <= 319 And spY >= 0 And spY <= 199 Then

If Pixels(spX, spY) <> spColor Then

If DitherMode = True Then
If pattern(spX / ResoDiv And 1, spY And 1) = 1 Then spColor = LeftColN
If pattern(spX / ResoDiv And 1, spY And 1) = 0 Then spColor = RightColN
End If


    
 
 ActCol = c64cols(spColor)
 If PixelFits(spX, spY, spColor) Then
   Pixels(spX, spY) = spColor
   If ResoDiv = 2 Then Pixels(spX + 1, spY) = spColor
   If FastSetPixel = False Then
    ZoomWinX = spX - ZoomWinLeft: ZoomWinY = spY - ZoomWinTop
    ZoomPic.Line (ZoomWinX, ZoomWinY)- _
    (ZoomWinX + ResoDiv, ZoomWinY + 1), ActCol, BF
   End If
 Else
 GoTo nofuck
  Max = 99999999
  cx = Int(spX / 8)
  cy = Int(spY / FliMul)
  dy = Int(spY / 8)
  If ScrBanks = 1 Then ScrBank = Ax And 1 Else ScrBank = 0
  c(0) = d800(cx, dy)
  c(1) = ScrRam(ScrBank, cx, cy, 0)
  c(2) = ScrRam(ScrBank, cx, cy, 1)
  c(3) = BackgrIndex
  For z = 0 To 3
  If cd(c(z), spColor) < Max Then Max = cd(c(z), spColor): Index = z
  Next z
   spColor = c(Index)
   Call PixelFits(spX, spY, spColor)
   Pixels(spX, spY) = spColor
   If ResoDiv = 2 Then Pixels(spX + 1, spY) = spColor
   If FastSetPixel = False Then
    ZoomWinX = spX - ZoomWinLeft: ZoomWinY = spY - ZoomWinTop
    ZoomPic.Line (ZoomWinX, ZoomWinY)- _
    (ZoomWinX + ResoDiv, ZoomWinY + 1), ActCol, BF
   End If
nofuck:
 End If
                    
 If FastSetPixel = False Then
 Call grid(ZoomWinLeft, ZoomWinTop)
 Call Form2.ReDrawDisp
 End If

End If

End If

End Sub
Private Function PixelFits(pfX As Integer, pfY As Integer, pfColor As Byte) As Boolean

Dim qbank As Byte

PixelFits = False
col2change = Pixels(pfX, pfY)

If GfxMode = "koala" Or _
GfxMode = "drazlace" Or _
GfxMode = "ifli" Or _
GfxMode = "fli" Or _
GfxMode = "drazlacespec" Then

If pfX >= XFliLimit Then 'And col2change <> pfColor Then (setpixel already checks')
 
 If ScrBanks = 1 Then ScrBank = pfX And 1 Else ScrBank = 0
 If BitBanks = 1 Then BitBank = pfX And 1 Else BitBank = 0
 
 'If GfxMode = "drazlace" Then Bank = 0 ' restrict colormem into 1 bank!!!
 
 cx = Int(pfX / 8)
 cy = Int(pfY / FliMul)
 dy = Int(pfY / 8)
 
 If ScrRam(ScrBank, cx, cy, 0) = pfColor Then
        Call PutBitmap(BitBank, pfX, pfY, 1): PixelFits = True
 ElseIf ScrRam(ScrBank, cx, cy, 1) = pfColor Then
        Call PutBitmap(BitBank, pfX, pfY, 2): PixelFits = True
 ElseIf d800(cx, dy) = pfColor Then
        Call PutBitmap(BitBank, pfX, pfY, 3): PixelFits = True
 ElseIf pfColor = background Then
        Call PutBitmap(BitBank, pfX, pfY, 0): PixelFits = True
 ElseIf 0 = 0 Then
 
 Bp00 = 0
 Bp01 = 0
 Bp10 = 0
 Bp11 = 0
 
 starty = Int(pfY / 8) * 8
 endy = (Int(pfY / 8) * 8) + 7

 For qbank = 0 To BitBanks
 For y = starty To endy
  Bp11 = Bp11 + BitCount11(BITMAP(qbank, cx, y))
  Bp00 = Bp00 + BitCount00(BITMAP(qbank, cx, y))
 Next y
 Next qbank
 
  
 starty = cy * FliMul
 endy = ((cy + 1) * FliMul) - 1
 
 If BitBanks = 1 And ScrBanks = 0 Then
 
  'ha 2 bitbank es 1 scrbank: mindket bitmapban nezzuk h befer-e
  
   For qbank = 0 To 1
    For y = starty To endy
     Bp01 = Bp01 + BitCount01(BITMAP(qbank, cx, y))
     Bp10 = Bp10 + BitCount10(BITMAP(qbank, cx, y))
    Next y
   Next qbank
   
 Else
  
  'ha 2 vagy 1 bitbank es 2 vagy 1 scrbank: rajzolando bitmapban nezzuk h befer-e
  
    For y = starty To endy
     Bp01 = Bp01 + BitCount01(BITMAP(BitBank, cx, y))
     Bp10 = Bp10 + BitCount10(BITMAP(BitBank, cx, y))
    Next y
 
 End If
 
 'Text1.Text = Str(Bp01) + " " + Str(Bp10) + " " + Str(Bp11)
 
 If Bp01 = 0 Or (Bp01 = 1 And col2change = ScrRam(ScrBank, cx, cy, 0)) Then
  Call PutBitmap(BitBank, pfX, pfY, 1)
  ScrRam(ScrBank, cx, cy, 0) = pfColor
  PixelFits = True
  'Text1.Text = "01:" + Str(Bp01)
 ElseIf Bp10 = 0 Or (Bp10 = 1 And col2change = ScrRam(ScrBank, cx, cy, 1)) Then
  Call PutBitmap(BitBank, pfX, pfY, 2)
  ScrRam(ScrBank, cx, cy, 1) = pfColor
  PixelFits = True
  'Text1.Text = "10:" + Str(Bp10)
 ElseIf Bp11 = 0 Or (Bp11 = 1 And col2change = d800(cx, dy)) Then
  Call PutBitmap(BitBank, pfX, pfY, 3)
  d800(cx, dy) = pfColor
  PixelFits = True
  'Text1.Text = "11:" + Str(Bp11) + " c2c:" + Str(col2change) _
  + " "
 End If

'Text1.Text = Str(Bp01) + " " + Str(Bp10) + " " + Str(Bp11)
'Text1.Text = Str(starty) + " " + Str(endy) + " " + Str(cy)

 End If
End If
End If

'******************* hires********************************

If GfxMode = "hires" Or _
GfxMode = "afli" Then

If pfX >= XFliLimit And col2change <> pfColor Then
 
 BitBank = 0
 ScrBank = 0
 cx = Int(pfX / 8)
 cy = Int(pfY / FliMul)
 dy = Int(pfY / 8)
 
 If ScrRam(ScrBank, cx, cy, 0) = pfColor Then
        Call PutBitmap(BitBank, pfX, pfY, 0): PixelFits = True
 ElseIf ScrRam(ScrBank, cx, cy, 1) = pfColor Then
        Call PutBitmap(BitBank, pfX, pfY, 1): PixelFits = True
 ElseIf 0 = 0 Then
 
 Bp1 = 0
 Bp0 = 0
  
 starty = cy * FliMul
 endy = ((cy + 1) * FliMul) - 1
 
 For y = starty To endy
  Bp1 = Bp1 + BitCount1(BITMAP(BitBank, cx, y))
  Bp0 = Bp0 + BitCount0(BITMAP(BitBank, cx, y))
 Next y
  
 'Text1.Text = Str(Bp0) + " x " + Str(Bp1) + " b:" + Str(BITMAP(Bank, cx, pfY))
 
 If Bp0 = 0 Or (Bp0 = 1 And col2change = ScrRam(ScrBank, cx, cy, 0)) Then
  Call PutBitmap(BitBank, pfX, pfY, 0)
  ScrRam(ScrBank, cx, cy, 0) = pfColor
  PixelFits = True
  'Text1.Text = "0:" + Str(Bp0)
 ElseIf Bp1 = 0 Or (Bp1 = 1 And col2change = ScrRam(ScrBank, cx, cy, 1)) Then
  Call PutBitmap(BitBank, pfX, pfY, 1)
  ScrRam(ScrBank, cx, cy, 1) = pfColor
  PixelFits = True
  'Text1.Text = "1:" + Str(Bp1)
 End If

'Text1.Text = Str(Bp01) + " " + Str(Bp10) + " " + Str(Bp11)
'Text1.Text = Str(starty) + " " + Str(endy) + " " + Str(cy)

 End If
End If
End If

End Function
Private Sub PutBitmap(pb, x, y, mask)

If GfxMode = "koala" Or _
   GfxMode = "drazlace" Or _
   GfxMode = "drazlacespec" Or _
   GfxMode = "fli" Or _
   GfxMode = "ifli" Then

    z = 3 - (Int(x / 2) And 3)
    BITMAP(pb, Int(x / 8), y) = BITMAP(pb, Int(x / 8), y) And Mask00(z)
    BITMAP(pb, Int(x / 8), y) = BITMAP(pb, Int(x / 8), y) Or Shift(mask, z)
    
End If

If GfxMode = "hires" Or _
   GfxMode = "afli" Then
   
    z = 7 - (Int(x / 1) And 7)
    BITMAP(pb, Int(x / 8), y) = BITMAP(pb, Int(x / 8), y) And Mask0(z)
    If mask = 1 Then
    BITMAP(pb, Int(x / 8), y) = BITMAP(pb, Int(x / 8), y) Or Shift1(z)
    End If
    
End If

'Text1.Text = Str(y)
End Sub
Private Sub InitAttribs()

For Sx = 0 To 39
 For Sy = 0 To 24
  Call SetColAttrib
 Next Sy
Next Sx

End Sub

Private Sub SetColAttrib()
  
If GfxMode = "koala" Or GfxMode = "drazlace" Or _
GfxMode = "ifli" Or GfxMode = "fli" Or _
GfxMode = "drazlacespec" Then
  
 For qx = Sx * 8 To ((Sx + 1) * 8) - 1 Step ResoDiv
  For qy = Sy * 8 To ((Sy + 1) * 8) - 1
     
   If ScrBanks = 1 Then ScrBank = qx And 1 Else ScrBank = 0
   If BitBanks = 1 Then BitBank = qx And 1 Else BitBank = 0
   
   If Pixels(qx, qy) = BackgrIndex Then
      Call PutBitmap(BitBank, qx, qy, 0)
   ElseIf Pixels(qx, qy) = ScrRam(ScrBank, Sx, Int(qy / FliMul), 0) Then
      Call PutBitmap(BitBank, qx, qy, 1)
   ElseIf Pixels(qx, qy) = ScrRam(ScrBank, Sx, Int(qy / FliMul), 1) Then
      Call PutBitmap(BitBank, qx, qy, 2)
   ElseIf Pixels(qx, qy) = d800(Sx, Sy) Then
      Call PutBitmap(BitBank, qx, qy, 3)
   End If
   
  Next qy
 Next qx


End If


If GfxMode = "hires" Or GfxMode = "afli" _
Or GfxMode = "ufli" Then
  
 For qx = Sx * 8 To ((Sx + 1) * 8) - 1 Step ResoDiv
  For qy = Sy * 8 To ((Sy + 1) * 8) - 1
     
   Bank = 0
   If Pixels(qx, qy) = ScrRam(Bank, Sx, Int(qy / FliMul), 0) Then
      Call PutBitmap(Bank, qx, qy, 0)
   ElseIf Pixels(qx, qy) = ScrRam(Bank, Sx, Int(qy / FliMul), 1) Then
      Call PutBitmap(Bank, qx, qy, 1)
   End If
   
  Next qy
 Next qx
  
  
End If

End Sub

Private Sub ChangeCol(ChangeTo)
'change2=replace color
Dim ChangeFrom As Byte
Dim FromIndex As Byte
Dim ToIndex As Byte

cx = Int(Ax / 8)
cy = Int(Ay / FliMul)
dy = Int(Ay / 8)

ChangeFrom = Pixels(Ax, Ay)


If (ChangeFrom <> ChangeTo And _
ChangeFrom <> BackgrIndex) Or _
GfxMode = "hires" Or GfxMode = "ufli" Or GfxMode = "afli" Then
 
 'MsgBox ("changecol'")
  'If BankNum = 1 Then Bank = Ax And 1 Else Bank = 0
  
  For Bank = 0 To ScrBanks
   If ScrRam(Bank, cx, cy, 0) = ChangeFrom Then
    ScrRam(Bank, cx, cy, 0) = ChangeTo
   ElseIf ScrRam(Bank, cx, cy, 1) = ChangeFrom Then
    ScrRam(Bank, cx, cy, 1) = ChangeTo
   ElseIf d800(cx, dy) = ChangeFrom Then
    d800(cx, dy) = ChangeTo
   End If
  Next Bank
  
  For x = cx * 8 To ((cx + 1) * 8) - 1 Step 1 'ResoDiv
   For y = cy * FliMul To ((cy + 1) * FliMul) - 1
     If ChangeFrom = Pixels(x, y) Then Pixels(x, y) = ChangeTo
   Next y
  Next x

Call Form2.ReDrawDisp
Call ZoomWin(ZoomWinLeft, ZoomWinTop)

End If



End Sub


Public Sub mnuPaste_Click()


If StretchPic = False Then

 Set Display.Picture = Clipboard.GetData()

Else
 
 If KeepAspect = False Then
  
  Set Buffer = Clipboard.GetData()
  Buffer.ScaleMode = vbPixels

  Display.PaintPicture Buffer.Image, 0, 0, _
  Display.Width, Display.Height, _
  0, 0, _
  Buffer.Width, Buffer.Height, vbSrcCopy
 
 Else

  Set Buffer = Clipboard.GetData()
  Buffer.ScaleMode = vbPixels
  
  If Buffer.Width / Buffer.Height < 1.6 Then
  
   'magasabb mint szelesebb
   
   z = 1 / (Buffer.Height / 200)
   x = (319 - (z * Buffer.Width)) / 2
   
   Display.PaintPicture Buffer.Image, _
   x, 0, _
   z * Buffer.Width, z * Buffer.Height, _
   0, 0, _
   Buffer.Width, Buffer.Height, vbSrcCopy
  
  Else

   z = 1 / (Buffer.Width / 320)
   y = (199 - (z * Buffer.Height)) / 2
   
   Display.PaintPicture Buffer.Image, _
   0, y, _
   z * Buffer.Width, z * Buffer.Height, _
   0, 0, _
   Buffer.Width, Buffer.Height, vbSrcCopy
   
  End If
 
 End If
 
End If

Call ConvertPic
Call ZoomWin(ZoomWinTop, ZoomWinLeft)
End Sub

Private Sub Load_Click()
Dim pattern As String

   Form2.Timer1.Enabled = False
   Call Form2.ReDrawDisp

CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' Set flags
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNPathMustExist
' Set filters
pattern = "Pictures (*.bmp;*.gif;*.jpg)"
pattern = pattern & "|*.bmp;*.gif;*.jpg"
CommonDialog1.Filter = pattern
' Specify default filter
CommonDialog1.FilterIndex = 1

CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
PicName = "..." + Right(Filename, 20)
Form2.Caption = Str(zoom) + "% " + Form1.PicName
On Error GoTo 0

 Display.BackColor = 0
 Display.Cls
 Display.Refresh
 Buffer.BackColor = 0
 Buffer.Cls
 Buffer.Refresh
 
 Display.Line (0, 0)-(320, 200), 0, BF
 Buffer.Line (0, 0)-(320, 200), 0, BF
 
If StretchPic = False Then

 Set Display = LoadPicture(Filename)

Else
 
 If KeepAspect = False Then
  
  Set Buffer = LoadPicture(Filename)
  Buffer.ScaleMode = vbPixels

  Display.PaintPicture Buffer.Image, 0, 0, _
  Display.Width, Display.Height, _
  0, 0, _
  Buffer.Width, Buffer.Height, vbSrcCopy
 
 Else

  Set Buffer = LoadPicture(Filename)
  Buffer.ScaleMode = vbPixels
  
  If Buffer.Width / Buffer.Height < 1.6 Then
  
   'magasabb mint szelesebb
   
   z = 1 / (Buffer.Height / 200)
   x = (319 - (z * Buffer.Width)) / 2
   
   Display.PaintPicture Buffer.Image, _
   x, 0, _
   z * Buffer.Width, z * Buffer.Height, _
   0, 0, _
   Buffer.Width, Buffer.Height, vbSrcCopy
  
  Else

   z = 1 / (Buffer.Width / 320)
   y = (199 - (z * Buffer.Height)) / 2
   
   Display.PaintPicture Buffer.Image, _
   0, y, _
   z * Buffer.Width, z * Buffer.Height, _
   0, 0, _
   Buffer.Width, Buffer.Height, vbSrcCopy
   
  End If
 
 End If
 
End If

MDIForm1.mnuInterlaceEmu.Checked = False
Form2.Refresh
Form1.Refresh

Call ConvertPic
Call ZoomWin(ZoomWinLeft, ZoomWinTop)

Call ResetUndo

ErrHandler:
  'User pressed the Cancel button
  Exit Sub
  
End Sub

Private Sub ConvertPic()

Call color_filter
Call Color_Restrict

End Sub

Private Sub Color_Restrict()

If GfxMode = "drazlace" Then Call Multicol_Attrib
If GfxMode = "koala" Then Call Multicol_Attrib
If GfxMode = "hires" Then Call Hires_Attrib
If GfxMode = "afli" Then Call Hires_Attrib
If GfxMode = "ifli" Then Call Multicol_Attrib
If GfxMode = "fli" Then Call Multicol_Attrib
If GfxMode = "drazlacespec" Then Call Multicol_Attrib
If GfxMode = "ufli" Then Call Ufli_Attrib

End Sub
Public Sub mnuSavePicture_Click()
 
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' Set flags
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
' Set filters
CommonDialog1.Filter = "Bmp Files (*.bmp)|*.bmp|"
' Specify default filter
CommonDialog1.FilterIndex = 1

 CommonDialog1.ShowSave
 Filename = CommonDialog1.Filename
 If Filename <> "" Then
 StretchBlt Display.hdc, 0, 0, 320, 200, PixelsDib.hdc, 0, 0, 320, 200, vbSrcCopy
 SavePicture Display.Image, Filename
 End If

ErrHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub Picture3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index <= 15 Then
  If Button = 2 Then
   RightColN = Index
   RightCol = c64cols(Index)
   'Picture3(21).BackColor = c64cols(Index)
  Else
   If Button = 1 Then
    LeftColN = Index
    LeftCol = c64cols(Index)
    'Picture3(19).BackColor = c64cols(Index)
   End If
  End If
End If

 MDIForm1.StatusBar1.Panels(9).Picture = Picture3(LeftColN).Image
 MDIForm1.StatusBar1.Panels(10).Picture = Picture3(RightColN).Image

End Sub


Private Sub ZoomPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
Timer1.Enabled = False 'turn off timer to avoid to have the color changed back by it

If PaintMode = "draw" Then

FastSetPixel = False
OldAx = Ax
OldAy = Ay

  PaintNow = False

   Index = Pixels(Ax, Ay)
   If Button = 1 Then
    ColNum = LeftColN
    If Shift = 0 Then
     Call Setpixel(Ax, Ay, LeftColN)
     PaintNow = True
    End If
    If Shift = 1 Then
     LeftCol = c64cols(Index): LeftColN = Index
     ColNum = LeftColN
     MDIForm1.StatusBar1.Panels(9).Picture = Picture3(LeftColN).Image
     MDIForm1.ColorPicker1.Color = c64cols(LeftColN)
    End If
    If Shift = 2 Then
     Call ChangeCol(LeftColN)
    End If
    If Shift = 4 Then
     DitherMode = True
     PaintNow = True
    End If
   End If
   
   If Button = 2 Then
   ColNum = RightColN
    If Shift = 0 Then
     Call Setpixel(Ax, Ay, RightColN)
     PaintNow = True
    End If
    If Shift = 1 Then
     RightCol = c64cols(Index): RightColN = Index
     ColNum = RightColN
     MDIForm1.StatusBar1.Panels(10).Picture = Picture3(RightColN).Image
     MDIForm1.ColorPicker2.Color = c64cols(RightColN)
    End If
    If Shift = 2 Then
     Call ChangeCol(RightColN)
    End If
    If Shift = 4 Then
     DitherMode = True
     PaintNow = True
    End If

   End If
   
   If Button = 4 Then
    PaintNow = True
    ColNum = BackgrIndex
    Call Setpixel(Ax, Ay, BackgrIndex)
   End If
   
   Call Form2.ReDrawDisp
   Call ZoomWin(ZoomWinLeft, ZoomWinTop)
   
ElseIf PaintMode = "fill" Then

   If Button = 1 Then FillCol = LeftColN
   If Button = 2 Then FillCol = RightColN
   Call StartFill
   
ElseIf PaintMode = "brush" Then

   PaintNow = True
   Call Paint(Ax, Ay, LeftColN)
   Call Form2.ReDrawDisp
   Call ZoomWin(ZoomWinLeft, ZoomWinTop)

ElseIf PaintMode = "debug" Then
   
   Call debugger(Ax, Ay)
   
End If

Timer1.Enabled = True
End Sub
Private Sub ZoomPic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PaintNow = False   ' Turn off painting.
   DitherMode = False
   Call WriteUndo
   'Call Form2.ShowCursor(Ax, Ay, OldAx, OldAy)
End Sub

Private Sub DrawLine(ByVal X1 As Integer, ByVal Y1 As Integer, _
ByVal X2 As Integer, ByVal Y2 As Integer, ByVal lColor As Byte)

Dim Xdist As Integer
Dim Ydist As Integer
Dim Xdir As Integer
Dim Ydir As Integer
Dim YY As Single
Dim XX As Single
Dim Xstep As Single
Dim Ystep As Single
Dim Xb As Integer
Dim Yb As Integer

Dim temp As Integer

FastSetPixel = True


Xdist = Abs(X1 - X2)
Ydist = Abs(Y1 - Y2)

If Xdist = 0 Then

  If Y1 > Y2 Then
   temp = Y1: Y1 = Y2: Y2 = temp
  End If
  
  Xb = X1
  Text1.Text = "vertical line!"
  For Yb = Y1 To Y2
    Call Paint(Xb, Yb, lColor)
  Next Yb

ElseIf Ydist = 0 Then

  If X1 > X2 Then
   temp = X1: X1 = X2: X2 = temp
  End If
  
  Yb = Y1
  For Xb = X1 To X2 Step ResoDiv
    Call Paint(Xb, Yb, lColor)
  Next Xb

ElseIf Xdist / ResoDiv > Ydist Then
 
 If X1 > X2 Then
  temp = X1: X1 = X2: X2 = temp
  temp = Y1: Y1 = Y2: Y2 = temp
 End If
 
 
 Ydir = Sgn(Y2 - Y1)
  
 Ystep = (Ydist / Xdist) * Ydir * ResoDiv
 YY = Y1
 For Xb = X1 To X2 Step ResoDiv
    Call Paint(Xb, Int(YY), lColor)
    YY = YY + Ystep
 Next Xb

Else
   
 If Y1 > Y2 Then
  temp = X1: X1 = X2: X2 = temp
  temp = Y1: Y1 = Y2: Y2 = temp
 End If
 
 Xdir = Sgn(X2 - X1)
 
 Xstep = (Xdist / Ydist) * Xdir
 XX = X1
 For Yb = Y1 To Y2 Step 1
    Call Paint(Int((XX / ResoDiv) + 0.5) * ResoDiv, Yb, lColor)
    XX = XX + Xstep
 Next Yb

End If

FastSetPixel = False

End Sub

Private Sub ZoomPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

ZoomPicAbsX = x: ZoomPicAbsY = y

If PaintMode = "brush" Then
 Call ZoomWin(ZoomWinLeft, ZoomWinTop)
 ZoomPic.ForeColor = RGB(255, 255, 255)
 ZoomPic.Circle (x, y), Form1.Bsize / 2
 'Call Form2.picture2_mousemove(Button, Shift, x, y)
End If


 
If PaintNow = True Then
 Call CalcZoomCords(x, y)
 If Sqr((Ax - OldAx) ^ 2 + (Ay - OldAy) ^ 2) < 2 Then
  Call Paint(Ax, Ay, ColNum)
 Else
  Call DrawLine(OldAx, OldAy, Ax, Ay, ColNum)
  Call ZoomWin(ZoomWinLeft, ZoomWinTop)
 End If
End If

Call CalcZoomCords(x, y)
Call Form2.ShowCursor(Ax, Ay)

'draw box around current pixel

If PixelBox = True Then
 Call CalcZoomCords(x, y)
 
 If PixelGrid = True Then
 
  If CharGrid = True Then
     temp = PixelGrid: PixelGrid = False
     Call grid(ZoomWinLeft, ZoomWinTop)
     PixelGrid = temp
  End If
 
  x = Ax - ZoomWinLeft: y = Ay - ZoomWinTop
  ZoomPic.Line (ox, oy)-(ox + ResoDiv, oy + 1), RGB(0, 0, 0), B
  ZoomPic.Line (x, y)-(x + ResoDiv, y + 1), RGB(255, 255, 255), B
  ZoomPic.Refresh
  ox = x: oy = y

 Else

  ActCol = c64cols(Pixels(ox + ZoomWinLeft, oy + ZoomWinTop))

  Call grid(ZoomWinLeft, ZoomWinTop)

  x = Ax - ZoomWinLeft: y = Ay - ZoomWinTop
  ZoomPic.Line (ox, oy)-(ox + ResoDiv - (1 / ZoomScale), oy + 1 - (1 / ZoomScale)), ActCol, B
  ZoomPic.Line (x, y)-(x + ResoDiv - (1 / ZoomScale), y + 1 - (1 / ZoomScale)), RGB(255, 255, 255), B

  ZoomPic.Refresh
  ox = x: oy = y
 
 End If
 
End If


cx = Int(Ax / 8)
cy = Int(Ay / FliMul)


Call CharCols

OldAx = Ax
OldAy = Ay
ZoomPic.Refresh

MDIForm1.StatusBar1.Panels(5).Text = "x:" + Str(Ax / ResoDiv)
MDIForm1.StatusBar1.Panels(6).Text = "y:" + Str(Ay)
MDIForm1.StatusBar1.Panels(7).Text = "c:" + Str(Int(Ax / 8))
MDIForm1.StatusBar1.Panels(8).Text = "r:" + Str(Int(Ay / 8))

End Sub

Private Sub CalcZoomCords(x, y)
 
 If x > ZoomWidth Then
 x = ZoomWidth
 ZoomWinLeft = ZoomWinLeft + ResoDiv: Call ZoomWin(ZoomWinLeft, ZoomWinTop)
 End If
 
 If y > ZoomHeight Then
 y = ZoomHeight
 ZoomWinTop = ZoomWinTop + 1: Call ZoomWin(ZoomWinLeft, ZoomWinTop)
 End If
 
 If x < 0 Then
 x = 0
 ZoomWinLeft = ZoomWinLeft - ResoDiv: Call ZoomWin(ZoomWinLeft, ZoomWinTop)
 End If

 If y < 0 Then
 y = 0
 ZoomWinTop = ZoomWinTop - 1: Call ZoomWin(ZoomWinLeft, ZoomWinTop)
 End If
 
 Ax = Int(ZoomWinLeft + Int(x / ResoDiv) * ResoDiv)
 Ay = Int(ZoomWinTop) + Int(y)
 
 If Ax > 319 Then Ax = 319 - ResoDiv
 If Ay > 199 Then Ay = 199
 If Ax < 0 Then Ax = 0
 If Ay < 0 Then Ay = 0
 
End Sub
Private Sub form_load()
    
'prevent close button to appear
DeleteMenu GetSystemMenu(Me.hwnd, False), SC_CLOSE, MF_BYCOMMAND


'setup toolbar

PaintMode = "draw"
Call MDIForm1.ResetButtonMenuCaption

'setup c64 color model supporting tables

 For x = 0 To 15
  IncColors(x) = 1
 Next x
 
For Sx = 0 To 1
 For x = 0 To 39
  For y = 0 To 199
   BITMAP(Sx, x, y) = 0
  Next y
 Next x
Next Sx

Bp1 = 1

For x = 0 To 7
 Shift1(x) = Bp1
 Mask0(x) = 255 - Bp1
 Bp1 = Bp1 * 2
Next x

Bp00 = 0
Bp01 = 1
Bp10 = 2
Bp11 = 3

For x = 0 To 3
 Shift(0, x) = Bp00
 Shift(1, x) = Bp01
 Shift(2, x) = Bp10
 Shift(3, x) = Bp11
 Mask00(x) = 255 - Bp11
 Bp00 = Bp00 * 4
 Bp01 = Bp01 * 4
 Bp10 = Bp10 * 4
 Bp11 = Bp11 * 4
Next x

For x = 0 To 255
 
 BitCount00(x) = 0
 BitCount01(x) = 0
 BitCount10(x) = 0
 BitCount11(x) = 0
 
 BitCount1(x) = 0
 BitCount0(x) = 0
 
 y = x
 For Sx = 0 To 3
  If (y And 3) = 0 Then BitCount00(x) = BitCount00(x) + 1
  If (y And 3) = 1 Then BitCount01(x) = BitCount01(x) + 1
  If (y And 3) = 2 Then BitCount10(x) = BitCount10(x) + 1
  If (y And 3) = 3 Then BitCount11(x) = BitCount11(x) + 1
  y = Int(y / 4)
 Next Sx
 
 y = x
 For Sx = 0 To 7
  If (y And 1) = 0 Then BitCount0(x) = BitCount0(x) + 1
  If (y And 1) = 1 Then BitCount1(x) = BitCount1(x) + 1
  y = Int(y / 2)
 Next Sx
 
 'MsgBox (Str(BitCount1(x)))
 
Next x

Dim cor(15) As Single
Dim cog(15) As Single
Dim cob(15) As Single
Dim s As Byte
   
   Call UnCheckModes
   MDIForm1.mnuGfxModeKoala.Checked = True
   GfxMode = "koala"
   ResoDiv = 2 '2
   FliMul = 8
   XFliLimit = 0
   ZoomHeight = 32
   ZoomWidth = 40
   ZoomScale = 8
   
   StretchPic = True
   KeepAspect = True
   
   Bank = 0
   Set PixelsDib = New cDIBSection256
   PixelsDib.Create 320, -200
   
   With tsain
       .cbElements = 1
       .cDims = 2
       .Bounds(0).lLbound = 0
       .Bounds(0).cElements = PixelsDib.Height
       .Bounds(1).lLbound = 0
       .Bounds(1).cElements = PixelsDib.BytesPerScanLine()
       .pvData = PixelsDib.DIBSectionBitsPtr
   End With
   
   CopyMemory ByVal VarPtrArray(Pixels), VarPtr(tsain), 4
            
   Form1.ScaleMode = vbPixels
   ZoomPic.ScaleMode = vbPixels
   ZoomPic.BackColor = RGB(0, 0, 0)
   Display.ScaleMode = vbPixels
   ZoomPic.Width = 640 'ScaleX(640, 3, 1)
   ZoomPic.Height = 256 'ScaleY(256, 3, 1)
   ZoomPic.Scale (0, 0)-(ZoomWidth, ZoomHeight)
   ZoomPic.AutoRedraw = True
   ZoomPic.AutoSize = True
   
   Display.BackColor = RGB(0, 0, 0)
   Display.BorderStyle = 0
   Display.Width = 320 'ScaleX(320, 3, 1)
   Display.Height = 200 'ScaleY(200, 3, 1)
   Display.Scale (0, 0)-(320, 200)
   Display.Move 0, 0
   Display.Visible = False
   Display.AutoRedraw = True
   
   Buffer.BackColor = RGB(0, 0, 0)
   Buffer.BorderStyle = 0
   Buffer.Width = 320 'ScaleX(320, 3, 1)
   Buffer.Height = 200 'ScaleY(200, 3, 1)
   Buffer.Scale (0, 0)-(320, 200)
   Buffer.Move 0, 0
   Buffer.Visible = False
   Buffer.AutoRedraw = True
      
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.fileexists(App.Path + "\default.txt") Then

    Open App.Path + "\default.txt" For Input As #1

    For y = 0 To 15
     Line Input #1, TextLine
     s = 1
     For x = 1 To 3
      U = s
      Do Until Mid(TextLine, U, 1) = "," Or U > Len(TextLine)
       U = U + 1
      Loop

      cor(x) = Mid(TextLine, s, U - s) '+"*"
      s = U + 1
     Next x
     c64cols(y) = RGB(Val(cor(1)), Val(cor(2)), Val(cor(3)))
    Next y
    
    Close #1
Else

c64cols(0) = RGB(0, 0, 0)
c64cols(1) = RGB(255, 255, 255)
c64cols(2) = RGB(104, 55, 43)
c64cols(3) = RGB(112, 164, 178)
c64cols(4) = RGB(111, 61, 134)
c64cols(5) = RGB(88, 141, 67)
c64cols(6) = RGB(53, 40, 121)
c64cols(7) = RGB(184, 199, 111)
c64cols(8) = RGB(111, 79, 37)
c64cols(9) = RGB(67, 57, 0)
c64cols(10) = RGB(154, 103, 89)
c64cols(11) = RGB(68, 68, 68)
c64cols(12) = RGB(108, 108, 108)
c64cols(13) = RGB(154, 210, 132)
c64cols(14) = RGB(108, 94, 181)
c64cols(15) = RGB(149, 149, 149)

End If

For x = 0 To 15
    Picture3(x).BackColor = c64cols(x)
    PixelsDib.Color(x) = c64cols(x)
Next x

'Picture3(16).BackColor = RGB(255, 0, 136)
PixelsDib.Color(16) = RGB(255, 0, 136)
  
'color picker setup

For z = 0 To 15
  MDIForm1.ColorPicker1.Items.Add c64cols(z), "color"
  MDIForm1.ColorPicker2.Items.Add c64cols(z), "color"
Next z

For z = 0 To 15
 cor(z) = c64cols(z) And 255
 cog(z) = Int(c64cols(z) / 256) And 255
 cob(z) = Int(c64cols(z) / 65536) And 255
Next z
    
For x = 0 To 15
     For y = 0 To 15
         rd = cor(x) - cor(y)
         gd = cog(x) - cog(y)
         bd = cob(x) - cob(y)
         cd(x, y) = Sqr(rd * rd + gd * gd + bd * bd)
     Next y
 Next x
   
   BackgrIndex = 0
   
   pattern(0, 0) = 0: pattern(0, 1) = 1
   pattern(1, 0) = 1: pattern(1, 1) = 0
   
   'code flags
   
   DitherFill = False
   MDIForm1.mnuDitherfill.Checked = False
   DitherMode = False
   FastSetPixel = False
   
   'Convert Options
   Reso_PixAvg = True
   mostfreqbackgr = True
   UsrBackgr = 0
   optchars = True
   FixWithBackgr = 1
   FillMode = "strict"
   MDIForm1.mnuCompensatingFill.Checked = False
   MDIForm1.mnuStrictFill.Checked = True
   PixelGrid = True
   CharGrid = True
   PixelBox = True
   
   MDIForm1.mnuPixelGrid.Checked = PixelGrid
   MDIForm1.mnuCharGrid.Checked = CharGrid
   MDIForm1.mnuPixelBox.Checked = PixelBox
   
   CopyMode = "reallynocopy"
   Force_Greys = True
   Tresh_Grey = 100
   ConvertDialog.Force_Greys.Value = 1
   Dither = False
   Uni_Dither = True
   Avg_Dither = False
   Tresh_Dith = 100
   Uni_Strength = 32
   
   'statusbar setup
    
   MDIForm1.StatusBar1.Height = 25 * 10
   For z = 2 To 10
   MDIForm1.StatusBar1.Panels.Add (z)
   Next z
   
   For z = 1 To 10
   MDIForm1.StatusBar1.Panels(z).Width = 47 * 15
   MDIForm1.StatusBar1.Panels(z).AutoSize = sbrNoAutoSize
   Next z
      
   MDIForm1.StatusBar1.Panels(1).Text = "ScrL"
   MDIForm1.StatusBar1.Panels(2).Text = "ScrH"
   MDIForm1.StatusBar1.Panels(3).Text = "D800"
   MDIForm1.StatusBar1.Panels(4).Text = "D021"
    
   MDIForm1.StatusBar1.Panels(5).Text = "x:"
   MDIForm1.StatusBar1.Panels(6).Text = "y:"
   MDIForm1.StatusBar1.Panels(7).Text = "c:"
   MDIForm1.StatusBar1.Panels(8).Text = "r:"
    
   MDIForm1.StatusBar1.Panels(9).Text = "rmb"
   MDIForm1.StatusBar1.Panels(10).Text = "lmb"
 
   MDIForm1.StatusBar1.Panels(9).Picture = Picture3(LeftColN).Image
   MDIForm1.StatusBar1.Panels(10).Picture = Picture3(RightColN).Image
   
   'x = (Frame1.Width - Picture3(0).Width * 2) / 3
   For z = 0 To 15 Step 1
    Picture3(z).Visible = False
    'Picture3(z).Left = Frame1.Left + x
   'Picture3(z + 1).Left = Frame1.Left + Picture3(z).Width + x + x
   Next
   Frame1.Visible = False
   
   Call InitAttribs
   Call ResetUndo: RedoCount = 0
   Call ConvertDialog.form_load
   Form2.Show
   Form2.zoom = 100
   Call Form2.ResizeForm
   Call Form2.resize
   Call Form2.ResizeForm
   Form1.Caption = "Zoom Window 1:" + Str(ZoomScale) + "  Mode: " + GfxMode
   Form1.ScaleMode = Form2.ScaleMode
   Form2.Top = 0
   Form2.Left = 0
   Form1.Top = 0
   Form1.Left = Form2.Width
   Form1.Width = MDIForm1.Width - 180 - Form2.Width
   Form1.Height = Form2.Height * 2.5 ' MDIForm1.Height - (MDIForm1.Toolbar1.Height)
   
   'Call DrawLine(100, 100, 115, 110, 1)
   'Call Form_Resize
   'Call Form2.Form_Resize
   
End Sub

Private Sub debugger(dbx, dby)

Dim temp As String
Dim db As Byte
Dim db2 As Byte
Dim tempy As Byte

cx = Int(dbx / 8)
cy = Int(dby / FliMul)
dy = Int(dby / 8)

Set fso = CreateObject("Scripting.FileSystemObject")

Open App.Path + "\debug.txt" For Output As #1

Print #1, "           x:" + Str(dbx And 7)
Print #1, "           y:" + Str(dby And 7)
Print #1, "      column:" + Str(cx)
Print #1, "         row:" + Str(dy)
Print #1, "    col left:" + Str(LeftCol)
Print #1, "    col rite:" + Str(RightCol)
Print #1, " "
Print #1, "        mode:" + GfxMode
Print #1, "bitmap banks:" + Str(BitBanks + 1)
Print #1, "screen banks:" + Str(ScrBanks + 1)
Print #1, "     ResoDiv:" + Str(ResoDiv)
Print #1, "      FliMul:" + Str(FliMul)
Print #1, "  "
Print #1, "Static Colors:"
Print #1, "d021(00):" + Str(BackgrIndex)
Print #1, "d800(11):" + Str(d800(cx, dy))
Print #1, " "
Print #1, " "

For BitBank = 0 To BitBanks

 Bp00 = 0
 Bp01 = 0
 Bp10 = 0
 Bp11 = 0
 
 For y = 0 To 7
  
  db = BITMAP(BitBank, cx, (dy * 8) + y)
  temp = Str(y) + ": " '+ Str(db) + " "
  Bp00 = Bp00 + BitCount00(db)
  Bp01 = Bp01 + BitCount01(db)
  Bp10 = Bp10 + BitCount10(db)
  Bp11 = Bp11 + BitCount11(db)
  
  For z = 3 To 0 Step -1
    db2 = (db And Shift(3, z)) / (4 ^ z)
    Select Case db2
    Case 0
     temp = temp + "00 "
    Case 1
     temp = temp + "01 "
    Case 2
     temp = temp + "10 "
    Case 3
     temp = temp + "11 "
    End Select
  Next z
  tempy = Int(dy + y / FliMul)
  If ScrBanks = 1 Then ScrBank = BitBank Else ScrBank = 0
  temp = temp + "|"
  Print #1, temp;
  Print #1, "  01:" + Str(ScrRam(ScrBank, cx, tempy, 0)); Tab;
  Print #1, "  10:" + Str(ScrRam(ScrBank, cx, tempy, 1)); Tab '; '" tempy: " + Str(tempy) + " dby: " + Str(dby) + " y: " + Str(y)
  'Print #1, temp; Tab
 Next y
 
Print #1, " "
Print #1, "count 00:" + Str(Bp00)
Print #1, "count 01:" + Str(Bp01)
Print #1, "count 10:" + Str(Bp10)
Print #1, "count 11:" + Str(Bp11)
Print #1, " "

Next BitBank

For y = 0 To 7
 temp = ""
 For x = 0 To 7
  temp = temp + Chr(Pixels(x + (cx * 8), y + (dy * 8)) + 32)
 Next x
 Print #1, temp
Next y
Close #1
    
MsgBox ("Debug information was written into the Application path")

End Sub

Public Sub Form_Resize()
   
 If Form1.Height > 2000 And Form1.Width > 2000 Then
   
   GetClientRect Form1.hwnd, Rectangle
   h = Rectangle.Bottom - HScroll1.Height
   w = Rectangle.Right
    
   If Int(h / ZoomScale) > 199 Then h = 199 * ZoomScale
   
   ZoomPic.Height = Int(h / ZoomScale) * ZoomScale
   ZoomHeight = Int(h / ZoomScale)
   
   ZoomPic.Top = 0
   ZoomPic.Left = 0
   
   HScroll1.Top = h
   HScroll1.Left = 0
   
   VScroll1.Top = 0
   VScroll1.Height = h
   
   VScroll1.Left = w - VScroll1.Width
   HScroll1.Width = VScroll1.Left
      
   zw = w - VScroll1.Width
   
   If Int(zw / ZoomScale) > 319 Then zw = 319 * ZoomScale
   
   ZoomPic.Width = Int(zw / ZoomScale) * ZoomScale
   ZoomWidth = Int(zw / ZoomScale)
   ZoomPic.Scale (0, 0)-(ZoomWidth, ZoomHeight)
   
   HScroll1.Max = 320 - ZoomWidth
   VScroll1.Max = 200 - ZoomHeight
   Call ZoomWin(ZoomWinLeft, ZoomWinTop)
   ZoomPic.Refresh
   
 End If
End Sub
Private Sub form_keydown(KeyCode As Integer, Shift As Integer)

If KeyCode = "38" Then
ZoomWinTop = ZoomWinTop - 8:
Call CoordLimit(ZoomWinLeft, ZoomWinTop)
VScroll1.Value = ZoomWinTop:
Call CalcZoomCords(ZoomPicAbsX, ZoomPicAbsY)
Call CharCols
End If

If KeyCode = "40" Then
ZoomWinTop = ZoomWinTop + 8:
Call CoordLimit(ZoomWinLeft, ZoomWinTop)
VScroll1.Value = ZoomWinTop:
Call CalcZoomCords(ZoomPicAbsX, ZoomPicAbsY)
Call CharCols
End If

If KeyCode = "37" Then
ZoomWinLeft = ZoomWinLeft - 8:
Call CoordLimit(ZoomWinLeft, ZoomWinTop)
HScroll1.Value = ZoomWinLeft:
Call CalcZoomCords(ZoomPicAbsX, ZoomPicAbsY)
Call CharCols
End If

If KeyCode = "39" Then
ZoomWinLeft = ZoomWinLeft + 8:
Call CoordLimit(ZoomWinLeft, ZoomWinTop)
HScroll1.Value = ZoomWinLeft:
Call CalcZoomCords(ZoomPicAbsX, ZoomPicAbsY)
Call CharCols
End If

If KeyCode = "8" Then ' backspace: clear char with actcol
   
 ColNum = LeftColN
 cx = Int(Ax / 8): cy = Int(Ay / 8)
 
 'find which color was cursor over of
  If ScrRam(Bank, cx, cy, 0) = ChangeFrom Then
  ScrRam(Bank, cx, cy, 0) = LeftColN
  ElseIf ScrRam(Bank, cx, cy, 1) = ChangeFrom Then
  ScrRam(Bank, cx, cy, 1) = LeftColN
  ElseIf d800(cx, cy) = ChangeFrom Then
  d800(cx, cy) = LeftColN
  End If
  
  For x = cx * 8 To ((cx + 1) * 8) - 1 Step 1
   For y = cy * 8 To ((cy + 1) * 8) - 1
     Pixels(x, y) = 1 'LeftColN
   Next y
  Next x

  Call InitAttribs
  
  Call ZoomWin(ZoomWinLeft, ZoomWinTop)
  
End If



End Sub
Private Sub Form_KeyPress(keyascii As Integer)
PressedKey = Chr(keyascii)

If PressedKey = "1" Then LeftColN = 0
If PressedKey = "2" Then LeftColN = 1
If PressedKey = "3" Then LeftColN = 2
If PressedKey = "4" Then LeftColN = 3
If PressedKey = "5" Then LeftColN = 4
If PressedKey = "6" Then LeftColN = 5
If PressedKey = "7" Then LeftColN = 6
If PressedKey = "8" Then LeftColN = 7
If PressedKey = "9" Then LeftColN = 8
If PressedKey = "q" Then LeftColN = 9
If PressedKey = "w" Then LeftColN = 10
If PressedKey = "e" Then LeftColN = 11
If PressedKey = "r" Then LeftColN = 12
If PressedKey = "t" Then LeftColN = 13
If PressedKey = "z" Then LeftColN = 14
If PressedKey = "u" Then LeftColN = 15

LeftCol = c64cols(LeftColN)
MDIForm1.ColorPicker1.Color = c64cols(LeftColN)



If PressedKey = "b" Then
 Call backgrchange
 BackgrIndex = LeftColN
End If

If PressedKey = " " Then
 ColNum = BackgrIndex
 Call Setpixel(Ax, Ay, BackgrIndex)
End If

End Sub

Private Sub StartFill()


 Area = Pixels(Ax, Ay)
 If Area <> FillCol Then
  FastSetPixel = True
  FillChangedPic = False
  Call Fill(Ax, Ay)
  Call ZoomWin(ZoomWinLeft, ZoomWinTop)
  If FillMode = "compensating" Then
   For x = 0 To 319
   For y = 0 To 199
   ColMap(x, y) = Pixels(x, y)
   Next y
   Next x
   Call Color_Restrict
   Call InitAttribs
  End If
  FastSetPixel = False
  Call ZoomWin(ZoomWinLeft, ZoomWinTop)
  ZoomPic.Refresh
  If FillChangedPic = True Then
   Call WriteUndo
   'MsgBox ("fillwritesundo")
  End If
  
 End If


End Sub
Private Sub backgrchange()
Dim buzi As Byte

For BitBank = 0 To BitBanks
For x = 0 To 39
For y = 0 To 199

 buzi = BITMAP(BitBank, x, y)
 For q = 3 To 0 Step -1
   If (buzi And 3) = 0 Then
    Pixels((x * 8) + (q * 2) + BitBank, y) = LeftColN
    If ResoDiv = 2 Then Pixels((x * 8) + (q * 2) + 1, y) = LeftColN
   End If
  buzi = Int(buzi / 4)
 Next q
 
Next y
Call Form2.ReDrawDisp
Next x
Next BitBank

Call ZoomWin(ZoomWinLeft, ZoomWinTop)

End Sub

Public Sub ReceiveStuff()
Dim z As Byte

For z = 0 To 15
ConvertDialog.Check1(z).Value = IncColors(z)
Next z
 
End Sub

Public Sub SubmitStuff()
Dim z As Byte

For z = 0 To 15
IncColors(z) = ConvertDialog.Check1(z).Value
Next z
 
End Sub
Private Sub color_filter()

Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim avg As Double
Dim dev As Single
Dim catch As Long
Dim dist As Long
Dim dist2 As Long
Dim Index As Integer
Dim bayer(3, 3) As Byte
Dim sbayer(2, 2) As Byte
Dim cor(15) As Single
Dim cog(15) As Single
Dim cob(15) As Single


For z = 0 To 15
 cor(z) = c64cols(z) And 255
 cog(z) = Int(c64cols(z) / 256) And 255
 cob(z) = Int(c64cols(z) / 65536) And 255
Next z

bayer(0, 0) = 0: bayer(0, 1) = 8: bayer(0, 2) = 2: bayer(0, 3) = 10
bayer(1, 0) = 12: bayer(1, 1) = 4: bayer(1, 2) = 14: bayer(1, 3) = 6
bayer(2, 0) = 3: bayer(2, 1) = 11: bayer(2, 2) = 1: bayer(2, 3) = 9
bayer(3, 0) = 15: bayer(3, 1) = 7: bayer(3, 2) = 13: bayer(3, 3) = 5

sbayer(0, 0) = 0: sbayer(0, 1) = 3:
sbayer(1, 0) = 3: sbayer(1, 1) = 0:



For y = 0 To 199
For x = 0 To 319 Step ResoDiv

 If Reso_PixRight = True And ResoDiv = 2 Then
  Color = Display.Point(x + 1, y)
  r = Color And 255
  g = Int(Color / 256) And 255
  b = Int(Color / 65536) And 255
 End If

 If Reso_PixAvg = True And ResoDiv = 2 Then
  Color = Display.Point(x, y)
  r = Color And 255
  g = Int(Color / 256) And 255
  b = Int(Color / 65536) And 255
 
  Color = Display.Point(x + 1, y)
  cr = Color And 255
  cg = Int(Color / 256) And 255
  cb = Int(Color / 65536) And 255
 
  r = (cr + r) / 2
  g = (cg + g) / 2
  b = (cb + b) / 2
 End If
  
 If Reso_PixLeft = True Or ResoDiv = 1 Then
  Color = Display.Point(x, y)
  r = Color And 255
  g = Int(Color / 256) And 255
  b = Int(Color / 65536) And 255
 End If

If Uni_Dither = True And Dither = True Then
 
 unic = Uni_Strength
 uand = &HFF - unic + 1
 If bayer((x / ResoDiv) And 3, y And 3) * unic < r Then _
 r = ((r And uand) + unic - 1) Else r = r And uand
 If bayer((x / ResoDiv) And 3, y And 3) * unic < g Then _
 g = ((g And uand) + unic - 1) Else g = g And uand
 If bayer((x / ResoDiv) And 3, y And 3) * unic < b Then _
 b = ((b And uand) + unic - 1) Else b = b And uand

End If

 '0 16 32 48 64 80 96 112 128 144 160 176 192 208 224 240 256
 
 If Force_Greys = False Then tresh = 0 Else tresh = Tresh_Grey
 avg = (r + g + b) / 3
 dev = (((r - avg) * (r - avg) + (g - avg) * (g - avg) + (b - avg) * (b - avg)) / 3)
 catch = 2147483647
 For z = 0 To 15
  If dev < tresh Then
    If z = 0 Or z = 11 Or z = 12 Or z = 15 Or z = 1 Then
     dist = ((cor(z) - r) * (cor(z) - r) + (cog(z) - g) * (cog(z) - g) + (cob(z) - b) * (cob(z) - b))
     If dist < catch And IncColors(z) = 1 Then catch = dist: Index = z
    End If
  Else
    dist = ((cor(z) - r) * (cor(z) - r) + (cog(z) - g) * (cog(z) - g) + (cob(z) - b) * (cob(z) - b))
    If (dist < catch) And IncColors(z) = 1 Then catch = dist: Index = z
  End If
 Next z
 final = Index

If Avg_Dither = True And Dither = True Then
   
  'find closest average
  catch2 = 2147483647
  For z = 0 To 15
   If z <> Index Then
    If dev < tresh Then
     If z = 0 Or z = 11 Or z = 12 Or z = 15 Or z = 1 Then
      dr = (cor(Index) + cor(z)) / 2
      dg = (cog(Index) + cog(z)) / 2
      db = (cob(Index) + cob(z)) / 2
      dist = ((dr - r) * (dr - r) + (dg - g) * (dg - g) + (db - b) * (db - b))
      If dist < catch2 And IncColors(z) = 1 Then catch2 = dist: index2 = z
     End If
    Else
     dr = (cor(Index) + cor(z)) / 2
     dg = (cog(Index) + cog(z)) / 2
     db = (cob(Index) + cob(z)) / 2
     dist = ((dr - r) * (dr - r) + (dg - g) * (dg - g) + (db - b) * (db - b))
     If dist < catch2 And IncColors(z) = 1 Then catch2 = dist: index2 = z
    End If
   End If
  Next z
  
 z = index2
 catch2 = ((cor(z) - r) * (cor(z) - r) + (cog(z) - g) * (cog(z) - g) + (cob(z) - b) * (cob(z) - b))
 limit = catch / (catch + catch2)
 
 
If limit > 0.5 - (Tresh_Dith / 200) And limit < 0.5 + (Tresh_Dith / 200) Then
' If catch < dist Then
   If Index < index2 Then Add = 0 Else Add = 1
   If bayer(((x / ResoDiv) + Add) And 3, y And 3) < (limit * 16) Then final = index2 Else final = Index
   'If sbayer(((x / ResoDiv) + Add) And 1, y And 1) < (limit * 4) * (Tresh_Dith / 100) Then final = index2 Else final = Index
   'If catch < dist Then final = Index
   ' Else
  'final = Index
  'MsgBox ("mu")
' End If
End If
 
End If ' if dither= true
 
 'Display.ForeColor = RGB(R, G, B)
 'Display.PSet (x, y)

 cols(final) = cols(final) + 1
 
 ColMap(x, y) = final
 Pixels(x, y) = final
 
 If ResoDiv = 2 Then
  ColMap(x + 1, y) = final
  Pixels(x + 1, y) = final
  'Display.PSet (x + 1, y)
 End If

Next x

 If y Mod 8 = 0 Then
  Call Form2.ReDrawDisp3(y)
  'Display.Refresh
 End If
 
Next y

Max = -1: MostFreqCol = -1
For z = 0 To 15
 If cols(z) > Max Then
  Max = cols(z)
  MostFreqCol = z
 End If
Next z


End Sub

Private Sub Multicol_Attrib()

If mostfreqbackgr = True Then BackgrIndex = MostFreqCol
If usrdefbackgr = True Then BackgrIndex = UsrBackgr

Draw = False

If optbackgr = True Then
 
 For tr = 0 To 15
  err = 0
  Call Mc_Attrib(tr)
  try(tr) = err
 Next tr

 Max = 2147483647
 For z = 0 To 15
  If try(z) < Max Then
   BackgrIndex = z
   Max = try(z)
  End If
 Next z
 
End If


Draw = True
Call Mc_Attrib(BackgrIndex)
Call Form2.ReDrawDisp
Call InitAttribs

End Sub
Private Sub Mc_Attrib(background)
Dim inchar(15) As Byte
Dim gotcha As Boolean
Dim charerr As Double
Dim errcount(7) As Long

'put used colors/char into usedcols(charx,chary,index) index 16 acts as colorcount
For x = 0 To 319 Step 8
For y = 0 To 199 Step 8
 
 cx = Int(x / 8)
 cy = Int(y / 8)
 
 For z = 0 To 15: cols(z) = -1: Next z
 
 z = 0
 
 For qx = x To x + 7 Step ResoDiv
 For qy = y To y + 7
 cols(ColMap(qx, qy)) = cols(ColMap(qx, qy)) + 1
 Next qy
 Next qx
 cols(BackgrIndex) = -1
 
 Index = 0
 For z = 0 To 15
  If cols(z) <> -1 Then UsedCols(cx, cy, Index) = z: Index = Index + 1
 Next z

 UsedCols(cx, cy, 16) = Index
 
Next y
Next x



BackgrIndex = background

For x = 0 To 319 Step 8
For y = 0 To 199 Step 8

cx = Int(x / 8)
cy = Int(y / 8)


If mostfrq3 = True Then
 
 For z = 0 To 15: cols(z) = -1: inchar(z) = BackgrIndex: Next z
 For qx = x To x + 7 Step ResoDiv
  For qy = y To y + FliMul - 1
   cols(ColMap(qx, qy)) = cols(ColMap(qx, qy)) + 1
  Next qy
 Next qx
 
 cols(BackgrIndex) = -1

 For U = 0 To 2
  Max = -1
  For z = 0 To 15
   If cols(z) > Max Then Max = cols(z): Index = z
  Next z
  c(U) = Index: cols(Index) = -1
 Next U
 c(3) = BackgrIndex

End If




If optchars = True Then
 
 maxcharerr = 2147483647
 
 
 For c0 = 0 To UsedCols(cx, cy, 16) 'usedcols
  
 charerr = 0
 
 For Bank = 0 To BitBanks
 
 For chary = y To y + 7 Step FliMul
  
 'used colors in fli box
 
 For z = 0 To 15: cols(z) = -1: inchar(z) = BackgrIndex: Next z
 
 If BitBanks = 1 Then
 Add = Bank
 Xstep = 2
 Else
 Add = 0
 Xstep = ResoDiv
 End If
 
 If (ScrBanks = 0 And BitBanks = 1) Then
 Add = 0
 Xstep = ResoDiv
 End If
 
 'used colors in char
 For qx = x + Add To x + 7 Step Xstep
  For qy = chary To chary + FliMul - 1
   cols(ColMap(qx, qy)) = cols(ColMap(qx, qy)) + 1
  Next qy
 Next qx
 
 cols(BackgrIndex) = -1 'set backgr as unused
 cols(UsedCols(cx, cy, c0)) = -1 'set d800 as unused
 
 'compact used colors into inchar array
 Max = 0
 For z = 0 To 15
  If cols(z) <> -1 Then inchar(Max) = z: Max = Max + 1
 Next z
 
 'number of colors in char, doesnt counts backgr and d800
 maxi = Max
 
 'set colors for fli box
 
 c(0) = UsedCols(cx, cy, c0)
 c(3) = BackgrIndex
  
 flilinerr = 2147483647
 
 'go through colors for scrrams
 For c1 = 0 To maxi - 1         '0400 #1
 For c2 = c1 + 1 To maxi        '0400 #2
 
  
 c(1) = inchar(c1)
 c(2) = inchar(c2)
 
  
 Merr = 0
 For qx = x + Add To x + 7 Step Xstep
 For qy = chary To chary + FliMul - 1
   
  pcol = ColMap(qx, qy)
  
   If (pcol <> c(0) And pcol <> c(1) And pcol <> c(2) And pcol <> c(3)) Then
    Max = 2147483647
     For z = 0 To 2 + FixWithBackgr
      If cd(pcol, c(z)) < Max Then Max = cd(pcol, c(z))
     Next z
     Merr = Merr + Max
   End If
 
 Next qy
 Next qx
 
 If Merr < flilinerr Then flilinerr = Merr
 
 
 Next c2
 Next c1
 
 charerr = charerr + flilinerr
 
 Next chary
 
 Next Bank
 
 If charerr < maxcharerr Then
 maxcharerr = charerr
 bestd800 = UsedCols(cx, cy, c0)
 End If

 Next c0


End If






  
 charerr = 0
 
 For Bank = 0 To BitBanks
 
 For chary = y To y + 7 Step FliMul
  
 'used colors in fli box
 
 For z = 0 To 15: cols(z) = -1: inchar(z) = BackgrIndex: Next z

 If BitBanks = 1 Then
 Add = Bank
 Xstep = 2
 Else
 Add = 0
 Xstep = ResoDiv
 End If
 
 If (ScrBanks = 0 And BitBanks = 1) Then
 Add = 0
 Xstep = ResoDiv
 End If

 For qx = x + Add To x + 7 Step Xstep
  For qy = chary To chary + FliMul - 1
   cols(ColMap(qx, qy)) = cols(ColMap(qx, qy)) + 1
  Next qy
 Next qx
 cols(BackgrIndex) = -1
 cols(bestd800) = -1
 
 Max = 0
 For z = 0 To 15
  If cols(z) <> -1 Then inchar(Max) = z: Max = Max + 1
 Next z
 
 maxi = Max
 
 'color variations / fli box
 
 c(0) = bestd800
 c(3) = BackgrIndex
 bestc1 = BackgrIndex
 bestc2 = BackgrIndex
 
 flilinerr = 2147483647
 
 For c1 = 0 To maxi - 1         '0400 #1
 For c2 = c1 + 1 To maxi        '0400 #2
 
 c(1) = inchar(c1)
 c(2) = inchar(c2)
 
 Merr = 0
 For qx = x + Add To x + 7 Step Xstep
 For qy = chary To chary + FliMul - 1
   
  pcol = ColMap(qx, qy)
  
   If (pcol <> c(0) And pcol <> c(1) And pcol <> c(2) And pcol <> c(3)) Then
    Max = 2147483647
     For z = 0 To 2 + FixWithBackgr
      If cd(pcol, c(z)) < Max Then Max = cd(pcol, c(z))
     Next z
     Merr = Merr + Max
   End If
 
 Next qy
 Next qx
 
 If Merr < flilinerr Then
 flilinerr = Merr
 bestc1 = c(1)
 bestc2 = c(2)
 End If
 

 Next c2
 Next c1
 
 c(0) = bestd800
 c(1) = bestc1
 c(2) = bestc2
 c(3) = BackgrIndex
 'If maxi = 0 Then c(1) = inchar(1)
 
 Merr = 0
 For qx = x + Add To x + 7 Step Xstep
 For qy = chary To chary + FliMul - 1
   
  pcol = ColMap(qx, qy)
  
   If (pcol <> c(0) And pcol <> c(1) And pcol <> c(2) And pcol <> c(3)) Then
    Max = 2147483647
     For z = 0 To 2 + FixWithBackgr
      If cd(pcol, c(z)) < Max Then Max = cd(pcol, c(z)): Index = z
     Next z
   final = c(Index)
   Merr = Merr + 1
   Else
   final = pcol
   End If
   If x < XFliLimit Then final = 0
   Pixels(qx, qy) = final
   If ResoDiv = 2 Then Pixels(qx + 1, qy) = final
 
 Next qy
 Next qx

 err = err + Merr
 If ScrBanks = 1 Then ScrBank = Bank Else ScrBank = 0
 ScrRam(ScrBank, cx, chary / FliMul, 0) = c(1)
 ScrRam(ScrBank, cx, chary / FliMul, 1) = c(2)
 d800(cx, Int(chary / 8)) = bestd800
 
 Next chary
 
 
 
 Next Bank
 
 

Next y

'Call Form2.ReDrawDisp
Call Form2.ReDrawDisp2(x)

Next x
End Sub
Public Sub SaveKoala()
Dim Filename As String
Dim ActAttrib(2) As Byte
Dim used(2) As Boolean
Dim found(2) As Integer
Dim Bmp As Byte
Dim Ff As Long

Ff = FreeFile

'get save filename
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
CommonDialog1.Filter = "(*.kla)|*.kla|(*.koa)|*.koa|(*.prg)|*.prg"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowSave
Filename = CommonDialog1.Filename
On Error GoTo 0


Set fso = CreateObject("Scripting.FileSystemObject")

If fso.fileexists(Filename) Then
 fso.deletefile (Filename)
 Open Filename For Binary As #Ff
Else
 Open Filename For Binary As #Ff
End If

Put #Ff, , &H6000

ActAttrib(0) = BackgrIndex
ActAttrib(1) = BackgrIndex
ActAttrib(2) = BackgrIndex

Bank = 0

For cy = 0 To 24
 For cx = 0 To 39
  
   ActAttrib(2) = d800(cx, cy)
 
 qx = cx * 8: qy = cy * 8
 For y = 0 To 7
   
   Bmp = BITMAP(0, cx, qy + y)
   Put #Ff, , Bmp
  Next y
  
Next cx
Next cy

For y = 0 To 24
 For x = 0 To 39
 Bmp = (ScrRam(Bank, x, y, 1) And 15) + (ScrRam(Bank, x, y, 0) And 15) * 16
 'bmp = 1 + 2 * 16
 Put #Ff, , Bmp
 Next x
Next y

For y = 0 To 24
 For x = 0 To 39
 Bmp = d800(x, y) And 15
 'bmp = 3
 Put #Ff, , Bmp
 Next x
Next y


Bmp = BackgrIndex
Put #Ff, , Bmp


Close #Ff

ErrHandler:
'User pressed the Cancel button
Exit Sub

End Sub

Public Sub mnuLoadKoala_click()
Dim Read As Byte
Dim filebuffer(10000) As Byte
Dim andtab(3), divtab(3)
Dim Bmp As Byte
Dim Ff As Long

Ff = FreeFile

andtab(3) = 3 * 1
andtab(2) = 3 * 4
andtab(1) = 3 * 16
andtab(0) = 3 * 64

divtab(3) = 1
divtab(2) = 4
divtab(1) = 16
divtab(0) = 64

CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNPathMustExist
'CommonDialog1.Filter = " |*.kla| |*.koa|"
CommonDialog1.Filter = "(*.kla)|*.kla|(*.koa)|*.koa|(*.prg)|*.prg"
CommonDialog1.FilterIndex = 1

'CommonDialog1.Flags = cdlCCFullOpen 'Or cdlCCPreventFullOpen

CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
On Error GoTo 0
'If Len(filename) = 0 Then GoTo ErrHandler

Open Filename For Binary As #Ff
'Open "c:\test.kla" For Binary As ff

Get #Ff, , Read
Get #Ff, , Read

Bank = 0
For x = 0 To 10000: Get #Ff, , filebuffer(x): Next

BackgrIndex = filebuffer(10000)

q = 0
For cy = 0 To 24
 For cx = 0 To 39
  'g = Int(actcolor / 256) And 255
  ScrRam(Bank, cx, cy, 0) = Int(filebuffer(8000 + q) / 16) And 15
  ScrRam(Bank, cx, cy, 1) = filebuffer(8000 + q) And 15
  q = q + 1
 Next cx
Next cy

q = 0
For cy = 0 To 24
 For cx = 0 To 39
  d800(cx, cy) = filebuffer(9000 + q) And 15
  q = q + 1
 Next cx
Next cy

q = 0
For cy = 0 To 24
For cx = 0 To 39
 qx = cx * 8: qy = cy * 8
 For y = 0 To 7
  For x = 0 To 3
   Bmp = filebuffer(q)
   temp = (Bmp And andtab(x)) / divtab(x)
   Select Case temp
    Case 0
     ColNum = BackgrIndex
    Case 1
     ColNum = ScrRam(Bank, cx, cy, 0)
    Case 2
     ColNum = ScrRam(Bank, cx, cy, 1)
    Case 3
     ColNum = d800(cx, cy)
   End Select
   
   Pixels(qx + x * 2, qy + y) = ColNum
   Pixels(qx + x * 2 + 1, qy + y) = ColNum
  Next x
  q = q + 1
 Next y
Next cx
'Display.Refresh
Call Form2.ReDrawDisp

Next cy

Close #Ff

'Text2.Text = "koala load end!"
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
'Picture3(20).BackColor = c64cols(BackgrIndex)
Call InitAttribs

GfxMode = "koala"
ResoDiv = 2
FliMul = 8
BitBanks = 0
ScrBanks = 0
XFliLimit = 0
Call ZoomWin(ZoomWinLeft, ZoomWinTop)

ErrHandler:
Close #Ff
  'User pressed the Cancel button
  Exit Sub

Close #Ff
'Display.Refresh
End Sub

Private Sub Fill(ByVal FillX As Integer, ByVal FillY As Integer)
Dim Tx As Integer

If FillMode = "strict" Then

Tx = FillX
Do While Tx <= 320 - ResoDiv

 If DitherFill = True Then
  If pattern(Tx / ResoDiv And 1, FillY And 1) = 1 Then FillCol = LeftColN
  If pattern(Tx / ResoDiv And 1, FillY And 1) = 0 Then FillCol = RightColN
 End If

 If (Pixels(Tx, FillY) = Area) Then 'And PixelFits(tx, FillY, FillCol) Then
  If PixelFits(Tx, FillY, FillCol) Then
   FillChangedPic = True
   Pixels(Tx, FillY) = FillCol
   If ResoDiv = 2 Then Pixels(Tx + 1, FillY) = FillCol
   Tx = Tx + ResoDiv
  Else
   Exit Do
  End If
 Else
  Exit Do
 End If
Loop

xMax = Tx - ResoDiv

Tx = FillX - ResoDiv 'tx > 0 And
Do While Tx >= 0

  If DitherFill = True Then
   If pattern(Tx / ResoDiv And 1, FillY And 1) = 1 Then FillCol = LeftColN
   If pattern(Tx / ResoDiv And 1, FillY And 1) = 0 Then FillCol = RightColN
  End If

If (Pixels(Tx, FillY) = Area) Then 'And PixelFits(tx, FillY, FillCol) Then
 If PixelFits(Tx, FillY, FillCol) Then
  FillChangedPic = True
  Pixels(Tx, FillY) = FillCol
  If ResoDiv = 2 Then Pixels(Tx + 1, FillY) = FillCol
  Tx = Tx - ResoDiv
 Else
  Exit Do
 End If
Else
 Exit Do
End If
Loop

minx = Tx + ResoDiv

End If

If FillMode = "compensating" Then

Tx = FillX
Do While Tx <= 320 - ResoDiv

If DitherFill = True Then
If pattern(Tx / ResoDiv And 1, FillY And 1) = 1 Then FillCol = LeftColN
If pattern(Tx / ResoDiv And 1, FillY And 1) = 0 Then FillCol = RightColN
End If

If (Pixels(Tx, FillY) = Area) Then
FillChangedPic = True
Pixels(Tx, FillY) = FillCol
If ResoDiv = 2 Then Pixels(Tx + 1, FillY) = FillCol
Tx = Tx + ResoDiv
Else
 Exit Do
End If
Loop

xMax = Tx - ResoDiv

Tx = FillX - ResoDiv
Do While Tx >= 0

If DitherFill = True Then
If pattern(Tx / ResoDiv And 1, FillY And 1) = 1 Then FillCol = LeftColN
If pattern(Tx / ResoDiv And 1, FillY And 1) = 0 Then FillCol = RightColN
End If

If (Pixels(Tx, FillY) = Area) Then
FillChangedPic = True
Pixels(Tx, FillY) = FillCol
If ResoDiv = 2 Then Pixels(Tx + 1, FillY) = FillCol
Tx = Tx - ResoDiv
Else
 Exit Do
End If
Loop

minx = Tx + ResoDiv

End If

'Display.Refresh
Call Form2.ReDrawDisp

If FillY - 1 >= 0 Then
 'tx = minx
 'Do While tx < xMax
 For Tx = minx To xMax Step ResoDiv
 If Pixels(Tx, FillY - 1) = Area Then Call Fill(Tx, FillY - 1)
 Next Tx
 'tx = tx + 1
 'Loop
End If

If FillY + 1 <= 199 Then
 'tx = minx
 'Do While tx < xMax
 For Tx = minx To xMax Step ResoDiv
 If Pixels(Tx, FillY + 1) = Area Then Call Fill(Tx, FillY + 1)
 Next Tx
 'tx = tx + 1
 'Loop
End If

End Sub


Private Sub form_unload(Cancel As Integer)

For i = Forms.Count - 1 To 0 Step -1
Unload Forms(i)
Next

CopyMemory ByVal VarPtrArray(Pixels), 0&, 4
End Sub

Private Sub Hires_Attrib()
Dim inchar(15) As Integer
Dim gotcha As Boolean
Dim charerr As Double

For x = 0 To 319 Step 8
For y = 0 To 199 Step FliMul

For z = 0 To 15: cols(z) = -1: inchar(z) = BackgrIndex: Next z

 For qx = x To x + 7
  For qy = y To y + FliMul - 1
   cols(ColMap(qx, qy)) = cols(ColMap(qx, qy)) + 1
  Next qy
 Next qx
 

If mostfrq3 = True Then
 For U = 0 To 1
  Max = -1
  For z = 0 To 15
   If cols(z) > Max Then Max = cols(z): Index = z
  Next z
  c(U) = Index: cols(Index) = -1
 Next U
End If

If optchars = True Then
 Max = 0
 For z = 0 To 15
  If cols(z) <> -1 Then inchar(Max) = z: Max = Max + 1
 Next z

 Maxerr = 2147483647
 maxi = Max

 For c0 = 0 To maxi - 1         '0400 #1
 For c1 = c0 + 1 To maxi        '0400 #2
 charerr = 0

 c(0) = inchar(c0)
 c(1) = inchar(c1)

 For qx = x To x + 7
 For qy = y To y + FliMul - 1
   
  pcol = ColMap(qx, qy)
  
   If (pcol <> c(0) And pcol <> c(1)) Then
    Max = 65536: Index = -1
     For z = 0 To 1
      If cd(pcol, c(z)) < Max Then Max = cd(pcol, c(z))
     Next z
    'err = err + Max
    charerr = charerr + Max
   End If
 
 Next qy
 Next qx
 
 If charerr < Maxerr Then
 ScrRam(Bank, x / 8, y / FliMul, 0) = inchar(c0)
 ScrRam(Bank, x / 8, y / FliMul, 1) = inchar(c1)
 Maxerr = charerr
 End If

 Next c1
 Next c0

 c(0) = ScrRam(Bank, x / 8, y / FliMul, 0)
 c(1) = ScrRam(Bank, x / 8, y / FliMul, 1)
 If maxi = 0 Or maxi = 1 Then c(0) = inchar(0)
End If

For qx = x To x + 7
 For qy = y To y + FliMul - 1
   
  pcol = ColMap(qx, qy)
  
   If (pcol <> c(0) And pcol <> c(1)) Then
    Max = 65536: Index = -1
    For z = 0 To 1
     If cd(pcol, c(z)) < Max Then Max = cd(pcol, c(z)): Index = z
    Next z
    err = err + Max
    final = c(Index)
   Else
    final = pcol
   End If
   If x < XFliLimit Then final = 0
   Pixels(qx, qy) = final
  Next qy
 Next qx
 
Next y

Call Form2.ReDrawDisp2(x)

Next x

Call InitAttribs
End Sub

Public Sub lacesetup()

For x = 0 To 319
 For y = 0 To 199
  ColMap(x, y) = Pixels(x, y)
 Next y
Next x


For x = 0 To 319 Step 2
 For y = 0 To 199
  Pixels(Int(x / 2), y) = ColMap(x, y)
 Next y
Next x

StretchBlt Buffer.hdc, 0, 0, _
320, 200, _
Form1.PixelsDib.hdc, _
0, 0, _
160, 200, _
vbSrcCopy


For x = 1 To 319 Step 2
 For y = 0 To 199
  Pixels(Int(x / 2), y) = ColMap(x, y)
 Next y
Next x

StretchBlt Display.hdc, 0, 0, _
320, 200, _
Form1.PixelsDib.hdc, _
0, 0, _
160, 200, _
vbSrcCopy

For x = 0 To 319
 For y = 0 To 199
  Pixels(x, y) = ColMap(x, y)
 Next y
Next x

End Sub
Private Sub Ufli_Attrib()
Dim inchar(15) As Integer
Dim gotcha As Boolean
Dim charerr As Double

BackgrIndex = MostFreqCol

For x = 0 To 319 Step 8
For y = 0 To 199 Step FliMul

For z = 0 To 15: cols(z) = -1: inchar(z) = BackgrIndex: Next z

 For qx = x To x + 7
  For qy = y To y + FliMul - 1
   cols(ColMap(qx, qy)) = cols(ColMap(qx, qy)) + 1
  Next qy
 Next qx
 

If mostfrq3 = True Then
 For U = 0 To 1
  Max = -1
  For z = 0 To 15
   If cols(z) > Max Then Max = cols(z): Index = z
  Next z
  c(U) = Index: cols(Index) = -1
 Next U
 c(2) = BackgrIndex
End If

If optchars = True Then
 Max = 0
 
 cols(BackgrIndex) = -1
 
 For z = 0 To 15
  If cols(z) <> -1 Then inchar(Max) = z: Max = Max + 1
 Next z

 Maxerr = 2147483647
 maxi = Max

 For c0 = 0 To maxi - 1         '0400 #1
 For c1 = c0 + 1 To maxi        '0400 #2
 charerr = 0

 c(0) = inchar(c0)
 c(1) = inchar(c1)
 c(2) = BackgrIndex
 
 For qx = x To x + 7
 For qy = y To y + FliMul - 1
   
  pcol = ColMap(qx, qy)
  
   If (pcol <> c(0) And pcol <> c(1) And pcol <> BackgrIndex) Then
    Max = 65536: Index = -1
     For z = 0 To 2
      If cd(pcol, c(z)) < Max Then Max = cd(pcol, c(z))
     Next z
    'err = err + Max
    charerr = charerr + Max
   End If
 
 Next qy
 Next qx
 
 If charerr < Maxerr Then
 ScrRam(Bank, x / 8, y / FliMul, 0) = inchar(c0)
 ScrRam(Bank, x / 8, y / FliMul, 1) = inchar(c1)
 Maxerr = charerr
 End If

 Next c1
 Next c0

 c(0) = ScrRam(Bank, x / 8, y / FliMul, 0)
 c(1) = ScrRam(Bank, x / 8, y / FliMul, 1)
 If maxi = 0 Or maxi = 1 Then c(0) = inchar(0)
End If

For qx = x To x + 7
 For qy = y To y + FliMul - 1
   
  pcol = ColMap(qx, qy)
  
   If (pcol <> c(0) And pcol <> c(1) And pcol <> BackgrIndex) Then
    Max = 65536: Index = -1
    For z = 0 To 2
     If cd(pcol, c(z)) < Max Then Max = cd(pcol, c(z)): Index = z
    Next z
    err = err + Max
    final = c(Index)
   Else
    final = pcol
   End If
   If x < XFliLimit Then final = 0
   
   If final <> BackgrIndex Then
    If Pixels(qx And 510, qy) <> BackgrIndex Then Pixels(qx, qy) = final
   Else
    Pixels(qx And 510, qy) = final
    Pixels((qx And 510) + 1, qy) = final
    If qx And 510 = 0 Then qx = qx + 1
   End If
   
  Next qy
 Next qx
 
Next y

Call Form2.ReDrawDisp2(x)

Next x

Call InitAttribs
End Sub


Public Sub SaveDrazlace()
Dim Filename As String
Dim ActAttrib(2) As Byte
Dim Bmp As Byte
Dim Ff As Long
Dim Fr As Byte

'5800- d800
'5c00- 0400
'6000- bmp1
'a000- bmp2
'7f40- background

Ff = FreeFile

'get save filename
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
CommonDialog1.Filter = "(*.drl)|*.drl|(*.prg)|*.prg"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowSave
Filename = CommonDialog1.Filename
On Error GoTo 0


Set fso = CreateObject("Scripting.FileSystemObject")

If fso.fileexists(Filename) Then
 fso.deletefile (Filename)
 Open Filename For Binary As #Ff
Else
 Open Filename For Binary As #Ff
End If

Put #Ff, , &H5800

' save d800

For y = 0 To 24
 For x = 0 To 39
 Bmp = d800(x, y) And 15
 'Bmp = 3
 Put #Ff, , Bmp
 Next x
Next y

'pad till 1024
For y = 0 To 23
 Put #Ff, , Bmp
Next y

' save 0400

For y = 0 To 24
 For x = 0 To 39
 Bmp = (ScrRam(Bank, x, y, 1) And 15) + (ScrRam(Bank, x, y, 0) And 15) * 16
 'Bmp = 1 + 2 * 16
 Put #Ff, , Bmp
 Next x
Next y

'pad till 1024
For y = 0 To 23
 Put #Ff, , Bmp
Next y

'save bitmaps

ActAttrib(0) = BackgrIndex
ActAttrib(1) = BackgrIndex
ActAttrib(2) = BackgrIndex

For Fr = 0 To 1

For cy = 0 To 24
 For cx = 0 To 39
  
   ActAttrib(2) = d800(cx, cy)
 
 qx = cx * 8: qy = cy * 8
 For y = 0 To 7
   
   FliY = Int((qy + y) / FliMul)
   ActAttrib(0) = ScrRam(Bank, cx, FliY, 0)
   ActAttrib(1) = ScrRam(Bank, cx, FliY, 1)

   Bmp = 0
   If Pixels(qx + 6 + Fr, qy + y) = ActAttrib(0) Then Bmp = Bmp Or 1
   If Pixels(qx + 6 + Fr, qy + y) = ActAttrib(1) Then Bmp = Bmp Or 2
   If Pixels(qx + 6 + Fr, qy + y) = ActAttrib(2) Then Bmp = Bmp Or 3
      
   If Pixels(qx + 4 + Fr, qy + y) = ActAttrib(0) Then Bmp = Bmp Or 1 * 4
   If Pixels(qx + 4 + Fr, qy + y) = ActAttrib(1) Then Bmp = Bmp Or 2 * 4
   If Pixels(qx + 4 + Fr, qy + y) = ActAttrib(2) Then Bmp = Bmp Or 3 * 4
      
   If Pixels(qx + 2 + Fr, qy + y) = ActAttrib(0) Then Bmp = Bmp Or 1 * 16
   If Pixels(qx + 2 + Fr, qy + y) = ActAttrib(1) Then Bmp = Bmp Or 2 * 16
   If Pixels(qx + 2 + Fr, qy + y) = ActAttrib(2) Then Bmp = Bmp Or 3 * 16
   
   If Pixels(qx + 0 + Fr, qy + y) = ActAttrib(0) Then Bmp = Bmp Or 1 * 64
   If Pixels(qx + 0 + Fr, qy + y) = ActAttrib(1) Then Bmp = Bmp Or 2 * 64
   If Pixels(qx + 0 + Fr, qy + y) = ActAttrib(2) Then Bmp = Bmp Or 3 * 64
   'Bmp = 255
   Put #Ff, , Bmp
  Next y
  
Next cx
Next cy

'skip gap till next bmp

If Fr = 0 Then
Bmp = BackgrIndex
Put #Ff, , Bmp    'store background color
Bmp = 0
For y = 0 To &HBE
 Put #Ff, , Bmp
Next y

End If

Next Fr

'bmp = BackgrIndex
'Put #ff, , bmp


Close #Ff

ErrHandler:
'User pressed the Cancel button
Exit Sub

End Sub

Public Sub mnuLoadDrazlace_click()
Dim Read As Byte
Dim filebuffer(18240) As Byte
Dim andtab(3), divtab(3)
Dim Bmp As Byte
Dim Ff As Long

Ff = FreeFile

andtab(3) = 3 * 1
andtab(2) = 3 * 4
andtab(1) = 3 * 16
andtab(0) = 3 * 64

divtab(3) = 1
divtab(2) = 4
divtab(1) = 16
divtab(0) = 64


CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNPathMustExist
'CommonDialog1.Filter = " |*.kla| |*.koa|"
CommonDialog1.Filter = "(*.drl)|*.drl"
CommonDialog1.FilterIndex = 1

'CommonDialog1.Flags = cdlCCFullOpen 'Or cdlCCPreventFullOpen

CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
On Error GoTo 0
'If Len(filename) = 0 Then GoTo ErrHandler

Open Filename For Binary As #Ff
'Open "c:\test.kla" For Binary As ff

Get #Ff, , Read
Get #Ff, , Read

'5800- d800
'5c00- 0400
'6000- bmp1
'a000- bmp2
'7f40- background

For x = 0 To 18240: Get #Ff, , filebuffer(x): Next

BackgrIndex = filebuffer(&H2740)
Bank = 0

'read d800
q = 0
For cy = 0 To 24
 For cx = 0 To 39
  d800(cx, cy) = filebuffer(q) And 15
  q = q + 1
 Next cx
Next cy

'read 0400
q = 1024
For cy = 0 To 24
 For cx = 0 To 39
  ScrRam(Bank, cx, cy, 0) = Int(filebuffer(q) / 16) And 15
  ScrRam(Bank, cx, cy, 1) = filebuffer(q) And 15
  q = q + 1
 Next cx
Next cy

q = 2048
For Fr = 0 To 1

For cy = 0 To 24
For cx = 0 To 39
 qx = cx * 8: qy = cy * 8
 For y = 0 To 7
  For x = 0 To 3
   Bmp = filebuffer(q)
   temp = (Bmp And andtab(x)) / divtab(x)
   Select Case temp
    Case 0
     ColNum = BackgrIndex
    Case 1
     ColNum = ScrRam(Bank, cx, cy, 0)
    Case 2
     ColNum = ScrRam(Bank, cx, cy, 1)
    Case 3
     ColNum = d800(cx, cy)
   End Select
   
   Pixels(qx + x * 2 + Fr, qy + y) = ColNum
   
  Next x
  q = q + 1
 Next y
Next cx
Call Form2.ReDrawDisp

Next cy
q = q + &HC0
Next Fr

Close #Ff

'Text2.Text = "koala load end!"
Call ZoomWin(ZoomWinLeft, ZoomWinTop)
'Picture3(20).BackColor = c64cols(BackgrIndex)
'Call mnugfxmodedrazlace_Click

Call InitAttribs
GfxMode = "drazlace"
ResoDiv = 1
FliMul = 8
BitBanks = 1
ScrBanks = 0
XFliLimit = 0

Call ZoomWin(ZoomWinLeft, ZoomWinTop)

ErrHandler:
Close #Ff
  'User pressed the Cancel button
  Exit Sub

Close #Ff
'Display.Refresh
End Sub
Private Sub MouseHelper1_MouseWheel(Ctrl As Variant, Direction As MBMouseHelper.mbDirectionConstants, Button As Long, Shift As Long, Cancel As Boolean)
    
    If Direction < 0 Then
     If ZoomScale <> 4 Then
      ZoomScale = ZoomScale - 1
      Call Form_Resize
     End If
    End If
    
    If Direction > 0 Then
     If ZoomScale <> 16 Then
      ZoomScale = ZoomScale + 1
      Call Form_Resize
     End If
    End If
    
    
    Form1.Caption = "Zoom Window 1:" + Str(ZoomScale) + "  Mode: " + GfxMode
    Cancel = True   'cancel the Window's default action
End Sub

'Private Sub MouseHelper1_MouseLeave(Ctrl As Variant, Button As Long, Shift As Long, Cancel As Boolean)

'If Ctrl = ZoomPic Then MouseHelper1.CursorVisible = True ' Text1.Text = "left zoompic"
'End Sub

'Private Sub MouseHelper1_MouseEnter(Ctrl As Variant, Button As Long, Shift As Long, Cancel As Boolean)

' If Ctrl = ZoomPic Then If Ctrl = ZoomPic Then MouseHelper1.CursorVisible = fale 'Text1.Text = "entered zoompic"

'End Sub

