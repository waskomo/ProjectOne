VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ZoomWindow 
   BorderStyle     =   0  'None
   Caption         =   "Zoom Window"
   ClientHeight    =   6825
   ClientLeft      =   450
   ClientTop       =   3465
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "Form1_new_scr_handling.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox PrevPic 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox LoadedPic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   975
      Left            =   3720
      ScaleHeight     =   975
      ScaleWidth      =   1575
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
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
   Begin VB.PictureBox ZoomPic 
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   480
      MouseIcon       =   "Form1_new_scr_handling.frx":1CCA
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   0
      Top             =   1920
      Width           =   2820
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4440
      Picture         =   "Form1_new_scr_handling.frx":3994
      Top             =   3000
      Width           =   240
   End
End
Attribute VB_Name = "ZoomWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'resodiv: 2 if 160x200, 1 if 320x200
'flimul: fli density (i.e.: 8 = koala, 1 = fli)

'Ax,Ay: absolute x, absolute y, c64 coordinates
'leftcol, rightcol: color index assignet to rmb, lmb (leftcoln,rightcoln = 24 bit)
'zoomwinleft,zoomwinright: self explaining..
'zx,zy,zcentered: wether zoomed pic is centered on mdi surface, and its topleft coords when so


'todo:

'top prior:

'- new fileformat for variable size pictures, only one load/save menu, fileformat selects save/load routine
'- plus4 support

'- 4 col mode
'- export sourcecode style txt, and raw 4/8 bit
'- option for 'new' picture
'- fill should be able to replace colors, if color doesnt fit
'- fix colors to certain bitpair, bitpair - color assigment manipulation / char / fli
'- FLI auto optimization of D800 usage

'wd todo list

'- rectangular brush, modern brush gui
'- blinking cursor when editing with keys
'- recolor tool
'-


'buglist

' fill: ditherfill cant detect fast enough area to be filled if area is same color as one of ditherfill color
' copy function: trashes display of main window when centered
' convert shouldnt change screenmode if result is not accepted
' custom save: strange behaviour of enabled disabled states of text boxes
' when saving gif the gif palette should be set to the active 16col palette

'CStr(Now) !!!!!!! to have a string with current date!


'notes:
'pixels(319,199): array of picture
'colmap(319,199): array of colors converted when loading pic



Private ClientW As Long
Private ClientH As Long


'___________________________________________________________________________________________



'drawing line with api
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'prevent hscroll/vscroll events redraw zoomwindow, when zoomresize changes scrollbar values
Private NoScroll As Boolean

'coords to center zoompic
Public Zx As Long
Public Zy As Long

'keyboard handling
Dim KeyX As Long
Dim KeyY As Long
Dim KeyLeftDown As Boolean 'for pixeling with keyboard lef/right pressed status
Dim KeyRightDown As Boolean

'shit
Dim Z As Long

Public PicName As String






'stuff for declaring an array over a 256 color dibsection bitmap data
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



Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                               lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'let shared mousedown, mouseup, and mousemove know where it was called from
Public EventSource As String

' move ZoomWinReDraw when
Dim MoveZoomWin As Boolean


'to remember last mouse position on zoomwindow, for brush redraw when mousewheeling
Dim LastZoomx As Single
Dim LastZoomy As Single

Dim ScrBank As Long
Dim BmpBank As Long



Private leftcol As Long
Private rightcol As Long


Dim ColNum As Long
Attribute ColNum.VB_VarUserMemId = 1073938503




Dim RightClick As Boolean, LeftClick As Boolean
Attribute RightClick.VB_VarUserMemId = 1073938526
Attribute LeftClick.VB_VarUserMemId = 1073938526







Dim FillCol As Long, Area As Long
Attribute FillCol.VB_VarUserMemId = 1073938548

Public xx1 As Long
Public xx2 As Long
Public xx3 As Long
Public yy1 As Long
Public yy2 As Long
Public yy3 As Long



Dim X As Long
Attribute X.VB_VarUserMemId = 1073938578
Dim Y As Long
Attribute Y.VB_VarUserMemId = 1073938579








Dim UnBitMap() As Byte
Dim UnScrRam() As Byte
Dim UnPixels() As Byte
Dim UnD800() As Byte
Dim UnD021() As Byte


Dim RedoCount As Long
Dim UndoPtr As Long
Dim UndoOffs As Long

'bitmap 16000
'scrram 32000
'pixels 64000
'd800    1000
'------------
'      113000


Dim FillChangedPic As Boolean
Attribute FillChangedPic.VB_VarUserMemId = 1073938603


' new attribute system



Public BP00 As Long
Attribute BP00.VB_VarUserMemId = 1073938614
Public BP01 As Long
Attribute BP01.VB_VarUserMemId = 1073938615
Public BP10 As Long
Attribute BP10.VB_VarUserMemId = 1073938616
Public BP11 As Long
Attribute BP11.VB_VarUserMemId = 1073938617




Public BitPairFits As Long
Attribute BitPairFits.VB_VarUserMemId = 1073938624

Public ActiveTool As String


Public Sub mnuRedo_Click()
        '<EhHeader>
        On Error GoTo mnuRedo_Click_Err
        '</EhHeader>

100 Call Redo
102 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
104 Call ZoomWinRefresh
106 Call Palett.UpdateColors
108 Call PrevWin.ReDraw

        '<EhFooter>
        Exit Sub

mnuRedo_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuRedo_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Private Sub Redo()
        '<EhHeader>
        On Error GoTo Redo_Err
        '</EhHeader>
    Dim Ptr As Long

100 If RedoCount > 0 Then
102     RedoCount = RedoCount - 1
104     UndoPtr = UndoPtr + 1
    End If

106 MainWin.mnuUndo.Enabled = True
108 If RedoCount <> 0 Then
110     MainWin.mnuRedo.Enabled = True
    Else
112     MainWin.mnuRedo.Enabled = False
    End If

114 Ptr = (UndoPtr + UndoOffs) And 15

116 CopyMemory Bitmap(0, 0, 0), ByVal VarPtr(UnBitMap(0, 0, 0, Ptr)), LenB(UnBitMap(0, 0, 0, 0)) * CW * PH * 2 ' 16000
118 CopyMemory ScrRam(0, 0, 0, 0, 0), ByVal VarPtr(UnScrRam(0, 0, 0, 0, 0, Ptr)), LenB(UnScrRam(0, 0, 0, 0, 0, 0)) * CW * CH * 8 * 2 * 2 ' 32000
120 CopyMemory Pixels(0, 0), ByVal VarPtr(UnPixels(0, 0, Ptr)), LenB(UnPixels(0, 0, 0)) * PW * PH '* 64000
122 CopyMemory D800(0, 0), ByVal VarPtr(UnD800(0, 0, Ptr)), LenB(UnD800(0, 0, 0)) * CW * CH '* 1000
124 CopyMemory BackgrIndex, ByVal VarPtr(UnD021(Ptr)), LenB(BackgrIndex) * 1

    'Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
        '<EhFooter>
        Exit Sub

Redo_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.Redo" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub mnuUndo_Click()
        '<EhHeader>
        On Error GoTo mnuUndo_Click_Err
        '</EhHeader>

100 Call ReadUndo

102 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
104 Call PrevWin.ReDraw
106 Call ZoomWinRefresh
108 Call Palett.UpdateColors

        '<EhFooter>
        Exit Sub

mnuUndo_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuUndo_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ReadUndo()
        '<EhHeader>
        On Error GoTo ReadUndo_Err
        '</EhHeader>
    Dim Ptr As Long

100 If UndoPtr > 0 Then
102     UndoPtr = UndoPtr - 1
104     RedoCount = RedoCount + 1
    End If

106 MainWin.mnuRedo.Enabled = True
108 If UndoPtr = 0 Then MainWin.mnuUndo.Enabled = False

110 Ptr = (UndoPtr + UndoOffs) And 15

112 CopyMemory Bitmap(0, 0, 0), ByVal VarPtr(UnBitMap(0, 0, 0, Ptr)), LenB(UnBitMap(0, 0, 0, 0)) * CW * PH * 2 ' 16000
114 CopyMemory ScrRam(0, 0, 0, 0, 0), ByVal VarPtr(UnScrRam(0, 0, 0, 0, 0, Ptr)), LenB(UnScrRam(0, 0, 0, 0, 0, 0)) * CW * CH * 8 * 2 * 2 ' 32000
116 CopyMemory Pixels(0, 0), ByVal VarPtr(UnPixels(0, 0, Ptr)), LenB(UnPixels(0, 0, 0)) * PW * PH '* 64000
118 CopyMemory D800(0, 0), ByVal VarPtr(UnD800(0, 0, Ptr)), LenB(UnD800(0, 0, 0)) * CW * CH '* 1000
120 CopyMemory BackgrIndex, ByVal VarPtr(UnD021(Ptr)), LenB(BackgrIndex) * 1

    'Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
        '<EhFooter>
        Exit Sub

ReadUndo_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ReadUndo" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub WriteUndo()
        '<EhHeader>
        On Error GoTo WriteUndo_Err
        '</EhHeader>
    Dim Ptr As Long

100 RedoCount = 0
102 MainWin.mnuRedo.Enabled = False
104 MainWin.mnuUndo.Enabled = True

106 If UndoPtr < 15 Then
108     UndoPtr = UndoPtr + 1
    Else
110     UndoOffs = (UndoOffs + 1) And 15
    End If

112 Ptr = (UndoPtr + UndoOffs) And 15

    'bitmap 16000
    'scrram 32000
    'pixels 64000
    'd800    1000
    '------------
    '      113000

114 CopyMemory UnBitMap(0, 0, 0, Ptr), ByVal VarPtr(Bitmap(0, 0, 0)), LenB(Bitmap(0, 0, 0)) * CW * PH * 2
116 CopyMemory UnScrRam(0, 0, 0, 0, 0, Ptr), ByVal VarPtr(ScrRam(0, 0, 0, 0, 0)), LenB(ScrRam(0, 0, 0, 0, 0)) * CW * CH * 8 * 2 * 2
118 CopyMemory UnPixels(0, 0, Ptr), ByVal VarPtr(Pixels(0, 0)), LenB(Pixels(0, 0)) * PW * PH
120 CopyMemory UnD800(0, 0, Ptr), ByVal VarPtr(D800(0, 0)), LenB(D800(0, 0)) * CW * CH
122 CopyMemory UnD021(Ptr), ByVal VarPtr(BackgrIndex), LenB(BackgrIndex) * 1

    'Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
        '<EhFooter>
        Exit Sub

WriteUndo_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.WriteUndo" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ResetUndo()
        '<EhHeader>
        On Error GoTo ResetUndo_Err
        '</EhHeader>
    Dim Ptr As Long

100 MainWin.mnuRedo.Enabled = False
102 MainWin.mnuUndo.Enabled = False

104 UndoPtr = 0
106 UndoOffs = 0
108 RedoCount = 0
110 Ptr = 0

112 CopyMemory UnBitMap(0, 0, 0, Ptr), ByVal VarPtr(Bitmap(0, 0, 0)), LenB(Bitmap(0, 0, 0)) * CW * PH * 2
114 CopyMemory UnScrRam(0, 0, 0, 0, 0, Ptr), ByVal VarPtr(ScrRam(0, 0, 0, 0, 0)), LenB(ScrRam(0, 0, 0, 0, 0)) * CW * CH * 8 * 2 * 2
116 CopyMemory UnPixels(0, 0, Ptr), ByVal VarPtr(Pixels(0, 0)), LenB(Pixels(0, 0)) * PW * PH
118 CopyMemory UnD800(0, 0, Ptr), ByVal VarPtr(D800(0, 0)), LenB(D800(0, 0)) * CW * CH
120 CopyMemory UnD021(Ptr), ByVal VarPtr(BackgrIndex), LenB(BackgrIndex) * 1

    'Text1.Text = "ptr:" + Str(UndoPtr) + " offs:" + Str(UndoOffs) + " redo:" + Str(RedoCount)
        '<EhFooter>
        Exit Sub

ResetUndo_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ResetUndo" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub grid(ZoomX As Long, ZoomY As Long)
        '<EhHeader>
        On Error GoTo grid_Err
        '</EhHeader>

    Dim XX As Single
    Dim YY As Single
    Dim gridx As Single
    Dim gridy As Single
    Dim blah As POINTAPI
    Dim Xstart As Single
    Dim xEnd As Single
    Dim Ystart As Single
    Dim Yend As Single

100 ZoomPic.DrawMode = 13
102 ZoomPic.DrawStyle = 0

104 gridy = CellHeigth - (Int(ZoomY) Mod CellHeigth)
106 gridx = CellWidth - (Int(ZoomX) Mod CellWidth)

108 ZoomX = Int(ZoomX / ResoDiv) * ResoDiv
    'MainWin.Text1.Text = "zx:" + Str(ZoomX) + " zy:" + Str(ZoomY)

110 If ZoomX < ZoomWidth Then
112     Xstart = ((ZoomWidth - ZoomX) * ZoomScaleX) + Zx
114     gridx = 0
    Else
116     Xstart = Zx
118     gridx = CellWidth - (Int(ZoomX - ZoomWidth) Mod CellWidth)
    End If

120 If ZoomX > PW Then
122     xEnd = ((PW - ZoomX + ZoomWidth) * ZoomScaleX) + Zx
124     gridx = CellWidth - (Int(ZoomX - ZoomWidth) Mod CellWidth)
    Else
126     xEnd = Zx + (ZoomWidth * ZoomScaleX)
    End If

128 If ZoomY < ZoomHeight Then
130     Ystart = ((ZoomHeight - ZoomY) * ZoomScaleY) + Zy
132     gridy = 0
    Else
134     Ystart = Zy
136     gridy = CellHeigth - (Int(ZoomY - ZoomHeight) Mod CellHeigth)
    End If

138 If ZoomY > PH Then
140     Yend = ((PH - ZoomY + ZoomHeight) * ZoomScaleY) + Zy
142     gridy = CellHeigth - (Int(ZoomY - ZoomHeight) Mod CellHeigth)
    Else
144     Yend = Zy + (ZoomHeight * ZoomScaleY)
    End If

146 If ZoomPicCenteredX = True Then
148     ZoomX = 0
150     Xstart = Zx
152     xEnd = Zx + (ZoomWidth * ZoomScaleX)
    End If

154 If ZoomPicCenteredY = True Then
156     ZoomY = 0
158     Ystart = Zy
160     Yend = Zy + (ZoomHeight * ZoomScaleY)
    End If

    'Draw pixel grid
162 If PixelGrid = 1 And ZoomScale > PixelGridLimit Then

164     ZoomPic.Forecolor = PixelGridColor

166     For XX = Xstart To xEnd Step ResoDiv * ZoomScaleX
168         MoveToEx ZoomPic.hdc, XX, Ystart, ByVal 0&
170         LineTo ZoomPic.hdc, XX, Yend
172     Next XX

174     For YY = Ystart To Yend Step ZoomScaleY
176         MoveToEx ZoomPic.hdc, Xstart, YY, ByVal 0&
178         LineTo ZoomPic.hdc, xEnd, YY
180     Next YY

    End If

    'Draw FLI lines
182 If ShowFliLines = 1 And ZoomScale > FliGridLimit Then
184     ZoomPic.DrawWidth = 1
186     ZoomPic.Forecolor = FliGridColor

188     If ZoomY < ZoomHeight Then XX = 0 Else XX = ZoomScaleX * CellWidth
190     For YY = Ystart + gridy * ZoomScaleY - XX To Yend Step FliMul * ZoomScaleY 'zoomscalemod
192         MoveToEx ZoomPic.hdc, Xstart, YY, ByVal 0&
194         LineTo ZoomPic.hdc, xEnd, YY
196     Next YY
    End If

    'Draw charborders grid

198 If CharGrid = 1 And ZoomScale > CharGridLimit Then

200     ZoomPic.DrawWidth = 1
202     ZoomPic.Forecolor = CharGridColor

204     For XX = Xstart + gridx * ZoomScaleX To xEnd Step CellWidth * ZoomScaleX
206         MoveToEx ZoomPic.hdc, XX, Ystart, ByVal 0&
208         LineTo ZoomPic.hdc, XX, Yend
210     Next XX

212     For YY = Ystart + gridy * ZoomScaleY To Yend Step CellHeigth * ZoomScaleY        '8
214         MoveToEx ZoomPic.hdc, Xstart, YY, ByVal 0&
216         LineTo ZoomPic.hdc, xEnd, YY
218     Next YY

    End If

        '<EhFooter>
        Exit Sub

grid_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.grid" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub CoordLimit(ByRef X As Long, ByRef Y As Long)
        '<EhHeader>
        On Error GoTo CoordLimit_Err
        '</EhHeader>

100 If X > PW + ZoomWidth Then X = PW + ZoomWidth
102 If X < 0 Then X = 0
104 If Y > PH + ZoomHeight Then Y = PH + ZoomHeight
106 If Y < 0 Then Y = 0

108 If ZoomPicCenteredY = True Then Y = ZoomHeight
110 If ZoomPicCenteredX = True Then X = ZoomWidth

112 X = Int(X / ResoDiv) * ResoDiv
114 Y = Int(Y)

        '<EhFooter>
        Exit Sub

CoordLimit_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.CoordLimit" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ZoomWinReDraw(ByVal ZoomX As Long, ByVal ZoomY As Long)
        '<EhHeader>
        On Error GoTo ZoomWinReDraw_Err
        '</EhHeader>
        Dim zw As Integer, zh As Integer, TempZoomX As Integer
    
100     Call CoordLimit(ZoomX, ZoomY)

102     TempZoomX = (Int(ZoomX / ResoDiv) * ResoDiv) - ZoomWidth

104     zw = ZoomWidth * ZoomScaleX
106     zh = ZoomHeight * ZoomScaleY

108     ZoomPic.Line (0, 0)-(ZoomPic.Width, ZoomPic.Height), &H8000000C, BF
        

110     If ZoomPicCenteredX = False Or ZoomPicCenteredY = False Then
112         StretchBlt ZoomPic.hdc, Zx, Zy, zw, zh, PixelsDib.hdc, TempZoomX, ZoomY - ZoomHeight, ZoomWidth, ZoomHeight, vbSrcCopy
        Else
114         StretchBlt ZoomPic.hdc, Zx, Zy, zw, zh, PixelsDib.hdc, 0, 0, ZoomWidth, ZoomHeight, vbSrcCopy
        End If


116     Call grid(ZoomX, ZoomY)

        '<EhFooter>
        Exit Sub

ZoomWinReDraw_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomWinReDraw" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ZoomWinRefresh()
        '<EhHeader>
        On Error GoTo ZoomWinRefresh_Err
        '</EhHeader>

100 MainWin.Picture = ZoomPic.Image
102 Call MainWin.invalidate

        '<EhFooter>
        Exit Sub

ZoomWinRefresh_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomWinRefresh" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub CalcPrevPicCords(X As Long, Y As Long)
        '<EhHeader>
        On Error GoTo CalcPrevPicCords_Err
        '</EhHeader>

100 Ax = (Int(X / ResoDiv) * ResoDiv): Ay = Int(Y)

102 If Ax > PW - ResoDiv Then Ax = PW - ResoDiv
104 If Ay > PH - 1 Then Ay = PH - 1
106 If Ax < 0 Then Ax = 0
108 If Ay < 0 Then Ay = 0


110 If KeyX > PW - 1 Then KeyX = PW - ResoDiv
112 If KeyY > PH - 1 Then KeyY = PH - 1
114 If KeyX < 0 Then KeyX = 0
116 If KeyY < 0 Then KeyY = 0

        '<EhFooter>
        Exit Sub

CalcPrevPicCords_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.CalcPrevPicCords" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub CalcZoomCords(X As Long, Y As Long)
        '<EhHeader>
        On Error GoTo CalcZoomCords_Err
        '</EhHeader>

100 If ZoomPicCenteredX = False Then
102     Ax = Int(ZoomWinLeft + Int(X / ResoDiv) * ResoDiv) - ZoomWidth
104     If Ax > ZoomWinLeft - ResoDiv Then Ax = ZoomWinLeft - ResoDiv
    Else
106     Ax = Int(X / ResoDiv) * ResoDiv
    End If

108 If ZoomPicCenteredY = False Then
110     Ay = Int(ZoomWinTop) + Int(Y) - ZoomHeight
112     If Ay > ZoomWinTop - 1 Then Ay = ZoomWinTop - 1
    Else
114     Ay = Int(Y)
    End If

116 If Ax > PW - 1 Then Ax = (PW - 1) - (ResoDiv - 1)
118 If Ay > PH - 1 Then Ay = PH - 1
120 If Ax < 0 Then Ax = 0
122 If Ay < 0 Then Ay = 0


        '<EhFooter>
        Exit Sub

CalcZoomCords_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.CalcZoomCords" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub Shared_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim db As Long

    'update keyboard coordinates
    If EventSource <> "keyboard" Then
        KeyX = Ax: KeyY = Ay
    End If

    'move ZoomWin if necessary
    If RightClick = True And EventSource = "PrevPic" Then
        ZoomWinLeft = Ax + ZoomWidth / 2
        ZoomWinLeft = Int(ZoomWinLeft / ResoDiv) * ResoDiv
        ZoomWinTop = Ay + ZoomHeight / 2
        Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
        Call ZoomWinRefresh
        MainWin.HScroll1.Value = ZoomWinLeft 'TempZoomX + ZoomWidth
        MainWin.VScroll1.Value = ZoomWinTop 'ZoomY
    End If

    'draw if mousebutton hold, and mode is draw
    If (ActiveTool = "draw" Or ActiveTool = "brush") And MoveZoomWin = False Then
        If LeftClick = True Or RightClick = True Then
            If Sqr((Ax - OldAx) ^ 2 + (Ay - OldAy) ^ 2) <= 1 Then
                Call Paint(Ax, Ay, ColNum)
                Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
            Else
                Call DrawLine(OldAx, OldAy, Ax, Ay, ColNum)
                Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
            End If
        End If
    End If


    If (LeftClick = False Or RightClick = False) And ActiveTool = "brush" Then
        'Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
    End If
    
    CX = Int(Ax / 8)
    CY = Int(Ay / 8)
    On Error Resume Next
    db = Bitmap(0, CX, Ay)
    On Error GoTo 0
    'Dec2Bin (db) & " "
    MainWin.Text1.Text = "a" 'Str(D800(CX, CY)) & Str(Screen(Ax, Ay, 0)) & Str(Screen(Ax, Ay, 1))

    MainWin.Text1.Text = Str(db) & " " & Str(D800(CX, CY)) & Str(ScrRam(0, 0, CX, CY, 0)) & Str(ScrRam(0, 0, CX, CY, 1)) '& Dec2Bin(db)
    
    'copy mode 1: show dynamic selected area
    If ActiveTool = "copy1" And LeftClick = True Then
        xx2 = Int((Ax / 8) + 0.5) * 8
        yy2 = Int((Ay / 8) + 0.5) * 8
    End If

    Call DrawCursors
    Call PrevWin.DrawCursors(Ax, Ay)
    Call ZoomWinRefresh
    
    If Not (Ax = OldAx And Ay = OldAy) Then Call StatusBarUpdate
    OldAx = Ax: OldAy = Ay

End Sub

Public Sub DrawCursors()
        '<EhHeader>
        On Error GoTo DrawCursors_Err
        '</EhHeader>

    'redraw various cursors if necessary
100 If Not (Ax = OldAx And Ay = OldAy) Then
        
102     Select Case ActiveTool
            Case "draw"
104             If PixelBox = 1 Then DrawPixelBox
106         Case "fill"
108             If PixelBox = 1 Then DrawPixelBox
110         Case "brush"
112             Call ShowBrushCursor(Ax, Ay)
        End Select
    
    End If

        '<EhFooter>
        Exit Sub

DrawCursors_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.DrawCursors" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
'draw box around current pixel
Private Sub DrawPixelBox()
    'VBCH Dim blah As POINTAPI
        '<EhHeader>
        On Error GoTo DrawPixelBox_Err
        '</EhHeader>
    Dim ox As Single
    Dim oy As Single
    Dim X As Single
    Dim Y As Single


        'colors of the pixelbox
100     If PixelBoxColor = 1 Then
102         If PixelFits(Ax, Ay, LeftColN) >= 0 Then leftcol = PaletteRGB(LeftColN) Else leftcol = RGB(255, 0, 0)
104         If PixelFits(Ax, Ay, RightColN) >= 0 Then rightcol = PaletteRGB(RightColN) Else rightcol = RGB(255, 0, 0)
        Else
106         leftcol = RGB(255, 255, 255)
108         rightcol = leftcol
        End If

110     If ZoomPicCenteredX = False Then
112         X = (Ax - ZoomWinLeft + ZoomWidth) * ZoomScaleX + Zx
114         ox = (OldAx - ZoomWinLeft + ZoomWidth) * ZoomScaleX + Zx
        Else
116         X = (Ax) * ZoomScaleX + Zx + 0
118         ox = (OldAx) * ZoomScaleX + Zx + 0
        End If

120     If ZoomPicCenteredY = False Then
122         Y = (Ay - ZoomWinTop + ZoomHeight) * ZoomScaleY + Zy
124         oy = (OldAy - ZoomWinTop + ZoomHeight) * ZoomScaleY + Zy
        Else
126         Y = (Ay) * ZoomScaleY + Zy + 0
128         oy = (OldAy) * ZoomScaleY + Zy + 0
        End If

130     ZoomPic.Forecolor = PaletteRGB(Pixels(OldAx, OldAy))
132     ZoomPic.Line (ox + 1, oy + 1)-((ox + ResoDiv * ZoomScaleX) - 1, (oy + 1 * ZoomScaleY) - 1), , BF
    
134     If PixelBoxColor = 1 Then
136         ZoomPic.DrawMode = 13
        Else
138         ZoomPic.DrawMode = 7
        End If

140     ZoomPic.Forecolor = leftcol
142     ZoomPic.Line (X + 1, (Y + 1 * ZoomScaleY) - 1)-(X + 1, Y + 1)
144     ZoomPic.Line (X + 1, Y + 1)-((X + ResoDiv * ZoomScaleX) - 1, Y + 1)
146     ZoomPic.Forecolor = rightcol
148     ZoomPic.Line ((X + ResoDiv * ZoomScaleX) - 1, Y + 1)-((X + ResoDiv * ZoomScaleX) - 1, (Y + 1 * ZoomScaleY) - 1)
150     ZoomPic.Line ((X + ResoDiv * ZoomScaleX) - 1, (Y + 1 * ZoomScaleY) - 1)-(X + 1, (Y + 1 * ZoomScaleY) - 1)
    
152     ZoomPic.DrawMode = 13


        '<EhFooter>
        Exit Sub

DrawPixelBox_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.DrawPixelBox" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

'draw box around current FLI limiter
Private Sub DrawFLIBox()
        '<EhHeader>
        On Error GoTo DrawFLIBox_Err
        '</EhHeader>
        Dim blah As POINTAPI
        Dim ox As Long
        Dim oy As Long
        Dim oCx As Long
        Dim oCy As Long



100     CX = Int(Ax / 8) * 8
102     CY = Int(Ay / FliMul) * FliMul
104     oCx = Int(OldAx / 8) * 8
106     oCy = Int(OldAy / FliMul) * FliMul

108     If ZoomPicCenteredX = False Then
110         X = (CX - ZoomWinLeft + ZoomWidth) * ZoomScaleX + Zx
112         ox = (oCx - ZoomWinLeft + ZoomWidth) * ZoomScaleX + Zx
        Else
114         X = (CX) * ZoomScaleX + Zx + 0
116         ox = (oCx) * ZoomScaleX + Zx + 0
        End If

118     If ZoomPicCenteredY = False Then
120         Y = (CY - ZoomWinTop + ZoomHeight) * ZoomScaleY + Zy
122         oy = (oCy - ZoomWinTop + ZoomHeight) * ZoomScaleY + Zy
        Else
124         Y = (CY) * ZoomScaleY + Zy + 0
126         oy = (oCy) * ZoomScaleY + Zy + 0
        End If

        'ZoomPic.ForeColor = PaletteRGB(pixels(OldAx, OldAy))
        'ZoomPic.Line (ox + 1, oy + 1)-((ox + resodiv * ZoomScale) - 1, (oy + 1 * ZoomScale) - 1), , BF

128     ZoomPic.DrawMode = 7
130     ZoomPic.Forecolor = RGB(255, 255, 255)

132     MoveToEx ZoomPic.hdc, ox, (oy + FliMul * ZoomScaleY), ByVal 0&
134     LineTo ZoomPic.hdc, ox, oy
136     LineTo ZoomPic.hdc, (ox + 8 * ZoomScaleX), oy

138     LineTo ZoomPic.hdc, (ox + 8 * ZoomScaleX), (oy + FliMul * ZoomScaleY)
140     LineTo ZoomPic.hdc, ox, (oy + FliMul * ZoomScaleY)

142     MoveToEx ZoomPic.hdc, X, (Y + FliMul * ZoomScaleY), ByVal 0&
144     LineTo ZoomPic.hdc, X, Y
146     LineTo ZoomPic.hdc, (X + 8 * ZoomScaleX), Y

148     LineTo ZoomPic.hdc, (X + 8 * ZoomScaleX), (Y + FliMul * ZoomScaleY)
150     LineTo ZoomPic.hdc, X, (Y + FliMul * ZoomScaleY)

152     ZoomPic.DrawMode = 13


        '<EhFooter>
        Exit Sub

DrawFLIBox_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.DrawFLIBox" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub



Public Sub Shared_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Shared_MouseDown_Err
        '</EhHeader>
    Dim Index As Long

    'stop lace when drawing
100 If MainWin.mnuInterlaceEmu.Checked = True Then StopInterlace

102 If Button = 2 Then RightClick = True
104 If Button = 1 Then LeftClick = True

    'drag start of ZoomWinReDraw
106 If Button = 2 And EventSource = "PrevPic" Then
108     MoveZoomWin = True
110     NoScroll = True
112     Call ZoomWinReDraw(Ax + (ZoomWidth / 2), Ay + (ZoomHeight / 2))
    End If

114 If ActiveTool = "draw" Then

        'OldAx = Ax
        'OldAy = Ay

116     Index = Pixels(Ax, Ay)
118     If Button = 1 Then
120         ColNum = LeftColN
122         If Shift = 0 Then
124             Call SetPixel(Ax, Ay, LeftColN)
            End If
126         If Shift = 1 Then
128             leftcol = PaletteRGB(Index): LeftColN = Index
130             ColNum = LeftColN
132             Call Palett.UpdateColors
            End If
134         If Shift = 2 Then
136             Call Drawing.ChangeCol(LeftColN)
138             Call PrevWin.ReDraw
            End If
140         If Shift = 4 Then
142             DitherMode = True
            End If
        End If

144     If Button = 2 And EventSource = "zoompic" Then
146         ColNum = RightColN
148         If Shift = 0 Then
150             Call SetPixel(Ax, Ay, RightColN)
            End If
152         If Shift = 1 Then
154             rightcol = PaletteRGB(Index): RightColN = Index
156             ColNum = RightColN
158             Call Palett.UpdateColors
            End If
160         If Shift = 2 Then
162             Call ChangeCol(RightColN)
164             Call PrevWin.ReDraw
            End If
166         If Shift = 4 Then
168             DitherMode = True
            End If

        End If

        'mouse middle
170     If Button = 4 Then
172         ColNum = BackgrIndex
174         Call SetPixel(Ax, Ay, BackgrIndex)
        End If

176 ElseIf ActiveTool = "fill" Then

178     If Button = 1 Then
180         FillCol = LeftColN
182         Call StartFill
        End If

184     If Button = 2 And EventSource = "zoompic" Then
186         FillCol = RightColN
188         Call StartFill
        End If

190 ElseIf ActiveTool = "brush" And Button = 1 Then

192     Call Paint(Ax, Ay, LeftColN)
194     Call PrevWin.ReDraw
196     Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

198 ElseIf ActiveTool = "debug" Then

200     Call debugger(Ax, Ay)

202 ElseIf ActiveTool = "copy" And Button = 1 Then

204     xx1 = Int((Ax / 8) + 0.5) * 8
206     yy1 = Int((Ay / 8) + 0.5) * 8
208     LeftClick = True
210     ActiveTool = "copy1"


212 ElseIf ActiveTool = "copy2" And Button = 1 Then

214     If Ax > xx1 And Ax < xx2 And Ay > yy1 And Ay < yy2 Then
216         ActiveTool = "copy3"
218         xx3 = Int((Ax / 8) + 0.5) * 8
220         yy3 = Int((Ay / 8) + 0.5) * 8
222         LeftClick = True
        End If

    End If


224 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
226 Call DrawPixelBox
228 Call ZoomWinRefresh


        '<EhFooter>
        Exit Sub

Shared_MouseDown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.Shared_MouseDown" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub Shared_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Shared_MouseUp_Err
        '</EhHeader>
    Dim Temp As Integer


    'this is from zoompic
100 DitherMode = False

    'fill does already writes undo on its own
102 If MoveZoomWin = False And ActiveTool <> "fill" Then Call WriteUndo

104 MoveZoomWin = False
106 NoScroll = False

108 If Button = 2 Then
110     RightClick = False        ' Turn off painting.
112 ElseIf Button = 1 Then
114     LeftClick = False

116     If ActiveTool = "copy1" Then
118         ActiveTool = "copy2"
120         xx2 = Int((Ax / 8) + 0.5) * 8
122         yy2 = Int((Ay / 8) + 0.5) * 8
124         If xx2 = xx1 Or yy2 = yy1 Then
126             ActiveTool = "copy"
            End If
128         If xx1 > xx2 And yy1 > yy2 Then
130             Temp = xx1: xx1 = xx2: xx2 = Temp
132             Temp = yy1: yy1 = yy2: yy2 = Temp
134         ElseIf xx1 > xx2 And yy1 < yy2 Then
136             Temp = xx1: xx1 = xx2: xx2 = Temp
138         ElseIf xx1 < xx2 And yy1 > yy2 Then
140             Temp = yy1: yy1 = yy2: yy2 = Temp
            End If
        End If

142     If ActiveTool = "copy3" Then
144         Call copy
146         ActiveTool = "copy"
148         Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
150         Call PrevWin.ReDraw
        End If
    End If

    'Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

        '<EhFooter>
        Exit Sub

Shared_MouseUp_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.Shared_MouseUp" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub mnuBrush_Click()
        '<EhHeader>
        On Error GoTo mnuBrush_Click_Err
        '</EhHeader>
100 BrushDialog.Show vbModal
        '<EhFooter>
        Exit Sub

mnuBrush_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuBrush_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub mnuDitherfill_Click()
        '<EhHeader>
        On Error GoTo mnuDitherfill_Click_Err
        '</EhHeader>
100 MainWin.mnuDitherfill.Checked = Not MainWin.mnuDitherfill.Checked
102 DitherFill = Not DitherFill
        '<EhFooter>
        Exit Sub

mnuDitherfill_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuDitherfill_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub mnustrictfill_Click()
        '<EhHeader>
        On Error GoTo mnustrictfill_Click_Err
        '</EhHeader>
100 MainWin.mnuCompensatingFill.Checked = False
102 MainWin.mnuStrictFill.Checked = True
104 FillMode = "strict"
        '<EhFooter>
        Exit Sub

mnustrictfill_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnustrictfill_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub mnuCompensatingFill_Click()
        '<EhHeader>
        On Error GoTo mnuCompensatingFill_Click_Err
        '</EhHeader>
100 MainWin.mnuCompensatingFill.Checked = True
102 MainWin.mnuStrictFill.Checked = False
104 FillMode = "compensating"
        '<EhFooter>
        Exit Sub

mnuCompensatingFill_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuCompensatingFill_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub ChangeGfxModeFromConvert()
        '<EhHeader>
        On Error GoTo ChangeGfxModeFromConvert_Err
        '</EhHeader>

100 If GfxMode = "custom" Then Call mnuGfxModeCustom_Click
102 If GfxMode = "hires" Then Call mnuGfxModeHires_Click
104 If GfxMode = "koala" Then Call mnuGfxModeKoala_Click
106 If GfxMode = "drazlace" Then Call mnuGfxModeDrazlace_Click
108 If GfxMode = "afli" Then Call mnuGfxModeAfli_Click
110 If GfxMode = "ifli" Then Call mnuGfxModeIFli_Click
112 If GfxMode = "fli" Then Call mnuGfxModeFli_Click
114 If GfxMode = "drazlacespec" Then Call mnuGfxModeDrazlaceSpec_Click
116 If GfxMode = "unrestricted hires" Then BaseMode = BaseModeTyp.unrestricted
118 If GfxMode = "unrestricted multi" Then BaseMode = BaseModeTyp.unrestricted

120 Call UpdateFileMenu

        '<EhFooter>
        Exit Sub

ChangeGfxModeFromConvert_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ChangeGfxModeFromConvert" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ChangeGfxMode()
        '<EhHeader>
        On Error GoTo ChangeGfxMode_Err
        '</EhHeader>

100 If GfxMode = "custom" Then Call mnuGfxModeCustom_Click
102 If GfxMode = "hires" Then Call mnuGfxModeHires_Click
104 If GfxMode = "koala" Then Call mnuGfxModeKoala_Click
106 If GfxMode = "drazlace" Then Call mnuGfxModeDrazlace_Click
108 If GfxMode = "afli" Then Call mnuGfxModeAfli_Click
110 If GfxMode = "ifli" Then Call mnuGfxModeIFli_Click
112 If GfxMode = "fli" Then Call mnuGfxModeFli_Click
114 If GfxMode = "drazlacespec" Then Call mnuGfxModeDrazlaceSpec_Click

116 Call UpdateFileMenu
118 Call ModeChangeReset
        '<EhFooter>
        Exit Sub

ChangeGfxMode_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ChangeGfxMode" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ModeChangeReset()
        '<EhHeader>
        On Error GoTo ModeChangeReset_Err
        '</EhHeader>

    Dim X As Long
    Dim Y As Long

100 Call UpdateFileMenu

102 For X = 0 To PW - 1
104     For Y = 0 To PH - 1
    
106     ColMap(X, Y) = Pixels(X, Y)
    
108     Next Y
110 Next X

112 Call Color_Restrict
114 Call InitAttribs

    'DrawPicFromMem
116 Call PrevWin.ReDraw
118 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
120 Call ZoomWinRefresh
122 Call SetMdiCaption

        '<EhFooter>
        Exit Sub

ModeChangeReset_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ModeChangeReset" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Private Sub UpdateFileMenu()
        '<EhHeader>
        On Error GoTo UpdateFileMenu_Err
        '</EhHeader>

100 MainWin.mnuLoadKoala.Visible = False
102 MainWin.mnuSaveKoala.Visible = False
104 MainWin.mnuLoadDrazlace.Visible = False
106 MainWin.mnuSaveDrazlace.Visible = False
108 MainWin.mnuLoadFli.Visible = False
110 MainWin.mnuSaveFli.Visible = False
112 MainWin.mnuSeparator.Visible = False

114 If GfxMode = "koala" Then
116     MainWin.mnuLoadKoala.Visible = True
118     MainWin.mnuSaveKoala.Visible = True
120     MainWin.mnuSeparator.Visible = True
    End If

122 If GfxMode = "drazlace" Then
124     MainWin.mnuLoadDrazlace.Visible = True
126     MainWin.mnuSaveDrazlace.Visible = True
128     MainWin.mnuSeparator.Visible = True
    End If

130 If GfxMode = "fli" Then
132     MainWin.mnuLoadFli.Visible = True
134     MainWin.mnuSaveFli.Visible = True
136     MainWin.mnuSeparator.Visible = True
    End If

        '<EhFooter>
        Exit Sub

UpdateFileMenu_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.UpdateFileMenu" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Private Sub UnCheckModes()
        '<EhHeader>
        On Error GoTo UnCheckModes_Err
        '</EhHeader>
100 MainWin.mnuGfxModeHires.Checked = False
102 MainWin.mnuGfxModeKoala.Checked = False
104 MainWin.mnuGfxModeDrazlace.Checked = False
106 MainWin.mnuGfxModeAfli.Checked = False
108 MainWin.mnuGfxModeIFli.Checked = False
110 MainWin.mnuGfxModeFli.Checked = False
112 MainWin.mnuGfxModeDrazlaceSpec.Checked = False
114 MainWin.mnuGfxModeUnrestrictedMulti.Checked = False
116 MainWin.mnuGfxModeUnrestrictedHires.Checked = False
118 MainWin.mnuGfxModeCustom.Checked = False
120 Call SetMdiCaption
        '<EhFooter>
        Exit Sub

UnCheckModes_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.UnCheckModes" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub mnuGfxModeCustom_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeCustom_Click_Err
        '</EhHeader>

100 BaseMode = BaseMode_cm
102 GfxMode = "custom"
104 ResoDiv = ResoDiv_cm
106 FliMul = FliMul_cm
108 ScrBanks = ScrBanks_cm
110 BmpBanks = BmpBanks_cm
112 XFliLimit = XFliLimit_cm
114 Call UnCheckModes
116 MainWin.mnuGfxModeCustom.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeCustom_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeCustom_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub mnuGfxModeHires_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeHires_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.hires
102 GfxMode = "hires"
104 ResoDiv = 1
106 FliMul = 8
108 ScrBanks = 0
110 BmpBanks = 0
112 XFliLimit = 0
114 Call UnCheckModes
116 MainWin.mnuGfxModeHires.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeHires_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeHires_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub mnuGfxModeKoala_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeKoala_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.multi
102 GfxMode = "koala"
104 ResoDiv = 2
106 FliMul = 8
108 ScrBanks = 0
110 BmpBanks = 0
112 XFliLimit = 0
114 Call UnCheckModes
116 MainWin.mnuGfxModeKoala.Checked = True
118 Debug.Print "mnugfxmodekoala"

        '<EhFooter>
        Exit Sub

mnuGfxModeKoala_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeKoala_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub mnuGfxModeDrazlace_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeDrazlace_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.multi
102 GfxMode = "drazlace"
104 ResoDiv = 1
106 FliMul = 8
108 ScrBanks = 0
110 BmpBanks = 1
112 XFliLimit = 0
114 Call UnCheckModes
116 MainWin.mnuGfxModeDrazlace.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeDrazlace_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeDrazlace_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub mnuGfxModeAfli_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeAfli_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.hires
102 GfxMode = "afli"
104 ResoDiv = 1
106 FliMul = 1
108 ScrBanks = 0
110 BmpBanks = 0
112 XFliLimit = 24
114 Call UnCheckModes
116 MainWin.mnuGfxModeAfli.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeAfli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeAfli_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub mnuGfxModeIFli_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeIFli_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.multi
102 GfxMode = "ifli"
104 ResoDiv = 1
106 FliMul = 1
108 ScrBanks = 1
110 BmpBanks = 1
112 XFliLimit = 24
114 Call UnCheckModes
116 MainWin.mnuGfxModeIFli.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeIFli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeIFli_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub mnuGfxModeFli_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeFli_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.multi
102 GfxMode = "fli"
104 ResoDiv = 2
106 FliMul = 1
108 ScrBanks = 0
110 BmpBanks = 0
112 XFliLimit = 24
114 Call UnCheckModes
116 MainWin.mnuGfxModeFli.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeFli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeFli_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub mnuGfxModeDrazlaceSpec_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeDrazlaceSpec_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.multi
102 GfxMode = "drazlacespec"
104 ResoDiv = 1
106 FliMul = 8
108 ScrBanks = 1
110 BmpBanks = 1
112 XFliLimit = 0
114 Call UnCheckModes
116 MainWin.mnuGfxModeDrazlaceSpec.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeDrazlaceSpec_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeDrazlaceSpec_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub mnuGfxModeUnrestrictedHires_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeUnrestrictedHires_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.unrestricted
102 GfxMode = "unrestrictedhires"
104 ResoDiv = 1
106 FliMul = 8
108 ScrBanks = 0
110 BmpBanks = 0
112 XFliLimit = 0
114 Call UnCheckModes
116 MainWin.mnuGfxModeUnrestrictedMulti.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeUnrestrictedHires_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeUnrestrictedHires_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub mnuGfxModeUnrestrictedMulti_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeUnrestrictedMulti_Click_Err
        '</EhHeader>

100 BaseMode = BaseModeTyp.unrestricted
102 GfxMode = "unrestrictedmulti"
104 ResoDiv = 2
106 FliMul = 8
108 ScrBanks = 0
110 BmpBanks = 0
112 XFliLimit = 0
114 Call UnCheckModes
116 MainWin.mnuGfxModeUnrestrictedMulti.Checked = True

        '<EhFooter>
        Exit Sub

mnuGfxModeUnrestrictedMulti_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuGfxModeUnrestrictedMulti_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub



Public Sub mnuCopy_Click()
        '<EhHeader>
        On Error GoTo mnuCopy_Click_Err
        '</EhHeader>

100     LoadedPic.Width = PW: LoadedPic.Height = PH
102     StretchBlt LoadedPic.hdc, 0, 0, PW, PH, PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy
104     Clipboard.Clear
106     Clipboard.SetData LoadedPic.Image

        '<EhFooter>
        Exit Sub

mnuCopy_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuCopy_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub VScroll1_scroll()
        '<EhHeader>
        On Error GoTo VScroll1_scroll_Err
        '</EhHeader>

100     Call VScroll1_Change

        '<EhFooter>
        Exit Sub

VScroll1_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.VScroll1_scroll" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub HScroll1_scroll()
        '<EhHeader>
        On Error GoTo HScroll1_scroll_Err
        '</EhHeader>

100     Call HScroll1_Change

        '<EhFooter>
        Exit Sub

HScroll1_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.HScroll1_scroll" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub VScroll1_Change()
        '<EhHeader>
        On Error GoTo VScroll1_Change_Err
        '</EhHeader>

100 ZoomWinTop = MainWin.VScroll1.Value

102 If NoScroll = False Then
104     Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
106     Call PrevWin.DrawCursors(Ax, Ay)
108     Call ZoomWinRefresh
    End If

        '<EhFooter>
        Exit Sub

VScroll1_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.VScroll1_Change" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub HScroll1_Change()
        '<EhHeader>
        On Error GoTo HScroll1_Change_Err
        '</EhHeader>

100 ZoomWinLeft = MainWin.HScroll1.Value
102 ZoomWinLeft = Int(ZoomWinLeft / ResoDiv) * ResoDiv

104 If NoScroll = False Then
106     Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
108     Call ZoomWinRefresh
110     Call PrevWin.DrawCursors(Ax, Ay)
    End If

        '<EhFooter>
        Exit Sub

HScroll1_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.HScroll1_Change" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub StatusBarUpdate()
        '<EhHeader>
        On Error GoTo StatusBarUpdate_Err
        '</EhHeader>

100 If ScrBanks = 1 Then ScrBank = Ax And 1 Else ScrBank = 0
102 ScrNum = Int((Ay And 7) / FliMul)

104 MainWin.StatusBar1.Panels(5).Text = "x:" + Str(Ax / ResoDiv)
106 MainWin.StatusBar1.Panels(6).Text = "y:" + Str(Ay)
108 MainWin.StatusBar1.Panels(7).Text = "col:" + Str(CX)
110 MainWin.StatusBar1.Panels(8).Text = "row:" + Str(CY)

    'MainWin.StatusBar1.Panels(9).Picture = PaletteRGB(LeftColN)
    'MainWin.StatusBar1.Panels(10).Picture = PaletteRGB(RightColN)

        '<EhFooter>
        Exit Sub

StatusBarUpdate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.StatusBarUpdate" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub Paint(drX As Long, drY As Long, ByVal drColor As Long)
        '<EhHeader>
        On Error GoTo Paint_Err
        '</EhHeader>

100     Select Case ActiveTool

            Case "draw"

102             Call SetPixel(drX, drY, drColor)

104         Case "brush"

106             Call PlotBrush(drX, drY)

        End Select

        '<EhFooter>
        Exit Sub

Paint_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.Paint" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub




Public Sub mnuPaste_Click()
        '<EhHeader>
        On Error GoTo mnuPaste_Click_Err
        '</EhHeader>


100 If Clipboard.GetFormat(vbCFBitmap) Then


        'disable interlace
102     StopInterlace

104     LoadedPic.Cls
106     LoadedPic.Forecolor = 0
108     LoadedPic.Line (0, 0)-(LoadedPic.Width, LoadedPic.Height), 0, BF

110     Set LoadedPic.Picture = Clipboard.GetData()
112     Set PrevPic.Picture = Clipboard.GetData()
114     LoadedPic.Refresh

        'Call ConvertDialog.ResizePic
        'ConvertDialog.Original.Picture = PrevPic.image
        'ConvertDialog.SrcPic.Picture = PrevPic.image
    
116     Call WriteUndo
118     ConvertDialog.Show vbModal

120     Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
122     Call ResetUndo

    End If


        '<EhFooter>
        Exit Sub

mnuPaste_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuPaste_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub StopInterlace()
        '<EhHeader>
        On Error GoTo StopInterlace_Err
        '</EhHeader>

100 PrevWin.Timer1.Enabled = False
102 PrevWin.ReDraw
104 MainWin.mnuInterlaceEmu.Checked = False

        '<EhFooter>
        Exit Sub

StopInterlace_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.StopInterlace" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub mnuLoadPicture_Click()
        '<EhHeader>
        On Error GoTo mnuLoadPicture_Click_Err
        '</EhHeader>
    Dim Pointer As Long

100 StopInterlace

102 If GetLoadName("(*.bmp)|*.bmp|(*.*)|*.*") = False Then Exit Sub

    'load pic with gdi
104 Set PrevPic.Picture = LoadPicture(FileName)
106 Set LoadedPic.Picture = LoadPicture(FileName)

108 If PrevPic.Picture = 0 Then
110     MsgBox "Error: Can't load picture, may be not a picture file." 'probably the user selected not a pic file
        Exit Sub
    End If

    'so we can return to old pic if user presses cancel on convertdialog
112 Call WriteUndo
114 ConvertDialog.Show vbModal
116 Call ResetUndo

        '<EhFooter>
        Exit Sub

mnuLoadPicture_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuLoadPicture_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub





Public Sub mnuExportPicture_Click()
        '<EhHeader>
        On Error GoTo mnuExportPicture_Click_Err
        '</EhHeader>

100 LoadedPic.Width = PW
102 LoadedPic.Height = PH

104 StretchBlt LoadedPic.hdc, 0, 0, PW, PH, PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy

106 If GetSaveName("(*.bmp)|*.bmp|(*.*)|*.*") = False Then Exit Sub

108 If FileExists(FileName) Then
110     Kill (FileName)
    End If

112 SavePicture LoadedPic.Image, FileName

        '<EhFooter>
        Exit Sub

mnuExportPicture_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.mnuExportPicture_Click" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Private Sub ZoomPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo ZoomPic_MouseDown_Err
        '</EhHeader>

100 EventSource = "zoompic"
102 Call Shared_MouseDown(Button, Shift, X, Y)

        '<EhFooter>
        Exit Sub

ZoomPic_MouseDown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomPic_MouseDown" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Private Sub ZoomPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo ZoomPic_MouseUp_Err
        '</EhHeader>

100 EventSource = "zoompic"
102 Call Shared_MouseUp(Button, Shift, X, Y)
        '<EhFooter>
        Exit Sub

ZoomPic_MouseUp_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomPic_MouseUp" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub DrawLine(ByVal x1 As Integer, ByVal y1 As Integer, _
                     ByVal x2 As Integer, ByVal y2 As Integer, ByVal lColor As Long)
        '<EhHeader>
        On Error GoTo DrawLine_Err
        '</EhHeader>

    Dim Xdist As Integer
    Dim Ydist As Integer
    Dim Xdir As Integer
    Dim Ydir As Integer
    Dim YY As Single
    Dim XX As Single
    Dim Xstep As Single
    Dim Ystep As Single
    Dim Xb As Long
    Dim Yb As Long

    Dim Temp As Integer


100 Xdist = Abs(x1 - x2)
102 Ydist = Abs(y1 - y2)

104 If Xdist = 0 Then

106     If y1 > y2 Then
108         Temp = y1: y1 = y2: y2 = Temp
        End If

110     Xb = x1
        'Text1.Text = "vertical line!"
112     For Yb = y1 To y2
114         Call Paint(Xb, Yb, lColor)
116     Next Yb

118 ElseIf Ydist = 0 Then

120     If x1 > x2 Then
122         Temp = x1: x1 = x2: x2 = Temp
        End If

124     Yb = y1
126     For Xb = x1 To x2 Step ResoDiv
128         Call Paint(Xb, Yb, lColor)
130     Next Xb

132 ElseIf Xdist / ResoDiv > Ydist Then

134     If x1 > x2 Then
136         Temp = x1: x1 = x2: x2 = Temp
138         Temp = y1: y1 = y2: y2 = Temp
        End If


140     Ydir = Sgn(y2 - y1)

142     Ystep = (Ydist / Xdist) * Ydir * ResoDiv
144     YY = y1
146     For Xb = x1 To x2 Step ResoDiv
148         Call Paint(Xb, Int(YY), lColor)
150         YY = YY + Ystep
152     Next Xb

    Else

154     If y1 > y2 Then
156         Temp = x1: x1 = x2: x2 = Temp
158         Temp = y1: y1 = y2: y2 = Temp
        End If

160     Xdir = Sgn(x2 - x1)

162     Xstep = (Xdist / Ydist) * Xdir
164     XX = x1
166     For Yb = y1 To y2 Step 1
168         Call Paint(Int((XX / ResoDiv) + 0.5) * ResoDiv, Yb, lColor)
170         XX = XX + Xstep
172     Next Yb

    End If

        '<EhFooter>
        Exit Sub

DrawLine_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.DrawLine" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub ShowBrushCursor(ByVal X As Long, ByVal Y As Long)
        '<EhHeader>
        On Error GoTo ShowBrushCursor_Err
        '</EhHeader>

100 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

102 If ZoomPicCenteredX = False Then
        X = (Ax - ZoomWinLeft + ZoomWidth) * ZoomScaleX + Zx:
    Else
        X = (Ax) * ZoomScaleX + Zx + 0:
    End If

104 If ZoomPicCenteredY = False Then
106     Y = (Ay - ZoomWinTop + ZoomHeight) * ZoomScaleY + Zy
    Else
108     Y = (Ay) * ZoomScaleY + Zy + 0
    End If

    'ZoomPic.DrawMode = 7
    'ZoomPic.DrawStyle = 2
110 ZoomPic.Forecolor = RGB(255, 255, 255)
112 ZoomPic.Circle (X, Y), (BrushSize / 2) * ZoomScale
114 ZoomPic.DrawStyle = 0

116 Call ZoomWinRefresh

        '<EhFooter>
        Exit Sub

ShowBrushCursor_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ShowBrushCursor" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ZoomPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo ZoomPic_MouseMove_Err
        '</EhHeader>

100 MouseOver = "zoompic"

    'only used by brush :P
102 LastZoomx = X
104 LastZoomy = Y

106 EventSource = "zoompic"
    'calc coordinates on zoompic move

108 X = ScaleX(X, vbTwips, vbPixels)
110 Y = ScaleY(Y, vbTwips, vbPixels)
112 X = (X / ZoomScaleX) - Zx / ZoomScaleX
114 Y = (Y / ZoomScaleY) - Zy / ZoomScaleY
116 Call CalcZoomCords(Int(X), Int(Y))

118 Call Shared_Mousemove(Button, Shift, X, Y)

        '<EhFooter>
        Exit Sub

ZoomPic_MouseMove_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomPic_MouseMove" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub



Public Sub ResizeCanvas(ByRef PW As Long, ByRef PH As Long)
        '<EhHeader>
        On Error GoTo ResizeCanvas_Err
        '</EhHeader>

100 CW = Int(PW / 8)
102 CH = Int(PH / 8)

104 ReDim ColMap(PW, PH) As Byte
106 ReDim Pixels(PW, PH) As Byte
108 ReDim D800(CW, CH) As Byte
110 ReDim ScrRam(1, 7, CW, CH, 1) As Byte
112 ReDim Bitmap(1, CW, PH) As Byte
114 ReDim Registers(1, 63, PH) As Byte
116 ReDim ColMap(PW, PH) As Byte

118 ReDim UnBitMap(1, CW, PH, 15) As Byte
120 ReDim UnScrRam(1, 7, CW, CH, 1, 15) As Byte
122 ReDim UnPixels(PW, PH, 15) As Byte
124 ReDim UnD800(CW, PW, 15) As Byte
126 ReDim UnD021(15) As Byte


128 ReDim BBitmap(1, CW, PH) As Byte
130 ReDim BScrRam(1, 7, CW, CH, 1) As Byte
132 ReDim Bd800(CW, PW) As Byte


        '<EhFooter>
        Exit Sub

ResizeCanvas_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ResizeCanvas" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub form_load()
        '<EhHeader>
        On Error GoTo form_load_Err
        '</EhHeader>
    Dim Out As Long

    'DLL TESTING
    'Out = 2
    'MsgBox modGraphical.GPX_DummyDouble(Out)
    'MsgBox Out

100 Debug.Print "zoomwindow form_load"

    'dither arry init for brush
102 Bayer(0, 0) = 1: Bayer(0, 1) = 9: Bayer(0, 2) = 3: Bayer(0, 3) = 11
104 Bayer(1, 0) = 13: Bayer(1, 1) = 5: Bayer(1, 2) = 15: Bayer(1, 3) = 7
106 Bayer(2, 0) = 4: Bayer(2, 1) = 12: Bayer(2, 2) = 2: Bayer(2, 3) = 10
108 Bayer(3, 0) = 16: Bayer(3, 1) = 8: Bayer(3, 2) = 14: Bayer(3, 3) = 6

    'set up icons
    'Set MainWin.Icon = ZoomWindow.Image1.Picture
110 SetAppIcon MainWin
112 Set Me.Icon = ZoomWindow.Image1.Picture



    'load default mouse icon
114 MainWin.MousePointer = vbCustom
116 MainWin.MouseIcon = LoadPicture(App.Path & "\Cursors\pencil.ico")


    'setup bitmask tables
118 Drawing.SetupBitTables



120 Call LoadSettings           '##################################### Load Settings
122 Call ResizeCanvas(PW, PH)


    'create our 8 bit paletted picture with 'pixels' array defined on top of it
124 Set PixelsDib = New cDIBSection256
126 PixelsDib.Create PW, -PH


128 BrushSize = BrushSize
130 BrushDither = BrushDither
132 BrushColor1 = BrushColor1
134 BrushColor2 = BrushColor2

136 With tsain
138     .cbElements = 1
140     .cDims = 2
142     .Bounds(0).lLbound = 0
144     .Bounds(0).cElements = PixelsDib.Height
146     .Bounds(1).lLbound = 0
148     .Bounds(1).cElements = PixelsDib.BytesPerScanLine()
150     .pvData = PixelsDib.DIBSectionBitsPtr
    End With

152 CopyMemory ByVal VarPtrArray(Pixels), VarPtr(tsain), 4


    'Non user GUI settings
154 ZoomWindow.ScaleMode = vbPixels

156 ZoomPic.ScaleMode = vbPixels
158 ZoomPic.BackColor = 0
160 ZoomPic.AutoRedraw = True
162 ZoomPic.BorderStyle = 0

164 PrevPic.ScaleMode = vbPixels
166 PrevPic.BackColor = 0
168 PrevPic.BorderStyle = 0
170 PrevPic.Width = PW
172 PrevPic.Height = PH
174 PrevPic.AutoRedraw = True

176 LoadedPic.ScaleMode = vbPixels
178 LoadedPic.BackColor = 0
180 LoadedPic.BorderStyle = 0
182 LoadedPic.AutoRedraw = True





184 pattern(0, 0) = 0: pattern(0, 1) = 1
186 pattern(1, 0) = 1: pattern(1, 1) = 0

188 DitherMode = False
190 ResizeScale = 100






192 If FillMode = "strict" Then
194     MainWin.mnuStrictFill.Checked = True
196 ElseIf FillMode = "compensating" Then
198     MainWin.mnuCompensatingFill.Checked = "true"
    End If

200 Call LoadPalette(pal_Selected(ChipType))

202 Call PaletteInit

204 Call InitAttribs
206 GfxMode = "koala"
208 Call ChangeGfxMode

210 Call BrushPreCalc        ' generate brush

212 Ax = Int(PW / 2)
214 Ay = Int(PH / 2)

216 Call ZoomResize
218 Call ZoomWindow.ZoomWinRefresh
220 Call PrevWin.ResizePrevPic
222 Call PrevWin.ChangeWinSize
224 Call Palett.UpdateColors

    'non user init

    'reset toolbar
226 ActiveTool = "draw"
228 Call MainWin.ResetButtonMenuCaption
230 MainWin.Toolbar1.Buttons(1).Value = tbrPressed


    'statusbar setup

232 MainWin.StatusBar1.Height = 25 * 10
234 For Z = 2 To 11
236     MainWin.StatusBar1.Panels.Add (Z)
238 Next Z

240 For Z = 1 To 10
242     MainWin.StatusBar1.Panels(Z).Width = 47 * 15
244     MainWin.StatusBar1.Panels(Z).AutoSize = sbrNoAutoSize
246 Next Z

    'for help txt
248 MainWin.StatusBar1.Panels(11).Width = 47 * 15
250 MainWin.StatusBar1.Panels(11).AutoSize = sbrSpring
252 MainWin.StatusBar1.Panels(11).Text = ""

254 MainWin.StatusBar1.Panels(1).Text = "ScrL"
256 MainWin.StatusBar1.Panels(2).Text = "ScrH"
258 MainWin.StatusBar1.Panels(3).Text = "D800"
260 MainWin.StatusBar1.Panels(4).Text = "D021"

262 MainWin.StatusBar1.Panels(5).Text = "x:"
264 MainWin.StatusBar1.Panels(6).Text = "y:"
266 MainWin.StatusBar1.Panels(7).Text = "c:"
268 MainWin.StatusBar1.Panels(8).Text = "r:"

270 MainWin.StatusBar1.Panels(9).Text = "lmb"
272 MainWin.StatusBar1.Panels(10).Text = "rmb"

    'MainWin.StatusBar1.Panels(9).Picture = PaletteRGB(LeftColN)
    'MainWin.StatusBar1.Panels(10).Picture = PaletteRGB(RightColN)

274 MainWin.mnuPaletteWindow.Checked = True
276 MainWin.mnuPreviewWindow.Checked = True

278 Call SetMdiCaption

280 PrevWin.Show
282 Palett.Show

        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.form_load" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub form_unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo form_unload_Err
        '</EhHeader>
    Dim i As Long
100 Call SaveSettings
    


    'kill array assigned over pixelsdib graphmem (???)
102 CopyMemory ByVal VarPtrArray(Pixels), 0&, 4

104 For i = Forms.Count - 1 To 0 Step -1
106     Unload Forms(i)
    Next


        '<EhFooter>
        Exit Sub

form_unload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.form_unload" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub SetMdiCaption()
        '<EhHeader>
        On Error GoTo SetMdiCaption_Err
        '</EhHeader>
    Dim lace As String
    Dim gfxchip As String
    Dim basmod As String

100 If BmpBanks = 1 Then lace = " Interlaced " Else lace = ""

102 Select Case ChipType
        Case Chip.vicii
104         gfxchip = "VICII"
106     Case Chip.ted
108         gfxchip = "TED"
110     Case Chip.vdc
112         gfxchip = "VDC"
    End Select

114 Select Case BaseMode
        Case BaseModeTyp.hires
116         basmod = "Hires"
118     Case BaseModeTyp.multi
120         basmod = "MultiColor"
    End Select


122 MainWin.Caption = "P1    Zoom Rate: 1:" & Str(ZoomScale) & "      Mode: " & GfxMode & "    (" & gfxchip & ", " & lace & " " & basmod & ", Fli lines per char:" & Str(Int(8 / FliMul)) & ")"
        '<EhFooter>
        Exit Sub

SetMdiCaption_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.SetMdiCaption" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub debugger(dbx, dby)
        '<EhHeader>
        On Error GoTo debugger_Err
        '</EhHeader>

    Dim Temp As String
    Dim db As Long
    Dim db2 As Long
    'VBCH Dim tempy As Long
    Dim dy As Long

100 CX = Int(dbx / 8)
102 CY = Int(dby / FliMul)
104 dy = Int(dby / 8)


106 FileName = App.Path + "debug_" + Str(Date) + "_" + ".txt"
108 Open FileName For Output As #1

110 Print #1, "           x:" + Str(dbx And 7)
112 Print #1, "           y:" + Str(dby And 7)
114 Print #1, "      column:" + Str(CX)
116 Print #1, "         row:" + Str(dy)
118 Print #1, "    col left:" + Str(LeftColN)
120 Print #1, "    col rite:" + Str(RightColN)
122 Print #1, " "
124 Print #1, "        mode:" + GfxMode
126 Print #1, "bitmap banks:" + Str(BmpBanks)
128 Print #1, "screen banks:" + Str(ScrBanks)
130 Print #1, "     ResoDiv:" + Str(ResoDiv)
132 Print #1, "      FliMul:" + Str(FliMul)
134 Print #1, "  "
136 Print #1, "Static Colors:"
138 Print #1, "d021(00):" + Str(BackgrIndex)
140 Print #1, "d800(11):" + Str(D800(CX, dy))
142 Print #1, " "
144 Print #1, " "

146 For BmpBank = 0 To BmpBanks

148     BP00 = 0
150     BP01 = 0
152     BP10 = 0
154     BP11 = 0

156     For Y = 0 To 7

158         db = Bitmap(BmpBank, CX, (dy * 8) + Y)
160         Temp = Str(Y) + ": "        '+ Str(db) + " "
162         BP00 = BP00 + BitCount00(db)
164         BP01 = BP01 + BitCount01(db)
166         BP10 = BP10 + BitCount10(db)
168         BP11 = BP11 + BitCount11(db)

170         For Z = 3 To 0 Step -1
172             db2 = (db And Shift(3, Z)) / (4 ^ Z)
174             Select Case db2
                Case 0
176                 Temp = Temp + "00 "
178             Case 1
180                 Temp = Temp + "01 "
182             Case 2
184                 Temp = Temp + "10 "
186             Case 3
188                 Temp = Temp + "11 "
                End Select
190         Next Z
192         ScrNum = Int(((dby + Y) And 7) / FliMul)
194         If ScrBanks = 1 Then ScrBank = BmpBank Else ScrBank = 0
196         Temp = Temp + "|"
198         Print #1, Temp;
200         Print #1, "  01:" + Str(ScrRam(ScrBank, ScrNum, CX, dy, 0)); Tab;
202         Print #1, "  10:" + Str(ScrRam(ScrBank, ScrNum, CX, dy, 1)); Tab        '; '" tempy: " + Str(tempy) + " dby: " + Str(dby) + " y: " + Str(y)
            'Print #1, temp; Tab
204     Next Y

206     Print #1, " "
208     Print #1, "count 00:" + Str(BP00)
210     Print #1, "count 01:" + Str(BP01)
212     Print #1, "count 10:" + Str(BP10)
214     Print #1, "count 11:" + Str(BP11)
216     Print #1, " "

218 Next BmpBank

220 For Y = 0 To 7
222     Temp = ""
224     For X = 0 To 7
226         Temp = Temp & Val(Pixels(X + (CX * 8), Y + (dy * 8))) & " "
228     Next X
230     Print #1, Temp
232 Next Y
234 Close #1

236 MsgBox ("Debug information was written into the Application path:" + FileName)

        '<EhFooter>
        Exit Sub

debugger_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.debugger" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub ZoomResize()
        '<EhHeader>
        On Error GoTo ZoomResize_Err
        '</EhHeader>
    Dim H As Long
    Dim W As Long
    Dim H2 As Long
    Dim W2 As Long

100 NoScroll = True

102 ZoomScaleX = ZoomScale * ARatioX
104 ZoomScaleY = ZoomScale * ARatioY

106 H = ScaleY(MainWin.ScaleHeight, vbTwips, vbPixels)
108 W = ScaleX(MainWin.ScaleWidth, vbTwips, vbPixels)
110 ClientH = H
112 ClientW = W
114 H2 = H
116 W2 = W

118 ZoomPicCenteredX = False
120 ZoomPicCenteredY = False

122 If Int(H / ZoomScaleY) > PH - 0 Then
124     H = (PH - 0) * ZoomScaleY
126     ZoomPicCenteredY = True
128     MainWin.VScroll1.Enabled = False
130     MainWin.VScroll1.Height = ScaleY(ClientH, vbPixels, vbTwips)
132     Zy = (H2 - Int(H / ZoomScaleY) * ZoomScaleY) / 2
    Else
134     MainWin.VScroll1.Enabled = True
136     MainWin.VScroll1.Height = ScaleY(H, vbPixels, vbTwips)
138     Zy = 0
    End If

140 If Int(W / ZoomScaleX) > PW - 0 Then
142     W = (PW - 0) * ZoomScaleX
144     ZoomPicCenteredX = True
146     MainWin.HScroll1.Enabled = False
148     MainWin.HScroll1.Width = ScaleX(ClientW, vbPixels, vbTwips)
150     Zx = (W2 - Int(W / ZoomScaleX) * ZoomScaleX) / 2
    Else
152     MainWin.HScroll1.Enabled = True
154     MainWin.HScroll1.Width = ScaleX(W, vbPixels, vbTwips)
156     Zx = 0
    End If

158 ZoomHeight = Int(H / ZoomScaleY) '+1 = partially visible pixels possible
160 ZoomWidth = Int(W / ZoomScaleX)  '+resodiv = partially visible pixels possible
162 ZoomWidth = Int(ZoomWidth / ResoDiv) * ResoDiv

164 ZoomPic.Move 0, 0, ClientW, ClientH

166 If ZoomPicCenteredX = True Or ZoomPicCenteredY = True Then

168     ZoomPic.Forecolor = &H8000000C
170     ZoomPic.Line (0, 0)-(ClientW, ClientH), , BF

172     ZoomPic.Forecolor = RGB(60, 60, 60)
174     ZoomPic.Line (Zx + 8, Zy + 8)-(Zx + (ZoomWidth * ZoomScaleX) + 8, Zy + (ZoomHeight * ZoomScaleY) + 8), , BF
176     ZoomPic.Line (Zx + 7, Zy + 7)-(Zx + (ZoomWidth * ZoomScaleX) + 7, Zy + (ZoomHeight * ZoomScaleY) + 7), 0, B

    End If

178 MainWin.HScroll1.Max = PW + ZoomWidth
180 MainWin.VScroll1.Max = PH + ZoomHeight

    'ReDraw Zoompic
182 ZoomWinTop = Ay + Int(ZoomHeight / 2)
184 ZoomWinLeft = Ax + Int(ZoomWidth / 2)
186 ZoomWinLeft = Int(ZoomWinLeft / ResoDiv) * ResoDiv

188 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

190 Call ZoomWinRefresh
192 Call PrevWin.DrawCursors(Ax, Ay)

194 NoScroll = False
        '<EhFooter>
        Exit Sub

ZoomResize_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomResize" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Private Sub form_keyup(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo form_keyup_Err
        '</EhHeader>
100 Call Shared_keyup(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

form_keyup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.form_keyup" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub Shared_keyup(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo Shared_keyup_Err
        '</EhHeader>

100 If KeyCode = 17 Then KeyRightDown = False
102 If KeyCode = 32 Then KeyLeftDown = False
        '<EhFooter>
        Exit Sub

Shared_keyup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.Shared_keyup" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Private Sub form_keydown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo form_keydown_Err
        '</EhHeader>
100 Call Shared_keydown(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

form_keydown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.form_keydown" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Private Sub SetPixelFromKeyboard()
        '<EhHeader>
        On Error GoTo SetPixelFromKeyboard_Err
        '</EhHeader>
        
100 If KeyRightDown = True Then
102         SetPixel Ax, Ay, RightColN
104         Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
106         Call DrawPixelBox
108         Call ZoomWinRefresh
110         Call PrevWin.DrawCursors(Ax, Ay)
    End If

112 If KeyLeftDown = True Then
114         SetPixel Ax, Ay, LeftColN
116         Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
118         Call DrawPixelBox
120         Call ZoomWinRefresh
122         Call PrevWin.DrawCursors(Ax, Ay)
    End If

        '<EhFooter>
        Exit Sub

SetPixelFromKeyboard_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.SetPixelFromKeyboard" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub Shared_keydown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo Shared_keydown_Err
        '</EhHeader>
    Dim byteColor As Byte

    'Text1.Text = Str(KeyCode)
    'Text1.Refresh

    'MsgBox (Str(KeyCode) + " " + Shift)

100 If KeyCode = 73 And Shift = 2 Then Call MainWin.mnuInterlaceEmu_Click

102 If Shift <> 2 Then

        'alt gr
104     If KeyCode = 17 Then
106         KeyRightDown = True
108         Call SetPixelFromKeyboard
        End If

        'space
110     If KeyCode = 32 Then
112         KeyLeftDown = True
114         Call SetPixelFromKeyboard
        End If

        '0-9, qwertz
116     If KeyCode = 48 And Shift = 0 Then LeftColN = 0
118     If KeyCode = 49 And Shift = 0 Then LeftColN = 1
120     If KeyCode = 50 And Shift = 0 Then LeftColN = 2
122     If KeyCode = 51 And Shift = 0 Then LeftColN = 3
124     If KeyCode = 52 And Shift = 0 Then LeftColN = 4
126     If KeyCode = 53 And Shift = 0 Then LeftColN = 5
128     If KeyCode = 54 And Shift = 0 Then LeftColN = 6
130     If KeyCode = 55 And Shift = 0 Then LeftColN = 7
132     If KeyCode = 56 And Shift = 0 Then LeftColN = 8
134     If KeyCode = 57 And Shift = 0 Then LeftColN = 9
136     If KeyCode = 65 And Shift = 0 Then LeftColN = 10
138     If KeyCode = 66 And Shift = 0 Then LeftColN = 11
140     If KeyCode = 67 And Shift = 0 Then LeftColN = 12
142     If KeyCode = 68 And Shift = 0 Then LeftColN = 13
144     If KeyCode = 69 And Shift = 0 Then LeftColN = 14
146     If KeyCode = 70 And Shift = 0 Then LeftColN = 15
148     If KeyCode = 71 And Shift = 0 Then LeftColN = 15

150     If KeyCode = 48 And Shift = 1 Then RightColN = 0
152     If KeyCode = 49 And Shift = 1 Then RightColN = 1
154     If KeyCode = 50 And Shift = 1 Then RightColN = 2
156     If KeyCode = 51 And Shift = 1 Then RightColN = 3
158     If KeyCode = 52 And Shift = 1 Then RightColN = 4
160     If KeyCode = 53 And Shift = 1 Then RightColN = 5
162     If KeyCode = 54 And Shift = 1 Then RightColN = 6
164     If KeyCode = 55 And Shift = 1 Then RightColN = 7
166     If KeyCode = 56 And Shift = 1 Then RightColN = 8
168     If KeyCode = 57 And Shift = 1 Then RightColN = 9
170     If KeyCode = 65 And Shift = 1 Then RightColN = 10
172     If KeyCode = 66 And Shift = 1 Then RightColN = 11
174     If KeyCode = 67 And Shift = 1 Then RightColN = 12
176     If KeyCode = 68 And Shift = 1 Then RightColN = 13
178     If KeyCode = 69 And Shift = 1 Then RightColN = 14
180     If KeyCode = 70 And Shift = 1 Then RightColN = 15
182     If KeyCode = 71 And Shift = 1 Then RightColN = 15

184     Call Palett.UpdateColors

186     If KeyCode = "38" And Shift = 0 Then
188         KeyY = KeyY - 1
190         Call CalcPrevPicCords(KeyX, KeyY)
192         EventSource = "keyboard"
194         Call Shared_Mousemove(0, 0, 0, 0)
196         Call SetPixelFromKeyboard
        End If

198     If KeyCode = "40" And Shift = 0 Then
200         KeyY = KeyY + 1
202         Call CalcPrevPicCords(KeyX, KeyY)
204         EventSource = "keyboard"
206         Call Shared_Mousemove(0, 0, 0, 0)
208         Call SetPixelFromKeyboard
        End If

210     If KeyCode = "37" And Shift = 0 Then
212         KeyX = KeyX - ResoDiv
214         Call CalcPrevPicCords(KeyX, KeyY)
216         EventSource = "keyboard"
218         Call Shared_Mousemove(0, 0, 0, 0)
220         Call SetPixelFromKeyboard
        End If

222     If KeyCode = "39" And Shift = 0 Then
224         KeyX = KeyX + ResoDiv
226         Call CalcPrevPicCords(KeyX, KeyY)
228         EventSource = "keyboard"
230         Call Shared_Mousemove(0, 0, 0, 0)
232         Call SetPixelFromKeyboard
        End If

234     If Shift = 1 Then

236         If KeyCode = "38" Then
238             KeyY = KeyY - 8
240             ZoomWinTop = ZoomWinTop - 8
242             Call CalcPrevPicCords(KeyX, KeyY)
244             EventSource = "keyboard"
246             Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
                'Call Shared_Mousemove(0, 0, KeyX, KeyY)
            End If

248         If KeyCode = "40" Then
250             KeyY = KeyY + 8
252             ZoomWinTop = ZoomWinTop + 8
254             Call CalcPrevPicCords(KeyX, KeyY)
256             EventSource = "keyboard"
258             Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
                'Call Shared_Mousemove(0, 0, KeyX, KeyY)
            End If

260         If KeyCode = "37" Then
262             KeyX = KeyX - 8
264             ZoomWinLeft = ZoomWinLeft - 8
266             Call CalcPrevPicCords(KeyX, KeyY)
268             EventSource = "keyboard"
270             Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
                'Call Shared_Mousemove(0, 0, KeyX, KeyY)
            End If

272         If KeyCode = "39" Then
274             KeyX = KeyX + 8
276             ZoomWinLeft = ZoomWinLeft + 8
278             Call CalcPrevPicCords(KeyX, KeyY)
280             EventSource = "keyboard"
282             Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
                'Call Shared_Mousemove(0, 0, KeyX, KeyY)
            End If

        End If

284     If KeyCode = "8" Then        ' backspace: clear char with actcol

            'ColNum = LeftColN
            'changefrom = pixels(Ax, Ay)
            'find which color was cursor over of
            'If ScrRam(ScrBank, ScrNum, cx, cy, 0) = changefrom Then
            'ScrRam(ScrBank, ScrNum, cx, cy, 0) = LeftColN
            'ElseIf ScrRam(ScrBank, ScrNum, cx, cy, 1) = changefrom Then
            'ScrRam(ScrBank, ScrNum, cx, cy, 1) = LeftColN
            'ElseIf d800(cx, cy) = changefrom Then
            'd800(cx, cy) = LeftColN
            'End If

            'For x = cx * 8 To ((cx + 1) * 8) - resodiv Step 1
            ' For y = cy * 8 To ((cy + 1) * 8) - 1
            '   'Pixels(x, y) = LeftColN
            '   Call Setpixel(x, y, LeftColN)
            ' Next y
            'Next x

            'Call InitAttribs
        
286         If ActiveTool = "copy2" Then
                    
288             If xx1 <> xx2 Or yy1 <> yy2 Then
290                 byteColor = BackgrIndex
292                 For X = xx1 To xx2 - ResoDiv Step ResoDiv
294                     For Y = yy1 To yy2 - 1
296                         Pixels(X, Y) = BackgrIndex
298                         If ResoDiv = 2 Then Pixels(X + 1, Y) = BackgrIndex
300                         Call PutBitmap(X, Y, 0, byteColor)
302                     Next Y
304                 Next X
306                 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop): ZoomWinRefresh
308                 PrevWin.ReDraw
310                 xx1 = 0: xx2 = 0: yy1 = 0: yy2 = 0: ActiveTool = "copy1"
                End If
            
            Else
        

312             Call SetPixel(Ax, Ay, BackgrIndex)
314             Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
316             Call ZoomWinRefresh
318             Call PrevWin.DrawCursors(Ax, Ay)

320             Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
322             Call PrevWin.ReDraw
        
            End If
        End If


    End If

        '<EhFooter>
        Exit Sub

Shared_keydown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.Shared_keydown" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub








Private Sub StartFill()
        '<EhHeader>
        On Error GoTo StartFill_Err
        '</EhHeader>


100 Area = Pixels(Ax, Ay)
102 If Area <> FillCol Then
104     FillChangedPic = False
106     Call Fill(Ax, Ay)
108     Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
110     If FillMode = "compensating" Then
112         For X = 0 To PW - 1
114             For Y = 0 To PH - 1
116                 ColMap(X, Y) = Pixels(X, Y)
118             Next Y
120         Next X
122         Call Color_Restrict
124         Call InitAttribs
        End If
126     Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
128     Call ZoomWinRefresh
130     If FillChangedPic = True Then
            'MsgBox ("fillwritesundo")
132         WriteUndo
        End If

    End If


        '<EhFooter>
        Exit Sub

StartFill_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.StartFill" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Private Sub Fill(ByVal Tx As Long, ByVal Ty As Long)
    'Dim Tx As Integer
    'Dim Ty As Integer
        '<EhHeader>
        On Error GoTo Fill_Err
        '</EhHeader>
    Dim xmax As Long
    Dim minx As Long
    Dim fillx As Integer

100 fillx = Tx
    'Tx = FillX
    'Ty = Filly

102 If FillMode = "strict" Then



104     Do While Tx <= PW - ResoDiv

106         If DitherFill = True Then
108             If pattern(Tx / ResoDiv And 1, Ty And 1) = 1 Then FillCol = LeftColN
110             If pattern(Tx / ResoDiv And 1, Ty And 1) = 0 Then FillCol = RightColN
            End If

112         If (Pixels(Tx, Ty) = Area) Then
114             If PixelFits(Tx, Ty, FillCol) <> -1 Then
116                 FillChangedPic = True
118                 Call SetPixel(Tx, Ty, FillCol)
120                 Tx = Tx + ResoDiv
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop

122     xmax = Tx - ResoDiv

124     Tx = fillx - ResoDiv        'tx > 0 And
126     Do While Tx >= 0

128         If DitherFill = True Then
130             If pattern(Tx / ResoDiv And 1, Ty And 1) = 1 Then FillCol = LeftColN
132             If pattern(Tx / ResoDiv And 1, Ty And 1) = 0 Then FillCol = RightColN
            End If

134         If (Pixels(Tx, Ty) = Area) Then
136             If PixelFits(Tx, Ty, FillCol) <> -1 Then
138                 FillChangedPic = True
140                 Call SetPixel(Tx, Ty, FillCol)
142                 Tx = Tx - ResoDiv
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop

144     minx = Tx + ResoDiv

    End If

146 If FillMode = "compensating" Then

148     Tx = fillx
150     Do While Tx <= PW - ResoDiv

152         If DitherFill = True Then
154             If pattern(Tx / ResoDiv And 1, Ty And 1) = 1 Then FillCol = LeftColN
156             If pattern(Tx / ResoDiv And 1, Ty And 1) = 0 Then FillCol = RightColN
            End If

158         If (Pixels(Tx, Ty) = Area) Then
160             FillChangedPic = True
162             Pixels(Tx, Ty) = FillCol
164             If ResoDiv = 2 Then Pixels(Tx + 1, Ty) = FillCol
166             Tx = Tx + ResoDiv
                'Call PutBitmap(Tx, Ty, BitPairFits, FillCol)
            Else
                Exit Do
            End If
        Loop

168     xmax = Tx - ResoDiv

170     Tx = fillx - ResoDiv
172     Do While Tx >= 0

174         If DitherFill = True Then
176             If pattern(Tx / ResoDiv And 1, Ty And 1) = 1 Then FillCol = LeftColN
178             If pattern(Tx / ResoDiv And 1, Ty And 1) = 0 Then FillCol = RightColN
            End If

180         If (Pixels(Tx, Ty) = Area) Then
182             FillChangedPic = True
184             Pixels(Tx, Ty) = FillCol
186             If ResoDiv = 2 Then Pixels(Tx + 1, Ty) = FillCol
188             Tx = Tx - ResoDiv
            Else
                Exit Do
            End If
        Loop

190     minx = Tx + ResoDiv

    End If

    'Display.Refresh
192 Call PrevWin.ReDraw

194 If Ty - 1 >= 0 Then
        'tx = minx
        'Do While tx < xMax
196     For Tx = minx To xmax Step ResoDiv
198         If Pixels(Tx, Ty - 1) = Area Then Call Fill(Tx, Ty - 1)
200     Next Tx
        'tx = tx + 1
        'Loop
    End If

202 If Ty + 1 <= PH - 1 Then
        'tx = minx
        'Do While tx < xMax
204     For Tx = minx To xmax Step ResoDiv
206         If Pixels(Tx, Ty + 1) = Area Then Call Fill(Tx, Ty + 1)
208     Next Tx
        'tx = tx + 1
        'Loop
    End If

        '<EhFooter>
        Exit Sub

Fill_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.Fill" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Public Sub lacesetup()
        '<EhHeader>
        On Error GoTo lacesetup_Err
        '</EhHeader>

100 For X = 0 To PW - 1
102     For Y = 0 To PH - 1
104         ColMap(X, Y) = Pixels(X, Y)
106     Next Y
108 Next X


110 For X = 0 To PW - 1 Step 2
112     For Y = 0 To PH - 1
114         Pixels(Int(X / 2), Y) = ColMap(X, Y)
116     Next Y
118 Next X

120 StretchBlt LoadedPic.hdc, 0, 0, _
               PW, PH, _
               PixelsDib.hdc, _
               0, 0, _
               PW / 2, PH, _
               vbSrcCopy


122 For X = 1 To PW - 1 Step 2
124     For Y = 0 To PH - 1
126         Pixels(Int(X / 2), Y) = ColMap(X, Y)
128     Next Y
130 Next X

132 StretchBlt PrevPic.hdc, 0, 0, _
               PW, PH, _
               PixelsDib.hdc, _
               0, 0, _
               PW / 2, PH, _
               vbSrcCopy

134 For X = 0 To PW - 1
136     For Y = 0 To PH - 1
138         Pixels(X, Y) = ColMap(X, Y)
140     Next Y
142 Next X

        '<EhFooter>
        Exit Sub

lacesetup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.lacesetup" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub




Public Sub ZoomWinZoomIn()
        '<EhHeader>
        On Error GoTo ZoomWinZoomIn_Err
        '</EhHeader>
100 If ZoomScale <> ZoomScaleMax Then
102         ZoomScale = ZoomScale + 1
104         Call ZoomResize
106         Call SetMdiCaption
    End If
        '<EhFooter>
        Exit Sub

ZoomWinZoomIn_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomWinZoomIn" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub ZoomWinZoomOut()
        '<EhHeader>
        On Error GoTo ZoomWinZoomOut_Err
        '</EhHeader>
100 If ZoomScale <> ZoomScaleMin Then
102         ZoomScale = ZoomScale - 1
104         Call ZoomResize
106         Call SetMdiCaption
    End If
        '<EhFooter>
        Exit Sub

ZoomWinZoomOut_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.ZoomWinZoomOut" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub MouseHelper1_MouseWheel(Ctrl As Variant, Direction As MBMouseHelper.mbDirectionConstants, Button As Long, Shift As Long, Cancel As Boolean)
        '<EhHeader>
        On Error GoTo MouseHelper1_MouseWheel_Err
        '</EhHeader>

100 If MouseOver = "PrevPic" And Shift = 0 Then

102     Call PrevWin.MouseHelper1_MouseWheel(Ctrl, Direction, Button, Shift, Cancel)

    Else


        'change zoom
104     If Shift = 0 Then
106         If Direction < 0 Then
108             If ZoomScale <> 1 Then
110                 ZoomScale = ZoomScale - 1
112                 Call ZoomResize
                End If
            End If

114         If Direction > 0 Then
116             If ZoomScale <> 16 Then
118                 ZoomScale = ZoomScale + 1
120                 Call ZoomResize
                End If
            End If
122         Call SetMdiCaption
        End If

124     If ActiveTool <> "brush" Then

            'change lmb color
126         If Shift = 1 Then

128             If Direction < 0 Then
130                 LeftColN = (LeftColN + 1) And 15
                Else
132                 LeftColN = (LeftColN - 1) And 15
                End If

134             Call Palett.UpdateColors

            End If

            'change rmb color
136         If Shift = 2 Then

138             If Direction < 0 Then
140                 RightColN = (RightColN + 1) And 15
                Else
142                 RightColN = (RightColN - 1) And 15
                End If

144             Call Palett.UpdateColors

            End If

        Else

            'change brush size
146         If Shift = 1 Then
148             If Direction < 0 Then
150                 If BrushSize < 32 Then BrushSize = BrushSize + 1
                Else
152                 If BrushSize > 1 Then BrushSize = BrushSize - 1
                End If

154             Call ShowBrushCursor(LastZoomx, LastZoomy)
156             ZoomPic.Refresh
158             Call ZoomWinRefresh
160             Call PrevWin.DrawCursors(Ax, Ay)
162             Call BrushPreCalc
            End If

        End If
164     Cancel = True        'cancel the Window's default action

    End If

        '<EhFooter>
        Exit Sub

MouseHelper1_MouseWheel_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.MouseHelper1_MouseWheel" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub


Private Sub copy()
        '<EhHeader>
        On Error GoTo copy_Err
        '</EhHeader>

    Dim Rx As Integer
    Dim Ry As Integer
    Dim Scr As Long
    Dim Bank As Long
    Dim char As Long


    'buffer current screen

100 For X = 0 To PW - 1
102     For Y = 0 To PH - 1
104         ColMap(X, Y) = Pixels(X, Y)
106     Next Y
108 Next X

110 For X = 0 To CW - 1
112     For Y = 0 To CH - 1
114         For Scr = 0 To 7
116             For Bank = 0 To 1
118                 For char = 0 To 1
120                     BScrRam(Bank, Scr, X, Y, char) = ScrRam(Bank, Scr, X, Y, char)
122                 Next char
124             Next Bank
126         Next Scr
128     Next Y
130 Next X

132 For X = 0 To CW - 1
134     For Y = 0 To CH - 1
136         Bd800(X, Y) = D800(X, Y)
138     Next Y
140 Next X

142 For Bank = 0 To 1
144     For X = 0 To CW - 1
146         For Y = 0 To PH - 1
148             BBitmap(Bank, X, Y) = Bitmap(Bank, X, Y)
150         Next Y
152     Next X
154 Next Bank

    'copy from buffer to screen

156 Rx = -1 * (xx3 - Int(Ax / 8) * 8)
158 Ry = -1 * (yy3 - Int(Ay / 8) * 8)

160 For X = xx1 To xx2 - 1
162     For Y = yy1 To yy2 - 1

164         If X + Rx >= 0 And X + Rx < PW And Y + Ry >= 0 And Y + Ry < PH Then
166             Pixels(X + Rx, Y + Ry) = ColMap(X, Y)
            End If

168     Next Y
170 Next X

172 Rx = -1 * (xx3 - Int(Ax / 8) * 8)
174 Ry = -1 * (yy3 - Int(Ay / 8) * 8)

176 Rx = Int(Rx / 8)
178 Ry = Int(Ry / 8)

180 For Bank = 0 To 1
182     For Scr = 0 To 7
184         For char = 0 To 1
186             For X = Int(xx1 / 8) To Int(xx2 / 8) - 1
188                 For Y = Int(yy1 / 8) To Int(yy2 / 8) - 1

190                     If X + Rx >= 0 And X + Rx < CW And Y + Ry >= 0 And Y + Ry < CH Then
192                         ScrRam(Bank, Scr, X + Rx, Y + Ry, char) = BScrRam(Bank, Scr, X, Y, char)
                        End If

194                 Next Y
196             Next X
198         Next char
200     Next Scr
202 Next Bank

204 Rx = -1 * (xx3 - Int(Ax / 8) * 8)
206 Ry = -1 * (yy3 - Int(Ay / 8) * 8)

208 Rx = Int(Rx / 8)
210 Ry = Int(Ry / 8)

212 For X = Int(xx1 / 8) To Int(xx2 / 8) - 1
214     For Y = Int(yy1 / 8) To Int(yy2 / 8) - 1
216         If X + Rx >= 0 And X + Rx < CW And Y + Ry >= 0 And Y + Ry < CH Then
218             D800(X + Rx, Y + Ry) = Bd800(X, Y)
            End If
220     Next Y
222 Next X

224 Rx = -1 * (xx3 - Int(Ax / 8) * 8)
226 Ry = -1 * (yy3 - Int(Ay / 8) * 8)

228 Rx = Int(Rx / 8)
    'Ry = Int(Ry / 8) * 8 WHAT THE FUCK???????

230 For Bank = 0 To 1
232     For X = Int(xx1 / 8) To Int(xx2 / 8) - 1
234         For Y = yy1 To yy2 - 1
236             If X + Rx >= 0 And X + Rx < CW And Y + Ry >= 0 And Y + Ry < PH Then
238                 Bitmap(Bank, X + Rx, Y + Ry) = BBitmap(Bank, X, Y)
                End If
240         Next Y
242     Next X
244 Next Bank



        '<EhFooter>
        Exit Sub

copy_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.copy" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub

Public Sub SaveOwnFormat()
        '<EhHeader>
        On Error GoTo SaveOwnFormat_Err
        '</EhHeader>
        Dim filebuffer(65535) As Byte
        Dim FileName          As String
        Dim ff                As Long
        Dim Qx                As Long
        Dim Qy                As Long
        Dim offset            As Long
        Dim bmp               As Byte
        Dim Bank              As Long

100     Call ZoomWindow.StopInterlace

102     ff = FreeFile
104     If GetSaveName("(*.p1p)|*.p1p|(*.*)|*.*") = False Then Exit Sub

106     If FileExists(FileName) Then
108         Kill (FileName)
110         Open FileName For Binary As #ff
        Else
112         Open FileName For Binary As #ff
        End If

        'clear filebuffer
114     For offset = 0 To 65535
116         filebuffer(offset) = 0
118     Next offset

        'offset = 0
        'write screens
120     For Bank = 0 To 1
122         If Bank = 0 Then
124             offset = &H6000&
            Else
126             offset = &HE000&
            End If
128         For ScrNum = 0 To 7
130             For CY = 0 To CH - 1
132                 For CX = 0 To CW - 1
134                     filebuffer(offset) = (ScrRam(Bank, ScrNum, CX, CY, 1) And 15) + (ScrRam(Bank, ScrNum, CX, CY, 0) And 15) * 16
136                     offset = offset + 1
138                 Next CX
140             Next CY
142             offset = offset + 24
144         Next ScrNum
146     Next Bank

        'write bitmaps
148     For Bank = 0 To 1
150         If Bank = 0 Then offset = &H4000& Else offset = &HC000&
152         For CY = 0 To CH - 1
154             For CX = 0 To CW - 1
156                 Qx = CX * 8: Qy = CY * 8
158                 For Y = 0 To 7
160                     filebuffer(offset) = Bitmap(Bank, CX, Qy + Y)
162                     offset = offset + 1
164                 Next Y
166             Next CX
168         Next CY
170     Next Bank

        'write d800
172     offset = &H8000&

174     For CY = 0 To CH - 1
176         For CX = 0 To CW - 1
178             filebuffer(offset) = D800(CX, CY)
180             offset = offset + 1
182         Next CX
184     Next CY

        'start addy
186     filebuffer(&H3FFE&) = 0
188     filebuffer(&H3FFF&) = &H40&

        'background
190     filebuffer(&H8400&) = BackgrIndex

        'gfxformat info

        'base mode
192     If BaseMode = BaseModeTyp.multi Then filebuffer(&H9000&) = 0 Else filebuffer(&H9000&) = 1

        'number of bitmaps (banks)
194     filebuffer(&H9001&) = BmpBanks

        'diff screens / bank ?
196     filebuffer(&H9002&) = ScrBanks

        'fli density
198     filebuffer(&H9003&) = FliMul

        'pixel size
200     filebuffer(&H9004&) = ResoDiv

        'FLI write limit
202     filebuffer(&H9005&) = XFliLimit

        'write buffer
204     For offset = &H3FFE& To 65535
206         bmp = filebuffer(offset)
208         Put #ff, , bmp
210     Next offset

212     Close ff

        '<EhFooter>
        Exit Sub

SaveOwnFormat_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.SaveOwnFormat" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub
Public Sub LoadOwnFormat()
        '<EhHeader>
        On Error GoTo LoadOwnFormat_Err
        '</EhHeader>
    Dim CX As Long
    Dim CY As Long
    Dim Qx As Long
    Dim Qy As Long
    Dim Y As Long
    Dim filebuffer(66000) As Byte
    Dim FileName As String
    Dim ff As Long
    Dim Q As Long
    Dim offset As Long
    Dim Bank As Long

100 Call ZoomWindow.StopInterlace

102 ff = FreeFile
104 If GetLoadName("(*.p1p)|*.p1p|(*.*)|*.*") = False Then Exit Sub

106 Open FileName For Binary As #ff

    'read file to buffer

108 For Q = &H3FFE To 65535
110     Get #ff, , filebuffer(Q)
112 Next Q



    'read screens
114 For Bank = 0 To 1
116     If Bank = 0 Then offset = &H6000& Else offset = &HE000&
118     For ScrNum = 0 To 7
120         For CY = 0 To CH - 1
122             For CX = 0 To CW - 1
124                 ScrRam(Bank, ScrNum, CX, CY, 1) = filebuffer(offset) And 15
126                 ScrRam(Bank, ScrNum, CX, CY, 0) = Int(filebuffer(offset) / 16)
128                 offset = offset + 1
130             Next CX
132         Next CY
134         offset = offset + 24
136     Next ScrNum
138 Next Bank

    'read bitmaps
140 For Bank = 0 To 1
142     If Bank = 0 Then offset = &H4000& Else offset = &HC000&
144     For CY = 0 To CH - 1
146         For CX = 0 To CW - 1
148             Qx = CX * 8: Qy = CY * 8
150             For Y = 0 To 7
152                 Bitmap(Bank, CX, Qy + Y) = filebuffer(offset)
154                 offset = offset + 1
156             Next Y
158         Next CX
160     Next CY
162 Next Bank

    'read d800
164 offset = &H8000&

166 For CY = 0 To CH - 1
168     For CX = 0 To CW - 1
170         D800(CX, CY) = filebuffer(offset) And 15
172         offset = offset + 1
174     Next CX
176 Next CY


    'read gfxformat info


    'background
178 BackgrIndex = filebuffer(&H8400&) And 15

    'base mode
180 If filebuffer(&H9000&) = 0 Then BaseMode = BaseModeTyp.multi Else BaseMode = BaseModeTyp.hires

    'number of bitmaps (banks)
182 BmpBanks = filebuffer(&H9001&)

    'diff screens / bank ?
184 ScrBanks = filebuffer(&H9002&)

    'fli density
186 FliMul = filebuffer(&H9003&)

    'pixel size
188 ResoDiv = filebuffer(&H9004&)

    'FLI write limit
190 XFliLimit = filebuffer(&H9005&)

192 GfxMode = "custom"
194 Call UnCheckModes
196 MainWin.mnuGfxModeCustom.Checked = True

198 Close ff

200 Call DrawPicFromMem
202 Call PrevWin.DrawCursors(Ax, Ay)
204 Call ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
206 Call ZoomWinRefresh
208 Call Palett.UpdateColors
210 Call ModeChangeReset

        '<EhFooter>
        Exit Sub

LoadOwnFormat_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ZoomWindow.LoadOwnFormat" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Sub



