VERSION 5.00
Object = "{DDA53BD0-2CD0-11D4-8ED4-00E07D815373}#1.0#0"; "MBMouse.ocx"
Begin VB.Form PrevWin 
   BackColor       =   &H8000000C&
   Caption         =   "Form2"
   ClientHeight    =   3450
   ClientLeft      =   510
   ClientTop       =   225
   ClientWidth     =   5505
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   5505
   Visible         =   0   'False
   Begin VB.Frame label1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   4920
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   3120
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2895
      LargeChange     =   50
      Left            =   5160
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   720
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2895
   End
   Begin VB.PictureBox PrevPic 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2295
      ScaleWidth      =   4695
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
   Begin VB.PictureBox ShadowPic 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2160
      Top             =   1440
   End
   Begin MBMouseHelper.MouseHelper MouseHelper1 
      Left            =   4320
      Top             =   2640
      _ExtentX        =   900
      _ExtentY        =   900
   End
End
Attribute VB_Name = "PrevWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DontSizeWindow As Boolean
Private Dummy As Byte

'Stretchblt
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'get border width
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CXFRAME = 32

Private BorderWidth As Long
Private MenuHeight As Long


Private Sub form_keyup(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo form_keyup_Err
        '</EhHeader>
100 Call ZoomWindow.Shared_keyup(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

form_keyup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.form_keyup" + " line: " + Str(Erl))

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
100 Call ZoomWindow.Shared_keydown(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

form_keydown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.form_keydown" + " line: " + Str(Erl))

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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err
        '</EhHeader>

    'user pressed X on preview window -> hide it
100 If UnloadMode = 0 Then
102     Cancel = True
104     PrevWin.Visible = False
106     MainWin.mnuPreviewWindow.Checked = PrevWin.Visible
    End If

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.Form_QueryUnload" + " line: " + Str(Erl))

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

    Dim H As Long
    Dim W As Long

100 Set Me.Icon = ZoomWindow.Image1.Picture

102 Timer1.Enabled = False

104 BorderWidth = GetSystemMetrics(SM_CXFRAME)

106 KeyPreview = True
108 PrevWin.Caption = Str(PreviewZoom) + "% " + ZoomWindow.PicName        ' - Normal View"

    ' Set ScaleMode to pixels.
110 PrevWin.ScaleMode = vbPixels
112 PrevPic.ScaleMode = vbPixels
114 ShadowPic.ScaleMode = vbPixels

116 PrevPic.AutoSize = True
118 PrevPic.AutoRedraw = True
120 ShadowPic.AutoRedraw = True

122 PrevPic.BorderStyle = 0
124 ShadowPic.BorderStyle = 0

126 PrevPic.Move 0, 0

128 H = PrevWin.ScaleHeight
130 W = PrevWin.ScaleWidth

    ' Set the Max property for the scroll bars.
132 HScroll1.Max = W - VScroll1.Width
134 VScroll1.Max = H - HScroll1.Height

    'determine if scrollbars needs to be visible
136 If ((PrevWin.Height < PrevPic.Height) _
        Or (PrevWin.Width < PrevPic.Width)) = True _
        Then
138     VScroll1.Visible = True
140     HScroll1.Visible = True
    Else
142     VScroll1.Visible = False
144     HScroll1.Visible = False
    End If
146 PrevPic.Scale (0, 0)-(PW, PH)


        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.form_load" + " line: " + Str(Erl))

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

Private Sub VScroll1_Change()
        '<EhHeader>
        On Error GoTo VScroll1_Change_Err
        '</EhHeader>
100     PrevPic.Top = -VScroll1.Value
        '<EhFooter>
        Exit Sub

VScroll1_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.VScroll1_Change" + " line: " + Str(Erl))

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

Private Sub VScroll1_scroll()
        '<EhHeader>
        On Error GoTo VScroll1_scroll_Err
        '</EhHeader>
100     PrevPic.Top = -VScroll1.Value
        '<EhFooter>
        Exit Sub

VScroll1_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.VScroll1_scroll" + " line: " + Str(Erl))

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

Private Sub HScroll1_Change()
        '<EhHeader>
        On Error GoTo HScroll1_Change_Err
        '</EhHeader>
100     PrevPic.Left = -HScroll1.Value
        '<EhFooter>
        Exit Sub

HScroll1_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.HScroll1_Change" + " line: " + Str(Erl))

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
Private Sub HScroll1_scroll()
        '<EhHeader>
        On Error GoTo HScroll1_scroll_Err
        '</EhHeader>
100     PrevPic.Left = -HScroll1.Value
        '<EhFooter>
        Exit Sub

HScroll1_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.HScroll1_scroll" + " line: " + Str(Erl))

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





Public Sub ReDraw()
        '<EhHeader>
        On Error GoTo ReDraw_Err
        '</EhHeader>

100 StretchBlt PrevPic.hdc, 0, 0, PW * (PreviewZoomX / 100), PH * (PreviewZoomY / 100), PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy
102 PrevPic.Refresh

        '<EhFooter>
        Exit Sub

ReDraw_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.ReDraw" + " line: " + Str(Erl))

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

Public Sub PrevPic_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo PrevPic_mousemove_Err
        '</EhHeader>

100 MouseOver = "PrevPic"
102 ZoomWindow.EventSource = "PrevPic"
104 Call ZoomWindow.CalcPrevPicCords(Int(X), Int(Y))
106 Call ZoomWindow.Shared_Mousemove(Button, Shift, X, Y)

        '<EhFooter>
        Exit Sub

PrevPic_mousemove_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.PrevPic_mousemove" + " line: " + Str(Erl))

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

Private Sub PrevPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo PrevPic_MouseDown_Err
        '</EhHeader>

100 If Button = 2 Then
102     Me.MousePointer = vbCustom
104     Me.MouseIcon = LoadPicture(App.Path & "\Cursors\handgrab.ico")
    End If

    'OldAx = ZoomWindow.Ax
    'OldAy = ZoomWindow.Ay

106 ZoomWindow.EventSource = "PrevPic"
108 Call ZoomWindow.Shared_MouseDown(Button, Shift, X, Y)


        '<EhFooter>
        Exit Sub

PrevPic_MouseDown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.PrevPic_MouseDown" + " line: " + Str(Erl))

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


Private Sub PrevPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo PrevPic_MouseUp_Err
        '</EhHeader>

100 Me.MousePointer = vbDefault
102 ZoomWindow.EventSource = "PrevPic"
104 Call ZoomWindow.Shared_MouseUp(Button, Shift, X, Y)

        '<EhFooter>
        Exit Sub

PrevPic_MouseUp_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.PrevPic_MouseUp" + " line: " + Str(Erl))

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

Private Sub Timer1_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>


If Dummy Then

    StretchBlt PrevPic.hdc, 0, 0, _
               PW * (PreviewZoomX / 100), PH * (PreviewZoomY / 100), _
               ZoomWindow.LoadedPic.hdc, _
               1, 0, _
               PW, PH, _
               vbSrcCopy

    PrevPic.Refresh

Else

    StretchBlt PrevPic.hdc, 0, 0, _
               PW * (PreviewZoomX / 100), PH * (PreviewZoomY / 100), _
               ZoomWindow.PrevPic.hdc, _
               0, 0, _
               PW, PH, _
               vbSrcCopy

    PrevPic.Refresh

End If

Dummy = Not Dummy

End Sub

Public Sub DrawCursors(X As Long, Y As Long)
        '<EhHeader>
        On Error GoTo DrawCursors_Err
        '</EhHeader>
    Dim Rx As Long
    Dim Ry As Long
    Dim x1 As Long
    Dim x2 As Long
    Dim y1 As Long
    Dim y2 As Long

    'dont redraw cursor or zoomkeret if we're emulating lace
100 If MainWin.mnuInterlaceEmu.Checked = False Then

102     StretchBlt PrevPic.hdc, 0, 0, PW * (PreviewZoomX / 100), PH * (PreviewZoomY / 100), PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy

        'zoom keret megrajzolás
104     If (ZoomPicCenteredX = False Or ZoomPicCenteredY = False) And ZoomKeret = 1 Then

106         PrevPic.DrawWidth = 1
108         PrevPic.DrawMode = 7        '7=eor,13=normal
110         PrevPic.Forecolor = RGB(255, 255, 255)
        
112         If ZoomPicCenteredX = True Then
114             PrevPic.Line (0, ZoomWinTop - ZoomHeight)-(ZoomWidth, ZoomWinTop), , B
116         ElseIf ZoomPicCenteredY = True Then
118             PrevPic.Line (ZoomWinLeft - ZoomWidth, 0)-(ZoomWinLeft + 0, ZoomHeight), , B
            Else
120             PrevPic.Line (ZoomWinLeft - ZoomWidth, ZoomWinTop - ZoomHeight)-(ZoomWinLeft, ZoomWinTop), , B
            End If
        
        End If

        'highlight pixel under mouse

122     If (ZoomWindow.ActiveTool = "draw" Or ZoomWindow.ActiveTool = "fill") And PixelBox = 1 Then

124         If PixelBoxColor = 1 Then
126             PrevPic.Forecolor = PaletteRGB(LeftColN)
128             PrevPic.DrawMode = 13        '7=eor,13=normal
            Else
130             PrevPic.Forecolor = RGB(255, 255, 255)
132             PrevPic.DrawMode = 7        '7=eor,13=normal
            End If

134         PrevPic.Line (X, Y)-((X + ResoDiv), Y + 1), , BF

        'draw brush outlines
136     ElseIf ZoomWindow.ActiveTool = "brush" Then
        
138         PrevPic.DrawStyle = 0
140         PrevPic.Forecolor = RGB(255, 255, 255)
142         PrevPic.DrawMode = 7        '7=eor,13=normal
144         PrevPic.Circle (X, Y), BrushSize / 2
146         PrevPic.DrawStyle = 0
148         PrevPic.DrawMode = 13

150     ElseIf ZoomWindow.ActiveTool = "copy1" Or ZoomWindow.ActiveTool = "copy2" Then

152         Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
154         PrevPic.DrawStyle = 2
156         PrevPic.DrawMode = 7        '7=eor,13=normal
158         PrevPic.Forecolor = RGB(255, 255, 255)
160         PrevPic.Line (ZoomWindow.xx1, ZoomWindow.yy1)-(ZoomWindow.xx2, ZoomWindow.yy2), RGB(255, 255, 255), B
162         PrevPic.DrawStyle = 0

164         ZoomWindow.ZoomPic.DrawWidth = 2
166         ZoomWindow.ZoomPic.DrawStyle = 1
168         ZoomWindow.ZoomPic.DrawMode = 13        '7=eor,13=normal
170         ZoomWindow.ZoomPic.Forecolor = RGB(0, 255, 0)
172         Rx = ZoomWindow.xx1
174         Ry = ZoomWindow.yy1
176         Call GetZoomCoords(Rx, Ry)
178         x1 = Rx
180         y1 = Ry
182         Rx = ZoomWindow.xx2
184         Ry = ZoomWindow.yy2
186         Call GetZoomCoords(Rx, Ry)
188         x2 = Rx
190         y2 = Ry
192         ZoomWindow.ZoomPic.Line (x1, y1)-(x2, y2), , B
194         ZoomWindow.ZoomPic.Refresh
196         ZoomWindow.ZoomPic.DrawStyle = 0
198         ZoomWindow.ZoomPic.DrawMode = 13        '7=eor,13=normal
200         ZoomWindow.ZoomPic.DrawWidth = 1

202     ElseIf ZoomWindow.ActiveTool = "copy3" Then
        
204         Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
206         Rx = -1 * (ZoomWindow.xx3 - Int(Ax / 8) * 8)
208         Ry = -1 * (ZoomWindow.yy3 - Int(Ay / 8) * 8)

210         PrevPic.DrawStyle = 2
212         PrevPic.DrawMode = 7        '7=eor,13=normal
214         PrevPic.Forecolor = RGB(255, 255, 255)
216         PrevPic.Line (ZoomWindow.xx1 + Rx - 1, ZoomWindow.yy1 + Ry - 1)-(ZoomWindow.xx2 + Rx - 1, ZoomWindow.yy2 + Ry - 1), RGB(255, 255, 255), B
218         PrevPic.DrawStyle = 0

220         ZoomWindow.ZoomPic.DrawStyle = 2
222         ZoomWindow.ZoomPic.DrawWidth = 2
224         ZoomWindow.ZoomPic.DrawMode = 13        '7=eor,13=normal
226         ZoomWindow.ZoomPic.Forecolor = RGB(0, 255, 0)
        
228         Rx = -1 * (ZoomWindow.xx3 - Int(Ax / 8) * 8)
230         Ry = -1 * (ZoomWindow.yy3 - Int(Ay / 8) * 8)
232         Rx = ZoomWindow.xx1 + Rx - 1
234         Ry = ZoomWindow.yy1 + Ry - 1
236         Call GetZoomCoords(Rx, Ry)
238         x1 = Rx
240         y1 = Ry
        
242         Rx = -1 * (ZoomWindow.xx3 - Int(Ax / 8) * 8)
244         Ry = -1 * (ZoomWindow.yy3 - Int(Ay / 8) * 8)
246         Rx = ZoomWindow.xx2 + Rx - 1
248         Ry = ZoomWindow.yy2 + Ry - 1
250         Call GetZoomCoords(Rx, Ry)
252         x2 = Rx
254         y2 = Ry
        
256         ZoomWindow.ZoomPic.Line (x1, y1)-(x2, y2), , B
258         ZoomWindow.ZoomPic.Refresh
260         ZoomWindow.ZoomPic.DrawStyle = 0
262         ZoomWindow.ZoomPic.DrawMode = 13        '7=eor,13=normal
264         ZoomWindow.ZoomPic.DrawWidth = 1

266     ElseIf ZoomWindow.ActiveTool = "copy" Then
        
268         Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
        
270         Rx = Int((Ax / 8) + 0.5) * 8
272         Ry = Int((Ay / 8) + 0.5) * 8
274         PrevPic.DrawMode = 7        '7=eor,13=normal
276         PrevPic.DrawStyle = 2
278         PrevPic.Forecolor = RGB(255, 255, 255)
280         PrevPic.Line (Rx, 0)-(Rx, PH - 1), RGB(255, 255, 255)
282         PrevPic.Line (0, Ry)-(PW - 1, Ry), RGB(255, 255, 255)
284         PrevPic.DrawStyle = 0

286         ZoomWindow.ZoomPic.DrawWidth = 2
288         ZoomWindow.ZoomPic.DrawStyle = 1
290         ZoomWindow.ZoomPic.DrawMode = 13        '7=eor,13=normal
292         ZoomWindow.ZoomPic.Forecolor = RGB(0, 255, 0)
        
294         Call GetZoomCoords(Rx, Ry)
296         x1 = Rx
298         y1 = ZoomWindow.Zy
300         x2 = Rx
302         y2 = ((PH - 1) * ZoomScaleY) + ZoomWindow.Zy
304         ZoomWindow.ZoomPic.Line (x1, y1)-(x2, y2)
306         x1 = ZoomWindow.Zx
308         y1 = Ry
310         x2 = ((PW - 1) * ZoomScaleX) + ZoomWindow.Zx
312         y2 = Ry
314         ZoomWindow.ZoomPic.Line (x1, y1)-(x2, y2)
316         ZoomWindow.ZoomPic.DrawStyle = 0
318         ZoomWindow.ZoomPic.Refresh
320         ZoomWindow.ZoomPic.DrawMode = 13        '7=eor,13=normal
322         ZoomWindow.ZoomPic.DrawWidth = 1

        End If

324     PrevPic.Refresh

    End If

        '<EhFooter>
        Exit Sub

DrawCursors_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.DrawCursors" + " line: " + Str(Erl))

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

Private Sub GetZoomCoords(ByRef X As Long, ByRef Y As Long)
        '<EhHeader>
        On Error GoTo GetZoomCoords_Err
        '</EhHeader>

100     If ZoomPicCenteredX = False Then
102         X = (X - ZoomWinLeft + ZoomWidth) * ZoomScaleX + ZoomWindow.Zx
        Else
104         X = (X) * ZoomScaleX + ZoomWindow.Zx + 0
        End If

106     If ZoomPicCenteredY = False Then
108         Y = (Y - ZoomWinTop + ZoomHeight) * ZoomScaleY + ZoomWindow.Zy
        Else
110         Y = (Y) * ZoomScaleY + ZoomWindow.Zy + 0
        End If
    
        '<EhFooter>
        Exit Sub

GetZoomCoords_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.GetZoomCoords" + " line: " + Str(Erl))

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
Public Sub ZoomInPrev()
        '<EhHeader>
        On Error GoTo ZoomInPrev_Err
        '</EhHeader>
100 If PreviewZoom <> PreviewZoomMax Then
102     PreviewZoom = PreviewZoom + 25
    End If
104 Call ResizePrevPic
106 Call ChangeWinSize
        '<EhFooter>
        Exit Sub

ZoomInPrev_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.ZoomInPrev" + " line: " + Str(Erl))

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
Public Sub ZoomOutPrev()
        '<EhHeader>
        On Error GoTo ZoomOutPrev_Err
        '</EhHeader>
100 If PreviewZoom <> PreviewZoomMin Then
102     PreviewZoom = PreviewZoom - 25
    End If
104 Call ResizePrevPic
106 Call ChangeWinSize
        '<EhFooter>
        Exit Sub

ZoomOutPrev_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.ZoomOutPrev" + " line: " + Str(Erl))

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

100     If MouseOver = "zoompic" Then

102         Call ZoomWindow.MouseHelper1_MouseWheel(Ctrl, Direction, Button, Shift, Cancel)

104     ElseIf Shift <> 1 Then

106         If Shift = 2 Then
108             DontSizeWindow = True
            Else
110             DontSizeWindow = False
            End If

            'this one is needed to prevent window ResizePrevPic when jailbird maximizes the preview window :)
112         If Me.WindowState = vbMaximized Then DontSizeWindow = True

114         If Direction < 0 Then
116             If PreviewZoom <> 400 Then
118                 PreviewZoom = PreviewZoom + 25
                End If
            Else
120             If PreviewZoom <> 25 Then
122                 PreviewZoom = PreviewZoom - 25
                End If
            End If

124         Call ResizePrevPic
126         Call ChangeWinSize
128         Call Form_Resize

130         Cancel = True        'cancel the Window's default action

        End If

        '<EhFooter>
        Exit Sub

MouseHelper1_MouseWheel_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.MouseHelper1_MouseWheel" + " line: " + Str(Erl))

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





Public Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
Dim H As Long
Dim W As Long

'safety margins in twips
If PrevWin.Height > 800 And PrevWin.Width > 800 Then

    'inner dimensions of window
    H = PrevWin.ScaleHeight
    W = PrevWin.ScaleWidth
    
    'place and size of sliders
    HScroll1.Move 0, H - HScroll1.Height, W - VScroll1.Width
    VScroll1.Move W - VScroll1.Width, 0, VScroll1.Width, HScroll1.Top
    
    'enable/disable scrollbars when needed
    VScroll1.Enabled = (H < PrevPic.Height)
    HScroll1.Enabled = (W < PrevPic.Width)
    
    'scrollbar visibility
    If (VScroll1.Enabled Or HScroll1.Enabled) Then
        VScroll1.Visible = True
        HScroll1.Visible = True
    Else
        VScroll1.Visible = False
        HScroll1.Visible = False
    End If
    
    'scrollbar max value
    If HScroll1.Enabled Then HScroll1.Max = PrevPic.Width - W + VScroll1.Width
    If VScroll1.Enabled Then VScroll1.Max = PrevPic.Height - H + HScroll1.Height

    'auto center picture on horizontally
    If (H - HScroll1.Height >= PrevPic.Height) Then
        PrevPic.Top = -(PrevPic.Height - H + HScroll1.Height) / 2
        ShadowPic.Top = PrevPic.Top + 5
    Else
        PrevPic.Top = 0
    End If
    
    'auto center picture on vertically
    If (W - VScroll1.Width >= PrevPic.Width) Then
        PrevPic.Left = -(PrevPic.Width - W + VScroll1.Width) / 2
        ShadowPic.Left = PrevPic.Left + 5
    Else
        PrevPic.Left = 0
    End If
    
    'hide hole on bottom right corner
    If VScroll1.Visible Then
        Label1.Visible = True
        Label1.Move VScroll1.Left, HScroll1.Top
    Else
        Label1.Visible = False
    End If

    PrevPic.Scale (0, 0)-(PW, PH)
  
End If

'hide mainwin scrollbars if prevwin is maxxed
If Me.WindowState = vbMaximized Then
    MainWin.Picture1.Visible = False
    MainWin.Picture2.Visible = False
Else
    MainWin.Picture1.Visible = True
    MainWin.Picture2.Visible = True
End If
    
End Sub

Public Sub ChangeWinSize()
        '<EhHeader>
        On Error GoTo ChangeWinSize_Err
        '</EhHeader>


100 If DontSizeWindow = False Then

102     MenuHeight = ScaleY(PrevWin.Height, vbTwips, vbPixels) - PrevWin.ScaleHeight

104     PrevWin.Move PrevWin.Left, PrevWin.Top, ScaleX(PW * (PreviewZoomX / 100), vbPixels, vbTwips) + ScaleX(BorderWidth * 2, vbPixels, vbTwips), ScaleY(PH * (PreviewZoomY / 100), vbPixels, vbTwips) + ScaleY(MenuHeight, vbPixels, vbTwips)
        'move palette window under preview window
106     Palett.Move PrevWin.Left, PrevWin.Top + PrevWin.Height
    
    End If

        '<EhFooter>
        Exit Sub

ChangeWinSize_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.ChangeWinSize" + " line: " + Str(Erl))

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

Public Sub ResizePrevPic()
        '<EhHeader>
        On Error GoTo ResizePrevPic_Err
        '</EhHeader>
        Dim W As Long
        Dim H As Long

100     PreviewZoomX = PreviewZoom * ARatioX
102     PreviewZoomY = PreviewZoom * ARatioY

104     PrevWin.Caption = Str(PreviewZoom) + "% " + ZoomWindow.PicName

106     With PrevPic
108         W = PW * (PreviewZoomX / 100)
110         H = PH * (PreviewZoomY / 100)
112         .Width = W
114         .Height = H
116         .ScaleMode = vbPixels
118         PrevPic.Scale (0, 0)-(W, H)
120         .Width = W
122         .Height = H
        End With

124     StretchBlt PrevPic.hdc, 0, 0, W, H, PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy

126     With ShadowPic
128         .Width = PrevPic.Width
130         .Height = PrevPic.Height
132         .Forecolor = RGB(60, 60, 60)
134         ShadowPic.Line (0, 0)-(.Width, .Height), , BF
136         ShadowPic.Line (0, 0)-(.Width - 1, .Height - 1), 0, B
138         .Refresh
        End With
    

        '<EhFooter>
        Exit Sub

ResizePrevPic_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PrevWin.ResizePrevPic" + " line: " + Str(Erl))

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
