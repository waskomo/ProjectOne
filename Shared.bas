Attribute VB_Name = "Shared"
Option Explicit

Public Enum BaseModeTyp
    hires = 0
    multi = 1
    unrestricted = 2
End Enum

Public Enum BrushTyp
    Round = 0
    Square = 1
End Enum

Public Enum Chip
    vicii = 0
    ted = 1
    vdc = 2
End Enum


'picture information
Public PW As Long     'size in pixels
Public PH As Long

Public CW As Long     'size in chars
Public CH As Long

Public PixelsDib As New cDIBSection256 'cDibSection holding our picture data
Public Pixels() As Byte
Public D800() As Byte
Public ScrRam() As Byte
Public Bitmap() As Byte
Public Registers() As Byte
Public BackgrIndex As Byte


'global (shared) variables

Public Ax As Long, Ay As Long

Public MouseOver As String

Public BitSwap(3, 3, 255) As Byte

Public ZoomHeight As Long 'in number of zoomed pixels
Public ZoomWidth As Long  'in number of zoomed pixels

Public DitherMode As Boolean ' used to draw in dither when alt is pressed
Public pattern(1, 1) As Long 'dither pattern
Public LeftColN As Byte     'user selected color for mouseleft
Public RightColN As Byte    'user selected color for mouseright

'palette stuff
Public PaletteRGB(255) As Long
Public coR() As Single
Public coG() As Single
Public coB() As Single

Public cd() As Long 'Color distance table
Public OldAx As Long
Public OldAy As Long

Public LastLoadPath As String
Public LastSavePath As String
Public LastLoadName As String
Public LastSaveName As String
Public LastLoadFilterIndex As Long
Public LastSaveFilterindex As Long



Public Type MemorySetup
    
    SkipStartAdress As Boolean
    HasStartAddress As Boolean
    
    StartAdressUser As Long
    ForceStartAddressFromUser As Boolean
    
    Start As Long
    End As Long
    Screen(15) As Long
    Bitmap(1) As Long
    D800 As Long
    D021 As Long
    Absolute As Boolean
End Type

Public Type SprTyp
    Xcoord As Long
    Color As Long
    MultiColor As Boolean
    Xstretch As Boolean
    Ystretch As Boolean
    Data() As Byte
End Type

Public adr_LoadorSave As String

Public CustomIOSetup As MemorySetup

'____________User Settings____________________________________________________________

Public Const PreviewZoomMax As Double = 400
Public Const PreviewZoomMin As Double = 25

Public PreviewZoom As Double     'zoom rate of prev window
Public PreviewZoomX As Double    'zoom rate of prev window
Public PreviewZoomY As Double    'zoom rate of prev window
Public ZoomKeret As Long         'wether zoom outline area to be shown in prev win

'drawing settings
Public EmulateFliBug As Boolean

Public FillMode As String
Public DitherFill As Boolean
Public BrushType As BrushTyp
Public BrushColor1 As Long
Public BrushColor2 As Long
Public BrushSize As Long
Public BrushDither As Long

'Converters settings

Public LastVisitedTab As Long

Public Gradient_Selected(2) As Long
Public BrLadderMax As Long
Public BrLadderTab(31) As Long

Public ColorTable_Filename(2) As String

Public Contrast As Single
Public Brightness As Long
Public Hue As Long
Public Saturation As Long

Public StretchPic As Boolean
Public KeepAspect As Boolean
Public ResizeScale As Long    'resize ratio

Public Reso_PixLeft As Boolean, Reso_PixRight As Boolean, Reso_PixAvg As Boolean
Public OptBackgr As Boolean, MostFreqBackgr As Boolean, UsrDefBackgr As Boolean
Public UsrBackgr As Long        'user selected backgr when converting
Public Mostfrq3 As Boolean, OptChars As Boolean
Public FixWithBackgr As Long

Public ColorFilterMode As Long  'Settings for the color filter
Public Bdither As Long
Public Hdither As Long
Public Sdither As Long
Public BditherVal As Long
Public Hditherval As Long
Public SditherVal As Long
Public BRnd As Long
Public HRnd As Long
Public SRnd As Long

'Zoom window settings
Public ZoomPicCenteredX As Boolean
Public ZoomPicCenteredY As Boolean
Public ZoomWinLeft As Long
Public ZoomWinTop As Long

Public Const ZoomScaleMax As Long = 16
Public Const ZoomScaleMin As Long = 1

Public ZoomScale As Long

Public ZoomScaleX As Single
Public ZoomScaleY As Single

Public Const ARatioMin As Single = 1
Public Const ARatioMax As Single = 4

Public ARatioX As Single
Public ARatioY As Single
Public ResizeScaleX As Single
Public ResizeScaleY As Single

'Grid Settings
Public PixelGrid As Long
Public CharGrid As Long
Public PixelBox As Long
Public ShowFlibox As Long
Public PixelBoxColor As Long
Public ShowFliLines As Long

Public PixelGridColor As Long
Public CharGridColor As Long
Public FliGridColor As Long

Public FliGridLimit As Long
Public PixelGridLimit As Long
Public CharGridLimit As Long

'Screenmode Definition

Public GfxMode As String




Public ChipType As Chip
Public BaseMode As BaseModeTyp
Public ResoDiv As Long
Public FliMul As Long
Public XFliLimit As Long
Public ScrBanks As Long
Public BmpBanks As Long
Public CellWidth As Long
Public CellHeigth As Long

Public ScrBank As Long
Public BmpBank As Long

Public ScrNum As Long


'custom screens settings
Public BaseMode_cm As BaseModeTyp
Public ResoDiv_cm As Long
Public FliMul_cm As Long
Public XFliLimit_cm As Long
Public ScrBanks_cm As Long
Public BmpBanks_cm As Long

'Brush
Public BrushArray(31, 31) As Long
Public Bayer(3, 3) As Long

'Handling of different palettes
Public m_cIni As New cInifile
Public pal_Count As Long
Public pal_Selected(2) As Long
Public pal_HexMode As Boolean
Public pal_Lastindex As Long




Public Sub LoadSettings()
        '<EhHeader>
        On Error GoTo LoadSettings_Err
        '</EhHeader>
    Dim X As Long

100 pal_HexMode = True
102 ChipType = Chip.vicii

104 With m_cIni
    
        'Brightness Ladder Load settings
106     .Path = App.Path & "\Ladders.ini"
108     .Section = "main"
110     .Key = "BrLadderSelected_VICII": .Default = 1: Gradient_Selected(Chip.vicii) = .Value
112     .Key = "BrLadderSelected_TED": .Default = 1: Gradient_Selected(Chip.ted) = .Value
114     .Key = "BrLadderSelected_VDC": .Default = 1: Gradient_Selected(Chip.vdc) = .Value
116     Call ConvertDialog.BrLadderLoadPreset(Gradient_Selected(ChipType))

        'Get Num of Palettes, and Palette Nr to be used, and Palette order
118     .Path = App.Path & "\Palettes.ini"
120     .Section = "Palettes"
122     .Key = "Count": .Default = 0: pal_Count = .Value
124     .Key = "Selected_VICII": .Default = 1: pal_Selected(Chip.vicii) = .Value
126     .Key = "Selected_TED": .Default = 1: pal_Selected(Chip.ted) = .Value
128     .Key = "Selected_VDC": .Default = 1: pal_Selected(Chip.vdc) = .Value
    
        'Over to our Main Ini
130     .Path = App.Path & "\ProjectOne.ini"
132     .Section = "General Settings"
134     .Key = "previewzoom": .Default = 100: PreviewZoom = .Value
        'Safety Barrier
136     If PreviewZoom > PreviewZoomMax Then
138         PreviewZoom = PreviewZoomMax
140         MsgBox "incorrect PreviewZoom value in ProjectOne.ini using " & Str(PreviewZoomMax) & " instead"
        End If
    
142     If PreviewZoom < PreviewZoomMin Then
144         PreviewZoom = PreviewZoomMin
146         MsgBox "incorrect PreviewZoom value in ProjectOne.ini using " & Str(PreviewZoomMin) & " instead"
        End If
    
148     .Key = "lastloadpath": .Default = "": LastLoadPath = .Value
150     .Key = "lastsavepath": .Default = "": LastSavePath = .Value
152     .Key = "lastloadname": .Default = "": LastLoadName = .Value
154     .Key = "lastsavename": .Default = "": LastSaveName = .Value
156     .Key = "lastloadfilterindex": .Default = 1: LastLoadFilterIndex = .Value
158     .Key = "lastsavefilterindex": .Default = 1: LastSaveFilterindex = .Value
    
        'gfx mode setup
160     .Key = "BaseMode": .Default = BaseModeTyp.multi: BaseMode = .Value
162     .Key = "GfxMode": .Default = "koala": GfxMode = .Value
        .Key = "resodiv": .Default = 2: ResoDiv = .Value:
164     If ResoDiv <> 1 Xor ResoDiv <> 2 Then
166         ResoDiv = 2
168         MsgBox "invalid ResoDiv value in ProjectOne.ini using 2 instead"
        End If
170     .Key = "flimul": .Default = 8: FliMul = .Value
172     If FliMul <> 1 Xor FliMul <> 2 Xor FliMul <> 4 Xor FliMul <> 8 Then
174         FliMul = 8
176         MsgBox "invalid FliMul value in ProjectOne.ini using 8 instead"
        End If
178     .Key = "xflilimit": .Default = 0: XFliLimit = .Value
180     .Key = "bmpbanks": .Default = 0: BmpBanks = .Value
182     .Key = "scrbanks": .Default = 0: ScrBanks = .Value
184     CellWidth = 8
186     CellHeigth = 8
    
        'custom mode set up to koala as default
188     .Key = "basemode_cm": .Default = BaseModeTyp.multi: BaseMode_cm = .Value
190     .Key = "resodiv_cm": .Default = 2: ResoDiv_cm = .Value
192     .Key = "flimul_cm": .Default = 8: FliMul_cm = .Value
194     .Key = "xflilimit_cm": .Default = 0: XFliLimit_cm = .Value
196     .Key = "bmpbanks_cm": .Default = 0: BmpBanks_cm = 0
198     .Key = "scrbanks_cm": .Default = 0: ScrBanks_cm = .Value
    
        'PW = 320
        'PH = 200
200     .Key = "picture_width": .Default = 320: PW = .Value
202     .Key = "picture_heigth": .Default = 200: PH = .Value
204     If PW <= 0 Or ((PW / 8) * 8 <> (Int(PW / 8) * 8)) Then
206         PW = 320
208         MsgBox "invalid picture width value in ProjectOne.ini using 320 instead"
        End If
210     If PH <= 0 Or ((PH / 8) * 8 <> (Int(PH / 8) * 8)) Then
212         PH = 200
214         MsgBox "invalid picture height value in ProjectOne.ini using 200 instead"
        End If
     
216      CW = PW / 8 '320
218      CH = PH / 8 '200
    
220     .Section = "View Settings"
        'view settings
222     .Key = "pixelgridcolor": .Default = RGB(50, 50, 50): PixelGridColor = .Value
224     .Key = "fligridcolor": .Default = RGB(40, 40, 232): FliGridColor = .Value
226     .Key = "chargridcolor": .Default = RGB(180, 180, 180): CharGridColor = .Value
228     .Key = "pixelgrid": .Default = 1: PixelGrid = .Value
230     .Key = "chargrid": .Default = 1: CharGrid = .Value
232     .Key = "pixelbox": .Default = 1: PixelBox = .Value
234     .Key = "showflibox": .Default = 0: ShowFlibox = .Value
236     .Key = "pixelboxcolor": .Default = 1: PixelBoxColor = .Value        'boolean, 1 = colored pixbox, 0= white eor box
238     .Key = "zoomkeret": .Default = 1: ZoomKeret = .Value
240     .Key = "showflilines": .Default = 1: ShowFliLines = .Value
242     .Key = "pixelgridlimit": .Default = 2: PixelGridLimit = .Value
244     .Key = "fligridlimit": .Default = 2: FliGridLimit = .Value
246     .Key = "chargridlimit": .Default = 2: CharGridLimit = .Value
248     .Key = "zoomwinleft": .Default = 160: ZoomWinLeft = .Value
250     .Key = "zoomwintop": .Default = 100: ZoomWinTop = .Value
252     .Key = "zoomscale": .Default = 8: ZoomScale = .Value
254     If ZoomScale < ZoomScaleMin Then
256         ZoomScale = ZoomScaleMin
258         MsgBox "incorrect ZoomScale value in ProjectOne.ini using " & Str(ZoomScaleMin) & " instead"
        End If
260     If ZoomScale > ZoomScaleMax Then
262         ZoomScale = ZoomScaleMax
264         MsgBox "incorrect ZoomScale value in ProjectOne.ini using " & Str(ZoomScaleMax) & " instead"
        End If
266     .Key = "aratiox": .Default = 1: ARatioX = .Value
268     If ARatioX < ARatioMin Then
270         ARatioX = ARatioMin
272         MsgBox "incorrect ARatioX value in ProjectOne.ini using " & Str(ARatioMin) & " instead"
        End If
274     If ARatioX > ARatioMax Then
276         ARatioX = ARatioMax
278         MsgBox "incorrect ARatioX value in ProjectOne.ini using " & Str(ARatioMax) & " instead"
        End If
280     .Key = "aratioy": .Default = 1: ARatioY = .Value
282     If ARatioY < ARatioMin Then
284         ARatioY = ARatioMin
286         MsgBox "incorrect ARatioY value in ProjectOne.ini using " & Str(ARatioMin) & " instead"
        End If
288     If ARatioY > ARatioMax Then
290         ARatioY = ARatioMax
292         MsgBox "incorrect ARatioy value in ProjectOne.ini using " & Str(ARatioMax) & " instead"
        End If
    
294     ZoomScaleX = ZoomScale * ARatioX
296     ZoomScaleY = ZoomScale * ARatioY
    
298     .Section = "Converter Settings"
        'Converter Settings:
300     .Key = "lastvisitedtab": .Default = 0: LastVisitedTab = .Value
302     .Key = "brightness": .Default = 0: Brightness = .Value
304     .Key = "contrast": .Default = 128: Contrast = .Value
306     .Key = "saturation": .Default = 128: Saturation = .Value
308     .Key = "hue": .Default = 0: Hue = .Value
310     .Key = "stretchpic": .Default = True: StretchPic = .Value
312     .Key = "keepaspect": .Default = True: KeepAspect = .Value
314     .Key = "reso_pixavg": .Default = True: Reso_PixAvg = .Value
316     .Key = "reso_pixleft": .Default = False: Reso_PixLeft = .Value
318     .Key = "reso_pixright": .Default = False: Reso_PixRight = .Value
320     .Key = "mostfreqbackgr": .Default = True: MostFreqBackgr = .Value
322     .Key = "usrdefbackgr": .Default = False: UsrDefBackgr = .Value
324     .Key = "usrbackgr": .Default = 0: UsrBackgr = .Value
326     .Key = "optchars": .Default = True: OptChars = .Value
328     .Key = "mostfrq3": .Default = False: Mostfrq3 = .Value
330     .Key = "fixwithbackgr": .Default = 1: FixWithBackgr = .Value
332     .Key = "colorfiltermode": .Default = 0: ColorFilterMode = .Value
334     .Key = "bdither": .Default = 2: Bdither = .Value
336     .Key = "hdither": .Default = 2: Hdither = .Value
338     .Key = "sdither": .Default = 2: Sdither = .Value
340     .Key = "bditherval": .Default = 0: BditherVal = .Value
342     .Key = "hditherval": .Default = 0: Hditherval = .Value
344     .Key = "sditherval": .Default = 0: SditherVal = .Value
346     .Key = "brnd": .Default = 1: BRnd = .Value
348     .Key = "hrnd": .Default = 1: HRnd = .Value
350     .Key = "srnd": .Default = 1: SRnd = .Value
    
352     .Key = "ColorTable_VICII": .Default = "hbtable_experimental.txt": ColorTable_Filename(Chip.vicii) = .Value
354     .Key = "ColorTable_TED": .Default = "table_ted6.txt": ColorTable_Filename(Chip.ted) = .Value
356     .Key = "ColorTable_VDC": .Default = "hbtable_vdc.txt": ColorTable_Filename(Chip.vdc) = .Value
    
358     .Section = "Drawing Settings"
        'Drawing Settings
360     .Key = "fillmode": .Default = "strict": FillMode = .Value
    
        'brush definition
362     .Key = "brushcolor1": .Default = 6: BrushColor1 = .Value
364     .Key = "brushcolor2": .Default = 14: BrushColor2 = .Value
366     .Key = "brushsize": .Default = 16: BrushSize = .Value
368     .Key = "brushdither": .Default = 8: BrushDither = .Value
370     .Key = "4thcolreplaces": .Default = "false": MainWin.mnu4thReplace.Checked = (.Value = True)
372     .Key = "ditherfill": .Default = "false": DitherFill = (.Value = True): MainWin.mnuDitherfill.Checked = DitherFill

        'PrevWin (preview window) position and size
374     .Section = PrevWin.Name
376     .LoadFormPosition PrevWin        ', PrevWin.Width, PrevWin.Height

        'main window position and size
378     .Section = MainWin.Name
380     .LoadFormPosition MainWin        ', PrevWin.Width, PrevWin.Height
    
        'main window position and size
382     .Section = Palett.Name
384     .LoadFormPosition Palett        ', PrevWin.Width, PrevWin.Height

        'Custom I/O setting (custom load/save)
386     .Section = "Custom Load Save"
388     For X = 0 To 15
390     .Key = "adr_screen(" & Format(X) & ")": .Default = 0: CustomIOSetup.Screen(X) = .Value
392     Next X
394     .Key = "adr_bitmap(0)":  .Default = 0: CustomIOSetup.Bitmap(0) = .Value
396     .Key = "adr_bitmap(1)": .Default = 0: CustomIOSetup.Bitmap(1) = .Value
398     .Key = "adr_d800":  .Default = 0: CustomIOSetup.D800 = .Value
400     .Key = "adr_d021": .Default = 0: CustomIOSetup.D021 = .Value
402     .Key = "adr_adressbyuser": .Default = 0: CustomIOSetup.StartAdressUser = .Value
404     .Key = "adr_savestart": .Default = 0: CustomIOSetup.Start = .Value
406     .Key = "adr_saveend": .Default = 0: CustomIOSetup.End = .Value
408     .Key = "adr_saveloadaddress": .Default = 0: CustomIOSetup.HasStartAddress = .Value
410     .Key = "adr_absolute": .Default = 0: CustomIOSetup.Absolute = .Value
412     .Key = "adr_overrideadress": .Default = 0: CustomIOSetup.ForceStartAddressFromUser = .Value
414     .Key = "adr_skipfirsttwobytes": .Default = 0: CustomIOSetup.HasStartAddress = .Value
    
    End With


        '<EhFooter>
        Exit Sub

LoadSettings_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.LoadSettings" + " line: " + Str(Erl))

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

Public Sub SaveSettings()
        '<EhHeader>
        On Error GoTo SaveSettings_Err
        '</EhHeader>
    Dim X As Long

100 With m_cIni


        'Brightness Ladder Load settings
102     .Path = App.Path & "\Ladders.ini"
104     .Section = "main"
106     .Key = "BrLadderSelected_VICII":   .Value = Gradient_Selected(Chip.vicii)
108     .Key = "BrLadderSelected_TED":  .Value = Gradient_Selected(Chip.ted)
110     .Key = "BrLadderSelected_VDC":  .Value = Gradient_Selected(Chip.vdc)
    
        'Save Num of Palette used
112     .Path = App.Path & "\Palettes.ini"
114     .Section = "Palettes"
116     .Key = "Count": pal_Count = .Value
118     .Key = "Selected_VICII": .Value = pal_Selected(Chip.vicii)
120     .Key = "Selected_VDC": .Value = pal_Selected(Chip.vdc)
122     .Key = "Selected_TED": .Value = pal_Selected(Chip.ted)
    
        'Main
124     .Path = App.Path & "\ProjectOne.ini"
126     .Section = PrevWin.Name
128     .SaveFormPosition PrevWin
    
130     .Section = MainWin.Name
132     .SaveFormPosition MainWin
    
134     .Section = Palett.Name
136     .SaveFormPosition Palett
    
138     .Section = "General Settings"
140     .Key = "previewzoom"
142     .Value = PreviewZoom
144     .Key = "lastloadpath": .Value = LastLoadPath
146     .Key = "lastsavepath": .Value = LastSavePath
148     .Key = "lastloadname": .Value = LastLoadName
150     .Key = "lastsavename": .Value = LastSaveName
152     .Key = "lastloadfilterindex": .Value = LastLoadFilterIndex
154     .Key = "lastsavefilterindex": .Value = LastSaveFilterindex
    
156     .Key = "picture_width":   .Value = PW
158     .Key = "picture_heigth":  .Value = PH

        'general gfxmode
160     .Key = "BaseMode": .Value = BaseMode
162     .Key = "GfxMode": .Value = GfxMode
164     .Key = "resodiv": .Value = ResoDiv
166     .Key = "flimul": .Value = FliMul
168     .Key = "xflilimit": .Value = XFliLimit
170     .Key = "bmpbanks": .Value = BmpBanks
172     .Key = "scrbanks": .Value = ScrBanks
    
        'custom gfx mode
174     .Key = "basemode_cm": .Value = BaseMode_cm
176     .Key = "resodiv_cm": .Value = ResoDiv_cm
178     .Key = "flimul_cm": .Value = FliMul_cm
180     .Key = "xflilimit_cm": .Value = XFliLimit_cm
182     .Key = "bmpbanks_cm": .Value = BmpBanks_cm
184     .Key = "scrbanks_cm": .Value = ScrBanks_cm
    
186     .Section = "View Settings"
        'view settings
188     .Key = "pixelgridcolor": .Value = PixelGridColor
190     .Key = "fligridcolor": .Value = FliGridColor
192     .Key = "chargridcolor": .Value = CharGridColor
194     .Key = "pixelgrid": .Value = PixelGrid
196     .Key = "chargrid": .Value = CharGrid
198     .Key = "pixelbox": .Value = PixelBox
200     .Key = "showflibox": .Value = ShowFlibox
202     .Key = "pixelboxcolor": .Value = PixelBoxColor   'boolean, 1 = colored pixbox, 0= white eor box
204     .Key = "zoomkeret": .Value = ZoomKeret
206     .Key = "showflilines": .Value = ShowFliLines
208     .Key = "pixelgridlimit": .Value = PixelGridLimit
210     .Key = "fligridlimit": .Value = FliGridLimit
212     .Key = "chargridlimit": .Value = CharGridLimit
214     .Key = "zoomwinleft": .Value = ZoomWinLeft
216     .Key = "zoomwintop": .Value = ZoomWinTop
218     .Key = "zoomscale": .Value = ZoomScale
220     .Key = "aratiox": .Value = ARatioX
222     .Key = "aratioy": .Value = ARatioY

224     .Section = "Converter Settings"
        'Converter Settings:
226     .Key = "lastvisitedtab": .Value = LastVisitedTab
228     .Key = "brightness": .Value = Brightness
230     .Key = "contrast": .Value = Contrast
232     .Key = "saturation": .Value = Saturation
234     .Key = "hue": .Value = Hue
236     .Key = "stretchpic": .Value = StretchPic
238     .Key = "keepaspect": .Value = KeepAspect
240     .Key = "reso_pixavg": .Value = Reso_PixAvg
242     .Key = "reso_pixleft": .Value = Reso_PixLeft
244     .Key = "reso_pixright": .Value = Reso_PixRight
246     .Key = "mostfreqbackgr": .Value = MostFreqBackgr
248     .Key = "usrdefbackgr": .Value = UsrDefBackgr
250     .Key = "usrbackgr": .Value = UsrBackgr
252     .Key = "optchars": .Value = OptChars
254     .Key = "mostfrq3": .Value = Mostfrq3
256     .Key = "fixwithbackgr": .Value = FixWithBackgr
258     .Key = "colorfiltermode": .Value = ColorFilterMode
260     .Key = "bdither": .Value = Bdither
262     .Key = "hdither": .Value = Hdither
264     .Key = "sdither": .Value = Sdither
266     .Key = "bditherval": .Value = BditherVal
268     .Key = "hditherval": .Value = Hditherval
270     .Key = "sditherval": .Value = SditherVal
272     .Key = "brnd": .Value = BRnd
274     .Key = "hrnd": .Value = HRnd
276     .Key = "srnd": .Value = SRnd
    
278     .Key = "ColorTable_VICII": .Value = ColorTable_Filename(Chip.vicii)
280     .Key = "ColorTable_TED": .Value = ColorTable_Filename(Chip.ted)
282     .Key = "ColorTable_VDC": .Value = ColorTable_Filename(Chip.vdc)
    
284     .Section = "Drawing Settings"
        'Drawing Settings
286     .Key = "fillmode": .Value = FillMode
    
        'brush settings
288     .Key = "brushcolor1": .Value = BrushColor1
290     .Key = "brushcolor2": .Value = BrushColor2
292     .Key = "brushsize": .Value = BrushSize
294     .Key = "brushdither": .Value = BrushDither
296     .Key = "4thcolreplaces": .Value = MainWin.mnu4thReplace.Checked
298     .Key = "ditherfill": .Value = DitherFill

        'Custom I/O setting (custom load/save)
300     .Section = "Custom Load Save"
    
  
302     For X = 0 To 15
304     .Key = "adr_screen(" & Format(X) & ")": .Value = CustomIOSetup.Screen(X)
306     Next X
308     .Key = "adr_bitmap(0)": .Value = CustomIOSetup.Bitmap(0)
310     .Key = "adr_bitmap(1)": .Value = CustomIOSetup.Bitmap(0)
312     .Key = "adr_d800": .Value = CustomIOSetup.D800
314     .Key = "adr_d021": .Value = CustomIOSetup.D021
316     .Key = "adr_adressbyuser": .Value = CustomIOSetup.StartAdressUser
318     .Key = "adr_savestart": .Value = CustomIOSetup.Start
320     .Key = "adr_saveend": .Value = CustomIOSetup.End
322     .Key = "adr_saveloadaddress": .Value = CustomIOSetup.HasStartAddress
324     .Key = "adr_absolute": .Value = CustomIOSetup.Absolute
326     .Key = "adr_overrideadress": .Value = CustomIOSetup.ForceStartAddressFromUser
328     .Key = "adr_skipfirsttwobytes": .Value = CustomIOSetup.HasStartAddress
    
    End With

        '<EhFooter>
        Exit Sub

SaveSettings_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.SaveSettings" + " line: " + Str(Erl))

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
'reads in requested palette
Public Sub LoadPalette(Index As Long)
        '<EhHeader>
        On Error GoTo LoadPalette_Err
        '</EhHeader>
    Dim PalCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long
    Dim Name As String
    Dim Colors As String

100 With m_cIni
    
102     .Path = App.Path & "\Palettes.ini"
104     .Section = "Palettes"
106     .Key = "count": .Default = 0: PalCount = .Value
    
108     If PalCount > 0 Then
110         .Key = "Name" & Format(Index): Name = .Value
112         .Section = .Value
114         .Key = "count": .Default = 0: ColorCount = .Value
116         pal_Count = .Value
118         If ColorCount <> 0 Then
120             For Y = 0 To ColorCount
122                 .Key = "c" & Format(Y): .Default = 0: Colors = .Value
124                 PaletteRGB(Y) = GetColor(Colors)
126             Next Y
            End If
        End If
 
    End With

        '<EhFooter>
        Exit Sub

LoadPalette_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.LoadPalette" + " line: " + Str(Erl))

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

Public Function GetPalName(Index As Long) As String
        '<EhHeader>
        On Error GoTo GetPalName_Err
        '</EhHeader>

100 With m_cIni
    
102     .Path = App.Path & "\Palettes.ini"
104     .Section = "Palettes"
106     .Key = "Name" & Format(Index)
108     .Default = ""
110     GetPalName = .Value
        
    End With

        '<EhFooter>
        Exit Function

GetPalName_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.GetPalName" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Function

Private Function GetColor(Colors As String) As Long
        '<EhHeader>
        On Error GoTo GetColor_Err
        '</EhHeader>
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim Color As Long

100 B = Val("&H" & Mid(Colors, 3, 2)) And &HFF
102 G = Val("&H" & Mid(Colors, 5, 2)) And &HFF
104 R = Val("&H" & Mid(Colors, 7, 2)) And &HFF
106 Color = RGB(R, G, B)

108 GetColor = Color
        '<EhFooter>
        Exit Function

GetColor_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.GetColor" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Function

Public Sub PaletteInit()
        '<EhHeader>
        On Error GoTo PaletteInit_Err
        '</EhHeader>
        Dim X As Long
        Dim Y As Long
        Dim Z As Long
        Dim yA As Single
        Dim Ua As Single
        Dim Va As Single
        Dim Yb As Single
        Dim Ub As Single
        Dim Vb As Single
        Dim Rd As Single
        Dim Gd As Single
        Dim bd As Single
100     ReDim coR(pal_Count)
102     ReDim coG(pal_Count)
104     ReDim coB(pal_Count)
106     ReDim cd(pal_Count, pal_Count)


108     For X = 0 To pal_Count
110         PixelsDib.Color(X) = PaletteRGB(X)
112         Palett.MyPalette1.PaletteRGB(X) = PaletteRGB(X)
114     Next X

116     Palett.MyPalette1.InitSurface
118     Palett.MyPalette1.BackgrColor = 0
120     Palett.MyPalette1.BoxCount = pal_Count

        'r g b components of c64 colors to get things faster
122     For Z = 0 To pal_Count
124         coR(Z) = PaletteRGB(Z) And 255
126         coG(Z) = Int(PaletteRGB(Z) / 256) And 255
128         coB(Z) = Int(PaletteRGB(Z) / 65536) And 255
130     Next Z

132     For X = 0 To pal_Count
134         For Y = 0 To pal_Count

136             yA = 0.299 * coR(X) + 0.587 * coG(X) + 0.114 * coB(X)
138             Ua = -0.147 * coR(X) - 0.289 * coG(X) + 0.436 * coB(X)
140             Va = 0.615 * coR(X) - 0.515 * coG(X) - 0.1 * coB(X)

142             Yb = 0.299 * coR(Y) + 0.587 * coG(Y) + 0.114 * coB(Y)
144             Ub = -0.147 * coR(Y) - 0.289 * coG(Y) + 0.436 * coB(Y)
146             Vb = 0.615 * coR(Y) - 0.515 * coG(Y) - 0.1 * coB(Y)

                'Ua = Ua / 2
                'Va = Va / 2
                'Ub = Ub / 2
                'Vb = Vb / 2

                'Rd = coR(X) - coR(Y)
                'Gd = coG(X) - coG(Y)
                'bd = coB(X) - coB(Y)

148             Rd = (yA - Yb) * 2
150             Gd = Ua - Ub
152             bd = Va - Vb

154             cd(X, Y) = Sqr(Rd * Rd + Gd * Gd + bd * bd)

156         Next Y
158     Next X

160     Palett.Caption = "Palette (" & GetPalName(pal_Selected(ChipType)) & ")"

        '<EhFooter>
        Exit Sub

PaletteInit_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.PaletteInit" + " line: " + Str(Erl))

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




 Public Function GetPathFromFullFileName(FullFileName As String) _
As String
    ' Passed a fully qualified filename, removes
    ' the filename part and returns the path information.
    ' Returns a blank string if there is no path information.
        '<EhHeader>
        On Error GoTo GetPathFromFullFileName_Err
        '</EhHeader>
    Dim idx As Integer
    ' Strip any spaces.
100 FullFileName = Trim(FullFileName)
    ' Check for empty argument.
102 If Len(FullFileName) = 0 Then
104 GetPathFromFullFileName = ""
    Exit Function
    End If

    ' Look for last / or \.
106 For idx = Len(FullFileName) To 1 Step -1
108 If Mid$(FullFileName, idx, 1) = "\" Or _
    Mid$(FullFileName, idx, 1) = "/" Then Exit For
110 Next idx
112 If idx = 1 Then
114 GetPathFromFullFileName = ""
    Else
116 GetPathFromFullFileName = Left$(FullFileName, idx)
    End If
        '<EhFooter>
        Exit Function

GetPathFromFullFileName_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.GetPathFromFullFileName" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Function

Public Function GetFileNameFromFullPath(sFile As String) As String
    ' by Peter Weighill, pweighill@btinternet.com, 20001020
    ' Only for VB6
        '<EhHeader>
        On Error GoTo GetFileNameFromFullPath_Err
        '</EhHeader>
      Dim iPos As Long
  
      ' search last backslash
100   iPos = InStrRev(sFile, "\", -1, vbBinaryCompare)
  
102   If iPos > 0 Then
104     GetFileNameFromFullPath = Mid$(sFile, iPos + 1)
      Else
106     GetFileNameFromFullPath = sFile
      End If
        '<EhFooter>
        Exit Function

GetFileNameFromFullPath_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Shared.GetFileNameFromFullPath" + " line: " + Str(Erl))

    Case vbAbort
       Resume ExitLine
    Case vbRetry
       Resume
    Case vbIgnore
       Resume Next

    End Select

ExitLine:

        '</EhFooter>
End Function

Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

