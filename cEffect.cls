VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' GDI+ Wrapper written by Paul
' http://www.syix.com/wpsjr1/index.html

' This class requires the gdiplus.dll version 6.x (or greater)
' Also required is my GDI+ 1.1 type library

Public Enum Effects
  Blur
  Sharpen
  Tint
  BrightnessContrast
  RedEyeCorrection
  ColorMatrixEffect
  ColorLUT
  HueSaturationLightness
  ColorBalance
  Levels
  ColorCurve
End Enum

Public gdip As cGDIP ' Reference for Error Handling

Private effect          As Long ' nativeEffect
Private lCurErr         As Long

Public Function CreateEffect(ByVal eEffect As Effects) As Long
  Select Case eEffect
    Case Blur:                   CreateEffect = Execute(GdipCreateEffect(&H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, effect))
    Case Sharpen:                CreateEffect = Execute(GdipCreateEffect(&H63CBF3EE, &H402CC526, &HC562718F, &H4251BF40, effect))
    Case Tint:                   CreateEffect = Execute(GdipCreateEffect(&H1077AF00, &H44412848, &HAD448994, &H2C7A2D4C, effect))
    Case BrightnessContrast:     CreateEffect = Execute(GdipCreateEffect(&HD3A1DBE1, &H4C178EC4, &H97EA4C9F, &H3D341CAD, effect))
    Case RedEyeCorrection:       CreateEffect = Execute(GdipCreateEffect(&H74D29D05, &H426669A4, &HC53C4995, &H32B63628, effect))
    Case ColorLUT:               CreateEffect = Execute(GdipCreateEffect(&HA7CE72A9, &H40D70F7F, &HC0D0CCB3, &H12325C2D, effect))
    Case ColorMatrixEffect:      CreateEffect = Execute(GdipCreateEffect(&H718F2615, &H40E37933, &H685F11A5, &H74DD14FE, effect))
    Case HueSaturationLightness: CreateEffect = Execute(GdipCreateEffect(&H8B2DD6C3, &H4D87EB07, &H871F0A5, &H5F9C6AE2, effect))
    Case ColorBalance:           CreateEffect = Execute(GdipCreateEffect(&H537E597D, &H48DA251E, &HCA296496, &HF8706B49, effect))
    Case Levels:                 CreateEffect = Execute(GdipCreateEffect(&H99C354EC, &H4F3A2A31, &HA817348C, &H253AB303, effect))
    Case ColorCurve:             CreateEffect = Execute(GdipCreateEffect(&HDD6A0022, &H4A6758E4, &H8ED49B9D, &H3DA581B8, effect))
  End Select
End Function

Public Function SetBlurParameters(ByVal radius As Single, ByVal bExpandEdge As Boolean) As Long
  Dim bp As BlurParams
  
  bp.radius = radius ' 0 to 255
  bp.expandEdge = Abs(bExpandEdge)
  SetBlurParameters = Execute(GdipSetEffectParameters(effect, bp, Len(bp)))
End Function

Public Function SetSharpenParameters(ByVal radius As Single, ByVal amount As Single) As Long
  Dim sp As SharpenParams
  
  sp.radius = radius ' 0 to 255
  sp.amount = amount ' 0 to 100
  SetSharpenParameters = Execute(GdipSetEffectParameters(effect, sp, Len(sp)))
End Function

Public Function SetTintParameters(ByVal hue As Long, ByVal amount As Long) As Long
  Dim tp As TintParams
  
  tp.hue = hue       ' -180 to 180
  tp.amount = amount ' -100 to 100
  SetTintParameters = Execute(GdipSetEffectParameters(effect, tp, Len(tp)))
End Function

Friend Function SetRedEyeCorrectionParameters(ByRef rc() As RECT) As Long
  Dim rp As RedEyeCorrectionParams
  Const RECT_SIZE As Long = 16
  
  If (Not rc) <> -1 Then ' not empty array
    rp.numberOfAreas = UBound(rc) - LBound(rc) + 1
    rp.prcAreas = VarPtr(rc(LBound(rc)))
    SetRedEyeCorrectionParameters = Execute(GdipSetEffectParameters(effect, rp, Len(rp) + (rp.numberOfAreas * RECT_SIZE)))
  End If
End Function

Friend Function SetColorMatrixEffectParameters(ByRef m As COLORMATRIX) As Long
  Const MATRIX_SIZE = 100 ' Len(Single) * 25
  SetColorMatrixEffectParameters = Execute(GdipSetEffectParameters(effect, m, MATRIX_SIZE))
End Function

Friend Function SetColorLUTParameters(ByRef lut As ColorLUTParams) As Long
  SetColorLUTParameters = Execute(GdipSetEffectParameters(effect, lut, Len(lut)))
End Function

Public Function SetBrightnessContrastParameters(ByVal brightness As Long, ByVal contrast As Long) As Long
  Dim bcp As BrightnessContrastParams
  
  bcp.brightnessLevel = brightness  ' -255 to 255
  bcp.contrastLevel = contrast      ' -100 to 100
  SetBrightnessContrastParameters = Execute(GdipSetEffectParameters(effect, bcp, Len(bcp)))
End Function

Public Function SetHueSaturationLightnessParameters(ByVal hue As Long, ByVal saturation As Long, ByVal lightness As Long) As Long
  Dim hsl  As HueSaturationLightnessParams
  
  hsl.hueLevel = hue                ' -180 to 180
  hsl.saturationLevel = saturation  ' -100 to 100
  hsl.lightnessLevel = lightness    ' -100 to 100
  SetHueSaturationLightnessParameters = Execute(GdipSetEffectParameters(effect, hsl, Len(hsl)))
End Function

Public Function SetColorBalanceParameters(ByVal cyan As Long, ByVal magneta As Long, ByVal yellow As Long) As Long
  Dim cmy  As ColorBalanceParams
  
  cmy.cyanRed = cyan         ' -100 to 100
  cmy.magentaGreen = magneta ' -100 to 100
  cmy.yellowBlue = yellow    ' -100 to 100
  SetColorBalanceParameters = Execute(GdipSetEffectParameters(effect, cmy, Len(cmy)))
End Function

Public Function SetLevelsParameters(ByVal highlight As Long, ByVal midtone As Long, ByVal shadow As Long) As Long
  Dim hms  As LevelsParams
  
  hms.highlight = highlight ' 0 to 100
  hms.midtone = midtone     ' -100 to 100
  hms.shadow = shadow       ' 0 to 100
  SetLevelsParameters = Execute(GdipSetEffectParameters(effect, hms, Len(hms)))
End Function

Public Function SetColorCurveParameters(ByVal eAdjustments As CurveAdjustments, ByVal eChannel As CurveChannel, ByVal adjustValue As Long) As Long
  Dim ccp  As ColorCurveParams
  
  ccp.adjustment = eAdjustments
  ccp.channel = eChannel
  ccp.adjustValue = adjustValue
  SetColorCurveParameters = Execute(GdipSetEffectParameters(effect, ccp, Len(ccp)))
End Function

Public Property Get Handle() As Long
  Handle = effect
End Property

Public Sub Dispose()
  If effect Then
    GdipDeleteEffect effect
    effect = 0
  End If
End Sub

Private Function Execute(ByVal lReturn As Long) As Long
  If lReturn = 0 Then
    lCurErr = 0
  Else
    lCurErr = lReturn
    If Not gdip Is Nothing Then gdip.Error lReturn, "cEffect"
  End If
  Execute = lCurErr
End Function

Friend Function StringToGuid(ByRef sGuid As String) As CLSID
  CLSIDFromString sGuid, StringToGuid
End Function

Private Sub Class_Terminate()
  Dispose
End Sub
