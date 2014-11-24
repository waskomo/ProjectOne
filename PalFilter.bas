Attribute VB_Name = "PalFilter"
Option Explicit

Dim coY(15) As Long
Dim coU(15) As Long
Dim cov(15) As Long

Dim coR(15) As Long
Dim coG(15) As Long
Dim coB(15) As Long


Private Sub InitPalFilter()
Dim z As Long

For z = 0 To 15

    coR(z) = PaletteRGB(z) And 255
    coG(z) = (PaletteRGB(z) \ 256) And 255
    coB(z) = (PaletteRGB(z) \ 65536) And 255
    
    'set up c64 y u v arrays
    coY(z) = 0.299 * coR(z) + 0.587 * coG(z) + 0.114 * coB(z)
    coU(z) = -0.147 * coR(z) - 0.289 * coG(z) + 0.436 * coB(z)
    cov(z) = 0.615 * coR(z) - 0.515 * coG(z) - 0.1 * coB(z)
Next z

ZoomWindow.Buffer.Width = PW * 2
ZoomWindow.Buffer.Height = PH * 2
ZoomWindow.Buffer.BackColor = 0
ZoomWindow.Buffer.Cls

End Sub

Public Sub ApplyPalFilter()
Dim X As Long
Dim Y As Long
Dim Uavg As Single
Dim Vavg As Single
Dim Yavg As Single
Dim R As Single
Dim G As Single
Dim B As Single

Dim Uavgold As Single
Dim Vavgold As Single

'Dim Uprev(PW) As Single
'Dim vprev(PW) As Single

Call InitPalFilter

'aabbccdd
'1 2 3 4
' 1 2 3 4

For Y = 1 To PH - 1
 For X = 8 To 312

    '
    '0aabbccdd
    ' 321 1234
    '321 1234
    
    Yavg = coY(pixels(X, Y))
    
    Uavg = (coU(pixels(X, Y)) + _
            coU(pixels(X - 1, Y - 1)) + _
            coU(pixels(X, Y - 1)) + _
            coU(pixels(X + 1, Y - 1)) + _
            coU(pixels(X + 1, Y)) + _
            coU(pixels(X + 1, Y + 1)) + _
            coU(pixels(X, Y + 1)) + _
            coU(pixels(X - 1, Y + 1))) / 8
            
 Vavg = (cov(pixels(X, Y)) + _
            cov(pixels(X - 1, Y - 1)) + _
            cov(pixels(X, Y - 1)) + _
            cov(pixels(X + 1, Y - 1)) + _
            cov(pixels(X + 1, Y)) + _
            cov(pixels(X + 1, Y + 1)) + _
            cov(pixels(X, Y + 1)) + _
            cov(pixels(X - 1, Y + 1))) / 8
            
    
    
    R = Yavg + 1.14 * Vavg
    G = Yavg - 0.395 * Uavg - 0.581 * Vavg
    B = Yavg + 2.032 * Uavg
    
        
    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If B < 0 Then B = 0
    
    ZoomWindow.Buffer.ForeColor = RGB(R, G, B)
    ZoomWindow.Buffer.PSet (X * 2, Y * 2)
    ZoomWindow.Buffer.PSet ((X * 2) + 1, Y * 2)
    
    R = R / 1.5
    G = G / 1.5
    B = B / 1.5
    ZoomWindow.Buffer.ForeColor = RGB(R, G, B)
    
    ZoomWindow.Buffer.PSet (X * 2, (Y * 2) + 1)
    ZoomWindow.Buffer.PSet ((X * 2) + 1, (Y * 2) + 1)
    
 Next X
Next Y

ZoomWindow.Buffer.Refresh
Call PrevWin.PalFilterRefresh

End Sub
