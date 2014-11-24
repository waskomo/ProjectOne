VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "P1Picture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

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

Private SA As SAFEARRAY2D




Private p_PW As Long     'size in pixels
Private p_PH As Long

Private p_CW As Long     'size in chars
Private p_CH As Long

Private p_ChipType As Chip
Private p_FliMul As Long
Private p_BaseMode As BaseModeTyp
Private p_ResoDiv As Long
Private p_BmpBanks As Long
Private p_ScrBanks As Long
Private p_HasSprites As Boolean
Private p_RegPerLine As Boolean
Private p_EmulateFliBug As Boolean
Private p_ZoomWinLeft
Private p_ZoomWinTop

'Brush settings
Private p_BrushSize As Long
Private p_BrushDither As Long
Private p_BrushType As BrushTyp
Private p_BrushArray(31, 31) As Long
Private p_BrushColor1 As Long
Private p_BrushColor2 As Long
Private p_Bayer(3, 3) As Long

'picture information
Private p_Pixels() As Byte
Private p_D800() As Byte
Private p_ScrRam() As Byte
Private p_Bitmap() As Byte
Private p_Registers() As Byte
Private p_BackGrIndex As Byte

'tables to speed up bitwise operations
Private BitCount00(255) As Long
Private BitCount01(255) As Long
Private BitCount10(255) As Long
Private BitCount11(255) As Long

Private BitCount1(255) As Long
Private BitCount0(255) As Long

Private Mask1(7) As Long
Private InvMask0(7) As Long

Private Mask11(3, 3) As Long
Private InvMask00(3) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" _
     (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
     (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
     


'public for speed (-> so these has to be calculated only once)
'variables to access attribute/bitmap information
Private CX As Long
Private CY As Long
Private ScrNum As Long
Private ScrBank As Long
Private BmpBank As Long


Public Sub Resize(ByVal NewWidth As Long, ByVal NewHeight As Long)



p_PW = NewWidth
p_PH = NewHeight
p_CW = Int(p_PW / 8)
p_CH = Int(p_PH / 8)

ReDim p_Pixels(p_PW, p_PH) As Byte
ReDim p_D800(p_CW, p_CH) As Byte
ReDim p_ScrRam(1, 7, p_CW, p_CH, 1) As Byte
ReDim p_Bitmap(1, p_CW, p_PH) As Byte
ReDim Registers(1, 63, p_PH) As Byte
ReDim ColMap(p_PW, p_PH) As Byte

End Sub


Public Sub DMAHack(ByVal DIBdata As Long)

    With SA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = p_PH
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = p_PW
        .pvData = DIBdata
    End With

    CopyMemory ByVal VarPtrArray(p_Pixels()), VarPtr(SA), 4
    
End Sub
Public Sub PutBitmap(ByVal X As Long, ByVal Y As Long, ByVal BitMask As Long, ByVal ColorIndex As Byte)

Dim Z As Long

Debug.Print "putbitmap chiptype"; ChipType

    Select Case ChipType

        Case Chip.vicii


            If p_BaseMode = BaseModeTyp.multi Then

                'set bitmap
                Z = 3 - (Int(X / 2) And 3)
                p_Bitmap(BmpBank, CX, Y) = (p_Bitmap(BmpBank, CX, Y) And InvMask00(Z)) Or Mask11(BitMask, Z)

                'set color
                Select Case BitMask
                Case 1
                    p_ScrRam(ScrBank, ScrNum, CX, CY, 0) = ColorIndex
                Case 2
                    p_ScrRam(ScrBank, ScrNum, CX, CY, 1) = ColorIndex
                Case 3
                    p_D800(CX, CY) = ColorIndex
                End Select
                
                Debug.Print "putbitmap multi. invmask mask11"; InvMask00(Z); Mask11(BitMask, Z)
            End If


            If p_BaseMode = BaseModeTyp.hires Then

                'set bitmap
                Z = 7 - (Int(X / 1) And 7)
                p_Bitmap(BmpBank, CX, Y) = p_Bitmap(BmpBank, CX, Y) And InvMask0(Z)
                If BitMask = 1 Then
                    p_Bitmap(BmpBank, CX, Y) = p_Bitmap(BmpBank, CX, Y) Or Mask1(Z)
                End If

                'set colors
                Select Case BitMask
                Case 0
                    p_ScrRam(ScrBank, ScrNum, CX, CY, 0) = ColorIndex
                Case 1
                    p_ScrRam(ScrBank, ScrNum, CX, CY, 1) = ColorIndex
                End Select

            End If

    End Select

End Sub

Public Sub BrushPreCalc()
Dim X As Single
Dim Y As Single
Dim D As Single
Dim Round As Single


For X = 0 To 31
    For Y = 0 To 31
        p_BrushArray(X, Y) = 0
    Next Y
Next X

Round = ((p_BrushSize \ 2) / p_ResoDiv) * p_ResoDiv

For X = -Round To Round Step p_ResoDiv
    For Y = -Round To Round Step 1

        D = Sqr(X * X + Y * Y)

        If D < Int((p_BrushSize / 2) + 0.5) Then

            p_BrushArray((31 \ 2) + X, (31 \ 2) + Y) = 1

            If p_ResoDiv = 2 Then _
               p_BrushArray((31 \ 2) + X + 1, (31 \ 2) + Y) = 1

        End If

    Next Y
Next X


End Sub
Public Sub PlotBrush(ByVal pbx As Long, ByVal pby As Long)
Dim px As Long
Dim py As Long
Dim pcol As Long
Dim Qx As Long
Dim Qy As Long

For px = 0 To 31 Step p_ResoDiv
    For py = 0 To 31

        If p_BrushArray(px, py) = 1 Then

            Qx = px + pbx - 16
            Qy = py + pby - 16

            If Qx >= 0 And Qx <= PW - 1 And _
               Qy >= 0 And Qy <= PH - 1 Then

                If p_BrushDither < p_Bayer((Qx \ p_ResoDiv) And 3, Qy And 3) Then pcol = p_BrushColor1 Else pcol = p_BrushColor2
                Call SetPixel(Qx, Qy, pcol)

            End If

        End If

    Next py
Next px

End Sub


Private Sub DrawPicFromMem()

Dim Zx As Long
Dim Zy As Long

Dim Qx As Long
Dim Qy As Long
Dim bmp As Byte
Dim Temp As Byte
Dim andtab(7) As Long
Dim divtab(7) As Long
Dim ColNum As Byte
Dim BmpX As Byte

Dim X As Long
Dim Y As Long


Select Case p_ChipType
    Case vicii

        Select Case p_BaseMode
            Case BaseModeTyp.multi
            
                andtab(3) = 3 * 1
                andtab(2) = 3 * 4
                andtab(1) = 3 * 16
                andtab(0) = 3 * 64

                divtab(3) = 1
                divtab(2) = 4
                divtab(1) = 16
                divtab(0) = 64

                For Zy = 0 To p_CH - 1
                    For Zx = 0 To p_CW - 1
                        Qx = Zx * 8: Qy = Zy * 8
                        For Y = 0 To 7
                            For X = 0 To 7 Step p_ResoDiv
                                BmpX = Int(X / 2)
                                Call AccesSetup(X, Y)
                                bmp = p_Bitmap(BmpBank, Zx, Qy + Y)
                                Temp = (bmp And andtab(BmpX)) / divtab(BmpX)
                    
                                    Select Case Temp
                                    Case 0
                                        ColNum = p_BackGrIndex
                                    Case 1
                                        ColNum = p_ScrRam(ScrBank, ScrNum, Zx, Zy, 0)
                                    Case 2
                                        ColNum = p_ScrRam(ScrBank, ScrNum, Zx, Zy, 1)
                                    Case 3
                                        ColNum = p_D800(Zx, Zy)
                                    End Select

                                p_Pixels(Qx + X, Qy + Y) = ColNum
                                If p_ResoDiv = 2 Then p_Pixels(Qx + X + 1, Qy + Y) = ColNum
                                
                            Next X
                        Next Y
                    Next Zx
                Next Zy

            Case BaseModeTyp.hires

                andtab(7) = 1
                andtab(6) = 2
                andtab(5) = 4
                andtab(4) = 8
                andtab(3) = 16
                andtab(2) = 32
                andtab(1) = 64
                andtab(0) = 128

                divtab(7) = 1
                divtab(6) = 2
                divtab(5) = 4
                divtab(4) = 8
                divtab(3) = 16
                divtab(2) = 32
                divtab(1) = 64
                divtab(0) = 128

                BmpBank = 0
                ScrBank = 0
    
                For Zy = 0 To CH - 1
                    For Zx = 0 To CW - 1
                        Qx = Zx * 8: Qy = Zy * 8
                        For Y = 0 To 7
                            For X = 0 To 7
                                ScrNum = Int(((Qy + Y) And 7) / p_FliMul)
                                bmp = p_Bitmap(BmpBank, Zx, Qy + Y)
                                Temp = Int(bmp / divtab(X)) And 1
                                Select Case Temp
                                Case 0
                                    ColNum = p_ScrRam(ScrBank, ScrNum, Zx, Zy, 0)
                                Case 1
                                    ColNum = p_ScrRam(ScrBank, ScrNum, Zx, Zy, 1)
                                End Select
                                
                                p_Pixels(Qx + X, Qy + Y) = ColNum

                            Next X

                        Next Y
                    Next Zx
                Next Zy

        End Select

End Select

End Sub








Public Sub PutPixel(ByVal X As Long, ByVal Y As Long, ByVal ColorIndex As Byte)

    'p_Pixels(X, Y) = ColorIndex

    MsgBox "dont use putpixel in p1picture "
    
End Sub

Public Function GetPixel(ByVal X As Long, ByVal Y As Long) As Byte

    GetPixel = p_Pixels(X, Y)

End Function

Public Sub SetD800(ByVal X As Long, ByVal Y As Long, ByVal ColorIndex As Byte)

p_D800(X, Y) = ColorIndex

End Sub

Public Function GetD800(ByVal X As Long, ByVal Y As Long) As Byte

GetD800 = p_D800(X, Y)

End Function

Public Sub SetScreen(ByVal X As Long, ByVal Y As Long, ByVal Nibble As Long, ByVal ColorIndex As Byte)

Call AccesSetup(X, Y)
p_ScrRam(ScrBank, ScrNum, CX, CY, Nibble) = ColorIndex

End Sub

Public Function GetScreen(ByVal X As Long, ByVal Y As Long, ByVal Nibble As Long) As Byte

Call AccesSetup(X, Y)
GetScreen = p_ScrRam(ScrBank, ScrNum, CX, CY, Nibble)

End Function

Public Sub SetBitmap(ByVal Bank As Long, ByVal X As Long, ByVal Y As Long, ByVal Mask As Byte)

    p_Bitmap(Bank, X, Y) = Mask
    
End Sub

Public Function GetBitmap(ByVal Bank As Long, ByVal ScrIndex As Long, ByVal X As Long, ByVal Y As Long) As Byte

    'If X < 40 Then
    GetBitmap = p_Bitmap(Bank, X, Y)

End Function

Public Property Let EmulateFliBug(ByVal vData As Boolean)
p_EmulateFliBug = vData
End Property


Public Property Get EmulateFliBug() As Boolean
EmulateFliBug = p_EmulateFliBug
End Property



Public Property Let RegPerLine(ByVal vData As Boolean)

End Property


Public Property Get RegPerLine() As Boolean

End Property



Public Property Let HasSprites(ByVal vData As Boolean)

End Property


Public Property Get HasSprites() As Boolean

End Property



Public Property Let ScrBanks(ByVal vData As Long)
p_ScrBanks = vData
End Property


Public Property Get ScrBanks() As Long
ScrBanks = p_ScrBanks
End Property



Public Property Let BmpBanks(ByVal vData As Long)
p_BmpBanks = vData
End Property


Public Property Get BmpBanks() As Long
BmpBanks = p_BmpBanks
End Property



Public Property Let ResoDiv(ByVal vData As Long)
p_ResoDiv = vData
End Property


Public Property Get ResoDiv() As Long
ResoDiv = p_ResoDiv
End Property




Public Property Let BaseMode(ByVal vData As BaseModeTyp)
p_BaseMode = vData
End Property


Public Property Get BaseMode() As BaseModeTyp
BaseMode = p_BaseMode
End Property

Public Property Let ZoomWinTop(ByVal vData As Long)
p_ZoomWinTop = vData
End Property


Public Property Get ZoomWinTop() As Long
ZoomWinTop = p_ZoomWinTop
End Property

Public Property Let ZoomWinLeft(ByVal vData As Long)
p_ZoomWinLeft = vData
End Property


Public Property Get ZoomWinLeft() As Long
ZoomWinLeft = p_ZoomWinLeft
End Property



Public Property Let FliMul(ByVal vData As Long)
p_FliMul = vData
End Property


Public Property Get FliMul() As Long
FliMul = p_FliMul
End Property



Public Property Let ChipType(ByVal vData As Chip)
p_ChipType = vData
End Property


Public Property Get ChipType() As Chip
ChipType = p_ChipType
End Property



Public Property Let Height(ByVal vData As Long)
p_PH = vData
End Property


Public Property Get Height() As Long
Height = p_PH
End Property



Public Property Let Width(ByVal vData As Long)
p_PW = vData
End Property


Public Property Get Width() As Long
Attribute Width.VB_UserMemId = 0
Width = p_PW
End Property
Public Property Let BrushDither(ByVal vData As Long)
p_BrushDither = vData
End Property

Public Property Get BrushDither() As Long
BrushDither = p_BrushDither
End Property

Public Property Let BrushSize(ByVal vData As Long)
p_BrushSize = vData
End Property

Public Property Get BrushSize() As Long
BrushSize = p_BrushSize
End Property

Public Property Let BrushColor1(ByVal vData As Long)
p_BrushColor1 = vData
End Property

Public Property Get BrushColor1() As Long
BrushColor1 = p_BrushColor1
End Property
Public Property Let BrushColor2(ByVal vData As Long)
p_BrushColor2 = vData
End Property

Public Property Get BrushColor2() As Long
BrushColor2 = p_BrushColor2
End Property

Public Property Let BackgrIndex(ByVal vData As Long)
p_BackGrIndex = vData
End Property

Public Property Get BackgrIndex() As Long
BackgrIndex = p_BackGrIndex
End Property

Public Sub Cls()

    Dim X As Long
    Dim Y As Long
    Dim ScrNum As Long
    Dim Nibble As Long
    Dim Bank As Long
    
    'clear pixels
    For X = 0 To p_PW - 1
        For Y = 0 To p_PH - 1
            p_Pixels(X, Y) = 255
        Next Y
    Next X

    'clear d800
    For X = 0 To Int((p_PW - 1) / 8)
        For Y = 0 To Int((p_PH - 1) / 8)
            p_D800(X, Y) = 0
        Next Y
    Next X

    'clear scrram
    For Bank = 0 To p_BmpBanks
        For ScrNum = 0 To 7
            For X = 0 To Int((p_PW - 1) / 8)
                For Y = 0 To Int((p_PH - 1) / 8)
                        p_ScrRam(Bank, ScrNum, X, Y, 0) = 0
                        p_ScrRam(Bank, ScrNum, X, Y, 1) = 0
                Next Y
            Next X
        Next ScrNum
    Next Bank
    
    'clear bitmap
    For Bank = 0 To p_BmpBanks
            For X = 0 To Int((p_PW - 1) / 8)
                For Y = 0 To p_PH - 1
                    p_Bitmap(Bank, X, Y) = 0
                Next Y
            Next X
    Next Bank
    
End Sub

Private Sub Class_Initialize()
    
    
    'this all are needed to setup bitwise speedup tables
    Dim BP00 As Long
    Dim BP01 As Long
    Dim BP10 As Long
    Dim BP11 As Long

    Dim bp1 As Long
    Dim Bp0 As Long

    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    
    
    Call Resize(319, 399)

    'generating basic hires bitmask table
    bp1 = 1
    For X = 0 To 7
        Mask1(X) = bp1
        InvMask0(X) = 255 - bp1
        bp1 = bp1 * 2
    Next X

    'generating basic multi bitmask table
    BP00 = 0
    BP01 = 1
    BP10 = 2
    BP11 = 3

    For X = 0 To 3
        Mask11(0, X) = BP00
        Mask11(1, X) = BP01
        Mask11(2, X) = BP10
        Mask11(3, X) = BP11
        InvMask00(X) = 255 - BP11
        BP00 = BP00 * 4
        BP01 = BP01 * 4
        BP10 = BP10 * 4
        BP11 = BP11 * 4
    Next X


    'generating bit / bitpair count tables
    For X = 0 To 255

        BitCount00(X) = 0
        BitCount01(X) = 0
        BitCount10(X) = 0
        BitCount11(X) = 0

        BitCount1(X) = 0
        BitCount0(X) = 0

        Y = X
        For Z = 0 To 3
            If (Y And 3) = 0 Then BitCount00(X) = BitCount00(X) + 1
            If (Y And 3) = 1 Then BitCount01(X) = BitCount01(X) + 1
            If (Y And 3) = 2 Then BitCount10(X) = BitCount10(X) + 1
            If (Y And 3) = 3 Then BitCount11(X) = BitCount11(X) + 1
            Y = Int(Y / 4)
        Next Z

        Y = X
        For Z = 0 To 7
            If (Y And 1) = 0 Then BitCount0(X) = BitCount0(X) + 1
            If (Y And 1) = 1 Then BitCount1(X) = BitCount1(X) + 1
            Y = Int(Y / 2)
        Next Z


    Next X

    For X = 0 To 3
        For Y = 0 To 3
            For Z = 0 To 255
                If Z And InvMask00(0) = X * 1 Then BitSwap(X, Y, Z) = (Z And InvMask00(0)) + Y * 1
                If Z And InvMask00(1) = X * 4 Then BitSwap(X, Y, Z) = (Z And InvMask00(1)) + Y * 4
                If Z And InvMask00(2) = X * 16 Then BitSwap(X, Y, Z) = (Z And InvMask00(2)) + Y * 16
                If Z And InvMask00(3) = X * 64 Then BitSwap(X, Y, Z) = (Z And InvMask00(3)) + Y * 64
            Next Z
        Next Y
    Next X

'brush dither array
p_Bayer(0, 0) = 1: p_Bayer(0, 1) = 9: p_Bayer(0, 2) = 3: p_Bayer(0, 3) = 11
p_Bayer(1, 0) = 13: p_Bayer(1, 1) = 5: p_Bayer(1, 2) = 15: p_Bayer(1, 3) = 7
p_Bayer(2, 0) = 4: p_Bayer(2, 1) = 12: p_Bayer(2, 2) = 2: p_Bayer(2, 3) = 10
p_Bayer(3, 0) = 16: p_Bayer(3, 1) = 8: p_Bayer(3, 2) = 14: p_Bayer(3, 3) = 6

End Sub


Private Function GetVICIIScrIndex(ByVal Y As Long)

GetVICIIScrIndex = Int((Y And 7) / p_FliMul)

End Function

Private Sub GetVICIIBanks(ByVal X As Long)

Select Case p_BaseMode
    Case multi
        If ScrBanks = 1 Then ScrBank = X And 1 Else ScrBank = 0
        If BmpBanks = 1 Then BmpBank = X And 1 Else BmpBank = 0
    Case hires
        BmpBank = 0
        ScrBank = 0
End Select

End Sub

Private Sub AccesSetup(ByVal X As Long, ByVal Y As Long)



    Select Case ChipType

        Case Chip.vicii
        
            CX = Int(X / 8)
            CY = Int(Y / 8)
            ScrNum = GetVICIIScrIndex(Y)
            GetVICIIBanks (X)
        
        Case Chip.ted
        
            MsgBox "p1picture: accesssetup unimplemented tedmode"
            
        Case Chip.vdc
            
            MsgBox "p1picture: accesssetup unimplemented vdcmode"
            
    End Select

End Sub
