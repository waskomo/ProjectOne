Attribute VB_Name = "Drawing"
Option Explicit

Public CX As Long
Public CY As Long

Public BitCount00(255) As Long
Public BitCount01(255) As Long
Public BitCount10(255) As Long
Public BitCount11(255) As Long

Public BitCount1(255) As Long
Public BitCount0(255) As Long

Public Mask11(3, 3) As Long

Public BP00 As Long
Public BP01 As Long
Public BP10 As Long
Public BP11 As Long

Public Mask1(7) As Long

Public InvMask0(7) As Long
Public InvMask00(3) As Long

Public Shift(3, 3) As Long




Public Sub SetPixel(ByVal X As Long, ByVal Y As Long, ByVal ColorIndex As Byte)
        '<EhHeader>
        On Error GoTo SetPixel_Err
        '</EhHeader>

        Dim BitMask As Integer

100     If DitherMode = True Then
102         If pattern(X / ResoDiv And 1, Y And 1) = 1 Then ColorIndex = LeftColN
104         If pattern(X / ResoDiv And 1, Y And 1) = 0 Then ColorIndex = RightColN
        End If
                
106     BitMask = PixelFits(X, Y, ColorIndex)
        'Debug.Print "pixelfits returned"; BitMask
108     If BitMask >= 0 Then
           ' Debug.Print "putbitmap x y bitmask color"; X; Y; BitMask; ColorIndex
110         Call PutBitmap(X, Y, BitMask, ColorIndex)
112         Pixels(X, Y) = ColorIndex
114         If ResoDiv = 2 Then
116             Pixels(X + 1, Y) = ColorIndex
            End If
118     ElseIf MainWin.mnu4thReplace.Checked = True Then
120         Call ChangeCol(ColorIndex)
        End If


        '<EhFooter>
        Exit Sub

SetPixel_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.SetPixel" + " line: " + Str(Erl))

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
Public Sub AccesSetup(ByVal X As Long, ByVal Y As Long)
        '<EhHeader>
        On Error GoTo AccesSetup_Err
        '</EhHeader>

100     Select Case ChipType

            Case Chip.vicii
                    
102             CX = X \ 8
104             CY = Y \ 8
106             ScrNum = GetVICIIScrIndex(Y)
108             GetVICIIBanks (X)
                'Debug.Print "acessetup vicII"
            
110         Case Chip.ted
        
112             MsgBox "p1picture: accesssetup unimplemented for TED mode"
            
114         Case Chip.vdc
            
116             MsgBox "p1picture: accesssetup unimplemented for VDC mode"
            
        End Select

        '<EhFooter>
        Exit Sub

AccesSetup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.AccesSetup" + " line: " + Str(Erl))

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
Private Function GetVICIIScrIndex(ByVal Y As Long)
        '<EhHeader>
        On Error GoTo GetVICIIScrIndex_Err
        '</EhHeader>

100 GetVICIIScrIndex = (Y And 7) \ FliMul

        '<EhFooter>
        Exit Function

GetVICIIScrIndex_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.GetVICIIScrIndex" + " line: " + Str(Erl))

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

Private Sub GetVICIIBanks(ByVal X As Long)
        '<EhHeader>
        On Error GoTo GetVICIIBanks_Err
        '</EhHeader>


100 Select Case BaseMode
        Case BaseModeTyp.multi
102         If ScrBanks = 1 Then ScrBank = X And 1 Else ScrBank = 0
104         If BmpBanks = 1 Then BmpBank = X And 1 Else BmpBank = 0
106     Case BaseModeTyp.hires
108         BmpBank = 0
110         ScrBank = 0
    End Select

        '<EhFooter>
        Exit Sub

GetVICIIBanks_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.GetVICIIBanks" + " line: " + Str(Erl))

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

Public Sub PutBitmap(ByVal X As Long, ByVal Y As Long, ByVal BitMask As Long, ByVal ColorIndex As Byte)
        '<EhHeader>
        On Error GoTo PutBitmap_Err
        '</EhHeader>

    Dim Z As Long

100 Call AccesSetup(X, Y)

102     Select Case ChipType

            Case Chip.vicii


104             If BaseMode = BaseModeTyp.multi Then

                    'set bitmap
106                 Z = 3 - (Int(X / 2) And 3)
108                 Bitmap(BmpBank, CX, Y) = (Bitmap(BmpBank, CX, Y) And InvMask00(Z)) Or Mask11(BitMask, Z)

                    'set color
110                 Select Case BitMask
                    Case 1
112                     ScrRam(ScrBank, ScrNum, CX, CY, 0) = ColorIndex
114                 Case 2
116                     ScrRam(ScrBank, ScrNum, CX, CY, 1) = ColorIndex
118                 Case 3
120                     D800(CX, CY) = ColorIndex
                    End Select
                
                End If


122             If BaseMode = BaseModeTyp.hires Then

                    'set bitmap
124                 Z = 7 - (X And 7)
                
126                 Bitmap(BmpBank, CX, Y) = Bitmap(BmpBank, CX, Y) And InvMask0(Z)
128                 If BitMask = 1 Then
130                     Bitmap(BmpBank, CX, Y) = Bitmap(BmpBank, CX, Y) Or Mask1(Z)
                        'Debug.Print "oraing mask1(z) bmpbank cx y"; Mask1(Z); Z; X
                    End If
                
                    'Debug.Print "bitmap byte changed to:"; Dec2Bin(Bitmap(BmpBank, CX, Y))
                    'set colors
132                 Select Case BitMask
                    Case 1
134                     ScrRam(ScrBank, ScrNum, CX, CY, 1) = ColorIndex
136                 Case 0
138                     ScrRam(ScrBank, ScrNum, CX, CY, 0) = ColorIndex
                    End Select

                End If

        End Select

        '<EhFooter>
        Exit Sub

PutBitmap_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.PutBitmap" + " line: " + Str(Erl))

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
Public Function PixelFits(ByVal X As Long, ByVal Y As Long, ByVal ColorIndex As Byte) As Integer
        '<EhHeader>
        On Error GoTo PixelFits_Err
        '</EhHeader>

        Dim Col2Change As Long
        Dim Qbank As Long
    
        Dim BP00 As Long
        Dim BP01 As Long
        Dim BP10 As Long
        Dim BP11 As Long

        Dim bp1 As Long
        Dim Bp0 As Long

        Dim YY As Long
    
        Dim Ls As Long
        Dim Le As Long
    
100     PixelFits = -1
    
102     If Y >= PH Or X >= PW Then
            Exit Function
        End If
    
104     Call AccesSetup(X, Y)
106     Col2Change = Pixels(X, Y)

108     Select Case ChipType

            Case vicii

110             If BaseMode = BaseModeTyp.multi Then


112                 Select Case ColorIndex
                        Case BackgrIndex
114                         PixelFits = 0
                            Exit Function
116                     Case ScrRam(ScrBank, ScrNum, CX, CY, 1)
118                         PixelFits = 2
                            Exit Function
120                     Case D800(CX, CY)
122                         PixelFits = 3
                            Exit Function
124                     Case ScrRam(ScrBank, ScrNum, CX, CY, 0)
126                         PixelFits = 1
                            Exit Function
128                     Case Else

130                         BP00 = 0
132                         BP01 = 0
134                         BP10 = 0
136                         BP11 = 0

                            'count d800/d021 pixels (11/00)/ char block
138                         For Qbank = 0 To BmpBanks
140                             Ls = CY * 8
142                             Le = (CY * 8) + 7
144                             For YY = Ls To Le
146                                 BP11 = BP11 + BitCount11(Bitmap(Qbank, CX, YY))
                                    'Bp00 = Bp00 + BitCount00(Bitmap(Qbank, Cx, Yy))
148                             Next YY
150                         Next Qbank


152                         If BmpBanks = 1 And ScrBanks = 0 Then
                                'ha 2 BmpBank es 1 scrbank: mindket bitmapban nezzuk h befer-e
154                             For Qbank = 0 To 1
156                                 Ls = Int(Y / FliMul) * FliMul
158                                 Le = ((Int(Y / FliMul) + 1) * FliMul) - 1
160                                 For YY = Ls To Le
162                                     BP01 = BP01 + BitCount01(Bitmap(Qbank, CX, YY))
164                                     BP10 = BP10 + BitCount10(Bitmap(Qbank, CX, YY))
166                                 Next YY
168                             Next Qbank
                            Else
                                'ha 2 vagy 1 BmpBank es 2 vagy 1 scrbank: rajzolando bitmapban nezzuk h befer-e
170                             Ls = Int(Y / FliMul) * FliMul
172                             Le = ((Int(Y / FliMul) + 1) * FliMul) - 1
174                             For YY = Ls To Le
176                                 BP01 = BP01 + BitCount01(Bitmap(BmpBank, CX, YY))
178                                 BP10 = BP10 + BitCount10(Bitmap(BmpBank, CX, YY))
180                             Next YY
                            End If


182                         If BP01 = 0 Or (BP01 = 1 And Col2Change = ScrRam(ScrBank, ScrNum, CX, CY, 0)) Then
184                             PixelFits = 1
                                Exit Function
186                         ElseIf BP10 = 0 Or (BP10 = 1 And Col2Change = ScrRam(ScrBank, ScrNum, CX, CY, 1)) Then
188                             PixelFits = 2
                                Exit Function
190                         ElseIf BP11 = 0 Or (BP11 = 1 And Col2Change = D800(CX, CY)) Then
192                             PixelFits = 3
                                Exit Function
                            End If

                    End Select
                
                End If

                '******************* hires********************************

194             If BaseMode = BaseModeTyp.hires Then




196                 Select Case ColorIndex
                        Case ScrRam(ScrBank, ScrNum, CX, CY, 0)
198                         PixelFits = 0
                            Exit Function
200                     Case ScrRam(ScrBank, ScrNum, CX, CY, 1)
202                         PixelFits = 1
                            Exit Function
204                     Case Else

206                         bp1 = 0
208                         Bp0 = 0

210                         Ls = Int(Y / FliMul) * FliMul
212                         Le = ((Int(Y / FliMul) + 1) * FliMul) - 1
214                         For YY = Ls To Le
216                             bp1 = bp1 + BitCount1(Bitmap(BmpBank, CX, YY))
218                             Bp0 = Bp0 + BitCount0(Bitmap(BmpBank, CX, YY))
220                         Next YY

222                         If Bp0 = 0 Or (Bp0 = 1 And Col2Change = ScrRam(ScrBank, ScrNum, CX, CY, 0)) Then
224                             PixelFits = 0
226                         ElseIf bp1 = 0 Or (bp1 = 1 And Col2Change = ScrRam(ScrBank, ScrNum, CX, CY, 1)) Then
228                             PixelFits = 1
                            End If

                        'Debug.Print "bp0: "; Bp0; "bp1"; bp1; " pixfits"; PixelFits
                    
                    End Select
                End If

                '******************* unrestricted ********************************

230             If BaseMode = BaseModeTyp.unrestricted Then
232                 PixelFits = 0
                End If
        End Select


        '<EhFooter>
        Exit Function

PixelFits_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.PixelFits" + " line: " + Str(Erl))

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
Public Sub PlotBrush(ByVal pbx As Long, ByVal pby As Long)
        '<EhHeader>
        On Error GoTo PlotBrush_Err
        '</EhHeader>
    Dim px As Long
    Dim py As Long
    Dim pcol As Long
    Dim Qx As Long
    Dim Qy As Long

100 For px = 0 To 31 Step ResoDiv
102     For py = 0 To 31

104         If BrushArray(px, py) = 1 Then

106             Qx = px + pbx - 16
108             Qy = py + pby - 16

110             If Qx >= 0 And Qx <= PW - 1 And _
                   Qy >= 0 And Qy <= PH - 1 Then

112                 If BrushDither < Bayer((Qx \ ResoDiv) And 3, Qy And 3) Then pcol = BrushColor1 Else pcol = BrushColor2
114                 Call SetPixel(Qx, Qy, pcol)

                End If

            End If

116     Next py
118 Next px

        '<EhFooter>
        Exit Sub

PlotBrush_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.PlotBrush" + " line: " + Str(Erl))

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

Public Sub BrushPreCalc()
        '<EhHeader>
        On Error GoTo BrushPreCalc_Err
        '</EhHeader>
    Dim X As Single
    Dim Y As Single
    Dim D As Single
    Dim Round As Single


100 For X = 0 To 31
102     For Y = 0 To 31
104         BrushArray(X, Y) = 0
106     Next Y
108 Next X

110 Round = ((BrushSize \ 2) / ResoDiv) * ResoDiv

112 For X = -Round To Round Step ResoDiv
114     For Y = -Round To Round Step 1

116         D = Sqr(X * X + Y * Y)

118         If D < Int((BrushSize / 2) + 0.5) Then

120             BrushArray((31 \ 2) + X, (31 \ 2) + Y) = 1

122             If ResoDiv = 2 Then _
                   BrushArray((31 \ 2) + X + 1, (31 \ 2) + Y) = 1

            End If

124     Next Y
126 Next X


        '<EhFooter>
        Exit Sub

BrushPreCalc_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.BrushPreCalc" + " line: " + Str(Erl))

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


'CSEH: ErrMsgBox
Public Sub DrawPicFromMem()
        '<EhHeader>
        On Error GoTo DrawPicFromMem_Err
        '</EhHeader>

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


100 Select Case ChipType
        Case vicii

102         Select Case BaseMode
                Case BaseModeTyp.multi
            
104                 andtab(3) = 3 * 1
106                 andtab(2) = 3 * 4
108                 andtab(1) = 3 * 16
110                 andtab(0) = 3 * 64

112                 divtab(3) = 1
114                 divtab(2) = 4
116                 divtab(1) = 16
118                 divtab(0) = 64

120                 For Zy = 0 To CH - 1
122                     For Zx = 0 To CW - 1
124                         Qx = Zx * 8: Qy = Zy * 8
126                         For Y = 0 To 7
128                             For X = 0 To 7 Step ResoDiv
130                                 BmpX = Int(X / 2)
132                                 Call AccesSetup(X, Y)
134                                 bmp = Bitmap(BmpBank, Zx, Qy + Y)
136                                 Temp = (bmp And andtab(BmpX)) / divtab(BmpX)
                    
138                                     Select Case Temp
                                        Case 0
140                                         ColNum = BackgrIndex
142                                     Case 1
144                                         ColNum = ScrRam(ScrBank, ScrNum, Zx, Zy, 0)
146                                     Case 2
148                                         ColNum = ScrRam(ScrBank, ScrNum, Zx, Zy, 1)
150                                     Case 3
152                                         ColNum = D800(Zx, Zy)
                                        End Select

154                                 Pixels(Qx + X, Qy + Y) = ColNum
156                                 If ResoDiv = 2 Then Pixels(Qx + X + 1, Qy + Y) = ColNum
                                
158                             Next X
160                         Next Y
162                     Next Zx
164                 Next Zy

166             Case BaseModeTyp.hires

168                 andtab(7) = 1
170                 andtab(6) = 2
172                 andtab(5) = 4
174                 andtab(4) = 8
176                 andtab(3) = 16
178                 andtab(2) = 32
180                 andtab(1) = 64
182                 andtab(0) = 128

184                 divtab(7) = 1
186                 divtab(6) = 2
188                 divtab(5) = 4
190                 divtab(4) = 8
192                 divtab(3) = 16
194                 divtab(2) = 32
196                 divtab(1) = 64
198                 divtab(0) = 128

200                 BmpBank = 0
202                 ScrBank = 0
    
204                 For Zy = 0 To CH - 1
206                     For Zx = 0 To CW - 1
208                         Qx = Zx * 8: Qy = Zy * 8
210                         For Y = 0 To 7
212                             For X = 0 To 7
214                                 ScrNum = Int(((Qy + Y) And 7) / FliMul)
216                                 bmp = Bitmap(BmpBank, Zx, Qy + Y)
218                                 Temp = Int(bmp / divtab(X)) And 1
220                                 Select Case Temp
                                    Case 0
222                                     ColNum = ScrRam(ScrBank, ScrNum, Zx, Zy, 0)
224                                 Case 1
226                                     ColNum = ScrRam(ScrBank, ScrNum, Zx, Zy, 1)
                                    End Select

228                                 Pixels(Qx + X, Qy + Y) = ColNum

230                             Next X

232                         Next Y
234                     Next Zx
236                 Next Zy

            End Select

    End Select

        '<EhFooter>
        Exit Sub

DrawPicFromMem_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.DrawPicFromMem" + " line: " + Str(Erl))

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

Public Sub SetupBitTables()
        '<EhHeader>
        On Error GoTo SetupBitTables_Err
        '</EhHeader>
    Dim bp1 As Long
    Dim X As Long
    Dim Y As Long
    Dim Z As Long

        'generating basic hires bitmask table
100     bp1 = 1
102     For X = 0 To 7
104         Mask1(X) = bp1
106         InvMask0(X) = 255 - bp1
108         bp1 = bp1 * 2
            'Debug.Print Mask1(X); X
110     Next X

        'generating basic multi bitmask table
112     BP00 = 0
114     BP01 = 1
116     BP10 = 2
118     BP11 = 3

120     For X = 0 To 3
122         Mask11(0, X) = BP00
124         Mask11(1, X) = BP01
126         Mask11(2, X) = BP10
128         Mask11(3, X) = BP11
130         InvMask00(X) = 255 - BP11
132         BP00 = BP00 * 4
134         BP01 = BP01 * 4
136         BP10 = BP10 * 4
138         BP11 = BP11 * 4
140     Next X


        'generating bit / bitpair count tables
142     For X = 0 To 255

144         BitCount00(X) = 0
146         BitCount01(X) = 0
148         BitCount10(X) = 0
150         BitCount11(X) = 0

152         BitCount1(X) = 0
154         BitCount0(X) = 0

156         Y = X
158         For Z = 0 To 3
160             If (Y And 3) = 0 Then BitCount00(X) = BitCount00(X) + 1
162             If (Y And 3) = 1 Then BitCount01(X) = BitCount01(X) + 1
164             If (Y And 3) = 2 Then BitCount10(X) = BitCount10(X) + 1
166             If (Y And 3) = 3 Then BitCount11(X) = BitCount11(X) + 1
168             Y = Int(Y / 4)
170         Next Z

172         Y = X
174         For Z = 0 To 7
176             If (Y And 1) = 0 Then BitCount0(X) = BitCount0(X) + 1
178             If (Y And 1) = 1 Then BitCount1(X) = BitCount1(X) + 1
180             Y = Int(Y / 2)
182         Next Z


184     Next X

186     For X = 0 To 3
188         For Y = 0 To 3
190             For Z = 0 To 255
192                 If Z And InvMask00(0) = X * 1 Then BitSwap(X, Y, Z) = (Z And InvMask00(0)) + Y * 1
194                 If Z And InvMask00(1) = X * 4 Then BitSwap(X, Y, Z) = (Z And InvMask00(1)) + Y * 4
196                 If Z And InvMask00(2) = X * 16 Then BitSwap(X, Y, Z) = (Z And InvMask00(2)) + Y * 16
198                 If Z And InvMask00(3) = X * 64 Then BitSwap(X, Y, Z) = (Z And InvMask00(3)) + Y * 64
200             Next Z
202         Next Y
204     Next X
        '<EhFooter>
        Exit Sub

SetupBitTables_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.SetupBitTables" + " line: " + Str(Erl))

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


Public Sub backgrchange(backcol)
        '<EhHeader>
        On Error GoTo backgrchange_Err
        '</EhHeader>

100 Call Drawing.DrawPicFromMem


102 Call PrevWin.ReDraw
104 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
106 Call ZoomWindow.ZoomWinRefresh

        '<EhFooter>
        Exit Sub

backgrchange_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.backgrchange" + " line: " + Str(Erl))

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

Public Sub ChangeCol(ChangeTo As Byte)
    'change2=replace color
        '<EhHeader>
        On Error GoTo ChangeCol_Err
        '</EhHeader>
    Dim ChangeFrom As Long
    Dim Starty As Long
    Dim Endy As Long
    Dim Mask As Long
    Dim X As Long
    Dim Y As Long

100 ChangeFrom = Pixels(Ax, Ay)

102 If BaseMode = BaseModeTyp.multi Then

104     If (ChangeFrom <> ChangeTo And _
            ChangeFrom <> BackgrIndex) Then

106         Call Drawing.AccesSetup(Ax, Ay)

108         If ScrRam(ScrBank, ScrNum, CX, CY, 0) = ChangeFrom Then
110             Mask = 1
112         ElseIf ScrRam(ScrBank, ScrNum, CX, CY, 1) = ChangeFrom Then
114             Mask = 2
116         ElseIf D800(CX, CY) = ChangeFrom Then
118             Mask = 3
120         ElseIf BackgrIndex = ChangeFrom Then
122             Mask = 0
            End If
        
124         If Mask <> 3 Then
        
126             Starty = Int(Ay / FliMul) * FliMul
128             Endy = ((Int(Ay / FliMul) + 1) * FliMul) - 1

130             For X = CX * 8 To ((CX + 1) * 8) - 1 Step 1 'resodiv
132                For Y = Starty To Endy
134                   If ChangeFrom = Pixels(X, Y) Then
136                         Pixels(X, Y) = ChangeTo
138                         Call PutBitmap(X, Y, Mask, ChangeTo And 15)
                      End If
140                 Next Y
142             Next X
            
            Else
            
                'MsgBox "here"
144             Starty = Int(Ay / 8) * 8
146             Endy = ((Int(Ay / 8) + 1) * 8) - 1

148             For X = CX * 8 To ((CX + 1) * 8) - 1 Step 1 'resodiv
150                For Y = Starty To Endy
152                   If ChangeFrom = Pixels(X, Y) Then
154                         Pixels(X, Y) = ChangeTo
156                         Call PutBitmap(X, Y, Mask, ChangeTo And 15)
                      End If
158                 Next Y
160             Next X
            
            End If

        End If
    End If

162 If BaseMode = BaseModeTyp.hires Then

164     If ChangeFrom <> ChangeTo Then

166         Call Drawing.AccesSetup(Ax, Ay)
            'If ScrBanks = 1 Then ScrBank = Ax And 1 Else ScrBank = 0
            'CX = Int(Ax / 8)
            'CY = Int(Ay / 8)
            'ScrNum = Int((Ay And 7) / FliMul)

168         Mask = 3

170         If ScrRam(ScrBank, ScrNum, CX, CY, 0) = ChangeTo Then
172             Mask = 0
174         ElseIf ScrRam(ScrBank, ScrNum, CX, CY, 1) = ChangeTo Then
176             Mask = 1
            End If

178         If Mask = 3 Then
180             If ScrRam(ScrBank, ScrNum, CX, CY, 0) = ChangeFrom Then
182                 Mask = 0
184             ElseIf ScrRam(ScrBank, ScrNum, CX, CY, 1) = ChangeFrom Then
186                 Mask = 1
                End If
            End If

188         Starty = Int(Ay / FliMul) * FliMul
190         Endy = ((Int(Ay / FliMul) + 1) * FliMul) - 1

192         For X = CX * 8 To ((CX + 1) * 8) - 1 Step 1        'resodiv
194             For Y = Starty To Endy
196                 If ChangeFrom = Pixels(X, Y) Then
198                     Pixels(X, Y) = ChangeTo
200                     Call PutBitmap(X, Y, Mask, ChangeTo And 15)
                    End If
202             Next Y
204         Next X

        End If

    End If

206 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

        '<EhFooter>
        Exit Sub

ChangeCol_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Drawing.ChangeCol" + " line: " + Str(Erl))

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


