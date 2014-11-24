Attribute VB_Name = "IO"
Option Explicit

'new fileformat
'0-15 projectone - in ascii
'16 graphicmode
'bitmap(s)
'screen(s) (if any)
'd800 (if any)
'sprites (if any)
'registers (if any)
Public FileName As String
Dim X As Long
Dim Y As Long
Dim ColNum As Long


Public Function GetSaveName(Filter As String) As Boolean

'get save filename
ZoomWindow.CommonDialog1.CancelError = True
On Error GoTo errhandler
ZoomWindow.CommonDialog1.InitDir = LastSavePath
ZoomWindow.CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
ZoomWindow.CommonDialog1.Filter = Filter
ZoomWindow.CommonDialog1.FilterIndex = 1
ZoomWindow.CommonDialog1.ShowSave
FileName = ZoomWindow.CommonDialog1.FileName
LastSaveName = ZoomWindow.CommonDialog1.FileName
LastSavePath = ZoomWindow.CommonDialog1.InitDir
On Error GoTo 0

GetSaveName = True
Exit Function

errhandler:

GetSaveName = False
End Function
Public Function GetLoadName(Filter As String) As Boolean

'get save filename
ZoomWindow.CommonDialog1.CancelError = True
On Error GoTo errhandlersave
ZoomWindow.CommonDialog1.InitDir = LastLoadPath
ZoomWindow.CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
ZoomWindow.CommonDialog1.Filter = Filter
ZoomWindow.CommonDialog1.FilterIndex = LastLoadFilterIndex
ZoomWindow.CommonDialog1.FileName = LastLoadName
ZoomWindow.CommonDialog1.ShowOpen

FileName = ZoomWindow.CommonDialog1.FileName
LastLoadPath = ZoomWindow.CommonDialog1.InitDir
LastLoadName = ZoomWindow.CommonDialog1.FileName
LastLoadFilterIndex = ZoomWindow.CommonDialog1.FilterIndex
On Error GoTo 0

GetLoadName = True
Exit Function

errhandlersave:

GetLoadName = False

End Function

Public Sub CustomLoad()
        '<EhHeader>
        On Error GoTo CustomLoad_Err
        '</EhHeader>
    Dim CX As Long
    Dim CY As Long
    Dim Qx As Long
    Dim Qy As Long
    Dim Y As Long
    Dim bmp As Byte
    Dim filebuffer(66000) As Byte
    Dim ff As Long
    Dim Q As Long
    Dim offset As Long
    Dim LoadAdress As Long
    Dim Bank As Long
    Dim ScrNum As Long

    'On Error GoTo errorhandler

    'get load adress
100 ff = FreeFile
102 Open FileName For Binary As #ff
104 Get #ff, , bmp:  LoadAdress = bmp
106 Get #ff, , bmp:  LoadAdress = LoadAdress + (bmp * 256)
108 LoadAdress = LoadAdress And 65535
110 Close #ff

    'clear filebuffer
112 For Q = 0 To 66000: filebuffer(Q) = 0: Next

    'read whole file into filebuffer
114 ff = FreeFile
116 Open FileName For Binary As #ff

118 If CustomIOSetup.Absolute = False Then LoadAdress = 0
120 If CustomIOSetup.ForceStartAddressFromUser = True Then LoadAdress = CustomIOSetup.StartAdressUser
    'If adr_SkipFirstTwoBytes = True And adr_OverrideAdress = False Then LoadAdress = 0

    'skip first two bytes
122 If CustomIOSetup.SkipStartAdress = True Then
124     Get #ff, , bmp
126     Get #ff, , bmp
128     For Q = 0 To FileLen(FileName) - 2
130         Get #ff, , filebuffer(Q + LoadAdress)
132     Next Q
    Else
134     For Q = 0 To FileLen(FileName)
136         Get #ff, , filebuffer(Q + LoadAdress)
138     Next Q
    End If

    'read screens
140 For Bank = 0 To 1
142     For ScrNum = 0 To 7
144         offset = CustomIOSetup.Screen(ScrNum + Bank * 8) '+ LoadAdress
146         For CY = 0 To CH - 1
148             For CX = 0 To CW - 1
150                 ScrRam(Bank, ScrNum, CX, CY, 1) = filebuffer(offset) And 15
152                 ScrRam(Bank, ScrNum, CX, CY, 0) = Int(filebuffer(offset) / 16)
154                 offset = offset + 1
156             Next CX
158         Next CY
160     Next ScrNum
162 Next Bank

    'read bitmaps
164 For Bank = 0 To 1
166     offset = CustomIOSetup.Bitmap(Bank) '+ LoadAdress
168     For CY = 0 To CH - 1
170         For CX = 0 To CW - 1
172             Qx = CX * 8: Qy = CY * 8
174             For Y = 0 To 7
176                 Bitmap(Bank, CX, Qy + Y) = filebuffer(offset)
178                 offset = offset + 1
180             Next Y
182         Next CX
184     Next CY
186 Next Bank

    'read d800
188 offset = CustomIOSetup.D800 '+ LoadAdress

190 For CY = 0 To CH - 1
192     For CX = 0 To CW - 1
194         D800(CX, CY) = filebuffer(offset) And 15
196         offset = offset + 1
198     Next CX
200 Next CY


    'background
202 offset = CustomIOSetup.D021 '+ LoadAdress
204 BackgrIndex = filebuffer(offset) And 15

206 Close ff

    Exit Sub

ErrorHandler:

208 MsgBox "Error while trying to load c64 binary, doublecheck your settings."

        '<EhFooter>
        Exit Sub

CustomLoad_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IO.CustomLoad" + " line: " + Str(Erl))

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


Public Sub CustomSave()


Dim CX As Long
Dim CY As Long
Dim Qx As Long
Dim Qy As Long
Dim Y As Long
Dim bmp As Byte
Dim filebuffer(66000) As Byte
Dim ff As Long
Dim Q As Long
Dim offset As Long
Dim Bank As Long
Dim ScrNum As Long


On Error GoTo ErrorHandler


ff = FreeFile
If FileExists(FileName) Then
    Kill (FileName)
End If

Open FileName For Binary As #ff

'write load adress
If CustomIOSetup.HasStartAddress Then
    CustomIOSetup.Start = CustomIOSetup.Start And 65535
    bmp = CustomIOSetup.Start And 255
    Put #ff, , bmp
    bmp = CustomIOSetup.Start \ 256
    Put #ff, , bmp
End If

'clear filebuffer
For Q = 0 To 66000: filebuffer(Q) = 0: Next

'write screens to virtual mem
For Bank = 0 To 1
    For ScrNum = 0 To 7
        offset = CustomIOSetup.Screen(ScrNum + Bank * 8) '+ LoadAdress
        For CY = 0 To CH - 1
            For CX = 0 To CW - 1
                filebuffer(offset) = (ScrRam(Bank, ScrNum, CX, CY, 1) And 15) + (ScrRam(Bank, ScrNum, CX, CY, 0) And 15) * 16
                offset = offset + 1
            Next CX
        Next CY
    Next ScrNum
Next Bank

'write bitmaps
For Bank = 0 To 1
    offset = CustomIOSetup.Bitmap(Bank) '+ LoadAdress
    For CY = 0 To CH - 1
        For CX = 0 To CW - 1
            Qx = CX * 8: Qy = CY * 8
            For Y = 0 To 7
                filebuffer(offset) = Bitmap(Bank, CX, Qy + Y)
                offset = offset + 1
            Next Y
        Next CX
    Next CY
Next Bank

'write d800 if multicolor
If BaseMode_cm = multi Then

    offset = CustomIOSetup.D800 '+ LoadAdress

    For CY = 0 To CH - 1
        For CX = 0 To CW - 1
            filebuffer(offset) = D800(CX, CY) And 15
            offset = offset + 1
        Next CX
    Next CY

End If

'write background
offset = CustomIOSetup.D021 '+ LoadAdress
filebuffer(offset) = BackgrIndex And 15


For Q = CustomIOSetup.Start To CustomIOSetup.End
    Put #ff, , filebuffer(Q)
Next Q

Close ff

Exit Sub

ErrorHandler:

MsgBox "Error while trying to save c64 binary, doublecheck your settings."

End Sub


Public Sub SaveKoala()
        '<EhHeader>
        On Error GoTo SaveKoala_Err
        '</EhHeader>

    Dim bmp As Byte
    Dim ff As Long
    Dim Qy As Long



100 ff = FreeFile


102 If FileExists(FileName) Then
104     Kill (FileName)
106     Open FileName For Binary As #ff
    Else
108     Open FileName For Binary As #ff
    End If

110 Put #ff, , &H6000


112 ScrNum = 0

114 For CY = 0 To CH - 1
116     For CX = 0 To CW - 1

118         Qy = CY * 8
120         For Y = 0 To 7

122             bmp = Bitmap(0, CX, Qy + Y)
124             Put #ff, , bmp
126         Next Y

128     Next CX
130 Next CY

132 For Y = 0 To CH - 1
134     For X = 0 To CW - 1
136         bmp = (ScrRam(0, ScrNum, X, Y, 1) And 15) + (ScrRam(0, ScrNum, X, Y, 0) And 15) * 16
138         Put #ff, , bmp
140     Next X
142 Next Y

144 For Y = 0 To CH - 1
146     For X = 0 To CW - 1
148         bmp = D800(X, Y) And 15
150         Put #ff, , bmp
152     Next X
154 Next Y


156 bmp = BackgrIndex
158 Put #ff, , bmp

160 Close #ff

        '<EhFooter>
        Exit Sub

SaveKoala_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IO.SaveKoala" + " line: " + Str(Erl))

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

Public Sub LoadKoala()
        '<EhHeader>
        On Error GoTo LoadKoala_Err
        '</EhHeader>
    Dim Read As Byte
    Dim filebuffer(10000) As Byte
    Dim andtab(3), divtab(3)
    Dim bmp As Byte
    Dim ff As Long
    Dim Qy As Long
    Dim Qx As Long
    Dim Q As Long
    Dim Temp As Long

100 If GetLoadName("(*.kla)|*.kla|(*.*)|*.*") = False Then Exit Sub

102 ff = FreeFile

104 andtab(3) = 3 * 1
106 andtab(2) = 3 * 4
108 andtab(1) = 3 * 16
110 andtab(0) = 3 * 64

112 divtab(3) = 1
114 divtab(2) = 4
116 divtab(1) = 16
118 divtab(0) = 64

120 Open FileName For Binary As #ff

122 Get #ff, , Read
124 Get #ff, , Read


126 ScrNum = 0

128 For X = 0 To 10000: Get #ff, , filebuffer(X): Next

130 BackgrIndex = filebuffer(10000) And 15

132 Q = 0
134 For CY = 0 To CH - 1
136     For CX = 0 To CW - 1
            'g = Int(actcolor / 256) And 255
138         ScrRam(0, ScrNum, CX, CY, 0) = Int(filebuffer(8000 + Q) / 16) And 15
140         ScrRam(0, ScrNum, CX, CY, 1) = filebuffer(8000 + Q) And 15
142         Q = Q + 1
144     Next CX
146 Next CY

148 Q = 0
150 For CY = 0 To CH - 1
152     For CX = 0 To CW - 1
154         D800(CX, CY) = filebuffer(9000 + Q) And 15
156         Q = Q + 1
158     Next CX
160 Next CY

162 Q = 0
164 For CY = 0 To CH - 1
166     For CX = 0 To CW - 1
168         Qx = CX * 8: Qy = CY * 8
170         For Y = 0 To 7
172             For X = 0 To 3
174                 bmp = filebuffer(Q)
176                 Temp = (bmp And andtab(X)) / divtab(X)
178                 Select Case Temp
                    Case 0
180                     ColNum = BackgrIndex
182                 Case 1
184                     ColNum = ScrRam(0, ScrNum, CX, CY, 0)
186                 Case 2
188                     ColNum = ScrRam(0, ScrNum, CX, CY, 1)
190                 Case 3
192                     ColNum = D800(CX, CY)
                    End Select

194                 Pixels(Qx + X * 2, Qy + Y) = ColNum
196                 Pixels(Qx + X * 2 + 1, Qy + Y) = ColNum
198             Next X
200             Q = Q + 1
202         Next Y
204     Next CX
        'PrevPic.Refresh
206     Call PrevWin.ReDraw

208 Next CY

210 Close #ff

    'Text2.Text = "koala load end!"
212 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
    'Picture3(20).BackColor = PaletteRGB(BackgrIndex)
214 Call InitAttribs

216 BaseMode = BaseModeTyp.multi
218 GfxMode = "koala"
220 ResoDiv = 2
222 FliMul = 8
224 BmpBanks = 0
226 ScrBanks = 0
228 XFliLimit = 0
230 Call ZoomWindow.ResetUndo
232 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

234 Close #ff

        '<EhFooter>
        Exit Sub

LoadKoala_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IO.LoadKoala" + " line: " + Str(Erl))

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

Public Sub SaveDrazlace()
        '<EhHeader>
        On Error GoTo SaveDrazlace_Err
        '</EhHeader>
    Dim FileName As String
    'VBCH Dim ActAttrib(2) As Long
    Dim bmp As Byte
    Dim ff As Long
    Dim fr As Long
    Dim Qx As Long
    Dim Qy As Long
    Dim uu As Long

    '5800- d800
    '5c00- 0400
    '6000- bmp1
    'a000- bmp2
    '7f40- background

100 ff = FreeFile

102 If FileExists(FileName) Then
104     Kill (FileName)
106     Open FileName For Binary As #ff
    Else
108     Open FileName For Binary As #ff
    End If

110 Put #ff, , &H5800

    ' save d800

112 For Y = 0 To CH - 1
114     For X = 0 To CW - 1
116         bmp = D800(X, Y) And 15
118         Put #ff, , bmp
120     Next X
122 Next Y

    'pad till 1024
124 For Y = 0 To 23
126     Put #ff, , bmp
128 Next Y

    ' save 0400


130 ScrNum = 0
132 For Y = 0 To CH - 1
134     For X = 0 To CW - 1
136         bmp = (ScrRam(0, ScrNum, X, Y, 1) And 15) + (ScrRam(0, ScrNum, X, Y, 0) And 15) * 16
            'Bmp = 1 + 2 * 16
138         Put #ff, , bmp
140     Next X
142 Next Y

    'pad till 1024
144 For Y = 0 To 23
146     Put #ff, , bmp
148 Next Y

    'save bitmaps

150 For fr = 0 To 1

152     For CY = 0 To CH - 1
154         For CX = 0 To CW - 1

156             Qx = CX * 8: Qy = CY * 8
158             For Y = 0 To 7
160                 bmp = Bitmap(fr, CX, Qy + Y)
162                 Put #ff, , bmp
164             Next Y

166         Next CX
168     Next CY

        'skip gap till next bmp

170     If fr = 0 Then
172         bmp = BackgrIndex
174         Put #ff, , bmp        'store background color
176         bmp = 0
178         For uu = 0 To &HBE
180             Put #ff, , bmp
182         Next uu
        End If

184 Next fr

186 Close #ff


        '<EhFooter>
        Exit Sub

SaveDrazlace_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IO.SaveDrazlace" + " line: " + Str(Erl))

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

Public Sub LoadDrazlace()
        '<EhHeader>
        On Error GoTo LoadDrazlace_Err
        '</EhHeader>
    Dim Read As Byte
    Dim filebuffer(18240) As Byte
    Dim andtab(3), divtab(3)
    Dim bmp As Byte
    Dim ff As Long
    Dim Q As Long
    Dim fr As Long
    Dim Qx As Long
    Dim Qy As Long
    Dim Temp As Long

100 ff = FreeFile
102 If GetLoadName("(*.drl)|*.drl|(*.*)|*.*") = False Then Exit Sub

104 andtab(3) = 3 * 1
106 andtab(2) = 3 * 4
108 andtab(1) = 3 * 16
110 andtab(0) = 3 * 64

112 divtab(3) = 1
114 divtab(2) = 4
116 divtab(1) = 16
118 divtab(0) = 64


120 Open FileName For Binary As #ff

122 Get #ff, , Read
124 Get #ff, , Read

    '5800- d800
    '5c00- 0400
    '6000- bmp1
    'a000- bmp2
    '7f40- background

126 For X = 0 To 18240: Get #ff, , filebuffer(X): Next

128 BackgrIndex = filebuffer(&H2740) And 15

130 ScrNum = 0

    'read d800
132 Q = 0
134 For CY = 0 To CH - 1
136     For CX = 0 To CW - 1
138         D800(CX, CY) = filebuffer(Q) And 15
140         Q = Q + 1
142     Next CX
144 Next CY

    'read 0400
146 Q = 1024
148 For CY = 0 To CH - 1
150     For CX = 0 To CW - 1
152         ScrRam(0, ScrNum, CX, CY, 0) = Int(filebuffer(Q) / 16) And 15
154         ScrRam(0, ScrNum, CX, CY, 1) = filebuffer(Q) And 15
156         Q = Q + 1
158     Next CX
160 Next CY

162 Q = 2048
164 For fr = 0 To 1

166     For CY = 0 To CH - 1
168         For CX = 0 To CW - 1
170             Qx = CX * 8: Qy = CY * 8
172             For Y = 0 To 7
174                 For X = 0 To 3
176                     bmp = filebuffer(Q)
178                     Temp = (bmp And andtab(X)) / divtab(X)
180                     Select Case Temp
                        Case 0
182                         ColNum = BackgrIndex
184                     Case 1
186                         ColNum = ScrRam(0, ScrNum, CX, CY, 0)
188                     Case 2
190                         ColNum = ScrRam(0, ScrNum, CX, CY, 1)
192                     Case 3
194                         ColNum = D800(CX, CY)
                        End Select

196                     Pixels(Qx + X * 2 + fr, Qy + Y) = ColNum

198                 Next X
200                 Q = Q + 1
202             Next Y
204         Next CX
206         Call PrevWin.ReDraw

208     Next CY
210     Q = Q + &HC0
212 Next fr

214 Close #ff


216 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

218 Call InitAttribs
220 BaseMode = BaseModeTyp.multi
222 GfxMode = "drazlace"
224 ResoDiv = 1
226 FliMul = 8
228 BmpBanks = 1
230 ScrBanks = 0
232 XFliLimit = 0
234 Call ZoomWindow.ResetUndo
236 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

238 Close #ff
        '<EhFooter>
        Exit Sub

LoadDrazlace_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IO.LoadDrazlace" + " line: " + Str(Erl))

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

Public Sub SaveFli()
        '<EhHeader>
        On Error GoTo SaveFli_Err
        '</EhHeader>

    Dim FileName As String
    'VBCH Dim ActAttrib(2) As Long
    Dim byt As Byte
    Dim ff As Long
    Dim Qx As Long
    Dim Qy As Long
    'VBCH Dim bmp As Long

100 If GetSaveName("(*.fli)|*.fli|(*.*)|*.*") = False Then Exit Sub

102 ff = FreeFile


104 If FileExists(FileName) Then
106     Kill (FileName)
108     Open FileName For Binary As #ff
    Else
110     Open FileName For Binary As #ff
    End If


112 BmpBank = 0
114 ScrBank = 0

    'start addy
116 Put #ff, , &H3C00

    'save d800

118 For Y = 0 To CH - 1
120     For X = 0 To CW - 1
122         byt = D800(X, Y) And 15
124         Put #ff, , byt
126     Next X
128 Next Y

130 byt = 0
132 For CX = 0 To 23
134     Put #ff, , byt
136 Next CX

    'save screens

138 For ScrNum = 0 To 7
140     For Y = 0 To CH - 1
142         For X = 0 To CW - 1
144             byt = ScrRam(ScrBank, ScrNum, X, Y, 1) + ScrRam(ScrBank, ScrNum, X, Y, 0) * 16
146             Put #ff, , byt
148         Next X
150     Next Y

152     byt = 0
154     For CX = 0 To 23
156         Put #ff, , byt
158     Next CX

160 Next ScrNum

    'save bitmap

162 For CY = 0 To CH - 1
164     For CX = 0 To CW - 1

166         Qx = CX * 8: Qy = CY * 8
168         For Y = 0 To 7

170             byt = Bitmap(0, CX, Qy + Y)
172             Put #ff, , byt
174         Next Y

176     Next CX
178 Next CY

180 byt = BackgrIndex
182 Put #ff, , byt

184 byt = 0
186 For X = 0 To 191 - 2
188     Put #ff, , byt
190 Next X


192 Close #ff


        '<EhFooter>
        Exit Sub

SaveFli_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IO.SaveFli" + " line: " + Str(Erl))

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

Public Sub LoadFli()
        '<EhHeader>
        On Error GoTo LoadFli_Err
        '</EhHeader>

    Dim Read As Byte
    Dim filebuffer(17408) As Byte
    Dim andtab(3), divtab(3)
    Dim bmp As Byte
    Dim ff As Long
    Dim Qx As Long
    Dim Q As Long
    Dim Qy As Long

100 ff = FreeFile
102 If GetLoadName("(*.fli)|*.fli|(*.*)|*.*") = False Then Exit Sub

104 andtab(3) = 3 * 1
106 andtab(2) = 3 * 4
108 andtab(1) = 3 * 16
110 andtab(0) = 3 * 64

112 divtab(3) = 1
114 divtab(2) = 4
116 divtab(1) = 16
118 divtab(0) = 64

120 Open FileName For Binary As #ff

122 Get #ff, , Read
124 Get #ff, , Read

126 ScrBank = 0
128 BmpBank = 0

130 For X = 0 To 17408: Get #ff, , filebuffer(X): Next

132 BackgrIndex = filebuffer(&H4340) And 15        '???
    'MsgBox (Str(filebuffer(&H4342)))
134 Q = 0
136 For ScrNum = 0 To 7
138     For CY = 0 To CH - 1
140         For CX = 0 To CW - 1
142             ScrRam(ScrBank, ScrNum, CX, CY, 0) = Int(filebuffer(1024 + Q) / 16) And 15
144             ScrRam(ScrBank, ScrNum, CX, CY, 1) = filebuffer(1024 + Q) And 15
146             Q = Q + 1
148         Next CX
150     Next CY
152     Q = Q + 24
154 Next ScrNum

156 Q = 0
158 For CY = 0 To CH - 1
160     For CX = 0 To CW - 1
162         D800(CX, CY) = filebuffer(0 + Q) And 15
164         Q = Q + 1
166     Next CX
168 Next CY

170 Q = 1024 + (8 * 1024)
172 For CY = 0 To CH - 1
174     For CX = 0 To CW - 1
176         Qx = CX * 8: Qy = CY * 8
178         For Y = 0 To 7
                'For x = 0 To 3
180                 bmp = filebuffer(Q)
182                 Bitmap(BmpBank, CX, Qy + Y) = bmp
                    'temp = (bmp And andtab(x)) / divtab(x)
                    'Select Case temp
                    'Case 0
                    '    ColNum = BackgrIndex
                    'Case 1
                    '    ColNum = ScrRam(ScrBank, y, cx, cy, 0)
                    'Case 2
                    '    ColNum = ScrRam(ScrBank, y, cx, cy, 1)
                    'Case 3
                    '    ColNum = d800(cx, cy)
                    'End Select
                    'If cx < 3 Then ColNum = BackgrIndex
                    'pixels(qx + x * 2, qy + y) = ColNum
                    'pixels(qx + x * 2 + 1, qy + y) = ColNum
                'Next x
184             Q = Q + 1
186         Next Y
188     Next CX
        'PrevPic.Refresh
        'Call PrevWin.ReDrawDisp

190 Next CY

192 Close #ff

    'Text2.Text = "koala load end!"
194 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
196 Call Palett.UpdateColors


198 BaseMode = BaseModeTyp.multi
200 GfxMode = "fli"
202 ResoDiv = 2
204 FliMul = 1
206 BmpBanks = 0
208 ScrBanks = 0
210 XFliLimit = 24

212 Call DrawPicFromMem
214 Call InitAttribs
216 Call ZoomWindow.ResetUndo
218 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
220 Call ZoomWindow.ZoomWinRefresh

222 Close #ff

        '<EhFooter>
        Exit Sub

LoadFli_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.IO.LoadFli" + " line: " + Str(Erl))

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

