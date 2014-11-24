Attribute VB_Name = "Convert"
Option Explicit

Public ColMap() As Byte        ', pixels(319, 199) As Byte

Public UsedCols() As Byte
Public Bd800() As Byte
Public BScrRam() As Byte
Public BBitmap() As Byte

Dim Cols(255) As Long

'converter shit
Dim MostFreqCol As Long
Dim Max As Long
Dim Z As Long
Dim Add As Long
Dim Final As Long
Dim Sx As Long
Dim Sy As Long
Dim Err As Long

Dim Y As Long
Dim X As Long

Dim cr As Long
Dim cg As Long
Dim cb As Long



'this is only called from convertdialog
Public Sub ConvertPic()
        '<EhHeader>
        On Error GoTo ConvertPic_Err
        '</EhHeader>

100 Call Color_Filter
102 Call Color_Restrict
104 Call InitAttribs

        '<EhFooter>
        Exit Sub

ConvertPic_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.ConvertPic" + " line: " + Str(Erl))

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

Public Sub Color_Filter()
        '<EhHeader>
        On Error GoTo Color_Filter_Err
        '</EhHeader>

        Dim catch As Long
        Dim dist As Long
        Dim Index As Single
        Dim R As Long, G As Long, B As Long
        Dim Ys As Single
        Dim Us As Single
        Dim Vs As Single
        Dim coY(255) As Single
        Dim coU(255) As Single
        Dim cov(255) As Single
        Dim Chromas(8, 1) As Long
        Dim coH(255) As Single
        Dim coS(255) As Single
        Dim Col(255) As Single

        'to read in Mirage's table
        Dim FileNm As String
        Dim File_Contents As String
        Dim File_Lines As Variant
        Dim Line_Fields As Variant
        Dim iFile As Integer
        Dim Rw As Integer
        Dim Cl As Integer
        Dim HBtable(640 + 32) As Long
        Dim HBgray(20) As Long

        'Variables for Mirage's method
        Dim H As Single
        Dim s As Single
        Dim l As Single
        Dim Max As Single
        Dim Min As Single
        Dim delta As Single
        Dim rR As Single
        Dim rG As Single
        Dim rB As Single
        Dim Hindex As Long
        Dim lIndex As Long
        Dim sIndex As Long
    
        Dim Hmod As Long
    
        Dim Ladd As Integer
        Dim Hadd As Integer
        Dim Sadd As Integer
        Dim L2 As Integer
        Dim H2 As Integer

        Dim Bayer2x2(1, 1) As Long
        Dim Bayer4x4(3, 3) As Long
        Dim Bayer4x4Odd(3, 3) As Long
        Dim Bayer4x4Even(3, 3) As Long
        Dim Bayer4x4Spotty(3, 3) As Long

        Dim z2 As Long

        Dim Color As Long
    
100     Chromas(0, 0) = 0: Chromas(0, 1) = 0
102     Chromas(1, 0) = 6: Chromas(1, 1) = 9
104     Chromas(2, 0) = 2: Chromas(2, 1) = 11
106     Chromas(3, 0) = 4: Chromas(3, 1) = 8
108     Chromas(4, 0) = 12: Chromas(4, 1) = 14
110     Chromas(5, 0) = 5: Chromas(5, 1) = 10
112     Chromas(6, 0) = 3: Chromas(6, 1) = 15
114     Chromas(7, 0) = 7: Chromas(7, 1) = 13
116     Chromas(8, 0) = 1: Chromas(8, 1) = 1

118     Bayer4x4Even(0, 0) = 1: Bayer4x4Even(0, 1) = 9: Bayer4x4Even(0, 2) = 4: Bayer4x4Even(0, 3) = 12
120     Bayer4x4Even(1, 0) = 5: Bayer4x4Even(1, 1) = 13: Bayer4x4Even(1, 2) = 6: Bayer4x4Even(1, 3) = 14
122     Bayer4x4Even(2, 0) = 3: Bayer4x4Even(2, 1) = 11: Bayer4x4Even(2, 2) = 2: Bayer4x4Even(2, 3) = 10
124     Bayer4x4Even(3, 0) = 7: Bayer4x4Even(3, 1) = 15: Bayer4x4Even(3, 2) = 8: Bayer4x4Even(3, 3) = 16

126     Bayer4x4Odd(0, 0) = 1: Bayer4x4Odd(0, 1) = 2: Bayer4x4Odd(0, 2) = 3: Bayer4x4Odd(0, 3) = 4
128     Bayer4x4Odd(1, 0) = 9: Bayer4x4Odd(1, 1) = 10: Bayer4x4Odd(1, 2) = 11: Bayer4x4Odd(1, 3) = 12
130     Bayer4x4Odd(2, 0) = 5: Bayer4x4Odd(2, 1) = 6: Bayer4x4Odd(2, 2) = 7: Bayer4x4Odd(2, 3) = 8
132     Bayer4x4Odd(3, 0) = 13: Bayer4x4Odd(3, 1) = 14: Bayer4x4Odd(3, 2) = 15: Bayer4x4Odd(3, 3) = 16

134     Bayer4x4Spotty(0, 0) = 10: Bayer4x4Spotty(0, 1) = 1: Bayer4x4Spotty(0, 2) = 12: Bayer4x4Spotty(0, 3) = 6
136     Bayer4x4Spotty(1, 0) = 4: Bayer4x4Spotty(1, 1) = 9: Bayer4x4Spotty(1, 2) = 3: Bayer4x4Spotty(1, 3) = 15
138     Bayer4x4Spotty(2, 0) = 14: Bayer4x4Spotty(2, 1) = 2: Bayer4x4Spotty(2, 2) = 13: Bayer4x4Spotty(2, 3) = 7
140     Bayer4x4Spotty(3, 0) = 8: Bayer4x4Spotty(3, 1) = 11: Bayer4x4Spotty(3, 2) = 5: Bayer4x4Spotty(3, 3) = 16

142     Bayer4x4(0, 0) = 1: Bayer4x4(0, 1) = 13: Bayer4x4(0, 2) = 4: Bayer4x4(0, 3) = 16
144     Bayer4x4(1, 0) = 9: Bayer4x4(1, 1) = 5: Bayer4x4(1, 2) = 12: Bayer4x4(1, 3) = 7
146     Bayer4x4(2, 0) = 3: Bayer4x4(2, 1) = 15: Bayer4x4(2, 2) = 2: Bayer4x4(2, 3) = 14
148     Bayer4x4(3, 0) = 11: Bayer4x4(3, 1) = 8: Bayer4x4(3, 2) = 10: Bayer4x4(3, 3) = 8

        'this was going from 1 to 4
        Bayer2x2(0, 0) = 4: Bayer2x2(0, 1) = 1:
        Bayer2x2(1, 0) = 2: Bayer2x2(1, 1) = 3:

150     For Z = 0 To pal_Count

            'set up c64 r g b arrays
152         coR(Z) = PaletteRGB(Z) And 255
154         coG(Z) = (PaletteRGB(Z) \ 256) And 255
156         coB(Z) = (PaletteRGB(Z) \ 65536) And 255

            'set up c64 y u v arrays
158         coY(Z) = 0.299 * coR(Z) + 0.587 * coG(Z) + 0.114 * coB(Z)
160         coU(Z) = -0.147 * coR(Z) - 0.289 * coG(Z) + 0.436 * coB(Z)
162         cov(Z) = 0.615 * coR(Z) - 0.515 * coG(Z) - 0.1 * coB(Z)

164         RGBToHSL coR(Z), coG(Z), coB(Z), coH(Z), coS(Z), Col(Z)

166         coH(Z) = coH(Z) * (255 / 6)
168         coS(Z) = coS(Z) * 255
170         Col(Z) = Col(Z) * 255

172         Cols(Z) = 0

174     Next Z

        'load conversion table for color sensitive mode
176     FileNm = App.Path & "\" & ColorTable_Filename(ChipType)

178     If FileExists(FileNm) Then

180         iFile = FreeFile

182         Open FileNm For Input As iFile
184         File_Contents = Input$(LOF(iFile), iFile)
186         Close iFile

188         File_Lines = Split(File_Contents, vbCrLf)
190         For Rw = 0 To UBound(File_Lines)
                ' Process this line.
192             Line_Fields = Split(File_Lines(Rw), ",")
194             For Cl = 0 To UBound(Line_Fields)
196                 If Cl = 32 Then
198                     HBgray(Rw) = Val(Line_Fields(Cl))
                    Else
200                     HBtable(Rw * 32 + Cl) = Val(Line_Fields(Cl))
                    End If
202             Next Cl
204         Next Rw

        Else

206         MsgBox ("Can't open " & FileNm + " please set up the correct filename at the 'Color Tables' tab. Please notice that the file must be in the application dir.")
            Exit Sub
        End If

    
208     For Y = 0 To PH - 1
210         For X = 0 To PW - 1 Step ResoDiv

212             If Reso_PixRight = True And ResoDiv = 2 Then
214                 Color = ConvertDialog.SrcPic.Point(X + 1, Y)
216                 R = Color And 255
218                 G = (Color \ 256) And 255
220                 B = (Color \ 65536) And 255
                End If

222             If Reso_PixAvg = True And ResoDiv = 2 Then
224                 Color = ConvertDialog.SrcPic.Point(X, Y)
226                 R = Color And 255
228                 G = (Color \ 256) And 255
230                 B = (Color \ 65536) And 255

232                 Color = ConvertDialog.SrcPic.Point(X + 1, Y)
234                 cr = Color And 255
236                 cg = (Color \ 256) And 255
238                 cb = (Color \ 65536) And 255

240                 R = (cr + R) \ 2
242                 G = (cg + G) \ 2
244                 B = (cb + B) \ 2
                End If

246             If Reso_PixLeft = True Or ResoDiv = 1 Then
248                 Color = ConvertDialog.SrcPic.Point(X, Y)
250                 R = Color And 255
252                 G = (Color \ 256) And 255
254                 B = (Color \ 65536) And 255

                End If

                '################################################### OWN TAB& MIRAGE

256             If ColorFilterMode = 0 Then

                    'RGB -->> HSL
258                 rR = R / 256: rG = G / 256: rB = B / 256

260                 If rR > rG Then
262                     If rR > rB Then Max = rR Else Max = rB
                    Else
264                     If rB > rG Then Max = rB Else Max = rG
                    End If

266                 If rR < rG Then
268                     If rR < rB Then Min = rR Else Min = rB
                    Else
270                     If rB < rG Then Min = rB Else Min = rG
                    End If

272                 l = ((Max + Min) + 0.214 * rB + 0.587 * rG + 0.199 * rR) / 3

274                 s = 0: H = 6

276                 If Max <> Min Then

278                     If l <= 0.5 Then
280                         s = (Max - Min) / (Max + Min)
                        Else
282                         s = (Max - Min) / (2 - Max - Min)
                        End If

284                     delta = Max - Min
286                     If rR = Max Then
288                         H = 0 + (rG - rB) / delta
290                     ElseIf rG = Max Then
292                         H = 2 + (rB - rR) / delta
294                     ElseIf rB = Max Then
296                         H = 4 + (rR - rG) / delta
                        End If

                    End If

                    'H = H + 0.3
298                 If H > 6 Then
300                     H = H - 6
302                 ElseIf H < 0 Then
304                     H = H + 6
                    End If
                    'end of rgb -->> hsl



306                 lIndex = Int((l * 8))

308                 If lIndex < 0 Then lIndex = 0
310                 L2 = ((l * 8) - Int(lIndex)) * 16

312                 Select Case BditherVal
                        Case 1
314                         L2 = L2 And 14
316                         L2 = L2 + Rnd * BRnd
318                     Case 2
320                         L2 = L2 And 12
322                         L2 = L2 + Rnd * 3 * BRnd
324                     Case 3
326                         L2 = L2 And 8
328                         L2 = L2 + Rnd * 7 * BRnd
                    End Select

330                 Ladd = 0
332                 Hadd = 0

334                 Select Case Bdither
                        Case 1
336                         L2 = L2 / 4
338                         If Bayer2x2((X / ResoDiv) And 1, Y And 1) > L2 Then Ladd = 0 Else Ladd = 1
340                     Case 2
342                         If Bayer4x4((X / ResoDiv) And 3, Y And 3) > L2 Then Ladd = 0 Else Ladd = 1
344                     Case 3
346                         If Bayer4x4Odd((X / ResoDiv) And 3, Y And 3) > L2 Then Ladd = 0 Else Ladd = 1
348                     Case 4
350                         If Bayer4x4Even((X / ResoDiv) And 3, Y And 3) > L2 Then Ladd = 0 Else Ladd = 1
352                     Case 5
354                         If Bayer4x4Spotty((X / ResoDiv) And 3, Y And 3) > L2 Then Ladd = 0 Else Ladd = 1
                    End Select

356                 Select Case ChipType
                
                        Case Chip.vicii
                    
                            'h2 0-16
                        
358                         Hindex = Int((H / 6) * 7.75) * 4
360                         H2 = (((H / 6) * 7.75) - Int(Hindex / 4)) * 16
362                         Hmod = 2
                        
364                     Case Chip.ted
                        
                            'Hindex = Int(H / 6) * 31
                            'H2 = (((H / 6) * 31) - Hindex) * 16
                            'Hmod = 1

366                         Hindex = Int((H / 6) * 7.75) * 4
368                         H2 = (((H / 6) * 7.75) - Int(Hindex / 4)) * 4
370                         Hmod = 4

                    End Select
                
372                 Select Case Hditherval
                        Case 1
374                         H2 = H2 And 14
376                         H2 = H2 + Rnd * HRnd
378                     Case 2
380                         H2 = H2 And 12
382                         H2 = H2 + Rnd * HRnd * 3
384                     Case 3
386                         H2 = H2 And 8
388                         H2 = H2 + Rnd * HRnd * 7
                    End Select

390                 Select Case Hdither
                        Case 1
392                         H2 = H2 / 4
394                         If Bayer2x2((X / ResoDiv) And 1, Y And 1) > H2 Then Hadd = 0 Else Hadd = Hmod
396                     Case 2
398                         If Bayer4x4((X / ResoDiv) And 3, Y And 3) > H2 Then Hadd = 0 Else Hadd = Hmod
400                     Case 3
402                         If Bayer4x4Odd((X / ResoDiv) And 3, Y And 3) > H2 Then Hadd = 0 Else Hadd = Hmod
404                     Case 4
406                         If Bayer4x4Even((X / ResoDiv) And 3, Y And 3) > H2 Then Hadd = 0 Else Hadd = Hmod
408                     Case 5
410                         If Bayer4x4Spotty((X / ResoDiv) And 3, Y And 3) > H2 Then Hadd = 0 Else Hadd = Hmod
                    End Select

412                 lIndex = lIndex + Ladd: If lIndex > 19 Then lIndex = 19
414                 Hindex = Hindex + Hadd: If Hindex > 31 Then Hindex = Hindex - 31

    

                              
416                 L2 = s * 32
                
418                 Select Case SditherVal
                        Case 1
420                         L2 = L2 And 14
422                         L2 = L2 + Rnd * SRnd
424                     Case 2
426                         L2 = L2 And 12
428                         L2 = L2 + Rnd * SRnd * 3
430                     Case 3
432                         L2 = L2 And 8
434                         L2 = L2 + Rnd * SRnd * 7
                    End Select

436                 Select Case Sdither
                        Case 1
438                         L2 = L2 / 4
440                         If Bayer2x2((X / ResoDiv) And 1, Y And 1) > L2 Then Sadd = 0 Else Sadd = 8
442                     Case 2
444                         If Bayer4x4((X / ResoDiv) And 3, Y And 3) > L2 Then Sadd = 0 Else Sadd = 8
446                     Case 3
448                         If Bayer4x4Odd((X / ResoDiv) And 3, Y And 3) > L2 Then Sadd = 0 Else Sadd = 8
450                     Case 4
452                         If Bayer4x4Even((X / ResoDiv) And 3, Y And 3) > L2 Then Sadd = 0 Else Sadd = 8
454                     Case 5
456                         If Bayer4x4Spotty((X / ResoDiv) And 3, Y And 3) > L2 Then Sadd = 0 Else Sadd = 8
                    End Select

458                 s = s * 16
460                 s = s + Sadd

462                 If s >= 8 Then  '0.2 for c64
464                     Final = HBtable(lIndex * 32 + Hindex)
                    Else
466                     Final = HBgray(lIndex)
                    End If


                End If


                '############################################################## no tables

468             If ColorFilterMode = 1 Then

470                 Ys = 0.249 * R + 0.587 * G + 0.164 * B        'Ys = 0.299 * r + 0.587 * g + 0.114 * b
472                 Us = -0.147 * R - 0.289 * G + 0.436 * B
474                 Vs = 0.615 * R - 0.515 * G - 0.1 * B

476                 RGBToHSL R, G, B, H, s, l

478                 Ys = (Ys - 128) * 1.2
480                 Ys = Ys + 128        '- 16
482                 If Ys < 0 Then Ys = 0
484                 If Ys > 256 Then Ys = 256


486                 Ladd = Int(Ys / 2) And 14
488                 Index = Int(Ys / (256 / 8))

490                 Select Case Bdither
                        Case 1
492                         Ladd = Ladd / 4
494                         If Bayer2x2((X / ResoDiv) And 1, Y And 1) < Ladd Then Index = Index + 1
496                     Case 2
498                         If Bayer4x4((X / ResoDiv) And 3, Y And 3) < Ladd Then Index = Index + 1
500                     Case 3
502                         If Bayer4x4Odd((X / ResoDiv) And 3, Y And 3) < Ladd Then Index = Index + 1
504                     Case 4
506                         If Bayer4x4Even((X / ResoDiv) And 3, Y And 3) < Ladd Then Index = Index + 1
508                     Case 5
510                         If Bayer4x4Spotty((X / ResoDiv) And 3, Y And 3) < Ladd Then Index = Index + 1
                    End Select

512                 If Index > 8 Then Index = 8

514                 RGBToHSL R, G, B, H, s, l

516                 Ladd = s * 32
518                 Ladd = Ladd And 15

520                 Sadd = 0
522                 Select Case Sdither
                        Case 1
524                         Ladd = Ladd / 4
526                         If Bayer2x2((X / ResoDiv) And 1, Y And 1) > Ladd Then Sadd = 0.7
528                     Case 2
530                         If Bayer4x4((X / ResoDiv) And 3, Y And 3) > Ladd Then Sadd = 0.7
532                     Case 3
534                         If Bayer4x4Odd((X / ResoDiv) And 3, Y And 3) > Ladd Then Sadd = 0.7
536                     Case 4
538                         If Bayer4x4Even((X / ResoDiv) And 3, Y And 3) > Ladd Then Sadd = 0.7
540                     Case 5
542                         If Bayer4x4Spotty((X / ResoDiv) And 3, Y And 3) > Ladd Then Sadd = 0.7
                    End Select

544                 s = s + Sadd


546                 If s > 0.45 Then        's>0.45
548                     catch = 2147483647
550                     For Z = 0 To 1
552                         dist = (Us - coU(Chromas(Index, Z))) * (Us - coU(Chromas(Index, Z))) + (Vs - cov(Chromas(Index, Z))) * (Vs - cov(Chromas(Index, Z)))
554                         If dist < catch Then catch = dist: Final = Chromas(Index, Z): z2 = Z
556                     Next Z

                    Else

                        'Ladd = Int(Ys / 2) And 15
                        'Index = Int(Ys / (256 / 8))
558                     l = Index

560                     If Bdither = True Then
562                         If Bayer4x4((X / ResoDiv) And 3, Y And 3) > Ladd Then l = l + 0 Else l = l + 1
                        End If


564                     If l = 0 Then Final = 0
566                     If l = 1 Then Final = 0
568                     If l = 2 Then Final = 11
570                     If l = 3 Then Final = 11
572                     If l = 4 Then Final = 12
574                     If l = 5 Then Final = 12
576                     If l = 6 Then Final = 15
578                     If l = 7 Then Final = 1
580                     If l = 8 Then Final = 1

                    End If

                End If

                '############################################# YUV DISTANCE

582             If ColorFilterMode = 3 Then


584                 Ys = 0.299 * R + 0.587 * G + 0.114 * B
586                 Us = (-0.147 * R - 0.289 * G + 0.436 * B)
588                 Vs = (0.615 * R - 0.515 * G - 0.1 * B)

590                 catch = 2147483647

592                 For Z = 0 To pal_Count
594                     dist = ((coY(Z) - Ys) * (coY(Z) - Ys) + (coU(Z) - Us) * (coU(Z) - Us) + (cov(Z) - Vs) * (cov(Z) - Vs))
596                     If (dist < catch) Then catch = dist: Index = Z
598                 Next Z
600                 Final = Index

                End If

                '####################################### RGB DISTANCE
602             If ColorFilterMode = 2 Then

604                 catch = 2147483647

606                 For Z = 0 To pal_Count
608                     dist = ((coR(Z) - R) * (coR(Z) - R) + (coG(Z) - G) * (coG(Z) - G) + (coB(Z) - B) * (coB(Z) - B))
610                     If (dist < catch) Then catch = dist: Index = Z
612                 Next Z
614                 Final = Index

                End If

                '######################### brightness ladder!
616             If ColorFilterMode = 4 Then


618                 Ys = 0.299 * R + 0.587 * G + 0.114 * B

620                 Index = Ys / (255 / (BrLadderMax - 1))

622                 Ladd = Int(((Index - Int(Index))) * 15)
624                 Index = Int(Index)

626                 Select Case BditherVal
                        Case 1
628                         Ladd = Ladd And 14
630                         Ladd = Ladd + Rnd * BRnd
632                     Case 2
634                         Ladd = Ladd And 12
636                         Ladd = Ladd + Rnd * 3 * BRnd
638                     Case 3
640                         Ladd = Ladd And 8
642                         Ladd = Ladd + Rnd * 7 * BRnd
                    End Select

644                 Ladd = Ladd And 15

646                 Select Case Bdither
                        Case 1
648                         Ladd = (Ladd / 4)
650                         If Bayer2x2((X / ResoDiv) And 1, Y And 1) <= Ladd Then Index = Index + 1
652                     Case 2
654                         If Bayer4x4((X / ResoDiv) And 3, Y And 3) <= Ladd Then Index = Index + 1
656                     Case 3
658                         If Bayer4x4Odd((X / ResoDiv) And 3, Y And 3) <= Ladd Then Index = Index + 1
660                     Case 4
662                         If Bayer4x4Even((X / ResoDiv) And 3, Y And 3) <= Ladd Then Index = Index + 1
664                     Case 5
666                         If Bayer4x4Spotty((X / ResoDiv) And 3, Y And 3) <= Ladd Then Index = Index + 1
                    End Select

668                 If Index > BrLadderMax - 1 Then Index = BrLadderMax - 1
670                 Final = BrLadderTab(Index)
                End If



                '###########################################################


672             Cols(Final) = Cols(Final) + 1

674             ColMap(X, Y) = Final
676             Pixels(X, Y) = Final

678             If ResoDiv = 2 Then
680                 ColMap(X + 1, Y) = Final
682                 Pixels(X + 1, Y) = Final
                End If

684         Next X



686         If Y Mod 8 = 0 Then
688             Call ConvertDialog.ReDrawDstPicHor(Y)
690             ZoomWindow.PrevPic.Refresh
            End If

692     Next Y

694     Max = -1: MostFreqCol = -1
696     For Z = 0 To 15
698         If Cols(Z) > Max Then
700             Max = Cols(Z)
702             MostFreqCol = Z
            End If
704     Next Z


        '<EhFooter>
        Exit Sub

Color_Filter_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.Color_Filter" + " line: " + Str(Erl))

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

Public Sub Color_Restrict()
        '<EhHeader>
        On Error GoTo Color_Restrict_Err
        '</EhHeader>

100     Select Case ChipType

            Case Chip.vicii
           
102             If BaseMode = BaseModeTyp.hires Then Call Hires_Attrib
104             If BaseMode = BaseModeTyp.multi Then Call Multicol_Attrib
            
106         Case Chip.ted
        
                'If GfxMode <> "unrestricted" Then Call TEDMulti_Attrib
            
108             If BaseMode = BaseModeTyp.hires Then Call TEDHires_Attrib
110             If BaseMode = BaseModeTyp.multi Then Call TedMulti_Attrib(0)
            
112         Case Chip.vdc
            
114             Call VDC_Attrib
            
        End Select

        '<EhFooter>
        Exit Sub

Color_Restrict_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.Color_Restrict" + " line: " + Str(Erl))

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

Public Sub Hires_Attrib()
        '<EhHeader>
        On Error GoTo Hires_Attrib_Err
        '</EhHeader>
    Dim inchar(16) As Integer
    Dim charerr As Double
    Dim Qx As Long
    Dim Qy As Long
    Dim Index As Long
    Dim maxerr As Long
    Dim maxi As Long
    Dim c0 As Long
    Dim c1 As Long
    Dim pcol As Long
    Dim X As Long
    Dim Y As Long
    Dim U As Long
    Dim c(3) As Long
    Dim Bank As Long

100 Bank = 0

102 For X = 0 To PW - 1 Step 8
104     For Y = 0 To PH - 1 Step FliMul

106         For Z = 0 To 15: Cols(Z) = -1: inchar(Z) = BackgrIndex: Next Z

108         For Qx = X To X + 7
110             For Qy = Y To Y + FliMul - 1
112                 Cols(ColMap(Qx, Qy)) = Cols(ColMap(Qx, Qy)) + 1
114             Next Qy
116         Next Qx


118         If Mostfrq3 = True Then
120             For U = 0 To 1
122                 Max = -1
124                 For Z = 0 To 15
126                     If Cols(Z) > Max Then Max = Cols(Z): Index = Z
128                 Next Z
130                 c(U) = Index: Cols(Index) = -1
132             Next U
            End If

134         If OptChars = True Then
136             Max = 0
138             For Z = 0 To 15
140                 If Cols(Z) <> -1 Then inchar(Max) = Z: Max = Max + 1
142             Next Z

144             maxerr = 2147483647
146             maxi = Max

148             For c0 = 0 To maxi - 1        '0400 #1
150                 For c1 = c0 + 1 To maxi        '0400 #2
152                     charerr = 0

154                     c(0) = inchar(c0)
156                     c(1) = inchar(c1)

158                     For Qx = X To X + 7
160                         For Qy = Y To Y + FliMul - 1

162                             pcol = ColMap(Qx, Qy)

164                             If (pcol <> c(0) And pcol <> c(1)) Then
166                                 Max = 65536: Index = -1
168                                 For Z = 0 To 1
170                                     If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z))
172                                 Next Z
                                    'err = err + Max
174                                 charerr = charerr + Max
                                End If

176                         Next Qy
178                     Next Qx

180                     If charerr < maxerr Then

182                         ScrNum = Int((Y And 7) / FliMul)
184                         ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 0) = inchar(c0)
186                         ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 1) = inchar(c1)

188                         maxerr = charerr
                        End If

190                 Next c1
192             Next c0

194             ScrNum = Int((Y And 7) / FliMul)
196             c(0) = ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 0)
198             c(1) = ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 1)

200             If maxi = 0 Or maxi = 1 Then c(0) = inchar(0)
            End If

202         For Qx = X To X + 7
204             For Qy = Y To Y + FliMul - 1

206                 pcol = ColMap(Qx, Qy) And 15

208                 If (pcol <> c(0) And pcol <> c(1)) Then
210                     Max = 6553666: Index = -1
212                     For Z = 0 To 1
214                         If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z)): Index = Z
216                     Next Z
218                     Err = Err + Max
220                     Final = c(Index)
                    Else
222                     Final = pcol
                    End If
224                 If X < XFliLimit Then Final = 0
226                 Pixels(Qx, Qy) = Final
228             Next Qy
230         Next Qx

232     Next Y

234     Call ConvertDialog.ReDrawDstPicVert(X)

236 Next X

    'Call InitAttribs
        '<EhFooter>
        Exit Sub

Hires_Attrib_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.Hires_Attrib" + " line: " + Str(Erl))

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

Public Sub TEDHires_Attrib()
        '<EhHeader>
        On Error GoTo TEDHires_Attrib_Err
        '</EhHeader>
    Dim inchar(255) As Long
    Dim charerr As Double
    Dim Qx As Long
    Dim Qy As Long
    Dim Index As Long
    Dim maxerr As Long
    Dim maxi As Long
    Dim c0 As Long
    Dim c1 As Long
    Dim pcol As Long
    Dim X As Long
    Dim Y As Long
    Dim U As Long
    Dim c(3) As Long
    Dim Bank As Long

100 Bank = 0

102 For X = 0 To PW - 1 Step 8
104     For Y = 0 To PH - 1 Step FliMul

106         For Z = 0 To 127: Cols(Z) = -1: inchar(Z) = BackgrIndex: Next Z

108         For Qx = X To X + 7
110             For Qy = Y To Y + FliMul - 1
112                 Cols(ColMap(Qx, Qy)) = Cols(ColMap(Qx, Qy)) + 1
114             Next Qy
116         Next Qx


118         If Mostfrq3 = True Then
120             For U = 0 To 1
122                 Max = -1
124                 For Z = 0 To 127
126                     If Cols(Z) > Max Then Max = Cols(Z): Index = Z
128                 Next Z
130                 c(U) = Index: Cols(Index) = -1
132             Next U
            End If

134         If OptChars = True Then
136             Max = 0
138             For Z = 0 To 127
140                 If Cols(Z) <> -1 Then inchar(Max) = Z: Max = Max + 1
142             Next Z

144             maxerr = 2147483647
146             maxi = Max

148             For c0 = 0 To maxi - 1        '0400 #1
150                 For c1 = c0 + 1 To maxi        '0400 #2
152                     charerr = 0

154                     c(0) = inchar(c0)
156                     c(1) = inchar(c1)

158                     For Qx = X To X + 7
160                         For Qy = Y To Y + FliMul - 1

162                             pcol = ColMap(Qx, Qy)

164                             If (pcol <> c(0) And pcol <> c(1)) Then
166                                 Max = 65536: Index = -1
168                                 For Z = 0 To 1
170                                     If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z))
172                                 Next Z
                                    'err = err + Max
174                                 charerr = charerr + Max
                                End If

176                         Next Qy
178                     Next Qx

180                     If charerr < maxerr Then

182                         ScrNum = Int((Y And 7) / FliMul)
184                         ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 0) = inchar(c0)
186                         ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 1) = inchar(c1)

188                         maxerr = charerr
                        End If

190                 Next c1
192             Next c0

194             ScrNum = Int((Y And 7) / FliMul)
196             c(0) = ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 0)
198             c(1) = ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 1)

200             If maxi = 0 Or maxi = 1 Then c(0) = inchar(0)
            End If

202         For Qx = X To X + 7
204             For Qy = Y To Y + FliMul - 1

206                 pcol = ColMap(Qx, Qy)

208                 If (pcol <> c(0) And pcol <> c(1)) Then
210                     Max = 6553666: Index = -1
212                     For Z = 0 To 1
214                         If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z)): Index = Z
216                     Next Z
218                     Err = Err + Max
220                     Final = c(Index)
                    Else
222                     Final = pcol
                    End If
224                 If X < XFliLimit Then Final = 0
226                 Pixels(Qx, Qy) = Final
228             Next Qy
230         Next Qx

232     Next Y

        'Call PrevWin.ReDrawConvertedPicVert(x)
234     Call ConvertDialog.ReDrawDstPicVert(X)

236 Next X

    'Call InitAttribs
        '<EhFooter>
        Exit Sub

TEDHires_Attrib_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.TEDHires_Attrib" + " line: " + Str(Erl))

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

Public Sub VDC_Attrib()
        '<EhHeader>
        On Error GoTo VDC_Attrib_Err
        '</EhHeader>
    Dim inchar(16) As Integer
    Dim charerr As Double
    Dim Qx As Long
    Dim Qy As Long
    Dim Index As Long
    Dim maxerr As Long
    Dim maxi As Long
    Dim c0 As Long
    Dim c1 As Long
    'VBCH Dim c2 As Long
    Dim pcol As Long
    Dim X As Long
    Dim Y As Long
    Dim U As Long
    Dim c(3) As Long
    Dim Bank As Long

100 Bank = 0

102 For X = 0 To PW - 1 Step 8
104     For Y = 0 To PH - 1 Step CellHeigth

106         For Z = 0 To 15: Cols(Z) = -1: inchar(Z) = BackgrIndex: Next Z

108         For Qx = X To X + 7
110             For Qy = Y To Y + CellHeigth - 1
112                 Cols(ColMap(Qx, Qy)) = Cols(ColMap(Qx, Qy)) + 1
114             Next Qy
116         Next Qx


118         If Mostfrq3 = True Then
120             For U = 0 To 1
122                 Max = -1
124                 For Z = 0 To 15
126                     If Cols(Z) > Max Then Max = Cols(Z): Index = Z
128                 Next Z
130                 c(U) = Index: Cols(Index) = -1
132             Next U
            End If

134         If OptChars = True Then
136             Max = 0
138             For Z = 0 To 15
140                 If Cols(Z) <> -1 Then inchar(Max) = Z: Max = Max + 1
142             Next Z

144             maxerr = 2147483647
146             maxi = Max

148             For c0 = 0 To maxi - 1        '0400 #1
150                 For c1 = c0 + 1 To maxi        '0400 #2
152                     charerr = 0

154                     c(0) = inchar(c0)
156                     c(1) = inchar(c1)

158                     For Qx = X To X + 7
160                         For Qy = Y To Y + CellHeigth - 1

162                             pcol = ColMap(Qx, Qy)

164                             If (pcol <> c(0) And pcol <> c(1)) Then
166                                 Max = 65536: Index = -1
168                                 For Z = 0 To 1
170                                     If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z))
172                                 Next Z
                                    'err = err + Max
174                                 charerr = charerr + Max
                                End If

176                         Next Qy
178                     Next Qx

180                     If charerr < maxerr Then

182                         ScrNum = Y \ CellHeigth
184                         ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 0) = inchar(c0)
186                         ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 1) = inchar(c1)

188                         maxerr = charerr
                        End If

190                 Next c1
192             Next c0

194             ScrNum = Int((Y And 7) / FliMul)
196             c(0) = ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 0)
198             c(1) = ScrRam(Bank, ScrNum, Int(X / 8), Int(Y / 8), 1)

200             If maxi = 0 Or maxi = 1 Then c(0) = inchar(0)
            End If

202         For Qx = X To X + 7
204             For Qy = Y To Y + CellHeigth - 1

206                 pcol = ColMap(Qx, Qy)

208                 If (pcol <> c(0) And pcol <> c(1)) Then
210                     Max = 6553666: Index = -1
212                     For Z = 0 To 1
214                         If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z)): Index = Z
216                     Next Z
218                     Err = Err + Max
220                     Final = c(Index)
                    Else
222                     Final = pcol
                    End If
224                 If X < XFliLimit Then Final = 0
226                 Pixels(Qx, Qy) = Final
228             Next Qy
230         Next Qx

232     Next Y

        'Call PrevWin.ReDrawConvertedPicVert(x)
234     Call ConvertDialog.ReDrawDstPicVert(X)

236 Next X

    'Call InitAttribs
        '<EhFooter>
        Exit Sub

VDC_Attrib_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.VDC_Attrib" + " line: " + Str(Erl))

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

Private Sub Multicol_Attrib()
        '<EhHeader>
        On Error GoTo Multicol_Attrib_Err
        '</EhHeader>
        Dim tr As Long
        Dim try(15) As Long


100     If MostFreqBackgr = True Then BackgrIndex = MostFreqCol
102     If UsrDefBackgr = True Then BackgrIndex = UsrBackgr




104     If OptBackgr = True Then

106         For tr = 0 To 15
108             Err = 0
110             try(tr) = 0
112             ConvertDialog.Backgr.BackColor = PaletteRGB(tr)
114             Call Mc_Attrib(tr)
116             try(tr) = Err
118         Next tr

120         Max = 2147483647
122         For Z = 0 To 15
124             If try(Z) < Max Then
126                 BackgrIndex = Z
128                 Max = try(Z)
                End If
130         Next Z

        End If

 
132  BackgrIndex = BackgrIndex
134  Call Palett.UpdateColors
136  Call Mc_Attrib(BackgrIndex)
138  Call ConvertDialog.ReDrawDstPic
     'Call InitAttribs

140  ConvertDialog.Backgr.BackColor = PaletteRGB(BackgrIndex)
        '<EhFooter>
        Exit Sub

Multicol_Attrib_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.Multicol_Attrib" + " line: " + Str(Erl))

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

Private Sub TedMulti_Attrib(ByVal background As Long)
        '<EhHeader>
        On Error GoTo TedMulti_Attrib_Err
        '</EhHeader>
        Dim inchar(255) As Long
        Dim charerr As Double
        Dim Qx As Long
        Dim Qy As Long
        Dim Index As Long
        Dim MaxCharErr As Double
        Dim c0 As Long
        Dim CharY As Long
        Dim Xstep As Long
        Dim maxi As Long
        Dim FliLinErr As Long
        Dim c1 As Long
        Dim c2 As Long
        Dim Merr As Long
        Dim pcol As Long
        Dim BestC1 As Long
        Dim BestC2 As Long
    '    Dim Cols(255) As Long
        Dim X As Long
        Dim Y As Long

        'plus4
        Dim ColsPerLine() As Long
        Dim TempCols(127) As Long
        Dim Q As Long

        Dim TedBack1 As Byte
        Dim TedBack2 As Byte

        Dim TedFore1 As Byte
        Dim TedFore2 As Byte

        Dim c3 As Byte
        Dim c4 As Byte

        Dim Max2 As Long

        Dim MaxMerr As Long

        Dim BestFore1 As Long
        Dim BestFore2 As Long
        Dim BestBack1 As Long
        Dim BestBack2 As Long

        Dim LineError As Long

        Dim BoxErr As Long
        Dim MinErr As Long
        Dim MinBoxErr As Long

        Dim Color(3) As Long
        Dim BestCol(3) As Long

        Dim Index1 As Long
        Dim Index2 As Long

        Dim LoopEntered As Boolean
        Dim Bank As Long
    

        'colmap has the color filtered picture

        Dim BestCols() As Byte

100     ReDim BestCols(CW - 1, PH - 1, 3)
102     ReDim ColsPerLine(PH - 1, 128)

104     ReDim UsedCols(CW - 1, PH - 1, 128)

        'init colsperline
106     For Z = 0 To PH - 1
108         For Qx = 0 To 128
110             ColsPerLine(Z, Qx) = -1
112         Next Qx
114     Next Z

        'collect used cols / char into usedcols array, index 128 = nr of color indexes
116     For X = 0 To PW - 1 Step 8
118         For Y = 0 To PH - 1 Step FliMul

120             CX = Int(X / 8)
122             CY = Int(Y / FliMul)

124             For Z = 0 To 127
126                 Cols(Z) = -1
128             Next Z

130             For Qx = X To X + 7 Step ResoDiv
132                 For Qy = Y To Y + FliMul - 1
134                     Cols(ColMap(Qx, Qy)) = 1
136                 Next Qy
138             Next Qx

140             Index = 0
142             For Z = 0 To 127
144                 If Cols(Z) = 1 Then
146                     UsedCols(CX, CY, Index) = Z
148                     Index = Index + 1
                    End If
150             Next Z

152             UsedCols(CX, CY, 128) = Index

154         Next Y
156     Next X


158     If BmpBanks = 1 Then
160         Add = Bank
162         Xstep = 2
        Else
164         Add = 0
166         Xstep = ResoDiv
        End If

168     If (ScrBanks = 0 And BmpBanks = 1) Then
170         Add = 0
172         Xstep = ResoDiv
        End If


        'brute force check per line------------------------------------------------------------------------
174     For Y = 0 To PH - 1 Step FliMul
176         CY = Int(Y / FliMul)

            'For Y = 0 To PH - 1 Step flimul
178         For X = 0 To PW - 1 Step 8

180             CX = Int(X / 8)

                'go through color combos / fli box

182             maxi = UsedCols(CX, CY, 128) - 1
184             MinBoxErr = 2147483647
186             LoopEntered = False
188             For c0 = 0 To maxi
190                 For c1 = c0 + 1 To maxi
192                     For c2 = c1 + 1 To maxi
194                         For c3 = c2 + 1 To maxi

196                             LoopEntered = True
198                             Color(0) = UsedCols(CX, CY, c0)
200                             Color(1) = UsedCols(CX, CY, c1)
202                             Color(2) = UsedCols(CX, CY, c2)
204                             Color(3) = UsedCols(CX, CY, c3)

206                             BoxErr = 0

208                             For Qx = X + Add To X + 7 Step Xstep
210                                 For Qy = Y To Y + FliMul - 1

212                                     pcol = ColMap(Qx, Qy)

214                                     If (pcol <> Color(0) And pcol <> Color(1) _
                                                And pcol <> Color(2) And pcol <> Color(3)) Then

216                                         MinErr = 2147483647
218                                         For Z = 0 To 3
220                                             If cd(pcol, Color(Z)) < MinErr Then MinErr = cd(pcol, Color(Z))
222                                         Next Z
224                                         BoxErr = BoxErr + MinErr

                                        End If

226                                 Next Qy
228                             Next Qx


230                             If BoxErr < MinBoxErr Then
232                                 MinBoxErr = BoxErr
234                                 For Z = 0 To 3
236                                     BestCol(Z) = Color(Z)
238                                 Next Z
                                End If

240                         Next c3
242                     Next c2
244                 Next c1
246             Next c0

248             If LoopEntered = False Then
250                 For Z = 0 To UsedCols(CX, CY, 128)
252                     BestCol(Z) = UsedCols(CX, CY, Z)
254                 Next Z
                End If

256             If LoopEntered = True Then
258                 For Z = 0 To 3
260                     ColsPerLine(CY, BestCol(Z)) = ColsPerLine(CY, BestCol(Z)) + 1
262                 Next Z
                End If

264             For Z = 0 To 3
266                 BestCols(CX, CY, Z) = BestCol(Z)
268             Next Z

270             For Qx = X + Add To X + 7 Step Xstep
272                 For Qy = Y To Y + FliMul - 1

274                     pcol = ColMap(Qx, Qy)

276                     If (pcol <> BestCol(0) And pcol <> BestCol(1) _
                                And pcol <> BestCol(2) And pcol <> BestCol(3)) Then

278                         MinErr = 2147483647
280                         For Z = 0 To 3
282                             If cd(pcol, BestCol(Z)) < MinErr Then
284                                 Final = BestCol(Z)
286                                 MinErr = cd(pcol, BestCol(Z))
                                End If
288                         Next Z
                        
290                         Final = 0
                        Else

292                         Final = pcol

                        End If

                        'pixels(qx, qy) = Final
                        'If Resodiv = 2 Then pixels(qx + 1, qy) = Final

294                 Next Qy
296             Next Qx



298         Next X

300         Call ConvertDialog.ReDrawDstPicHor(Y)
302     Next Y




        'find most used colors on lines##########################xx
304     For Y = 0 To PH - 1 Step FliMul
306         CY = Int(Y / FliMul)

            Max = -1:
308         For Z = 0 To 127
310             If ColsPerLine(CY, Z) > Max Then
312                 Max = ColsPerLine(CY, Z)
314                 Index1 = Z
                End If
316         Next Z

318         ColsPerLine(CY, Index1) = -1

            Max = -1:
320         For Z = 0 To 127
322             If ColsPerLine(CY, Z) > Max Then
324                 Max = ColsPerLine(CY, Z)
326                 Index2 = Z
                End If
328         Next Z

330         ColsPerLine(CY, 0) = Index1
332         ColsPerLine(CY, 1) = Index2

334     Next Y


        ' SECOND PASS###########################################

336     For Y = 0 To PH - 1 Step FliMul
338         CY = Int(Y / FliMul)

340         For X = 0 To PW - 1 Step 8
342             CX = Int(X / 8)

344             MinBoxErr = 2147483647
346             For c0 = 0 To 3
348                 For c1 = c0 + 1 To 3


350                     Color(0) = BestCols(CX, CY, c0)
352                     Color(1) = BestCols(CX, CY, c1)
354                     Color(2) = ColsPerLine(CY, 0)
356                     Color(3) = ColsPerLine(CY, 1)

358                     BoxErr = 0
360                     For Qx = X + Add To X + 7 Step Xstep
362                         For Qy = Y To Y + FliMul - 1

364                             pcol = ColMap(Qx, Qy)

366                             If (pcol <> Color(0) And pcol <> Color(1) _
                                        And pcol <> Color(2) And pcol <> Color(3)) Then

368                                 MinErr = 2147483647
370                                 For Z = 0 To 3
372                                     If cd(pcol, Color(Z)) < MinErr Then MinErr = cd(pcol, Color(Z))
374                                 Next Z
376                                 BoxErr = BoxErr + MinErr

                                End If

378                         Next Qy
380                     Next Qx

382                     If BoxErr < MinBoxErr Then
384                         MinBoxErr = BoxErr
386                         For Z = 0 To 3
388                             BestCol(Z) = Color(Z)
390                         Next Z
                        End If

392                 Next c1
394             Next c0

396             For Qx = X + Add To X + 7 Step Xstep
398                 For Qy = Y To Y + FliMul - 1

400                     pcol = ColMap(Qx, Qy)

402                     If (pcol <> BestCol(0) And pcol <> BestCol(1) _
                                And pcol <> BestCol(2) And pcol <> BestCol(3)) Then

404                         MinErr = 2147483647
406                         For Z = 0 To 3
408                             If cd(pcol, BestCol(Z)) < MinErr Then
410                                 Final = BestCol(Z)
412                                 MinErr = cd(pcol, BestCol(Z))
                                End If
414                         Next Z

                        Else

416                         Final = pcol

                        End If

418                     Pixels(Qx, Qy) = Final
420                     If ResoDiv = 2 Then Pixels(Qx + 1, Qy) = Final

422                 Next Qy
424             Next Qx

426             For Z = 0 To 16
428                 Pixels(Z, Y) = BestCol(2)
430             Next Z

432             For Z = 16 To 32
434                 Pixels(Z, Y) = BestCol(3)
436             Next Z

438         Next X
440         Call ConvertDialog.ReDrawDstPicHor(Y)
442     Next Y
        '<EhFooter>
        Exit Sub

TedMulti_Attrib_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.TedMulti_Attrib" + " line: " + Str(Erl))

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
Private Sub Mc_Attrib(ByVal background As Long)
        '<EhHeader>
        On Error GoTo Mc_Attrib_Err
        '</EhHeader>
    Dim inchar(15) As Long
    Dim charerr As Double
    Dim Qx As Long
    Dim Qy As Long
    Dim Index As Long
    Dim MaxCharErr As Double
    Dim c0 As Long
    Dim CharY As Long
    Dim Xstep As Long
    Dim maxi As Long
    Dim FliLinErr As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim Merr As Long
    Dim pcol As Long
    Dim BestD800 As Long
    Dim BestC1 As Long
    Dim BestC2 As Long
    Dim Cols(255) As Long
    Dim X As Long
    Dim Y As Long
    Dim U As Long
    Dim c(3) As Long
    Dim Bank As Long


        'ReDim ColsPerLine(PH - 1, 128)
100     ReDim UsedCols(CW - 1, PH - 1, 128)


    'put used colors/char into usedcols(charx,chary,index) index 16 acts as colorcount
102 For X = 0 To PW - 1 Step 8
104     For Y = 0 To PH - 1 Step 8

106         CX = Int(X / 8)
108         CY = Int(Y / 8)

110         For Z = 0 To 15: Cols(Z) = -1: Next Z

112         Z = 0

114         For Qx = X To X + 7 Step ResoDiv
116             For Qy = Y To Y + 7
118                 Cols(ColMap(Qx, Qy)) = Cols(ColMap(Qx, Qy)) + 1
120             Next Qy
122         Next Qx
124         Cols(background) = -1

126         Index = 0
128         For Z = 0 To 15
130             If Cols(Z) <> -1 Then UsedCols(CX, CY, Index) = Z: Index = Index + 1
132         Next Z

134         UsedCols(CX, CY, 16) = Index

136     Next Y
138 Next X



140 BackgrIndex = background

142 For X = 0 To PW - 1 Step 8
144     For Y = 0 To PH - 1 Step 8

146         CX = Int(X / 8)
148         CY = Int(Y / 8)


150         If Mostfrq3 = True Then

152             For Z = 0 To 15: Cols(Z) = -1: inchar(Z) = background: Next Z
154             For Qx = X To X + 7 Step ResoDiv
156                 For Qy = Y To Y + 7 'flimul - 1
158                     Cols(ColMap(Qx, Qy)) = Cols(ColMap(Qx, Qy)) + 1
160                 Next Qy
162             Next Qx

164             Cols(background) = -1

166             For U = 0 To 2
168                 Max = -1
170                 For Z = 0 To 15
172                     If Cols(Z) > Max Then Max = Cols(Z): Index = Z
174                 Next Z
176                 c(U) = Index: Cols(Index) = -1
178             Next U
180             c(3) = background
182             BestD800 = c(0)
            End If




184         If OptChars = True Then

186             MaxCharErr = 2147483647

                'go through all colors to pick best d800
188             For c0 = 0 To UsedCols(CX, CY, 16) - 1

190                 charerr = 0        'error counter for character

                    'now we loop through the char, first through banks
192                 For Bank = 0 To BmpBanks

                        'then through the bytes
194                     For CharY = Y To Y + 7 Step FliMul

196                         If BmpBanks = 1 Then
198                             Add = Bank
200                             Xstep = 2
                            Else
202                             Add = 0
204                             Xstep = ResoDiv
                            End If

206                         If (ScrBanks = 0 And BmpBanks = 1) Then
208                             Add = 0
210                             Xstep = ResoDiv
                            End If

                            'reset arrays b4:check what colors are used in the current char
212                         For Z = 0 To 15: Cols(Z) = -1: inchar(Z) = background: Next Z

                            'check what colors are used in current fli box
214                         For Qx = X + Add To X + 7 Step Xstep
216                             For Qy = CharY To CharY + FliMul - 1
218                                 Cols(ColMap(Qx, Qy)) = Cols(ColMap(Qx, Qy)) + 1
220                             Next Qy
222                         Next Qx

224                         Cols(background) = -1        'set backgr as unused
226                         Cols(UsedCols(CX, CY, c0)) = -1        'set current d800 as unused

                            'compact used colors into inchar array
228                         Max = 0
230                         For Z = 0 To 15
232                             If Cols(Z) <> -1 Then inchar(Max) = Z: Max = Max + 1
234                         Next Z

                            'number of colors in fli box, doesnt counts backgr and d800
236                         maxi = Max

                            'set d800 and d021 colors for fli box

238                         c(0) = UsedCols(CX, CY, c0)
240                         c(3) = background

242                         FliLinErr = 2147483647

                            'go through possible color combinations for scrrams
244                         For c1 = 0 To maxi - 1        '0400 #1
246                             For c2 = c1 + 1 To maxi        '0400 #2

248                                 c(1) = inchar(c1)
250                                 c(2) = inchar(c2)

252                                 Merr = 0
254                                 For Qx = X + Add To X + 7 Step Xstep
256                                     For Qy = CharY To CharY + FliMul - 1

258                                         pcol = ColMap(Qx, Qy)
260                                         If EmulateFliBug = False Or Qx >= XFliLimit Then
262                                             If (pcol <> c(0) And pcol <> c(1) And pcol <> c(2) And pcol <> c(3)) Then
264                                                 Max = 2147483647
266                                                 For Z = 0 To 2 + FixWithBackgr
268                                                     If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z))
270                                                 Next Z
272                                                 Merr = Merr + Max
                                                End If
                                            Else
274                                             If (pcol <> c(0) And pcol <> c(3) And pcol <> 15) Then
276                                                 Max = 2147483647
278                                                 If cd(pcol, c(0)) < Max Then Max = cd(pcol, c(0))
280                                                 If cd(pcol, c(3)) < Max Then Max = cd(pcol, c(3))
282                                                 If cd(pcol, 15) < Max Then Max = cd(pcol, 15)
284                                                 Merr = Merr + Max
                                                End If
                                            End If
286                                     Next Qy
288                                 Next Qx

290                                 If Merr < FliLinErr Then FliLinErr = Merr


292                             Next c2
294                         Next c1

296                         If FliLinErr <> 2147483647 Then charerr = charerr + FliLinErr

298                     Next CharY

300                 Next Bank

302                 If charerr < MaxCharErr Then
304                     MaxCharErr = charerr
306                     BestD800 = UsedCols(CX, CY, c0)
                    End If

308             Next c0


            End If







310         charerr = 0

312         For Bank = 0 To BmpBanks

314             For CharY = Y To Y + 7 Step FliMul

316                 If BmpBanks = 1 Then
318                     Add = Bank
320                     Xstep = 2
                    Else
322                     Add = 0
324                     Xstep = ResoDiv
                    End If

326                 If (ScrBanks = 0 And BmpBanks = 1) Then
328                     Add = 0
330                     Xstep = ResoDiv
                    End If

                    'used colors in fli box
332                 For Z = 0 To 15: Cols(Z) = -1: inchar(Z) = background: Next Z

334                 For Qx = X + Add To X + 7 Step Xstep
336                     For Qy = CharY To CharY + FliMul - 1
338                         Cols(ColMap(Qx, Qy)) = Cols(ColMap(Qx, Qy)) + 1
340                     Next Qy
342                 Next Qx
344                 Cols(background) = -1
346                 Cols(BestD800) = -1

348                 Max = 0
350                 For Z = 0 To 15
352                     If Cols(Z) <> -1 Then inchar(Max) = Z: Max = Max + 1
354                 Next Z

                    'If cx = 21 And cy = 15 Then
                    'cx = cx
                    'End If

356                 maxi = Max

                    'color variations / fli box

358                 c(0) = BestD800
360                 c(3) = background
362                 BestC1 = background
364                 BestC2 = background

366                 FliLinErr = 2147483647

368                 For c1 = 0 To maxi - 1        '0400 #1
370                     For c2 = c1 + 1 To maxi        '0400 #2

372                         c(1) = inchar(c1)
374                         c(2) = inchar(c2)

376                         Merr = 0
378                         For Qx = X + Add To X + 7 Step Xstep
380                             For Qy = CharY To CharY + FliMul - 1

382                                 pcol = ColMap(Qx, Qy)
384                                 If EmulateFliBug = False Or X >= XFliLimit Then
386                                     If (pcol <> c(0) And pcol <> c(1) And pcol <> c(2) And pcol <> c(3)) Then
388                                         Max = 2147483647
390                                         For Z = 0 To 2 + FixWithBackgr
392                                             If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z))
394                                         Next Z
396                                         Merr = Merr + Max
                                        End If
                                    Else
                                    
398                                     If (pcol <> c(0) And pcol <> 15 And pcol <> c(3)) Then
400                                         Max = 2147483647

402                                         If cd(pcol, c(0)) < Max Then Max = cd(pcol, c(0))
404                                         If cd(pcol, c(3)) < Max Then Max = cd(pcol, c(3))
406                                         If cd(pcol, 15) < Max Then Max = cd(pcol, 15)

408                                         Merr = Merr + Max
                                        End If
                                    End If
410                             Next Qy
412                         Next Qx


414                         If Merr < FliLinErr Then
416                             FliLinErr = Merr
418                             BestC1 = c(1)
420                             BestC2 = c(2)
422                             If X < XFliLimit And EmulateFliBug = True Then
424                                 BestC1 = 15
426                                 BestC2 = 15
                                End If
                            End If

428                     Next c2
430                 Next c1

432                 c(0) = BestD800
434                 c(1) = BestC1
436                 c(2) = BestC2
438                 c(3) = background
                    'If maxi = 0 Then c(1) = inchar(1)

440                 Merr = 0
442                 For Qx = X + Add To X + 7 Step Xstep
444                     For Qy = CharY To CharY + FliMul - 1

446                         pcol = ColMap(Qx, Qy) And 15

448                         If EmulateFliBug = False Or X >= XFliLimit Then
450                             If (pcol <> c(0) And pcol <> c(1) And pcol <> c(2) And pcol <> c(3)) Then
452                                 Max = 2147483647
454                                 For Z = 0 To 2 + FixWithBackgr
456                                     If cd(pcol, c(Z)) < Max Then Max = cd(pcol, c(Z)): Index = Z
458                                 Next Z
460                                 Final = c(Index)
462                                 Merr = Merr + 1
                                Else
464                                 Final = pcol
                                End If
                            Else
466                             Max = 2147483647
468                             If cd(pcol, c(0)) < Max Then Max = cd(pcol, c(0)): Final = c(0)
470                             If cd(pcol, c(3)) < Max Then Max = cd(pcol, c(3)): Final = c(3)
472                             If cd(pcol, 15) < Max Then Max = cd(pcol, 15): Final = 15
474                             Merr = Merr + 1
                            End If

                       
476                         If EmulateFliBug = False Then
478                             Pixels(Qx, Qy) = Final
480                             If ResoDiv = 2 Then Pixels(Qx + 1, Qy) = Final
                            Else
482                             If X < XFliLimit Then Final = 0
484                             Pixels(Qx, Qy) = Final
486                             If ResoDiv = 2 Then Pixels(Qx + 1, Qy) = Final
                            End If

488                     Next Qy
490                 Next Qx

492                 Err = Err + Merr
494                 If ScrBanks = 1 Then ScrBank = Bank Else ScrBank = 0
496                 ScrNum = Int((CharY And 7) / FliMul)
498                 ScrRam(ScrBank, ScrNum, CX, CY, 0) = c(1)
500                 ScrRam(ScrBank, ScrNum, CX, CY, 1) = c(2)
502                 If EmulateFliBug = True And X < XFliLimit Then
504                     ScrRam(ScrBank, ScrNum, CX, CY, 0) = 15
506                     ScrRam(ScrBank, ScrNum, CX, CY, 1) = 15
                    End If
508                 D800(CX, CY) = BestD800

510             Next CharY



512         Next Bank



514     Next Y

516     Call ConvertDialog.ReDrawDstPicVert(X)

518 Next X
        '<EhFooter>
        Exit Sub

Mc_Attrib_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.Mc_Attrib" + " line: " + Str(Erl))

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

Public Sub InitAttribs()
        '<EhHeader>
        On Error GoTo InitAttribs_Err
        '</EhHeader>

100 Debug.Print "initattribs invoked"

102 Select Case ChipType

        Case Chip.vicii

104         For Sx = 0 To CW - 1
106             For Sy = 0 To CH - 1
108                 Call SetColAttrib
110             Next Sy
112         Next Sx

114     Case ted


116     Case vdc

    End Select

        '<EhFooter>
        Exit Sub

InitAttribs_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.InitAttribs" + " line: " + Str(Erl))

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


Public Sub SetColAttrib()
        '<EhHeader>
        On Error GoTo SetColAttrib_Err
        '</EhHeader>
    Dim Qx As Long
    Dim Qy As Long
    Dim Nib As Long

100 If BaseMode = BaseModeTyp.multi Then

    
102     For Qy = Sy * 8 To ((Sy + 1) * 8) - 1
104         For Qx = Sx * 8 To ((Sx + 1) * 8) - 1 Step ResoDiv
                                    
106             Call AccesSetup(Qx, Qy)
    
108             If Pixels(Qx, Qy) = BackgrIndex Then
110                 Call PutBitmap(Qx, Qy, 0, Pixels(Qx, Qy))
112             ElseIf Pixels(Qx, Qy) = ScrRam(ScrBank, ScrNum, Sx, Sy, 0) Then
114                 Call PutBitmap(Qx, Qy, 1, Pixels(Qx, Qy))
116             ElseIf Pixels(Qx, Qy) = ScrRam(ScrBank, ScrNum, Sx, Sy, 1) Then
118                 Call PutBitmap(Qx, Qy, 2, Pixels(Qx, Qy))
120             ElseIf Pixels(Qx, Qy) = D800(Sx, Sy) Then
122                 Call PutBitmap(Qx, Qy, 3, Pixels(Qx, Qy))
                End If
                       
            
124         Next Qx
126     Next Qy


    End If


128 If BaseMode = BaseModeTyp.hires Then

130     For Qx = Sx * 8 To ((Sx + 1) * 8) - 1 Step ResoDiv
132         For Qy = Sy * 8 To ((Sy + 1) * 8) - 1

134             Call AccesSetup(Qx, Qy)
136             If Pixels(Qx, Qy) = ScrRam(ScrBank, ScrNum, CX, CY, 1) Then
138                 Call PutBitmap(Qx, Qy, 1, Pixels(Qx, Qy))
140             ElseIf Pixels(Qx, Qy) = ScrRam(ScrBank, ScrNum, CX, CY, 0) Then
142                 Call PutBitmap(Qx, Qy, 0, Pixels(Qx, Qy))
                End If
            
144         Next Qy
146     Next Qx


    End If

        '<EhFooter>
        Exit Sub

SetColAttrib_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.SetColAttrib" + " line: " + Str(Erl))

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

Private Sub RGBToHSL( _
        ByVal R As Long, ByVal G As Long, ByVal B As Long, _
        H As Single, s As Single, l As Single _
                                  )
        '<EhHeader>
        On Error GoTo RGBToHSL_Err
        '</EhHeader>
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single, rG As Single, rB As Single

100 rR = R / 255: rG = G / 255: rB = B / 255

    '{Given: rgb each in [0,1].
    ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
102 Max = Maximum(rR, rG, rB)
104 Min = Minimum(rR, rG, rB)
106 l = (Max + Min) / 2        '{This is the lightness}
    '{Next calculate saturation}
108 If Max = Min Then
        'begin {Acrhomatic case}
110     s = 0
112     H = 0
        'end {Acrhomatic case}
    Else
        'begin {Chromatic case}
        '{First calculate the saturation.}
114     If l <= 0.5 Then
116         s = (Max - Min) / (Max + Min)
        Else
118         s = (Max - Min) / (2 - Max - Min)
        End If
        '{Next calculate the hue.}
120     delta = Max - Min
122     If rR = Max Then
124         H = (rG - rB) / delta        '{Resulting color is between yellow and magenta}
126     ElseIf rG = Max Then
128         H = 2 + (rB - rR) / delta        '{Resulting color is between cyan and yellow}
130     ElseIf rB = Max Then
132         H = 4 + (rR - rG) / delta        '{Resulting color is between magenta and cyan}
        End If
        'Debug.Print h
        'h = h * 60
        'If h < 0# Then
        '     h = h + 360            '{Make degrees be nonnegative}
        'End If
        'end {Chromatic Case}
    End If
    'end {RGB_to_HLS}
        '<EhFooter>
        Exit Sub

RGBToHSL_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.RGBToHSL" + " line: " + Str(Erl))

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

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
        '<EhHeader>
        On Error GoTo Maximum_Err
        '</EhHeader>
100 If (rR > rG) Then
102     If (rR > rB) Then
104         Maximum = rR
        Else
106         Maximum = rB
        End If
    Else
108     If (rB > rG) Then
110         Maximum = rB
        Else
112         Maximum = rG
        End If
    End If
        '<EhFooter>
        Exit Function

Maximum_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.Maximum" + " line: " + Str(Erl))

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
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
        '<EhHeader>
        On Error GoTo Minimum_Err
        '</EhHeader>
100 If (rR < rG) Then
102     If (rR < rB) Then
104         Minimum = rR
        Else
106         Minimum = rB
        End If
    Else
108     If (rB < rG) Then
110         Minimum = rB
        Else
112         Minimum = rG
        End If
    End If
        '<EhFooter>
        Exit Function

Minimum_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Convert.Minimum" + " line: " + Str(Erl))

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

