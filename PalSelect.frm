VERSION 5.00
Begin VB.Form PalSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Palette Setup"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3978
      TabIndex        =   13
      Top             =   2808
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selected Entry:"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1638
      Width           =   6135
      Begin VB.OptionButton chkDecInput 
         Caption         =   "Dec Input"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton chkHexInput 
         Caption         =   "Hex Input"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   5400
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSelectedColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   598
         Left            =   234
         TabIndex        =   17
         Top             =   234
         Width           =   1066
      End
      Begin VB.Label Label1 
         Caption         =   "B:"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "G:"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "R:"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton btnFinished 
      Caption         =   "Done"
      Height          =   375
      Left            =   5265
      TabIndex        =   5
      Top             =   2808
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Palette: "
      Height          =   1417
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      Begin Project_One.MyPalette MyPalette1 
         Height          =   1065
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   1879
         BoxCount        =   15
         BoxWidth        =   24
         BoxHSpacing     =   8
         BoxTop          =   4
         BoxLeft         =   6
         BoxPerLine      =   8
         BackgrColor     =   16711680
         BoxborderColor  =   16777215
         BoxHasBorders   =   -1  'True
         BoxTextColor    =   16777215
         BoxHasText      =   0
         BoxBorderStyle  =   2
      End
   End
   Begin VB.CommandButton SaveOverSelected 
      Caption         =   "Save"
      Height          =   375
      Left            =   2691
      TabIndex        =   3
      Top             =   2808
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   6360
      TabIndex        =   2
      Top             =   468
      Width           =   1695
   End
   Begin VB.CommandButton NewPreset 
      Caption         =   "New"
      Height          =   375
      Left            =   117
      TabIndex        =   1
      Top             =   2808
      Width           =   975
   End
   Begin VB.CommandButton DeleteSelected 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1404
      TabIndex        =   0
      Top             =   2808
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Presets:"
      Height          =   247
      Left            =   6435
      TabIndex        =   18
      Top             =   234
      Width           =   1417
   End
End
Attribute VB_Name = "PalSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldPalNum As Long
Option Explicit


Private Sub btnCancel_Click()
        '<EhHeader>
        On Error GoTo btnCancel_Click_Err
        '</EhHeader>

100 pal_Selected(ChipType) = OldPalNum
102 Call LoadPalette(pal_Selected(ChipType))
104 Call PaletteInit
106 Call ZoomWindow.ZoomResize
108 Call PrevWin.ReDraw
110 Call Palett.UpdateColors
112 Unload Me

        '<EhFooter>
        Exit Sub

btnCancel_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.btnCancel_Click" + " line: " + Str(Erl))

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

Private Sub btnFinished_Click()
        '<EhHeader>
        On Error GoTo btnFinished_Click_Err
        '</EhHeader>

100 pal_Selected(ChipType) = List1.ItemData(List1.ListIndex)
102 Call LoadPalette(pal_Selected(ChipType))
104 Call PaletteInit
106 Call ZoomWindow.ZoomResize
108 Call PrevWin.ReDraw
110 Call Palett.UpdateColors
112 Unload Me

        '<EhFooter>
        Exit Sub

btnFinished_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.btnFinished_Click" + " line: " + Str(Erl))

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


Private Sub chkDecInput_Click()
        '<EhHeader>
        On Error GoTo chkDecInput_Click_Err
        '</EhHeader>
100 pal_HexMode = False
102 RefreshRGBVals
        '<EhFooter>
        Exit Sub

chkDecInput_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.chkDecInput_Click" + " line: " + Str(Erl))

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

Private Sub chkHexInput_Click()
        '<EhHeader>
        On Error GoTo chkHexInput_Click_Err
        '</EhHeader>
100 pal_HexMode = True
102 RefreshRGBVals
        '<EhFooter>
        Exit Sub

chkHexInput_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.chkHexInput_Click" + " line: " + Str(Erl))

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

Private Sub DeleteSelected_Click()
        '<EhHeader>
        On Error GoTo DeleteSelected_Click_Err
        '</EhHeader>
100 Call DeletePreset
        '<EhFooter>
        Exit Sub

DeleteSelected_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.DeleteSelected_Click" + " line: " + Str(Erl))

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

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err
        '</EhHeader>

100 If UnloadMode = 0 Then
102     pal_Selected(ChipType) = OldPalNum
104     Call LoadPalette(pal_Selected(ChipType))
106     Call PaletteInit
108     Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
110     Call ZoomWindow.ZoomWinRefresh
112     Call PrevWin.ReDraw
114     Call Palett.UpdateColors
    End If

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.Form_QueryUnload" + " line: " + Str(Erl))

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





Private Sub RefreshRGBVals()
        '<EhHeader>
        On Error GoTo RefreshRGBVals_Err
        '</EhHeader>
        Dim Color As Long
        Dim R As Byte
        Dim G As Byte
        Dim B As Byte

    
100     Color = MyPalette1.PaletteRGB(pal_Lastindex)

102     R = Color And 255
104     G = (Color \ 256) And 255
106     B = (Color \ 65536) And 255

108     Select Case pal_HexMode
            Case True
110             Text1(0).Text = Hex(R)
112             Text1(1).Text = Hex(G)
114             Text1(2).Text = Hex(B)
116         Case False
118             Text1(0).Text = R
120             Text1(1).Text = G
122             Text1(2).Text = B
        End Select

        '<EhFooter>
        Exit Sub

RefreshRGBVals_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.RefreshRGBVals" + " line: " + Str(Erl))

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





Private Sub List1_Click()
        '<EhHeader>
        On Error GoTo List1_Click_Err
        '</EhHeader>

100 pal_Selected(ChipType) = List1.ItemData(List1.ListIndex)
102 Call LoadPalette(pal_Selected(ChipType))
104 Call PaletteInit
106 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
108 Call ZoomWindow.ZoomWinRefresh
110 Call PrevWin.ReDraw
112 Call Palett.UpdateColors
114 RedrawPalette
116 RefreshRGBVals


        '<EhFooter>
        Exit Sub

List1_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.List1_Click" + " line: " + Str(Erl))

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

100 Set Me.Icon = ZoomWindow.Image1.Picture

        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.form_load" + " line: " + Str(Erl))

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
Private Sub RedrawPalette()
        '<EhHeader>
        On Error GoTo RedrawPalette_Err
        '</EhHeader>
    Dim X As Long
100 For X = 0 To pal_Count
102     MyPalette1.PaletteRGB(X) = PaletteRGB(X)
104 Next X
106 MyPalette1.BoxCount = pal_Count
108 MyPalette1.InitSurface

110 lblSelectedColor.BackColor = PaletteRGB(pal_Lastindex)


112 Frame1.Caption = "Selected Entry: #" & Str(pal_Lastindex)
114 Frame2.Caption = "Palette: " & GetPalName(pal_Selected(ChipType))

        '<EhFooter>
        Exit Sub

RedrawPalette_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.RedrawPalette" + " line: " + Str(Erl))

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


Private Sub form_activate()
        '<EhHeader>
        On Error GoTo form_activate_Err
        '</EhHeader>

100 MyPalette1.BackgrColor = Me.BackColor

102 OldPalNum = pal_Selected(ChipType)



104 Call RefreshRGBVals
106 If pal_HexMode = True Then
108  chkHexInput.Value = True
    Else
110  chkDecInput.Value = True
    End If
112 GetPalList
    'List1.ListIndex = pal_Default
114 RedrawPalette
        '<EhFooter>
        Exit Sub

form_activate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.form_activate" + " line: " + Str(Erl))

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

Private Sub MyPalette1_BoxClicked(Button As Integer, Index As Byte)
        '<EhHeader>
        On Error GoTo MyPalette1_BoxClicked_Err
        '</EhHeader>

100 pal_Lastindex = Index
102 RedrawPalette
104 RefreshRGBVals
106 lblSelectedColor.BackColor = MyPalette1.PaletteRGB(Index)

        '<EhFooter>
        Exit Sub

MyPalette1_BoxClicked_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.MyPalette1_BoxClicked" + " line: " + Str(Erl))

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

Private Sub NewPreset_Click()
        '<EhHeader>
        On Error GoTo NewPreset_Click_Err
        '</EhHeader>
    Dim NewName As String

100 NewName = InputBox("Please enter a name for the new preset:")
102 If NewName = "" Then
104     MsgBox "You must enter a name.", vbExclamation
        Exit Sub
    End If

106 Call SavePreset(NewName)

108 If NewName <> "alreadyexisted" Then
110     List1.AddItem NewName
112     List1.ListIndex = List1.ListCount - 1
114     List1.ItemData(List1.ListIndex) = List1.ListIndex
    Else
    
    End If


        '<EhFooter>
        Exit Sub

NewPreset_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.NewPreset_Click" + " line: " + Str(Erl))

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
Private Sub SavePreset(ByRef Name As String)
        '<EhHeader>
        On Error GoTo SavePreset_Err
        '</EhHeader>

    Dim PalCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long
    Dim Temp As String

100 If Name = "overwriteoldindex" Then


102 With m_cIni
    
104     .Path = App.Path & "\Palettes.ini"
106     .Section = "Palettes"
    
108     If Count > 0 Then
110         .Key = "Name" & Format(List1.ItemData(List1.ListIndex)): .Default = "": Name = .Value
112         If Name <> "" Then
114             .Section = Name
116             .DeleteSection
118             .Key = "count": .Value = 15
120             .Key = "type": .Value = ChipType
122             For Y = 0 To pal_Count
124                 Temp = Hex(PaletteRGB(Y))
126                 If Len(Temp) < 6 Then
128                     Temp = Mid("000000", 1, 6 - Len(Temp)) & Temp
                    End If
130                 .Key = "c" & Format(Y): .Value = "&H" & Temp
132             Next Y
            End If
        End If
    
    End With


    Else


134 With m_cIni
    
136     .Path = App.Path & "\Palettes.ini"
138     .Section = "Palettes"
    
140     .Key = "count":  .Value = .Value + 1
142     PalCount = .Value
    

144     For X = 0 To PalCount
146         .Key = "Name" & Format(X)
148         If StrComp(.Value, Name, vbTextCompare) = 0 Then
150             MsgBox "A preset with this name does already exist", vbExclamation, "Warning"
152             Name = "alreadyexisted"
                Exit Sub
            End If
154     Next X
    
156     .Key = "Name" & Format(List1.ListCount)
158     .Value = Name
    
160     .Section = Name
162     .DeleteSection
164     .Key = "count": .Value = pal_Count
166     .Key = "type": .Value = ChipType
168         For Y = 0 To pal_Count
170             Temp = Hex(PaletteRGB(Y))
172             If Len(Temp) < 6 Then
174                 Temp = Mid("000000", 1, 6 - Len(Temp)) & Temp
                End If
176             .Key = "c" & Format(Y): .Value = "&H" & Temp
178         Next Y

    End With


    End If

        '<EhFooter>
        Exit Sub

SavePreset_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.SavePreset" + " line: " + Str(Erl))

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

Private Function HexPad(ByRef Number As Long) As String
        '<EhHeader>
        On Error GoTo HexPad_Err
        '</EhHeader>
    Dim X As String

100 X = Hex(Number)
102 X = Mid("0000", 1, 6 - Len(X)) & X

104 HexPad = X
        '<EhFooter>
        Exit Function

HexPad_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.HexPad" + " line: " + Str(Erl))

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

Public Sub GetPalList()
        '<EhHeader>
        On Error GoTo GetPalList_Err
        '</EhHeader>
    Dim PalCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long
    Dim Name As String
    Dim Colors As String

100 List1.Clear

102 With m_cIni
    
104     .Path = App.Path & "\Palettes.ini"
106     .Section = "Palettes"
108     .Key = "count": .Default = 0: PalCount = .Value
    
110     If PalCount <> 0 Then
112         X = 0
114         For Y = 0 To PalCount
116             .Key = "Name" & Format(Y): .Default = "Error!": Name = .Value
            
118             .Section = Name
120             .Key = "type": .Default = "error"
122             If .Value = ChipType Then
124                 List1.AddItem Name
126                 List1.ItemData(X) = Y
128                 X = X + 1
                End If
             
130             .Section = "Palettes"
            
132         Next Y
        End If
 
    End With

        '<EhFooter>
        Exit Sub

GetPalList_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.GetPalList" + " line: " + Str(Erl))

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




Private Sub SaveOverSelected_Click()
        '<EhHeader>
        On Error GoTo SaveOverSelected_Click_Err
        '</EhHeader>

100 Call SavePreset("overwriteoldindex")

        '<EhFooter>
        Exit Sub

SaveOverSelected_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.SaveOverSelected_Click" + " line: " + Str(Erl))

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

Private Sub DeletePreset()
        '<EhHeader>
        On Error GoTo DeletePreset_Err
        '</EhHeader>

    Dim PalCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long
    Dim Name As String
    Dim Temp As String
100 With m_cIni
    
102     .Path = App.Path & "\Palettes.ini"
104     .Section = "Palettes"
106     .Key = "count": .Default = 0: PalCount = .Value
    
108     If PalCount <> 0 Then
110         .Key = "Name" & Format(List1.ItemData(List1.ListIndex)):    Name = .Value
112         .Section = Name
114         .DeleteSection
116         .Section = "Palettes"
118         .Key = "Name" & Format(List1.ItemData(List1.ListIndex))
120         .DeleteKey
122         PalCount = PalCount - 1
        
124         Y = List1.ItemData(List1.ListIndex) + 1
126         For X = Y To PalCount + 1
128             .Key = "Name" & Format(X)
130             Temp = .Value
132             .DeleteKey
134             .Key = "Name" & Format(X - 1)
136             .Value = Temp
138         Next X
140         .Section = "Palettes"
142         .Key = "count"
144         .Value = PalCount
        End If

    End With

146 Call GetPalList
        '<EhFooter>
        Exit Sub

DeletePreset_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.DeletePreset" + " line: " + Str(Erl))

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


Private Sub Text1_LostFocus(Index As Integer)
        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err
        '</EhHeader>
100 Text1(Index).Text = HexConvert(Text1(Index).Text)

102 If pal_HexMode = True Then
104     PaletteRGB(pal_Lastindex) = RGB(Val("&h" & Text1(0).Text), Val("&h" & Text1(1).Text), Val("&h" & Text1(2).Text))
    Else
106     PaletteRGB(pal_Lastindex) = RGB(Val(Text1(0).Text), Val(Text1(1).Text), Val(Text1(2).Text))
    End If


108 RedrawPalette
        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.Text1_LostFocus" + " line: " + Str(Erl))

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
Private Sub text1_gotfocus(Index As Integer)
        '<EhHeader>
        On Error GoTo text1_gotfocus_Err
        '</EhHeader>

100 Text1(Index).SelStart = 0
102 Text1(Index).SelLength = Len(Text1(Index).Text)

        '<EhFooter>
        Exit Sub

text1_gotfocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.text1_gotfocus" + " line: " + Str(Erl))

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
Private Function HexConvert(ByRef Number As String) As String
        '<EhHeader>
        On Error GoTo HexConvert_Err
        '</EhHeader>
    Dim X As Single

100 If pal_HexMode = True Then

102     X = Val("&H" & Number)
104     If X < 0 Then X = X + 255
106     Number = Hex(X)
    
108     If X < 0 Or X > 255 Then Number = "00"
    
110     If Len(Number) < 2 Then
112        Number = Mid("00", 1, 2 - Len(Number)) & Number
        End If
    
114 HexConvert = Number

    Else

116  X = Val(Number)
118  If X < 0 Or X > 255 Then X = 0

120  HexConvert = Str(X)
    End If
        '<EhFooter>
        Exit Function

HexConvert_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.PalSelect.HexConvert" + " line: " + Str(Erl))

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
