VERSION 5.00
Begin VB.Form MemMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Load"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Adress Handling:"
      Height          =   1575
      Left            =   120
      TabIndex        =   49
      Top             =   5040
      Width           =   4935
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   54
         Text            =   "Text5"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkOverrideAdress 
         Caption         =   "Override Original LoadAdress:"
         Height          =   495
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optAbsoluteAdress 
         Caption         =   "Adresses are absolute in memory"
         Height          =   495
         Left            =   2760
         TabIndex        =   52
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton optRelativeAdress 
         Caption         =   "Adresses are relative to file start"
         Height          =   375
         Left            =   2760
         TabIndex        =   51
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkSkipAdress 
         Caption         =   "Skip first 2 bytes (load address)"
         Height          =   495
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Background:"
      Height          =   855
      Left            =   2640
      TabIndex        =   44
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Text            =   "Text4"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "$D021:"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   47
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview:"
      Height          =   3855
      Left            =   5160
      TabIndex        =   42
      Top             =   1080
      Width           =   5535
      Begin VB.PictureBox Preview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   360
         ScaleHeight     =   217
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   337
         TabIndex        =   43
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Setup ScreenMode"
      Height          =   375
      Left            =   3480
      TabIndex        =   48
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color Memory:"
      Height          =   855
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "$D800:"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Done"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Frame Bank2 
      Caption         =   "Bank [B]:"
      Height          =   3855
      Index           =   1
      Left            =   2640
      TabIndex        =   30
      Top             =   1080
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1080
         TabIndex        =   17
         Text            =   "43e5"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1080
         TabIndex        =   16
         Text            =   "43e5"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   15
         Text            =   "43e5"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1080
         TabIndex        =   14
         Text            =   "43e5"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   13
         Text            =   "43e5"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1080
         TabIndex        =   12
         Text            =   "43e5"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1080
         TabIndex        =   11
         Text            =   "43e5"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   10
         Text            =   "43e5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 8"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   39
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 7"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   38
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 6"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   37
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 5"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   36
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 4"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   35
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 3"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 2"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 1"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Bitmap:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Bank1 
      Caption         =   "Bank [A]:"
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   8
         Text            =   "43e5"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   7
         Text            =   "43e5"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   6
         Text            =   "43e5"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   5
         Text            =   "43e5"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   4
         Text            =   "43e5"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   3
         Text            =   "43e5"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Text            =   "43e5"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Text            =   "43e5"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Bitmap:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 8"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 7"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 6"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Screen 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Adress Handling:"
      Height          =   1575
      Left            =   120
      TabIndex        =   55
      Top             =   5040
      Width           =   4935
      Begin VB.CheckBox chkSaveLoadAddress 
         Caption         =   "Include LoadAddress"
         Height          =   375
         Left            =   2640
         TabIndex        =   60
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   57
         Text            =   "Text7"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1320
         TabIndex        =   56
         Text            =   "Text6"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Last address:"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "First address:"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "MemMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Dim OldBackgrIndex As Long

Private Sub RefreshPreview()
        '<EhHeader>
        On Error GoTo RefreshPreview_Err
        '</EhHeader>
        Dim X As Long

        'first transform textbox addies into our variables

100     With CustomIOSetup

102         For X = 0 To 15
104             .Screen(X) = MyVal("&H" & Text1(X).Text)
106         Next X
        
108         .Bitmap(0) = MyVal("&H" & Text2(0).Text)
110         .Bitmap(1) = MyVal("&H" & Text2(1).Text)
112         .D800 = MyVal("&H" & Text3.Text)
114         .D021 = MyVal("&H" & Text4.Text)
116         .StartAdressUser = MyVal("&H" & Text5.Text)
118         .Start = MyVal("&H" & Text6.Text)
120         .End = MyVal("&H" & Text7.Text)
122          If chkSaveLoadAddress.Value = 1 Then .HasStartAddress = True Else .HasStartAddress = False
124         .ForceStartAddressFromUser = chkOverrideAdress.Value
126         .SkipStartAdress = chkSkipAdress.Value
        End With

128     If adr_LoadorSave = "load" Then
            'load binary data into our emulated c64 memory
130         Call CustomLoad
            'draw picture based on above
132         Call DrawPicFromMem
134         StretchBlt Preview.hdc, 0, 0, PW, PH, PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy
136         Preview.Refresh
        End If

        '<EhFooter>
        Exit Sub

RefreshPreview_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.RefreshPreview" + " line: " + Str(Erl))

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

Private Function MyVal(ByVal Number As String) As Long
        '<EhHeader>
        On Error GoTo MyVal_Err
        '</EhHeader>
    Dim X As Long

100 X = Val(Number)
102 X = X And 65535

104 MyVal = X
        '<EhFooter>
        Exit Function

MyVal_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.MyVal" + " line: " + Str(Erl))

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

Private Sub btnCancel_Click()
        '<EhHeader>
        On Error GoTo btnCancel_Click_Err
        '</EhHeader>

100 If adr_LoadorSave = "load" Then
102     Call ZoomWindow.ReadUndo
104     Call DrawPicFromMem
106     BackgrIndex = OldBackgrIndex
    End If

108 Unload Me
        '<EhFooter>
        Exit Sub

btnCancel_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.btnCancel_Click" + " line: " + Str(Erl))

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
102     If adr_LoadorSave = "load" Then
104         Call ZoomWindow.ReadUndo
106         Call DrawPicFromMem
108         BackgrIndex = OldBackgrIndex
        End If
    End If

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Form_QueryUnload" + " line: " + Str(Erl))

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
Private Sub btnOk_Click()
        '<EhHeader>
        On Error GoTo btnOk_Click_Err
        '</EhHeader>

100 If adr_LoadorSave = "load" Then
102     Me.ValidateControls
104     Call ZoomWindow.ZoomResize
106     GfxMode = "custom"
        'Call ZoomWindow.ChangeGfxMode
108     ZoomWindow.ZoomWinRefresh
    End If

110 If adr_LoadorSave = "save" Then
112     Me.ValidateControls
114     Call CustomSave
    End If

116 Unload Me

        '<EhFooter>
        Exit Sub

btnOk_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.btnOk_Click" + " line: " + Str(Erl))

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

Private Sub chkOverrideAdress_Click()
        '<EhHeader>
        On Error GoTo chkOverrideAdress_Click_Err
        '</EhHeader>
100 CustomIOSetup.ForceStartAddressFromUser = chkOverrideAdress.Value
102 Call RefreshPreview
        '<EhFooter>
        Exit Sub

chkOverrideAdress_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.chkOverrideAdress_Click" + " line: " + Str(Erl))

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

Private Sub chkSaveLoadAddress_Click()
        '<EhHeader>
        On Error GoTo chkSaveLoadAddress_Click_Err
        '</EhHeader>
100 If chkSaveLoadAddress.Value = 1 Then CustomIOSetup.HasStartAddress = True Else CustomIOSetup.HasStartAddress = False
        '<EhFooter>
        Exit Sub

chkSaveLoadAddress_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.chkSaveLoadAddress_Click" + " line: " + Str(Erl))

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

Private Sub chkSkipAdress_Click()
        '<EhHeader>
        On Error GoTo chkSkipAdress_Click_Err
        '</EhHeader>
100 CustomIOSetup.SkipStartAdress = chkSkipAdress.Value
102 Call RefreshPreview
        '<EhFooter>
        Exit Sub

chkSkipAdress_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.chkSkipAdress_Click" + " line: " + Str(Erl))

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

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>

100 Custom.Show vbModal
102 Call SetControlVisibility

104 Call RefreshPreview

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Command1_Click" + " line: " + Str(Erl))

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

Private Sub optAbsoluteAdress_Click()
        '<EhHeader>
        On Error GoTo optAbsoluteAdress_Click_Err
        '</EhHeader>
100 CustomIOSetup.Absolute = True
102 Call RefreshPreview
        '<EhFooter>
        Exit Sub

optAbsoluteAdress_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.optAbsoluteAdress_Click" + " line: " + Str(Erl))

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

Private Sub optRelativeAdress_Click()
        '<EhHeader>
        On Error GoTo optRelativeAdress_Click_Err
        '</EhHeader>
100 CustomIOSetup.Absolute = False
102 Call RefreshPreview
        '<EhFooter>
        Exit Sub

optRelativeAdress_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.optRelativeAdress_Click" + " line: " + Str(Erl))

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

'validate d021
Private Sub Text4_Validate(Cancel As Boolean)
        '<EhHeader>
        On Error GoTo Text4_Validate_Err
        '</EhHeader>
100 Text4.Text = HexConvert(Text4.Text)
102 Call RefreshPreview
        '<EhFooter>
        Exit Sub

Text4_Validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Text4_Validate" + " line: " + Str(Erl))

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

'validate bitmap adresses
Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
        '<EhHeader>
        On Error GoTo Text2_Validate_Err
        '</EhHeader>
    
100 Text2(Index).Text = HexConvert(Text2(Index).Text)

102 Call RefreshPreview
        '<EhFooter>
        Exit Sub

Text2_Validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Text2_Validate" + " line: " + Str(Erl))

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

'validate screen adresses
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
        '<EhHeader>
        On Error GoTo Text1_Validate_Err
        '</EhHeader>

100 Text1(Index).Text = HexConvert(Text1(Index).Text)
102 Call RefreshPreview
        '<EhFooter>
        Exit Sub

Text1_Validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Text1_Validate" + " line: " + Str(Erl))

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

'validate user base adresse
Private Sub Text5_Validate(Cancel As Boolean)
        '<EhHeader>
        On Error GoTo Text5_Validate_Err
        '</EhHeader>
    
100 Text5.Text = HexConvert(Text5.Text)
102 Call RefreshPreview

        '<EhFooter>
        Exit Sub

Text5_Validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Text5_Validate" + " line: " + Str(Erl))

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

'gotfocus events belove are to provide an autoselect all when clickid in textbox function

Private Sub text2_gotfocus(Index As Integer)
        '<EhHeader>
        On Error GoTo text2_gotfocus_Err
        '</EhHeader>

100 Text2(Index).SelStart = 0
102 Text2(Index).SelLength = Len(Text2(Index).Text)

        '<EhFooter>
        Exit Sub

text2_gotfocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.text2_gotfocus" + " line: " + Str(Erl))

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
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.text1_gotfocus" + " line: " + Str(Erl))

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
Private Sub text3_gotfocus()
        '<EhHeader>
        On Error GoTo text3_gotfocus_Err
        '</EhHeader>

100 Text3.SelStart = 0
102 Text3.SelLength = Len(Text3.Text)

        '<EhFooter>
        Exit Sub

text3_gotfocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.text3_gotfocus" + " line: " + Str(Erl))

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

'validate d800 adresses
Private Sub Text3_Validate(Cancel As Boolean)
        '<EhHeader>
        On Error GoTo Text3_Validate_Err
        '</EhHeader>
    
100 Text3.Text = HexConvert(Text3.Text)
102 Call RefreshPreview

        '<EhFooter>
        Exit Sub

Text3_Validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Text3_Validate" + " line: " + Str(Erl))

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
'start adress of save
Private Sub text6_gotfocus()
        '<EhHeader>
        On Error GoTo text6_gotfocus_Err
        '</EhHeader>

100 Text6.SelStart = 0
102 Text6.SelLength = Len(Text6.Text)

        '<EhFooter>
        Exit Sub

text6_gotfocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.text6_gotfocus" + " line: " + Str(Erl))

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

Private Sub Text6_Validate(Cancel As Boolean)
        '<EhHeader>
        On Error GoTo Text6_Validate_Err
        '</EhHeader>
    
100 Text6.Text = HexConvert(Text6.Text)
102 Call RefreshPreview

        '<EhFooter>
        Exit Sub

Text6_Validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Text6_Validate" + " line: " + Str(Erl))

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
'end adress of save
Private Sub text7_gotfocus()
        '<EhHeader>
        On Error GoTo text7_gotfocus_Err
        '</EhHeader>

100 Text7.SelStart = 0
102 Text7.SelLength = Len(Text7.Text)

        '<EhFooter>
        Exit Sub

text7_gotfocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.text7_gotfocus" + " line: " + Str(Erl))

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

Private Sub Text7_Validate(Cancel As Boolean)
        '<EhHeader>
        On Error GoTo Text7_Validate_Err
        '</EhHeader>
    
100 Text7.Text = HexConvert(Text7.Text)
102 Call RefreshPreview

        '<EhFooter>
        Exit Sub

Text7_Validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.Text7_Validate" + " line: " + Str(Erl))

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

Private Sub text4_gotfocus()
        '<EhHeader>
        On Error GoTo text4_gotfocus_Err
        '</EhHeader>

100 Text4.SelStart = 0
102 Text4.SelLength = Len(Text4.Text)

        '<EhFooter>
        Exit Sub

text4_gotfocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.text4_gotfocus" + " line: " + Str(Erl))

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
Private Sub text5_gotfocus()
        '<EhHeader>
        On Error GoTo text5_gotfocus_Err
        '</EhHeader>

100 Text5.SelStart = 0
102 Text5.SelLength = Len(Text5.Text)

        '<EhFooter>
        Exit Sub

text5_gotfocus_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.text5_gotfocus" + " line: " + Str(Erl))

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

100     X = Val("&H" & Number)
102     If X < 0 Then X = X + 65536
104     Number = Hex(X)
    
106     If X < 0 Or X > 65535 Then Number = "0000"
    
108     If Len(Number) < 4 Then
110        Number = Mid("0000", 1, 4 - Len(Number)) & Number
        End If
    
112 HexConvert = Number
        '<EhFooter>
        Exit Function

HexConvert_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.HexConvert" + " line: " + Str(Erl))

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
Private Function HexPad(ByRef Number As Long) As String
        '<EhHeader>
        On Error GoTo HexPad_Err
        '</EhHeader>
    Dim X As String

100 X = Hex(Number)
102 If Len(X) < 4 Then
104     X = Mid("0000", 1, 4 - Len(X)) & X
    End If

106 HexPad = X
        '<EhFooter>
        Exit Function

HexPad_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.HexPad" + " line: " + Str(Erl))

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

Private Sub form_activate()
        '<EhHeader>
        On Error GoTo form_activate_Err
        '</EhHeader>
    Dim X As Long

100 OldBackgrIndex = BackgrIndex

102 If adr_LoadorSave = "load" Then
104     Frame4.Visible = True
106     Frame5.Visible = False
108     Frame2.Visible = True
110     Me.Width = 10860
112     btnOk.Caption = "Load"
114 ElseIf adr_LoadorSave = "save" Then
116     Frame4.Visible = False
118     Frame5.Visible = True
120     Frame2.Visible = False
122     If CustomIOSetup.HasStartAddress = True Then chkSaveLoadAddress.Value = 1 Else chkSaveLoadAddress.Value = 0
        'Me.Width = (Bank1(0).Left + Bank1(0).Width - Bank2(1).Left) + Bank2(1).Left + Bank2(1).Width
124     Me.Width = 5220
126     btnOk.Caption = "Save"
    End If


128 Me.Refresh

    'set up gui based on settings

130 Call SetControlVisibility

    'fill up textboxes with the current adresses used
132 For X = 0 To 15
134  Text1(X).Text = HexPad(CustomIOSetup.Screen(X))
136 Next X
138 Text2(0).Text = HexPad(CustomIOSetup.Bitmap(0))
140 Text2(1).Text = HexPad(CustomIOSetup.Bitmap(1))
142 Text3.Text = HexPad(CustomIOSetup.D800)
144 Text4.Text = HexPad(CustomIOSetup.D021)
146 Text5.Text = HexPad(CustomIOSetup.StartAdressUser)
148 Text6.Text = HexPad(CustomIOSetup.Start)
150 Text7.Text = HexPad(CustomIOSetup.End)

    'checkbox setup based on current settings
152 If CustomIOSetup.Absolute = True Then
154     MemMap.optAbsoluteAdress.Value = 1
    Else
156     MemMap.optRelativeAdress = 0
    End If

158 If CustomIOSetup.SkipStartAdress = True Then
160     MemMap.chkSkipAdress.Value = 1
    Else
162     MemMap.chkSkipAdress.Value = 0
    End If

164 If CustomIOSetup.ForceStartAddressFromUser = True Then
166     MemMap.chkOverrideAdress.Value = 1
    Else
168     MemMap.chkOverrideAdress.Value = 0
    End If


    'set up of preview picture box
170 MemMap.ScaleMode = vbPixels
172 Preview.ScaleMode = vbPixels
174 Preview.Width = ScaleX(PW, vbPixels, vbTwips)
176 Preview.Height = ScaleY(PH, vbPixels, vbTwips)

178 Call RefreshPreview

    'so we can get back with readundo to the old picture if cancelled
180 Call ZoomWindow.WriteUndo

        '<EhFooter>
        Exit Sub

form_activate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.form_activate" + " line: " + Str(Erl))

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
'sets up visible adress settings textboxes based on custom gfxmode
'set visibility of controls according to screenmode
Private Sub SetControlVisibility()
        '<EhHeader>
        On Error GoTo SetControlVisibility_Err
        '</EhHeader>
    Dim X As Long
    Dim Max As Long
    Dim lace As String
    Dim Mode As String

100 If BaseMode_cm = BaseModeTyp.hires Then
102     Frame1.Enabled = False
104     Frame3.Enabled = False
106     Label3.Enabled = False
108     Label4.Enabled = False
110     Mode = "Hires"
112 ElseIf BaseMode_cm = BaseModeTyp.multi Then
114     Frame1.Enabled = True
116     Frame3.Enabled = True
118     Label3.Enabled = True
120     Label4.Enabled = True
122     Mode = "MultiColor"
    End If


124 If Not ((ResoDiv_cm = 1) And (BmpBanks_cm = 1)) Then
126     Text2(1).Enabled = False
128     Label2(1).Enabled = False
    Else
130     Text2(1).Enabled = True
132     Label2(1).Enabled = True
    End If


134 For X = 0 To 15
136     Text1(X).Enabled = False
138     Label1(X).Enabled = False
140 Next X

142 Max = 7 \ FliMul_cm

144 For X = 0 To Max
146     Text1(X).Enabled = True
148     Label1(X).Enabled = True
150 Next X

152 If ScrBanks_cm = 1 Then
154     For X = 8 To 8 + Max
156         Text1(X).Enabled = True
158         Label1(X).Enabled = True
160     Next X
    End If

    'set window caption

162 If BmpBanks_cm = 1 Then
164     lace = " interlaced"""
166     Bank2(1).Enabled = True
    Else
168     lace = ""
170     Bank2(1).Enabled = False
    End If

172 If adr_LoadorSave = "load" Then
174     MemMap.Caption = "Custom Load; Mode:" + _
                         lace + Mode + ", Fli lines per char:" + Str(Int(8 / FliMul_cm)) + ""
176 ElseIf adr_LoadorSave = "save" Then
178     MemMap.Caption = "Custom Save:" + _
                         lace + Mode + ", Fli lines per char:" + Str(Int(8 / FliMul_cm)) + ""
    End If

        '<EhFooter>
        Exit Sub

SetControlVisibility_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.SetControlVisibility" + " line: " + Str(Erl))

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
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MemMap.form_load" + " line: " + Str(Erl))

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
