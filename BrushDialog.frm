VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BrushDialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Brush Setup"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2355
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox col1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Brush 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   720
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin MSComctlLib.Slider SizeSlider 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   256
      SelStart        =   128
      TickStyle       =   3
      Value           =   128
   End
   Begin VB.PictureBox col2 
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin MSComctlLib.Slider DitherSlider 
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   476
      _Version        =   393216
      Max             =   16
      SelStart        =   8
      TickStyle       =   3
      Value           =   8
   End
   Begin VB.Label Label3 
      Caption         =   "Size"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Color 2"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Color 1"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "BrushDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As Double
Dim Color1 As Byte, color2 As Byte
Attribute color2.VB_VarUserMemId = 1073938434
Dim Dither As Byte
Attribute Dither.VB_VarUserMemId = 1073938436

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
100 BrushDither = Dither
102 BrushSize = R
104 BrushColor1 = Color1
106 BrushColor2 = color2

108 BrushSize = BrushSize
110 BrushDither = BrushDither
112 BrushColor1 = BrushColor1
114 BrushColor2 = BrushColor2

116 BrushPreCalc

118 Unload Me
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.Command1_Click" + " line: " + Str(Erl))

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

102 BrushDialog.ScaleMode = vbPixels
104 Brush.ScaleMode = vbPixels
106 Brush.AutoRedraw = True

108 Brush.BackColor = PaletteRGB(BackgrIndex)

110 Dither = BrushDither
112 R = BrushSize
114 Color1 = BrushColor1
116 color2 = BrushColor2

118 col1.BackColor = PaletteRGB(Color1)
120 col2.BackColor = PaletteRGB(color2)

122 SizeSlider.Value = BrushSize * 8
124 DitherSlider.Value = Dither

    'Call Generate_Brush



        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.form_load" + " line: " + Str(Erl))

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
Private Sub col1_Click()
        '<EhHeader>
        On Error GoTo col1_Click_Err
        '</EhHeader>
100 ColorSelect.Show vbModal
102 col1.BackColor = PaletteRGB(ColorSelect.SelectedColor)
104 Color1 = ColorSelect.SelectedColor
106 Call Generate_Brush
        '<EhFooter>
        Exit Sub

col1_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.col1_Click" + " line: " + Str(Erl))

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

Private Sub col2_Click()
        '<EhHeader>
        On Error GoTo col2_Click_Err
        '</EhHeader>
100 ColorSelect.Show vbModal
102 col2.BackColor = PaletteRGB(ColorSelect.SelectedColor)
104 color2 = ColorSelect.SelectedColor
106 Call Generate_Brush
        '<EhFooter>
        Exit Sub

col2_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.col2_Click" + " line: " + Str(Erl))

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
Private Sub ditherslider_scroll()
        '<EhHeader>
        On Error GoTo ditherslider_scroll_Err
        '</EhHeader>

100 Dither = DitherSlider.Value
102 Call Generate_Brush
        '<EhFooter>
        Exit Sub

ditherslider_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.ditherslider_scroll" + " line: " + Str(Erl))

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

Private Sub sizeslider_scroll()
        '<EhHeader>
        On Error GoTo sizeslider_scroll_Err
        '</EhHeader>

100 R = SizeSlider.Value / 8
102 Brush.Forecolor = RGB(255, 255, 255)
104 Brush.Cls
106 Brush.Circle (Brush.Width / 2, Brush.Height / 2), SizeSlider.Value * 2
    'Brush.Refresh

108 Call Generate_Brush
        '<EhFooter>
        Exit Sub

sizeslider_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.sizeslider_scroll" + " line: " + Str(Erl))

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


Private Sub Generate_Brush()
        '<EhHeader>
        On Error GoTo Generate_Brush_Err
        '</EhHeader>
    Dim X As Double
    Dim Y As Double
    Dim Bayer(3, 3) As Byte
    Dim D As Double
    Dim Round As Double

100 Brush.ScaleMode = vbPixels

102 Bayer(0, 0) = 1: Bayer(0, 1) = 9: Bayer(0, 2) = 3: Bayer(0, 3) = 11
104 Bayer(1, 0) = 13: Bayer(1, 1) = 5: Bayer(1, 2) = 15: Bayer(1, 3) = 7
106 Bayer(2, 0) = 4: Bayer(2, 1) = 12: Bayer(2, 2) = 2: Bayer(2, 3) = 10
108 Bayer(3, 0) = 16: Bayer(3, 1) = 8: Bayer(3, 2) = 14: Bayer(3, 3) = 6

110 Brush.Cls

112 Round = Int((R / 2) / ResoDiv) * ResoDiv

114 For X = -Round To Round Step ResoDiv
116     For Y = -Round To Round Step 1

118         D = Sqr(X * X + Y * Y)

120         If D < R / 2 Then

122             If Dither < Bayer((X / ResoDiv) And 3, Y And 3) Then _
                   Brush.Forecolor = PaletteRGB(Color1) Else _
                   Brush.Forecolor = PaletteRGB(color2)

124             Brush.PSet (Int(Brush.Width / 2) + X, Int(Brush.Height / 2) + Y)

126             If ResoDiv = 2 Then _
                   Brush.PSet (Int(Brush.Width / 2) + X + 1, Int(Brush.Height / 2) + Y)
            End If

128     Next Y
130 Next X


        '<EhFooter>
        Exit Sub

Generate_Brush_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.Generate_Brush" + " line: " + Str(Erl))

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

100 Call Generate_Brush


        '<EhFooter>
        Exit Sub

form_activate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.BrushDialog.form_activate" + " line: " + Str(Erl))

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

