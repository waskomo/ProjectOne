VERSION 5.00
Begin VB.Form ColorSelect 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Palette"
   ClientHeight    =   1350
   ClientLeft      =   4500
   ClientTop       =   2340
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Begin Project_One.MyPalette MyPalette1 
      Height          =   1066
      Left            =   117
      TabIndex        =   0
      Top             =   117
      Width           =   3055
      _ExtentX        =   5398
      _ExtentY        =   1879
      BoxCount        =   15
      BoxWidth        =   20
      BoxHSpacing     =   4
      BoxTop          =   4
      BoxLeft         =   4
      BoxPerLine      =   8
      BackgrColor     =   16711680
      BoxborderColor  =   0
      BoxHasBorders   =   -1  'True
      BoxTextColor    =   16777215
      BoxHasText      =   0
      BoxBorderStyle  =   0
   End
End
Attribute VB_Name = "ColorSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedColor As Byte
Option Explicit

Private Sub form_load()
        '<EhHeader>
        On Error GoTo form_load_Err
        '</EhHeader>
    Dim t As Byte

100 Set Me.Icon = ZoomWindow.Image1.Picture
102 Me.ScaleMode = vbPixels

104 For t = 0 To pal_Count
106     MyPalette1.PaletteRGB(t) = PaletteRGB(t)
    Next

108 MyPalette1.InitSurface
110 MyPalette1.BackgrColor = Me.BackColor
112 MyPalette1.BoxCount = 15
114 MyPalette1.Move 0, 0
        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ColorSelect.form_load" + " line: " + Str(Erl))

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



Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

MyPalette1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub MyPalette1_BoxClicked(Button As Integer, Index As Byte)
        '<EhHeader>
        On Error GoTo MyPalette1_BoxClicked_Err
        '</EhHeader>
100 SelectedColor = Index
102 Unload ColorSelect
        '<EhFooter>
        Exit Sub

MyPalette1_BoxClicked_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ColorSelect.MyPalette1_BoxClicked" + " line: " + Str(Erl))

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
