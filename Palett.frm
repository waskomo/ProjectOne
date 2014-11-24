VERSION 5.00
Begin VB.Form Palett 
   Caption         =   "Palette"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5700
   DrawMode        =   11  'Not Xor Pen
   FillColor       =   &H80000013&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000013&
   Icon            =   "Palett.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Palett"
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   2595
   ScaleWidth      =   5700
   Visible         =   0   'False
   Begin Project_One.MyPalette MyPalette1 
      Height          =   832
      Left            =   0
      TabIndex        =   1
      Top             =   117
      Width           =   2938
      _ExtentX        =   0
      _ExtentY        =   0
      BoxCount        =   15
      BoxWidth        =   20
      BoxHSpacing     =   4
      BoxTop          =   4
      BoxLeft         =   4
      BoxPerLine      =   8
      BackgrColor     =   4546465
      BoxborderColor  =   0
      BoxHasBorders   =   -1  'True
      BoxTextColor    =   16777215
      BoxHasText      =   -1
      BoxBorderStyle  =   2
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   135
      Left            =   4329
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Swap Colors"
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblLeftColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000000&
      Height          =   240
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   240
   End
   Begin VB.Label lblRightColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4215
      TabIndex        =   4
      Top             =   360
      Width           =   240
   End
   Begin VB.Label lblBackground 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   598
      Left            =   3978
      TabIndex        =   2
      Top             =   117
      Width           =   598
   End
End
Attribute VB_Name = "Palett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
    Dim Temp As Long
100 Temp = LeftColN
102 LeftColN = RightColN
104 RightColN = Temp
106 UpdateColors
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.Command1_Click" + " line: " + Str(Erl))

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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err
        '</EhHeader>

    'user pressed X on palett window -> hide it
100 If UnloadMode = 0 Then
102     Cancel = True
104     Palett.Visible = False
106     MainWin.mnuPaletteWindow.Checked = Palett.Visible
    End If

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.Form_QueryUnload" + " line: " + Str(Erl))

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

lblBackground.Move Palett.ScaleWidth - lblBackground.Width, (Me.ScaleHeight - lblBackground.Height) / 2

lblLeftColor.Move lblBackground.Left + lblBackground.Width / 6, lblBackground.Top + lblBackground.Height / 6
lblRightColor.Move lblLeftColor.Left + lblLeftColor.Width / 2, lblLeftColor.Top + lblLeftColor.Height / 2
Command1.Move lblLeftColor.Left + lblLeftColor.Width, lblLeftColor.Top


MyPalette1.Move 0, 0, Me.ScaleWidth - lblBackground.Width, Me.ScaleHeight
End Sub



Private Sub lblBackground_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo lblBackground_MouseUp_Err
        '</EhHeader>

100     If Button = 1 Then
102         BackgrIndex = LeftColN
104         If BaseMode = BaseModeTyp.multi Then Call backgrchange(LeftColN)
        
106     ElseIf Button = 2 Then
108         BackgrIndex = RightColN
110         If BaseMode = BaseModeTyp.multi Then Call backgrchange(RightColN)
        
        End If
112     lblBackground.BackColor = PaletteRGB(BackgrIndex)
114     Call ZoomWindow.WriteUndo

        '<EhFooter>
        Exit Sub

lblBackground_MouseUp_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.lblBackground_MouseUp" + " line: " + Str(Erl))

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

100     If Button = 1 Then LeftColN = Index
102     If Button = 2 Then RightColN = Index

104     Call UpdateColors

        '<EhFooter>
        Exit Sub

MyPalette1_BoxClicked_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.MyPalette1_BoxClicked" + " line: " + Str(Erl))

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



Public Sub UpdateColors()
        '<EhHeader>
        On Error GoTo UpdateColors_Err
        '</EhHeader>

100 lblBackground.BackColor = PaletteRGB(BackgrIndex And 15)
102 lblLeftColor.BackColor = PaletteRGB(LeftColN)
104 lblRightColor.BackColor = PaletteRGB(RightColN)

        '<EhFooter>
        Exit Sub

UpdateColors_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.UpdateColors" + " line: " + Str(Erl))

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


Private Sub form_keydown(KeyCode As Integer, Shift As Integer)
    'MsgBox "nyihi"
        '<EhHeader>
        On Error GoTo form_keydown_Err
        '</EhHeader>
100 Call ZoomWindow.Shared_keydown(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

form_keydown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.form_keydown" + " line: " + Str(Erl))

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
Private Sub form_keyup(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo form_keyup_Err
        '</EhHeader>
100 Call ZoomWindow.Shared_keyup(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

form_keyup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.form_keyup" + " line: " + Str(Erl))

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
    Dim X As Byte

100 Set Me.Icon = ZoomWindow.Image1.Picture

102 MyPalette1.BoxHasText = False
104 MyPalette1.BackgrColor = RGB(0, 255, 0)
106 MyPalette1.InitSurface

108 MyPalette1.Move 0, 0
110 Palett.ScaleMode = vbPixels


112 Palett.Top = PrevWin.Top + PrevWin.Height
114 Palett.Left = PrevWin.Left

116 UpdateColors

        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Palett.form_load" + " line: " + Str(Erl))

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




