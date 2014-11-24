VERSION 5.00
Begin VB.Form Custom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Mode Setup"
   ClientHeight    =   4320
   ClientLeft      =   4530
   ClientTop       =   3705
   ClientWidth     =   4035
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4035
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Base Mode: "
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optHires 
         Caption         =   "Hires"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optMultiColor 
         Caption         =   "Multicolor"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CheckBox chkSameScreen 
         Caption         =   "Same screen data used by both banks"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CheckBox chkInterlace 
         Caption         =   "Use interlace"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Frame lace 
         Caption         =   "Interlace settings:"
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fli Settings: "
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Custom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
        '<EhHeader>
        On Error GoTo btnCancel_Click_Err
        '</EhHeader>
100 Unload Me
        '<EhFooter>
        Exit Sub

btnCancel_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.btnCancel_Click" + " line: " + Str(Erl))

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

100 Select Case Combo1.ListIndex
    Case 0
102     FliMul_cm = 8
104 Case 1
106     FliMul_cm = 4
108 Case 2
110    FliMul_cm = 2
112 Case 3
114     FliMul_cm = 1
    End Select



116 If Combo1.ListIndex > 0 Then XFliLimit_cm = 24 Else XFliLimit_cm = 0

118 If optHires.Value = True Then
120     BaseMode_cm = BaseModeTyp.hires
122     ResoDiv_cm = 1
    End If


124 If optMultiColor.Value = True Then

126     BaseMode_cm = BaseModeTyp.multi
128     If chkInterlace.Value = 1 Then
130         ResoDiv_cm = 1
132         BmpBanks_cm = 1
134         If chkSameScreen.Value = 1 Then
136             ScrBanks_cm = 0
            Else
138             ScrBanks_cm = 1
            End If
        Else
140         ResoDiv_cm = 2
142         BmpBanks_cm = 0
144         ScrBanks_cm = 0
        End If

    End If

    'Call ZoomWindow.ChangeGfxMode
146 Unload Me
        '<EhFooter>
        Exit Sub

btnOk_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.btnOk_Click" + " line: " + Str(Erl))

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


Private Sub chkInterlace_Click()
        '<EhHeader>
        On Error GoTo chkInterlace_Click_Err
        '</EhHeader>

100 If chkInterlace.Value = 1 Then chkSameScreen.Enabled = True Else chkSameScreen.Enabled = False
    'MsgBox (Str(chkInterlace.value))
        '<EhFooter>
        Exit Sub

chkInterlace_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.chkInterlace_Click" + " line: " + Str(Erl))

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

Private Sub optHires_Click()
        '<EhHeader>
        On Error GoTo optHires_Click_Err
        '</EhHeader>

100 chkInterlace.Enabled = False
102 chkSameScreen.Enabled = False

        '<EhFooter>
        Exit Sub

optHires_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.optHires_Click" + " line: " + Str(Erl))

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

Private Sub optMultiColor_Click()
        '<EhHeader>
        On Error GoTo optMultiColor_Click_Err
        '</EhHeader>

100 chkInterlace.Enabled = True
102 If chkInterlace.Value = 1 Then chkSameScreen.Enabled = True Else chkSameScreen.Enabled = False

        '<EhFooter>
        Exit Sub

optMultiColor_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.optMultiColor_Click" + " line: " + Str(Erl))

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

Private Sub combo1_click()
        '<EhHeader>
        On Error GoTo combo1_click_Err
        '</EhHeader>

100 Select Case Combo1.ListIndex
        Case 0
102         FliMul_cm = 1
104     Case 1
106         FliMul_cm = 2
108     Case 2
110         FliMul_cm = 4
112     Case 3
114        FliMul_cm = 8
    End Select

        '<EhFooter>
        Exit Sub

combo1_click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.combo1_click" + " line: " + Str(Erl))

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


100 Combo1.AddItem "FLI every 8th line"
102 Combo1.AddItem "FLI every 4th line"
104 Combo1.AddItem "FLI every 2nd line"
106 Combo1.AddItem "FLI every line"


        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.form_load" + " line: " + Str(Erl))

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

100 Set Me.Icon = ZoomWindow.Image1.Picture


102 chkInterlace.Value = BmpBanks_cm
104 chkSameScreen.Value = ScrBanks_cm

106 If BaseMode_cm = BaseModeTyp.multi Then chkInterlace.Enabled = True Else chkInterlace.Enabled = False
108 If chkInterlace.Value = 1 Then chkSameScreen.Enabled = True Else chkSameScreen.Enabled = False

110 Select Case FliMul_cm
    Case 8
112     Combo1.ListIndex = 0
114 Case 4
116     Combo1.ListIndex = 1
118 Case 2
120     Combo1.ListIndex = 2
122 Case 1
124     Combo1.ListIndex = 3
    End Select

126 If BaseMode_cm = BaseModeTyp.hires Then
128     optHires.Value = True
130 ElseIf BaseMode_cm = BaseModeTyp.multi Then
132     optMultiColor.Value = True
    End If


134 If BmpBanks_cm = 1 Or (BmpBanks_cm = 0 And ResoDiv_cm = 2) Then

136     optMultiColor.Value = True

138     If BmpBanks_cm = 1 Then chkInterlace.Value = 1 Else chkInterlace.Value = 0
140     If ScrBanks_cm = 0 Then chkSameScreen.Value = 1 Else chkSameScreen.Value = 0

    End If

        '<EhFooter>
        Exit Sub

form_activate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Custom.form_activate" + " line: " + Str(Erl))

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
