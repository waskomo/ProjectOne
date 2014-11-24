VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form GridOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grid Settings"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "FLI Grid:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   2535
      Begin VB.PictureBox pbFliGridColor 
         Height          =   375
         Left            =   1200
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton btnFliGridColor 
         Caption         =   "Set Color"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chkShowFliLines 
         Caption         =   "Show"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin MSComctlLib.Slider sldFliGrid 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   16
         SelStart        =   1
         TickStyle       =   3
         Value           =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Hide if zoomrate smaller than:"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1793
      TabIndex        =   10
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Misc.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2640
      TabIndex        =   8
      Top             =   2280
      Width           =   2535
      Begin VB.CheckBox chkZoomKeret 
         Caption         =   "Box around zoomarea"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkFLIBox 
         Caption         =   "Show FLI Box"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkPixelBoxShowColor 
         Caption         =   "Extra Color Info"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox chkPixelBox 
         Caption         =   "Highlight active pixel"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComctlLib.Slider sldPixelGrid 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   16
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin MSComctlLib.Slider sldCharGrid 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   16
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.CheckBox chkPixelGrid 
      Caption         =   "Show"
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox chkChargrid 
      Caption         =   "Show"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Character Grid:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox pbCharGridColor 
         Height          =   375
         Left            =   1200
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton btnCharGridColor 
         Caption         =   "Set Color"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Hide if zoomrate smaller than:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pixel Grid:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox pbPixelGridColor 
         Height          =   375
         Left            =   1200
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton btnPixelGridColor 
         Caption         =   "Set Color"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Hide if zoomrate smaller than:"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
   End
End
Attribute VB_Name = "GridOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnCharGridColor_Click()
        '<EhHeader>
        On Error GoTo btnCharGridColor_Click_Err
        '</EhHeader>

100 pbCharGridColor.BackColor = ColorDialog()
102 CharGridColor = pbCharGridColor.BackColor

104 UpdateChanges
        '<EhFooter>
        Exit Sub

btnCharGridColor_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.btnCharGridColor_Click" + " line: " + Str(Erl))

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

Private Function ColorDialog() As Long

' Sets the Dialog Title to Save File
ZoomWindow.CommonDialog1.DialogTitle = "Select Desired Color:" 'Select Colour&quot;

ZoomWindow.CommonDialog1.CancelError = True
' Set flags - enabled the Custom Color button
ZoomWindow.CommonDialog1.Flags = cdlCCFullOpen
' Enables error handling to catch cancel error
On Error Resume Next
' display the set colour dialog box
ZoomWindow.CommonDialog1.ShowColor

'If err Then
'&nbsp;&nbsp;&nbsp; ' This code runs if the dialog was cancelled
'&nbsp;&nbsp;&nbsp; Msgbox &quot;Dialog Cancelled&quot;
'&nbsp;&nbsp;&nbsp; Exit Sub
'End If
' Sets ZoomWinReDraw background color to the selected colour

ColorDialog = ZoomWindow.CommonDialog1.Color

End Function

Private Sub btnPixelGridColor_Click()
        '<EhHeader>
        On Error GoTo btnPixelGridColor_Click_Err
        '</EhHeader>

100 pbPixelGridColor.BackColor = ColorDialog()
102 PixelGridColor = pbPixelGridColor.BackColor
104 UpdateChanges
        '<EhFooter>
        Exit Sub

btnPixelGridColor_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.btnPixelGridColor_Click" + " line: " + Str(Erl))

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

Private Sub btnfliGridColor_Click()
        '<EhHeader>
        On Error GoTo btnfliGridColor_Click_Err
        '</EhHeader>

100 pbFliGridColor.BackColor = ColorDialog()
102 FliGridColor = pbFliGridColor.BackColor
104 UpdateChanges
        '<EhFooter>
        Exit Sub

btnfliGridColor_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.btnfliGridColor_Click" + " line: " + Str(Erl))

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

Private Sub chkChargrid_Click()
        '<EhHeader>
        On Error GoTo chkChargrid_Click_Err
        '</EhHeader>
100 CharGrid = chkChargrid.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

chkChargrid_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.chkChargrid_Click" + " line: " + Str(Erl))

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

Private Sub chkPixelBox_Click()
        '<EhHeader>
        On Error GoTo chkPixelBox_Click_Err
        '</EhHeader>
100 PixelBox = chkPixelBox.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

chkPixelBox_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.chkPixelBox_Click" + " line: " + Str(Erl))

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

Private Sub chkfliBox_Click()
        '<EhHeader>
        On Error GoTo chkfliBox_Click_Err
        '</EhHeader>
100 ShowFlibox = chkFLIBox.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

chkfliBox_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.chkfliBox_Click" + " line: " + Str(Erl))

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
Private Sub chkPixelBoxShowColor_Click()
        '<EhHeader>
        On Error GoTo chkPixelBoxShowColor_Click_Err
        '</EhHeader>

100 PixelBoxColor = chkPixelBoxShowColor.Value

        '<EhFooter>
        Exit Sub

chkPixelBoxShowColor_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.chkPixelBoxShowColor_Click" + " line: " + Str(Erl))

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

Private Sub chkPixelGrid_Click()
        '<EhHeader>
        On Error GoTo chkPixelGrid_Click_Err
        '</EhHeader>

100 PixelGrid = chkPixelGrid.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

chkPixelGrid_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.chkPixelGrid_Click" + " line: " + Str(Erl))

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

Private Sub chkShowFliLines_Click()
        '<EhHeader>
        On Error GoTo chkShowFliLines_Click_Err
        '</EhHeader>

100 ShowFliLines = chkShowFliLines.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

chkShowFliLines_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.chkShowFliLines_Click" + " line: " + Str(Erl))

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

Private Sub chkZoomKeret_Click()
        '<EhHeader>
        On Error GoTo chkZoomKeret_Click_Err
        '</EhHeader>
100 ZoomKeret = chkZoomKeret.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

chkZoomKeret_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.chkZoomKeret_Click" + " line: " + Str(Erl))

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
100 Unload Me
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.Command1_Click" + " line: " + Str(Erl))

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
Private Sub UpdateChanges()
        '<EhHeader>
        On Error GoTo UpdateChanges_Err
        '</EhHeader>

100 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
102 Call ZoomWindow.ZoomWinRefresh
104 Call PrevWin.DrawCursors(OldAx, OldAy)

        '<EhFooter>
        Exit Sub

UpdateChanges_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.UpdateChanges" + " line: " + Str(Erl))

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


Private Sub sldCharGrid_change()
        '<EhHeader>
        On Error GoTo sldCharGrid_change_Err
        '</EhHeader>
100 CharGridLimit = sldCharGrid.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

sldCharGrid_change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.sldCharGrid_change" + " line: " + Str(Erl))

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

Private Sub sldCharGrid_scroll()
        '<EhHeader>
        On Error GoTo sldCharGrid_scroll_Err
        '</EhHeader>
100 CharGridLimit = sldCharGrid.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

sldCharGrid_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.sldCharGrid_scroll" + " line: " + Str(Erl))

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

Private Sub sldFliGrid_Click()
        '<EhHeader>
        On Error GoTo sldFliGrid_Click_Err
        '</EhHeader>
100 FliGridLimit = sldFliGrid.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

sldFliGrid_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.sldFliGrid_Click" + " line: " + Str(Erl))

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

Private Sub sldpixelGrid_change()
        '<EhHeader>
        On Error GoTo sldpixelGrid_change_Err
        '</EhHeader>
100 PixelGridLimit = sldPixelGrid.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

sldpixelGrid_change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.sldpixelGrid_change" + " line: " + Str(Erl))

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
Private Sub sldpixelGrid_scroll()
        '<EhHeader>
        On Error GoTo sldpixelGrid_scroll_Err
        '</EhHeader>
100 PixelGridLimit = sldPixelGrid.Value
102 UpdateChanges
        '<EhFooter>
        Exit Sub

sldpixelGrid_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.sldpixelGrid_scroll" + " line: " + Str(Erl))

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

100 Me.Refresh

102 pbCharGridColor.BackColor = CharGridColor
104 pbFliGridColor.BackColor = FliGridColor
106 pbPixelGridColor.BackColor = PixelGridColor

108 chkChargrid.Value = CharGrid
110 chkPixelBox.Value = PixelBox
112 chkFLIBox.Value = ShowFlibox
114 chkPixelGrid.Value = PixelGrid
116 chkShowFliLines.Value = ShowFliLines
118 chkZoomKeret.Value = ZoomKeret
120 chkPixelBoxShowColor = PixelBoxColor
 
122 sldPixelGrid.Value = PixelGridLimit
124 sldCharGrid.Value = CharGridLimit
126 sldFliGrid.Value = FliGridLimit






        '<EhFooter>
        Exit Sub

form_activate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.form_activate" + " line: " + Str(Erl))

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
Private Sub form_unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo form_unload_Err
        '</EhHeader>
100 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
102 Call ZoomWindow.ZoomWinRefresh
        '<EhFooter>
        Exit Sub

form_unload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.form_unload" + " line: " + Str(Erl))

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
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.GridOptions.form_load" + " line: " + Str(Erl))

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
