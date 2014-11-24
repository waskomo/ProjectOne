VERSION 5.00
Begin VB.Form Aspect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aspect Ratio Settings"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox fuckY 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox fuckX 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Aspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOk_Click()
        '<EhHeader>
        On Error GoTo btnOk_Click_Err
        '</EhHeader>

100 ARatioX = Val(fuckX.Text)
102 ARatioY = Val(fuckY.Text)

    
104 Call ZoomWindow.ZoomResize
106 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
108 Call ZoomWindow.ZoomWinRefresh
110 Call PrevWin.ResizePrevPic
112 Call PrevWin.ChangeWinSize
114 Call PrevWin.ReDraw
116 Call PrevWin.Form_Resize

118 Unload Me
        '<EhFooter>
        Exit Sub

btnOk_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Aspect.btnOk_Click" + " line: " + Str(Erl))

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
    'user pressed X
        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err
        '</EhHeader>
100 If UnloadMode = 0 Then

    End If

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Aspect.Form_QueryUnload" + " line: " + Str(Erl))

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

100 fuckX.Text = Str(ARatioX)
102 fuckY.Text = Str(ARatioY)

        '<EhFooter>
        Exit Sub

form_activate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Aspect.form_activate" + " line: " + Str(Erl))

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
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Aspect.form_load" + " line: " + Str(Erl))

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

Private Sub fuckx_validate(Cancel As Boolean)
        '<EhHeader>
        On Error GoTo fuckx_validate_Err
        '</EhHeader>

100 fuckX.Text = ClipVals(fuckX.Text)

102 ARatioX = Val(fuckX.Text)
104 ARatioY = Val(fuckY.Text)
106 Call ZoomWindow.ZoomResize
108 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
110 Call ZoomWindow.ZoomWinRefresh
112 Call PrevWin.ResizePrevPic
114 Call PrevWin.ChangeWinSize
116 Call PrevWin.ReDraw
118 Call PrevWin.Form_Resize

        '<EhFooter>
        Exit Sub

fuckx_validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Aspect.fuckx_validate" + " line: " + Str(Erl))

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

Private Sub fucky_validate(Cancel As Boolean)
        '<EhHeader>
        On Error GoTo fucky_validate_Err
        '</EhHeader>

100 fuckY.Text = ClipVals(fuckY.Text)

102 ARatioX = Val(fuckX.Text)
104 ARatioY = Val(fuckY.Text)
106 Call ZoomWindow.ZoomResize
108 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
110 Call ZoomWindow.ZoomWinRefresh
112 Call PrevWin.ResizePrevPic
114 Call PrevWin.ChangeWinSize
116 Call PrevWin.ReDraw
118 Call PrevWin.Form_Resize

        '<EhFooter>
        Exit Sub

fucky_validate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Aspect.fucky_validate" + " line: " + Str(Erl))

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

Private Function ClipVals(Value As String) As String
        '<EhHeader>
        On Error GoTo ClipVals_Err
        '</EhHeader>

100 ClipVals = Value
102 If Val(Value) < ARatioMin Then ClipVals = Str(1)
104 If Val(Value) > ARatioMax Then ClipVals = Str(4)


        '<EhFooter>
        Exit Function

ClipVals_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.Aspect.ClipVals" + " line: " + Str(Erl))

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
