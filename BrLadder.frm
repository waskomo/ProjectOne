VERSION 5.00
Begin VB.Form BrLadder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brightness Palette Setup"
   ClientHeight    =   4784
   ClientLeft      =   39
   ClientTop       =   325
   ClientWidth     =   6396
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4784
   ScaleWidth      =   6396
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DeleteSelected 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton NewPreset 
      Caption         =   "Save As"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   728
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   6135
   End
   Begin VB.CommandButton SaveOverSelected 
      Caption         =   "Overwrite Current"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      ScaleHeight     =   260
      ScaleWidth      =   260
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Palette: "
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      Begin VB.Image Image1 
         Height          =   255
         Index           =   47
         Left            =   5640
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   46
         Left            =   5280
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   45
         Left            =   4920
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   44
         Left            =   4560
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   43
         Left            =   4200
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   42
         Left            =   3840
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   41
         Left            =   3480
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   40
         Left            =   3120
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   39
         Left            =   2760
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   38
         Left            =   2400
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   37
         Left            =   2040
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   36
         Left            =   1680
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   35
         Left            =   1320
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   34
         Left            =   960
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   33
         Left            =   600
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   32
         Left            =   240
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preset: "
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   2
         Left            =   1920
         Max             =   32
         Min             =   2
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Value           =   2
         Width           =   3975
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   31
         Left            =   5640
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   30
         Left            =   5280
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   29
         Left            =   4920
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   28
         Left            =   4560
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   27
         Left            =   4200
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   26
         Left            =   3840
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   25
         Left            =   3480
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   24
         Left            =   3120
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   23
         Left            =   2760
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   22
         Left            =   2400
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   21
         Left            =   2040
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   20
         Left            =   1680
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   19
         Left            =   1320
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   18
         Left            =   960
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   17
         Left            =   600
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   16
         Left            =   240
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   15
         Left            =   5640
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   14
         Left            =   5280
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   4920
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   4560
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   4200
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   3840
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   3480
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   3120
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   2760
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   2400
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2040
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1680
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1320
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   960
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   600
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   240
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Entrys: "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Presets:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   2535
   End
End
Attribute VB_Name = "BrLadder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ActiveEntry As Long
Dim ActiveIndex As Long


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub InitLadder()
Dim X As Long

For X = 0 To 31
    Picture1.BackColor = PaletteRGB(BrLadderTab(X))
    Image1(X).Picture = Picture1.image
    Image1(X).Appearance = vb3D
    Image1(X).BorderStyle = vbFixedSingle
Next X
EnableEntrys
HScroll1.Value = BrLadderMax
SetCap

End Sub

Private Sub DeleteSelected_Click()
If List1.ListIndex >= 0 Then
Call DeletePreset
End If
End Sub

Private Sub form_activate()
Dim X As Long

Call LoadPreset

For X = 0 To 15
    Picture1.BackColor = PaletteRGB(X)
    Image1(32 + X).Picture = Picture1.image
    Image1(32 + X).Appearance = vb3D
    Image1(32 + X).BorderStyle = vbFixedSingle
Next X

Call InitLadder
LoadPresetList

End Sub

Private Sub EnableEntrys()
Dim X As Long

For X = 0 To 31
    Image1(X).Visible = False
Next X

For X = 0 To BrLadderMax - 1
    Image1(X).Visible = True
Next X

End Sub

Private Sub SetCap()

Label1.Caption = "Number of Entrys: " + Str(BrLadderMax)

End Sub
Private Sub HScroll1_Change()

BrLadderMax = HScroll1.Value
EnableEntrys
SetCap

End Sub
Private Sub HScroll1_scroll()

HScroll1_Change

End Sub

Private Sub Image1_Click(Index As Integer)
Dim X As Long

If Index >= 0 And Index <= 31 Then

    For X = 0 To 31
        Image1(X).Appearance = vb3D
        Image1(X).BorderStyle = vbFixedSingle
    Next X
    
    ActiveEntry = Index
    Image1(ActiveEntry).Appearance = vbFlat
    Image1(ActiveEntry).BorderStyle = vbFixedSingle
    
    Picture1.BackColor = PaletteRGB(ActiveIndex)
    Image1(ActiveEntry).Picture = Picture1.image
    BrLadderTab(ActiveEntry) = ActiveIndex

End If

If Index >= 32 And Index <= 47 Then
 
    For X = 32 To 47
        Image1(X).Appearance = vb3D
        Image1(X).BorderStyle = vbFixedSingle
    Next X
 
    ActiveIndex = Index - 32
    Image1(ActiveIndex + 32).Appearance = vbFlat
    Image1(ActiveIndex + 32).BorderStyle = vbFixedSingle

    Picture1.BackColor = PaletteRGB(ActiveIndex)
    Image1(ActiveEntry).Picture = Picture1.image
    BrLadderTab(ActiveEntry) = ActiveIndex
    
End If
End Sub

Private Sub form_load()

'    Set Me.Icon = ZoomWindow.Image1.Picture

End Sub

Private Sub LoadPreset()
Dim PresetCount As Long
Dim ColorCount As Long
Dim X As Long
Dim Y As Long
Dim Name As String

With m_cIni
    
    .Path = App.Path & "\Ladders.ini"
    .Section = "main"
    .Key = "presetcount": .Default = 0: PresetCount = .Value
    
    If PresetCount <> 0 Then
            .Key = "Name" & format(List1.ListIndex): .Default = "": Name = .Value
            If Name <> "" Then
                .Section = Name
                .Key = "colorcount": .Default = 0: ColorCount = .Value
                If ColorCount <> 0 Then
                    For Y = 1 To ColorCount
                        .Key = "colorindex" & format(Y): .Default = 0: BrLadderTab(Y) = .Value
                    Next Y
                    BrLadderMax = ColorCount
                End If
            End If
    End If

End With

Frame1.Caption = "Selected preset: " + Name

End Sub

Private Sub LoadPresetList()

Dim PresetCount As Long
Dim ColorCount As Long
Dim X As Long
Dim Y As Long
Dim Name As String

List1.Clear

With m_cIni
    
    .Path = App.Path & "\Ladders.ini"
    .Section = "main"
    .Key = "presetcount": .Default = 0: PresetCount = .Value
    
    If PresetCount <> 0 Then
        For X = 0 To PresetCount
            .Key = "Name" & format(X): .Default = "": Name = .Value
            If Name <> "" Then
            List1.AddItem Name
            End If
         Next X
    End If

End With

End Sub

Private Sub SavePreset(ByRef Name As String)

Dim PresetCount As Long
Dim ColorCount As Long
Dim X As Long
Dim Y As Long


If Name = "overwriteoldindex" Then


With m_cIni
    
    .Path = App.Path & "\Ladders.ini"
    .Section = "main"
    
    .Key = "Name" & format(List1.ListIndex):   Name = .Value
    .Section = Name
    .DeleteSection
    .Key = "colorcount": .Value = BrLadderMax
    For Y = 0 To BrLadderMax - 1
        .Key = "colorindex" & format(Y): .Value = BrLadderTab(Y)
    Next Y

End With


Else


With m_cIni
    
    .Path = App.Path & "\Ladders.ini"
    .Section = "main"
    .Key = "presetcount":
    .Default = 0
    PresetCount = .Value
    For X = 0 To PresetCount - 1
        .Key = "Name" & format(X)
        If StrComp(.Value, Name, vbTextCompare) = 0 Then
            MsgBox "A preset with this name does already exist", vbExclamation, "Warning"
            Name = "alreadyexisted"
            Exit Sub
        End If
    Next X
        
    .Key = "Name" & format(PresetCount):  .Value = Name
    PresetCount = PresetCount + 1
    .Key = "presetcount"
    .Value = PresetCount
        
    .Section = Name
    .DeleteSection
    .Key = "colorcount": .Value = BrLadderMax
    For Y = 0 To BrLadderMax - 1
        .Key = "colorindex" & format(Y): .Value = BrLadderTab(Y)
    Next Y

End With


End If

End Sub

Private Sub DeletePreset()

Dim PresetCount As Long
Dim ColorCount As Long
Dim X As Long
Dim Y As Long
Dim Name As String
Dim Temp As String
With m_cIni
    
    .Path = App.Path & "\Ladders.ini"
    .Section = "main"
    .Key = "presetcount": .Default = 0: PresetCount = .Value
    
    If PresetCount <> 0 Then
        .Key = "Name" & format(List1.ListIndex):   Name = .Value
        .Section = Name
        .DeleteSection
        .Section = "main"
        .Key = "Name" & format(List1.ListIndex)
        .DeleteKey
        PresetCount = PresetCount - 1
        
        Y = List1.ListIndex + 1
        For X = Y To PresetCount
            .Key = "Name" & format(X)
            Temp = .Value
            .DeleteKey
            .Key = "Name" & format(X - 1)
            .Value = Temp
        Next X
        .Section = "main"
        .Key = "presetcount"
        .Value = PresetCount
    End If

End With

Call LoadPresetList
End Sub

Private Sub List1_Click()

Call LoadPreset
Call InitLadder

End Sub

Private Sub NewPreset_Click()
Dim NewName As String

NewName = InputBox("Please enter a name for the new preset:")
If NewName = "" Then
    MsgBox "You must enter a name.", vbExclamation
    Exit Sub
End If

Call SavePreset(NewName)

If NewName <> "alreadyexisted" Then
    List1.AddItem NewName
    List1.ListIndex = List1.ListCount - 1
End If

End Sub

Private Sub SaveOverSelected_Click()

If List1.ListIndex >= 0 Then
    Call SavePreset("overwriteoldindex")
Else
    MsgBox "You must select a preset you whish to overwrite", vbExclamation
End If
End Sub
