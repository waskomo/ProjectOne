VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ZoomScaleDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1230
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider ZoomSlider 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Min             =   4
      Max             =   16
      SelStart        =   4
      TickStyle       =   3
      Value           =   4
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label ZoomScaleLabel 
      Caption         =   "16"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label ZoomScale 
      Alignment       =   2  'Center
      Caption         =   "ZoomScale"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "ZoomScaleDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldZoomScale As Long
'Option Explicit


Private Sub form_unload(Cancel As Integer)
'Form1.ZoomScale = OldZoomScale
'Unload ZoomScaleDialog
'Call CancelButton_Click
End Sub


Private Sub CancelButton_Click()
Form1.ZoomScale = OldZoomScale
Call Form1.Form_Resize
Unload ZoomScaleDialog
End Sub

Private Sub OKButton_Click()
Unload ZoomScaleDialog
End Sub

Private Sub ZoomSlider_Change()
ZoomScaleLabel.Caption = ZoomSlider.Value
Form1.ZoomScale = ZoomSlider.Value
Call Form1.Form_Resize
End Sub

Private Sub ZoomSlider_scroll()
ZoomScaleLabel.Caption = ZoomSlider.Value
Form1.ZoomScale = ZoomSlider.Value
Call Form1.Form_Resize
End Sub

Private Sub form_load()
ZoomSlider.Value = Form1.ZoomScale
ZoomScaleLabel.Caption = ZoomSlider.Value
OldZoomScale = Form1.ZoomScale
End Sub
