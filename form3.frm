VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ConvertDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Options"
   ClientHeight    =   7020
   ClientLeft      =   2490
   ClientTop       =   1950
   ClientWidth     =   11505
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11505
   Begin VB.Frame TabFrames 
      Caption         =   "Color Tables"
      Height          =   6015
      Index           =   3
      Left            =   6480
      TabIndex        =   76
      Top             =   240
      Width           =   6495
      Begin VB.CommandButton cmdSelectColorTabFname 
         Caption         =   "..."
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   85
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdSelectColorTabFname 
         Caption         =   "..."
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   84
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdSelectColorTabFname 
         Caption         =   "..."
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   83
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtCTabFnames 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtCTabFnames 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtCTabFnames 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lblTheseFiles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "These files must be inside the Application path."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   88
         Top             =   2640
         Width           =   4050
      End
      Begin VB.Label lblFnames 
         Caption         =   "VDC"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   82
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblFnames 
         Caption         =   "TED"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   81
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblFnames 
         Caption         =   "VICII"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   80
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.PictureBox TmpPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7800
      ScaleHeight     =   615
      ScaleWidth      =   1575
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame TabFrames 
      Caption         =   "Main"
      Height          =   6015
      Index           =   0
      Left            =   6720
      TabIndex        =   9
      Top             =   840
      Width           =   6375
      Begin VB.Frame Frame7 
         Caption         =   "Screen Mode "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   42
         Top             =   4320
         Width           =   5775
         Begin VB.ComboBox cmbGfxMode 
            Height          =   315
            Left            =   1920
            TabIndex        =   43
            Text            =   "Combo1"
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resize "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1937
         Left            =   3003
         TabIndex        =   24
         ToolTipText     =   "Defines how to resize the original picture"
         Top             =   2288
         Width           =   2886
         Begin VB.HScrollBar sldResize 
            Height          =   255
            LargeChange     =   20
            Left            =   360
            Max             =   100
            Min             =   1
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1320
            Value           =   100
            Width           =   1935
         End
         Begin VB.CheckBox Check_Stretch 
            Caption         =   "Resize to Fit"
            Height          =   255
            Left            =   720
            TabIndex        =   26
            ToolTipText     =   "Check this to resize the picture"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check_Aspect 
            Caption         =   "Keep Aspect "
            Height          =   375
            Left            =   720
            TabIndex        =   25
            ToolTipText     =   "Check this to keep the original aspect ratio"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Downscale Picture"
            Height          =   255
            Left            =   480
            TabIndex        =   62
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Colors Inside Char"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "This defines how to reduce the colors used in char"
         Top             =   2280
         Width           =   2775
         Begin VB.CheckBox chkFixWithBackgr 
            Caption         =   "Backgr Fixes Clash Bugs"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "Check if you want the background color to be used"
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.OptionButton chkMostFrq3 
            Caption         =   "Most frequent 3"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "Check this to use the most used colors"
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton chkOptChars 
            Caption         =   "Optimized (slow)"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            ToolTipText     =   "Check this for brute force mode - checks every possible combination and picks the best"
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   3000
         TabIndex        =   15
         ToolTipText     =   "Defines how to pick the background color"
         Top             =   240
         Width           =   2886
         Begin VB.CommandButton btnOptimizeNow 
            Caption         =   "Optimize Now"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton chkMostFreqBackgr 
            Caption         =   "Most Frequent Color"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "Check this to pick the most used color as background"
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton chkUsrDefBackgr 
            Caption         =   "User:"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "Check this to define a background color yourself"
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton SelColButton 
            Caption         =   "Select"
            Height          =   375
            Left            =   1680
            TabIndex        =   17
            Top             =   840
            Width           =   735
         End
         Begin VB.PictureBox UserColor 
            Height          =   375
            Left            =   960
            ScaleHeight     =   315
            ScaleWidth      =   555
            TabIndex        =   16
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Resolution Reduction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "This defines how to lower the resolution for multicolor pictures"
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton pixleft 
            Caption         =   "Pixels From Left"
            Height          =   375
            Left            =   360
            TabIndex        =   14
            ToolTipText     =   "Check this to pick every 2nd even pixel"
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton pixright 
            Caption         =   "Pixels From Right"
            Height          =   375
            Left            =   360
            TabIndex        =   13
            ToolTipText     =   "Check this to pick every 2nd odd pixel"
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton pixavg 
            Caption         =   "Average"
            Height          =   375
            Left            =   360
            TabIndex        =   12
            ToolTipText     =   "Check this to pick the average color of each pixel pair"
            Top             =   1320
            Width           =   1335
         End
      End
   End
   Begin VB.Frame TabFrames 
      Caption         =   "Colors"
      Height          =   6015
      Index           =   2
      Left            =   6120
      TabIndex        =   63
      Top             =   600
      Width           =   6495
      Begin Project_One.MyPalette MyPalette1 
         Height          =   1575
         Left            =   240
         TabIndex        =   73
         Top             =   720
         Width           =   3735
         _ExtentX        =   0
         _ExtentY        =   0
         BoxCount        =   15
         BoxWidth        =   20
         BoxHSpacing     =   4
         BoxTop          =   4
         BoxLeft         =   6
         BoxPerLine      =   8
         BackgrColor     =   16711680
         BoxborderColor  =   16777215
         BoxHasBorders   =   -1  'True
         BoxTextColor    =   16777215
         BoxHasText      =   0
         BoxBorderStyle  =   1
      End
      Begin VB.CommandButton SaveOverSelected 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   5160
         Width           =   1222
      End
      Begin VB.CommandButton DeleteSelected 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2760
         TabIndex        =   64
         Top             =   5160
         Width           =   1183
      End
      Begin VB.CommandButton NewPreset 
         Caption         =   "Save As"
         Height          =   375
         Left            =   1440
         TabIndex        =   65
         Top             =   5160
         Width           =   1222
      End
      Begin VB.Frame Frame8 
         Caption         =   "Preset: "
         Height          =   3341
         Left            =   0
         TabIndex        =   68
         Top             =   2340
         Width           =   4030
         Begin Project_One.MyPalette MyPalette2 
            Height          =   1770
            Left            =   240
            TabIndex        =   72
            Top             =   825
            Width           =   3495
            _ExtentX        =   0
            _ExtentY        =   0
            BoxCount        =   15
            BoxWidth        =   20
            BoxHSpacing     =   8
            BoxTop          =   4
            BoxLeft         =   6
            BoxPerLine      =   8
            BackgrColor     =   16711680
            BoxborderColor  =   0
            BoxHasBorders   =   -1  'True
            BoxTextColor    =   16777215
            BoxHasText      =   0
            BoxBorderStyle  =   1
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            LargeChange     =   5
            Left            =   1638
            Max             =   32
            Min             =   2
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   351
            Value           =   2
            Width           =   2223
         End
         Begin VB.Label Label7 
            Caption         =   "Number of Entrys: "
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.ListBox GradientList 
         Height          =   4935
         Left            =   4200
         TabIndex        =   66
         Top             =   600
         Width           =   2041
      End
      Begin VB.Label Label18 
         Caption         =   "Presets:"
         Height          =   260
         Left            =   4095
         TabIndex        =   71
         Top             =   234
         Width           =   1365
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   75
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Active Palette:"
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame TabFrames 
      Caption         =   "Paletted Brightness"
      Height          =   6015
      Index           =   1
      Left            =   -480
      TabIndex        =   10
      Top             =   120
      Width           =   6255
      Begin VB.Frame Cfefhgdf 
         Caption         =   "Color Filter: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   2895
         Begin VB.ComboBox cmbColorFilter 
            Height          =   315
            Left            =   600
            TabIndex        =   45
            Text            =   "Combo2"
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Dither 
         Caption         =   "Dither: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   3120
         TabIndex        =   33
         Top             =   240
         Width           =   2895
         Begin VB.CheckBox chkRndHue 
            Caption         =   "Random Dither"
            Height          =   255
            Left            =   1080
            TabIndex        =   60
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CheckBox chkRndSat 
            Caption         =   "Random Dither"
            Height          =   255
            Left            =   1080
            TabIndex        =   57
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkRndBright 
            Caption         =   "Random Dither"
            Height          =   255
            Left            =   1080
            TabIndex        =   55
            Top             =   1440
            Width           =   1455
         End
         Begin VB.HScrollBar sldHue 
            Height          =   255
            Left            =   960
            Max             =   3
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1575
         End
         Begin VB.HScrollBar sldSat 
            Height          =   255
            Left            =   960
            Max             =   3
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1575
         End
         Begin VB.HScrollBar sldBright 
            Height          =   255
            Left            =   960
            Max             =   3
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1575
         End
         Begin VB.ComboBox cmbDithHue 
            Height          =   315
            Left            =   960
            TabIndex        =   41
            Text            =   "Combo2"
            Top             =   4080
            Width           =   1575
         End
         Begin VB.ComboBox cmbDithSat 
            Height          =   315
            Left            =   960
            TabIndex        =   40
            Text            =   "Combo2"
            Top             =   2400
            Width           =   1575
         End
         Begin VB.ComboBox cmbDithBright 
            Height          =   315
            Left            =   960
            TabIndex        =   39
            Text            =   "Combo2"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Saturation"
            Height          =   255
            Left            =   840
            TabIndex        =   50
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Fineness:"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   4560
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Fineness:"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Fineness:"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Dither:"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   4080
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Dither:"
            Height          =   255
            Left            =   360
            TabIndex        =   53
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Dither:"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Hue"
            Height          =   255
            Left            =   960
            TabIndex        =   51
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Brightness"
            Height          =   255
            Left            =   600
            TabIndex        =   49
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Color Adjustments "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "This helps you to fine tune the colors"
         Top             =   2160
         Width           =   2895
         Begin MSComctlLib.Slider sldSaturation 
            Height          =   255
            Left            =   480
            TabIndex        =   35
            Top             =   2520
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            Min             =   -255
            Max             =   512
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
         End
         Begin MSComctlLib.Slider sldHuex 
            Height          =   255
            Left            =   480
            TabIndex        =   34
            Top             =   1920
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            Max             =   350
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
         End
         Begin MSComctlLib.Slider sldContrast 
            Height          =   255
            Left            =   480
            TabIndex        =   32
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            Max             =   255
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
         End
         Begin MSComctlLib.Slider sldBrightness 
            Height          =   255
            Left            =   480
            TabIndex        =   31
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            Min             =   -255
            Max             =   255
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
         End
         Begin VB.CommandButton DefaultColorControls 
            Caption         =   "Default"
            Height          =   375
            Left            =   840
            TabIndex        =   28
            ToolTipText     =   "Click this to reset the sliders"
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Saturation"
            Height          =   255
            Left            =   480
            TabIndex        =   37
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Hue"
            Height          =   255
            Left            =   480
            TabIndex        =   36
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Brightness"
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Contrast"
            Height          =   255
            Left            =   480
            TabIndex        =   29
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.Frame TabFrames 
      Caption         =   "4 color mode"
      Height          =   6015
      Index           =   4
      Left            =   -240
      TabIndex        =   86
      Top             =   600
      Width           =   6495
   End
   Begin VB.PictureBox Backgr 
      Height          =   375
      Left            =   8160
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.PictureBox DstPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   6480
      ScaleHeight     =   149.269
      ScaleMode       =   0  'User
      ScaleWidth      =   415.022
      TabIndex        =   4
      Top             =   3360
      Width           =   4455
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Apply 
      Caption         =   "Apply Changes"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
   End
   Begin VB.PictureBox SrcPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   6480
      MouseIcon       =   "form3.frx":0000
      ScaleHeight     =   149.269
      ScaleMode       =   0  'User
      ScaleWidth      =   415.022
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   6480
      Width           =   1695
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11033
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Main"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Colors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gradients"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color Tables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "4 color mode"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Background color:"
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Gfx Mode:"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "ConvertDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'conversion pipeline:

'1. zoomwin.mnuloadpicture, loads pic, puts picture into prevpic picture object
'2. convertdialog

'these two are for brightness ladder tab
Dim ActiveEntry As Long
Dim ActiveIndex As Long

Dim DragStart As Boolean
Dim Startx As Single
Dim Starty As Single

Public xD As Long
Public yD As Long
Public xA As Long
Public yA As Long

'Stretchblt
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub EnableControls()
        '<EhHeader>
        On Error GoTo EnableControls_Err
        '</EhHeader>
    Dim ocontrol As Object

100 For Each ocontrol In Me.Controls
102  ocontrol.Enabled = True
104 Next ocontrol



106 DithControlsEnabled
        '<EhFooter>
        Exit Sub

EnableControls_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.EnableControls" + " line: " + Str(Erl))

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
Private Sub DisableControls()
        '<EhHeader>
        On Error GoTo DisableControls_Err
        '</EhHeader>
    Dim ocontrol As Object

100 For Each ocontrol In Me.Controls
102  ocontrol.Enabled = False
104 Next ocontrol

        '<EhFooter>
        Exit Sub

DisableControls_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.DisableControls" + " line: " + Str(Erl))

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

Private Sub Apply_Click()
        '<EhHeader>
        On Error GoTo Apply_Click_Err
        '</EhHeader>

100 DisableControls

102 Call ConvertPic
104 Call ReDrawDstPic

106 EnableControls

        '<EhFooter>
        Exit Sub

Apply_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.Apply_Click" + " line: " + Str(Erl))

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



Private Sub cmdSelectColorTabFname_Click(Index As Integer)

ZoomWindow.CommonDialog1.CancelError = True
On Error GoTo errhandler
ZoomWindow.CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
ZoomWindow.CommonDialog1.Filter = "(*.txt)|*.txt|(*.*)|*.*"
ZoomWindow.CommonDialog1.FilterIndex = 1
ZoomWindow.CommonDialog1.FileName = ColorTable_Filename(Index)
ZoomWindow.CommonDialog1.ShowOpen

txtCTabFnames(Index).Text = GetFileNameFromFullPath(ZoomWindow.CommonDialog1.FileName)
ColorTable_Filename(Index) = GetFileNameFromFullPath(ZoomWindow.CommonDialog1.FileName)

On Error GoTo 0


errhandler:

End Sub

Private Sub SrcPic_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo SrcPic_mousedown_Err
        '</EhHeader>

    
100     If Button = 1 Then
102         DragStart = True
104         Startx = X
106         Starty = Y
108         SrcPic.MousePointer = vbCustom
110         SrcPic.MouseIcon = LoadPicture(App.Path & "\Cursors\handgrab.ico")
        End If

        '<EhFooter>
        Exit Sub

SrcPic_mousedown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SrcPic_mousedown" + " line: " + Str(Erl))

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


Private Sub SrcPic_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo SrcPic_mouseup_Err
        '</EhHeader>

100 DragStart = False

102 xA = xA + (X - Startx)
104 yA = yA + (Y - Starty)
106 Call ResizePic
108 Call PicAdjust

110 If Button = 1 Then
112     SrcPic.MousePointer = vbCustom
114     SrcPic.MouseIcon = LoadPicture(App.Path & "\Cursors\handflat.ico")
    End If

        '<EhFooter>
        Exit Sub

SrcPic_mouseup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SrcPic_mouseup" + " line: " + Str(Erl))

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
Private Sub SrcPic_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo SrcPic_mousemove_Err
        '</EhHeader>

100 If DragStart = True Then

102     xD = xA + (X - Startx)
104     yD = yA + (Y - Starty)

106     Call ResizePic

    End If

        '<EhFooter>
        Exit Sub

SrcPic_mousemove_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SrcPic_mousemove" + " line: " + Str(Erl))

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



Private Sub btnOptimizeNow_Click()
        '<EhHeader>
        On Error GoTo btnOptimizeNow_Click_Err
        '</EhHeader>

100 OptBackgr = True
102 Call Apply_Click
104 OptBackgr = False
        '<EhFooter>
        Exit Sub

btnOptimizeNow_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.btnOptimizeNow_Click" + " line: " + Str(Erl))

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

Private Sub chkRndBright_Click()
        '<EhHeader>
        On Error GoTo chkRndBright_Click_Err
        '</EhHeader>
100 BRnd = chkRndBright.Value
        '<EhFooter>
        Exit Sub

chkRndBright_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.chkRndBright_Click" + " line: " + Str(Erl))

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
Private Sub chkRndSat_Click()
        '<EhHeader>
        On Error GoTo chkRndSat_Click_Err
        '</EhHeader>
100 SRnd = chkRndSat.Value
        '<EhFooter>
        Exit Sub

chkRndSat_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.chkRndSat_Click" + " line: " + Str(Erl))

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
Private Sub chkRndhue_Click()
        '<EhHeader>
        On Error GoTo chkRndhue_Click_Err
        '</EhHeader>
100 HRnd = chkRndHue.Value
        '<EhFooter>
        Exit Sub

chkRndhue_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.chkRndhue_Click" + " line: " + Str(Erl))

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
Private Sub cmbColorFilter_Change()
        '<EhHeader>
        On Error GoTo cmbColorFilter_Change_Err
        '</EhHeader>
100     ColorFilterMode = cmbColorFilter.ListIndex
102     DithControlsEnabled
        '<EhFooter>
        Exit Sub

cmbColorFilter_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbColorFilter_Change" + " line: " + Str(Erl))

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
Private Sub cmbColorFilter_Click()
        '<EhHeader>
        On Error GoTo cmbColorFilter_Click_Err
        '</EhHeader>
100 Call cmbColorFilter_Change
        '<EhFooter>
        Exit Sub

cmbColorFilter_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbColorFilter_Click" + " line: " + Str(Erl))

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
Private Sub DithControlsEnabled()
        '<EhHeader>
        On Error GoTo DithControlsEnabled_Err
        '</EhHeader>
100 Select Case cmbColorFilter.ListIndex
    Case 0
102 Call BrightDithEnabled(True)
104 Call HueDithEnabled(True)
106 Call SatDithEnabled(True)
108 Case 1
110 Call BrightDithEnabled(True)
112 Call HueDithEnabled(False)
114 Call SatDithEnabled(True)
116 Case 2
118 Call BrightDithEnabled(False)
120 Call HueDithEnabled(False)
122 Call SatDithEnabled(False)
124 Case 3
126 Call BrightDithEnabled(False)
128 Call HueDithEnabled(False)
130 Call SatDithEnabled(False)
132 Case 4
134 Call BrightDithEnabled(True)
136 Call HueDithEnabled(False)
138 Call SatDithEnabled(False)
    End Select

    'set enable/disable state of picturedownsize slider & keep aspect

140 Check_Aspect.Enabled = StretchPic

142 sldResize.Enabled = Not StretchPic
144 Label17.Enabled = Not StretchPic
    
        '<EhFooter>
        Exit Sub

DithControlsEnabled_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.DithControlsEnabled" + " line: " + Str(Erl))

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
Private Sub cmbDithBright_Change()
        '<EhHeader>
        On Error GoTo cmbDithBright_Change_Err
        '</EhHeader>
100 Bdither = cmbDithBright.ListIndex
        '<EhFooter>
        Exit Sub

cmbDithBright_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbDithBright_Change" + " line: " + Str(Erl))

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
Private Sub cmbDithBright_Click()
        '<EhHeader>
        On Error GoTo cmbDithBright_Click_Err
        '</EhHeader>
100 Bdither = cmbDithBright.ListIndex
        '<EhFooter>
        Exit Sub

cmbDithBright_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbDithBright_Click" + " line: " + Str(Erl))

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
Private Sub cmbDithSat_Change()
        '<EhHeader>
        On Error GoTo cmbDithSat_Change_Err
        '</EhHeader>
100 Sdither = cmbDithSat.ListIndex
        '<EhFooter>
        Exit Sub

cmbDithSat_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbDithSat_Change" + " line: " + Str(Erl))

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
Private Sub cmbDithSat_Click()
        '<EhHeader>
        On Error GoTo cmbDithSat_Click_Err
        '</EhHeader>
100 Sdither = cmbDithSat.ListIndex
        '<EhFooter>
        Exit Sub

cmbDithSat_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbDithSat_Click" + " line: " + Str(Erl))

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
Private Sub cmbDithHue_Change()
        '<EhHeader>
        On Error GoTo cmbDithHue_Change_Err
        '</EhHeader>
100 Hdither = cmbDithHue.ListIndex
        '<EhFooter>
        Exit Sub

cmbDithHue_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbDithHue_Change" + " line: " + Str(Erl))

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
Private Sub cmbDithHue_Click()
        '<EhHeader>
        On Error GoTo cmbDithHue_Click_Err
        '</EhHeader>
100 Hdither = cmbDithHue.ListIndex
        '<EhFooter>
        Exit Sub

cmbDithHue_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbDithHue_Click" + " line: " + Str(Erl))

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
Private Sub DefaultColorControls_Click()
        '<EhHeader>
        On Error GoTo DefaultColorControls_Click_Err
        '</EhHeader>
100 sldBrightness.Value = 0
102 sldContrast.Value = 128
104 sldSaturation.Value = 128
106 sldHuex.Value = 0
108 Call PicAdjust
        '<EhFooter>
        Exit Sub

DefaultColorControls_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.DefaultColorControls_Click" + " line: " + Str(Erl))

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
Private Sub BrightDithEnabled(Setting As Boolean)
        '<EhHeader>
        On Error GoTo BrightDithEnabled_Err
        '</EhHeader>
100 Label3.Enabled = Setting
102 Label11.Enabled = Setting
104 Label14.Enabled = Setting
106 cmbDithBright.Enabled = Setting
108 sldBright.Enabled = Setting
110 chkRndBright.Enabled = Setting
        '<EhFooter>
        Exit Sub

BrightDithEnabled_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.BrightDithEnabled" + " line: " + Str(Erl))

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
Private Sub SatDithEnabled(Setting As Boolean)
        '<EhHeader>
        On Error GoTo SatDithEnabled_Err
        '</EhHeader>
100 Label4.Enabled = Setting
102 Label12.Enabled = Setting
104 Label15.Enabled = Setting
106 cmbDithSat.Enabled = Setting
108 sldSat.Enabled = Setting
110 chkRndSat.Enabled = Setting
        '<EhFooter>
        Exit Sub

SatDithEnabled_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SatDithEnabled" + " line: " + Str(Erl))

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
Private Sub HueDithEnabled(Setting As Boolean)
        '<EhHeader>
        On Error GoTo HueDithEnabled_Err
        '</EhHeader>
100 Label8.Enabled = Setting
102 Label13.Enabled = Setting
104 Label16.Enabled = Setting
106 cmbDithHue.Enabled = Setting
108 sldHue.Enabled = Setting
110 chkRndHue.Enabled = Setting
        '<EhFooter>
        Exit Sub

HueDithEnabled_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.HueDithEnabled" + " line: " + Str(Erl))

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
102     Call ZoomWindow.ReadUndo
104     Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
106     Call ZoomWindow.ZoomWinRefresh
    End If
        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.Form_QueryUnload" + " line: " + Str(Erl))

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
Private Sub Cancel_Click()
        '<EhHeader>
        On Error GoTo Cancel_Click_Err
        '</EhHeader>
100 Call ZoomWindow.ReadUndo
102 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
104 Call ZoomWindow.ZoomWinRefresh
106 Unload Me
        '<EhFooter>
        Exit Sub

Cancel_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.Cancel_Click" + " line: " + Str(Erl))

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
Private Sub chkMostFrq3_Click()
        '<EhHeader>
        On Error GoTo chkMostFrq3_Click_Err
        '</EhHeader>
100 MostFreqBackgr = chkMostFreqBackgr.Value
        '<EhFooter>
        Exit Sub

chkMostFrq3_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.chkMostFrq3_Click" + " line: " + Str(Erl))

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
Private Sub chkFixWithBackgr_Click()
        '<EhHeader>
        On Error GoTo chkFixWithBackgr_Click_Err
        '</EhHeader>
100 FixWithBackgr = chkFixWithBackgr.Value
        '<EhFooter>
        Exit Sub

chkFixWithBackgr_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.chkFixWithBackgr_Click" + " line: " + Str(Erl))

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
Private Sub chkOptChars_Click()
        '<EhHeader>
        On Error GoTo chkOptChars_Click_Err
        '</EhHeader>
100 OptChars = chkOptChars.Value
        '<EhFooter>
        Exit Sub

chkOptChars_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.chkOptChars_Click" + " line: " + Str(Erl))

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
Private Sub chkUsrDefBackgr_Click()
        '<EhHeader>
        On Error GoTo chkUsrDefBackgr_Click_Err
        '</EhHeader>
100 UsrDefBackgr = chkUsrDefBackgr.Value
        '<EhFooter>
        Exit Sub

chkUsrDefBackgr_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.chkUsrDefBackgr_Click" + " line: " + Str(Erl))

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
Private Sub pixavg_Click()
        '<EhHeader>
        On Error GoTo pixavg_Click_Err
        '</EhHeader>
100 Reso_PixAvg = pixavg.Value
102 Reso_PixRight = pixright.Value
104 Reso_PixLeft = pixleft.Value
        '<EhFooter>
        Exit Sub

pixavg_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.pixavg_Click" + " line: " + Str(Erl))

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
Private Sub pixleft_Click()
        '<EhHeader>
        On Error GoTo pixleft_Click_Err
        '</EhHeader>
100 Reso_PixAvg = pixavg.Value
102 Reso_PixRight = pixright.Value
104 Reso_PixLeft = pixleft.Value
        '<EhFooter>
        Exit Sub

pixleft_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.pixleft_Click" + " line: " + Str(Erl))

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
Private Sub pixright_Click()
        '<EhHeader>
        On Error GoTo pixright_Click_Err
        '</EhHeader>
100 Reso_PixAvg = pixavg.Value
102 Reso_PixRight = pixright.Value
104 Reso_PixLeft = pixleft.Value
        '<EhFooter>
        Exit Sub

pixright_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.pixright_Click" + " line: " + Str(Erl))

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
Private Sub sldBright_Change()
        '<EhHeader>
        On Error GoTo sldBright_Change_Err
        '</EhHeader>
100 BditherVal = sldBright.Value
        '<EhFooter>
        Exit Sub

sldBright_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldBright_Change" + " line: " + Str(Erl))

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
Private Sub sldBright_scroll()
        '<EhHeader>
        On Error GoTo sldBright_scroll_Err
        '</EhHeader>
100 BditherVal = sldBright.Value
        '<EhFooter>
        Exit Sub

sldBright_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldBright_scroll" + " line: " + Str(Erl))

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
Private Sub sldResize_Change()
        '<EhHeader>
        On Error GoTo sldResize_Change_Err
        '</EhHeader>
100 ResizeScale = sldResize.Value
102 ResizeScaleX = ResizeScale * ARatioY
104 ResizeScaleY = ResizeScale * ARatioX
106 Call ResizePic
108 Call PicAdjust
        '<EhFooter>
        Exit Sub

sldResize_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldResize_Change" + " line: " + Str(Erl))

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
Private Sub sldResize_scroll()
        '<EhHeader>
        On Error GoTo sldResize_scroll_Err
        '</EhHeader>
100 ResizeScale = sldResize.Value
102 ResizeScaleX = ResizeScale * ARatioY
104 ResizeScaleY = ResizeScale * ARatioX
106 Call ResizePic
108 Call PicAdjust
        '<EhFooter>
        Exit Sub

sldResize_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldResize_scroll" + " line: " + Str(Erl))

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
Private Sub sldHue_Change()
        '<EhHeader>
        On Error GoTo sldHue_Change_Err
        '</EhHeader>
100 Hditherval = sldHue.Value
        '<EhFooter>
        Exit Sub

sldHue_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldHue_Change" + " line: " + Str(Erl))

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
Private Sub sldsat_Change()
        '<EhHeader>
        On Error GoTo sldsat_Change_Err
        '</EhHeader>
100 SditherVal = sldSat.Value
        '<EhFooter>
        Exit Sub

sldsat_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldsat_Change" + " line: " + Str(Erl))

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
Private Sub sldsat_scroll()
        '<EhHeader>
        On Error GoTo sldsat_scroll_Err
        '</EhHeader>
100 SditherVal = sldSat.Value
        '<EhFooter>
        Exit Sub

sldsat_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldsat_scroll" + " line: " + Str(Erl))

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
Private Sub Check_Aspect_Click()
        '<EhHeader>
        On Error GoTo Check_Aspect_Click_Err
        '</EhHeader>
100 KeepAspect = Check_Aspect.Value
102 Call ResizePic
104 Call PicAdjust
        '<EhFooter>
        Exit Sub

Check_Aspect_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.Check_Aspect_Click" + " line: " + Str(Erl))

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
Private Sub Check_Stretch_Click()
        '<EhHeader>
        On Error GoTo Check_Stretch_Click_Err
        '</EhHeader>
100 StretchPic = (Check_Stretch.Value = vbChecked)
102 Check_Aspect.Enabled = (Check_Stretch.Value = vbChecked)


104 Call ResizePic
106 Call PicAdjust

108 DithControlsEnabled
        '<EhFooter>
        Exit Sub

Check_Stretch_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.Check_Stretch_Click" + " line: " + Str(Erl))

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

Private Sub tabstrip1_Click()
        '<EhHeader>
        On Error GoTo tabstrip1_Click_Err
        '</EhHeader>
100 TabFrames(TabStrip1.SelectedItem.Index - 1).ZOrder 0
102 LastVisitedTab = TabStrip1.SelectedItem.Index - 1

        '<EhFooter>
        Exit Sub

tabstrip1_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.tabstrip1_Click" + " line: " + Str(Erl))

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
Public Sub form_load()
        '<EhHeader>
        On Error GoTo form_load_Err
        '</EhHeader>
    Dim X As Byte



100 Label5.Visible = False

  
    '-------------------------------

102 Set Me.Icon = ZoomWindow.Image1.Picture

    'Tab control handling:

104 For X = 0 To TabFrames.Count - 1
106     With TabFrames(X)
108         .Move TabStrip1.ClientLeft, _
                  TabStrip1.ClientTop, _
                  TabStrip1.ClientWidth, _
                  TabStrip1.ClientHeight
110         .BorderStyle = 0
        End With
112 Next X

    ' Bring the first fraTab control to the front.
114 TabFrames(0).ZOrder 0

116 ConvertDialog.ScaleMode = vbPixels

118 DstPic.ScaleMode = vbPixels
120 SrcPic.ScaleMode = vbPixels
122 TmpPic.ScaleMode = vbPixels

124 DstPic.AutoRedraw = True
126 SrcPic.AutoRedraw = True
128 TmpPic.AutoRedraw = True


130 DstPic.AutoSize = True
132 DstPic.BorderStyle = 0

134 SrcPic.AutoSize = True
136 SrcPic.BorderStyle = 0

138 TmpPic.AutoSize = True
140 TmpPic.BorderStyle = 0

142     cmbDithBright.AddItem "none"
144     cmbDithBright.AddItem "2x2"
146     cmbDithBright.AddItem "4x4"
148     cmbDithBright.AddItem "4x4 odd"
150     cmbDithBright.AddItem "4x4 even"
152     cmbDithBright.AddItem "4x4 spotty"

154     cmbDithSat.AddItem "none"
156     cmbDithSat.AddItem "2x2"
158     cmbDithSat.AddItem "4x4"
160     cmbDithSat.AddItem "4x4 odd"
162     cmbDithSat.AddItem "4x4 even"
164     cmbDithSat.AddItem "4x4 spotty"

166     cmbDithHue.AddItem "none"
168     cmbDithHue.AddItem "2x2"
170     cmbDithHue.AddItem "4x4"
172     cmbDithHue.AddItem "4x4 odd"
174     cmbDithHue.AddItem "4x4 even"
176     cmbDithHue.AddItem "4x4 spotty"

178     cmbColorFilter.AddItem "Color Sensitive"
180     cmbColorFilter.AddItem "Brightness Sensitive"
182     cmbColorFilter.AddItem "RGB Distance"
184     cmbColorFilter.AddItem "YUV Distance"
186     cmbColorFilter.AddItem "Paletted Brightness"
    
188     cmbGfxMode.AddItem "Custom"
190     cmbGfxMode.AddItem "Hires"
192     cmbGfxMode.AddItem "Koala"
194     cmbGfxMode.AddItem "Drazlace"
196     cmbGfxMode.AddItem "Afli"
198     cmbGfxMode.AddItem "Ifli"
200     cmbGfxMode.AddItem "fli"
202     cmbGfxMode.AddItem "Drazlace Special"
204     cmbGfxMode.AddItem "unrestricted"
    
206 Debug.Print "convertdialog form load"

        '<EhFooter>
        Exit Sub

form_load_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.form_load" + " line: " + Str(Erl))

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
        Dim X As Long

100     Debug.Print "convertdialog form activate"
    
102     DstPic.Width = PW
104     SrcPic.Width = PW
106     TmpPic.Width = PW
    
108     DstPic.Height = PH
110     SrcPic.Height = PH
112     TmpPic.Height = PH


114     ConvertDialog.UserColor.BackColor = PaletteRGB(UsrBackgr)

        'colortable init -----------------
116     txtCTabFnames(Chip.vicii) = ColorTable_Filename(Chip.vicii)
118     txtCTabFnames(Chip.ted) = ColorTable_Filename(Chip.ted)
120     txtCTabFnames(Chip.vdc) = ColorTable_Filename(Chip.vdc)

        'brightness ladder init----------
122     LoadPresetList
        'activate selected gradient on listbox
124     GradientList.ListIndex = Gradient_Selected(ChipType)

        'load selected gradient
126     Call BrLadderLoadPreset(Gradient_Selected(ChipType))

        'init palette for gradient edit
128     For X = 0 To pal_Count
130         MyPalette1.PaletteRGB(X) = PaletteRGB(X)
132     Next X

        'init & redraw palettes
134     MyPalette1.BackgrColor = Me.BackColor
136     MyPalette1.BoxCount = pal_Count
138     MyPalette1.InitSurface
140     MyPalette2.BackgrColor = Me.BackColor
142     MyPalette2.InitSurface

        'put selected gradient into palette box
144     Call InitLadder


146     ResizeScaleX = ResizeScale * ARatioY
148     ResizeScaleY = ResizeScale * ARatioX
150     xA = 0: yA = 0
    
        'set up GUI ----------------------------------------------------------------------
    
        'set last visited tab to active
152     TabFrames(LastVisitedTab).ZOrder 0
154     TabStrip1.SelectedItem = TabStrip1.Tabs(LastVisitedTab + 1)
        
156     SrcPic.MousePointer = vbCustom 'set hand cursor to picSrc picbox
158     SrcPic.MouseIcon = LoadPicture(App.Path & "\Cursors\handflat.ico")

160     UserColor.BackColor = PaletteRGB(UsrBackgr)
    
162     If StretchPic = True Then Check_Stretch.Value = vbChecked Else Check_Stretch.Value = vbUnchecked
164     If KeepAspect = True Then Check_Aspect.Value = vbChecked Else Check_Aspect.Value = vbUnchecked

166     pixavg.Value = Reso_PixAvg
168     pixleft.Value = Reso_PixLeft
170     pixright.Value = Reso_PixRight

172     chkMostFreqBackgr.Value = MostFreqBackgr
174     chkUsrDefBackgr.Value = UsrDefBackgr

176     chkMostFrq3.Value = Mostfrq3
178     chkOptChars.Value = OptChars
    
180     chkFixWithBackgr.Value = FixWithBackgr

182     ConvertDialog.cmbDithBright.ListIndex = Bdither
184     ConvertDialog.cmbDithSat.ListIndex = Sdither
186     ConvertDialog.cmbDithHue.ListIndex = Hdither
    
188     chkRndBright.Value = BRnd
190     chkRndSat.Value = HRnd
192     chkRndHue.Value = SRnd

194     sldBright.Value = BditherVal
196     sldSat.Value = SditherVal
198     sldHue.Value = Hditherval
    
200     cmbColorFilter.ListIndex = ColorFilterMode
202     sldContrast.Value = Contrast
204     sldBrightness.Value = Brightness
206     sldHuex.Value = Hue
208     sldSaturation.Value = Saturation
    
210     Select Case GfxMode
            Case "custom"
212             cmbGfxMode.ListIndex = 0
214         Case "hires"
216             cmbGfxMode.ListIndex = 1
218         Case "koala"
220             cmbGfxMode.ListIndex = 2
222         Case "drazlace"
224             cmbGfxMode.ListIndex = 3
226         Case "afli"
228             cmbGfxMode.ListIndex = 4
230         Case "ifli"
232             cmbGfxMode.ListIndex = 5
234         Case "fli"
236             cmbGfxMode.ListIndex = 6
238         Case "drazlacespec"
240             cmbGfxMode.ListIndex = 7
242         Case "unrestricted"
244             cmbGfxMode.ListIndex = 8
        End Select
        

246     ConvertDialog.sldResize.Value = ResizeScale
248     DithControlsEnabled
    
    
250     Call PicAdjust
252     Call Apply_Click
        '<EhFooter>
        Exit Sub

form_activate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.form_activate" + " line: " + Str(Erl))

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
Private Sub cmbGfxMode_click()
        '<EhHeader>
        On Error GoTo cmbGfxMode_click_Err
        '</EhHeader>

100     If cmbGfxMode.ListIndex = 0 Then
102         GfxMode = "custom"
            'Custom.Show vbModal
        End If

104     If cmbGfxMode.ListIndex = 1 Then GfxMode = "hires"
106     If cmbGfxMode.ListIndex = 2 Then GfxMode = "koala"
108     If cmbGfxMode.ListIndex = 3 Then GfxMode = "drazlace"
110     If cmbGfxMode.ListIndex = 4 Then GfxMode = "afli"
112     If cmbGfxMode.ListIndex = 5 Then GfxMode = "ifli"
114     If cmbGfxMode.ListIndex = 6 Then GfxMode = "fli"
116     If cmbGfxMode.ListIndex = 7 Then GfxMode = "drazlacespec"
118     If cmbGfxMode.ListIndex = 8 Then GfxMode = "unrestricted"

120     Call ZoomWindow.ChangeGfxModeFromConvert


        '<EhFooter>
        Exit Sub

cmbGfxMode_click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbGfxMode_click" + " line: " + Str(Erl))

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

Private Sub cmbGfxMode_change()
        '<EhHeader>
        On Error GoTo cmbGfxMode_change_Err
        '</EhHeader>

100 Call cmbGfxMode_click

        '<EhFooter>
        Exit Sub

cmbGfxMode_change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.cmbGfxMode_change" + " line: " + Str(Erl))

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
Private Sub OKButton_Click()
        '<EhHeader>
        On Error GoTo OKButton_Click_Err
        '</EhHeader>

100 Reso_PixAvg = pixavg.Value
102 Reso_PixLeft = pixleft.Value
104 Reso_PixRight = pixright.Value

106 MostFreqBackgr = chkMostFreqBackgr.Value
108 UsrDefBackgr = chkUsrDefBackgr.Value

110 Mostfrq3 = chkMostFrq3.Value
112 OptChars = chkOptChars.Value
114 FixWithBackgr = chkFixWithBackgr.Value

116 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)
118 Call ZoomWindow.ZoomWinRefresh

120 Unload Me

        '<EhFooter>
        Exit Sub

OKButton_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.OKButton_Click" + " line: " + Str(Erl))

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



Private Sub SelColButton_Click()
        '<EhHeader>
        On Error GoTo SelColButton_Click_Err
        '</EhHeader>

100 ColorSelect.Show vbModal
102 UsrBackgr = ColorSelect.SelectedColor
104 ConvertDialog.UserColor.BackColor = PaletteRGB(ColorSelect.SelectedColor)

        '<EhFooter>
        Exit Sub

SelColButton_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SelColButton_Click" + " line: " + Str(Erl))

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
'redraw picDst pic preview from pixelsdib, with a vertical line to show progress
Public Sub ReDrawDstPicVert(X As Variant)
        '<EhHeader>
        On Error GoTo ReDrawDstPicVert_Err
        '</EhHeader>

100 StretchBlt DstPic.hdc, 0, 0, PW, PH, PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy
102 DstPic.DrawMode = 7        '7=eor,13=normal
104 DstPic.Forecolor = RGB(255, 255, 255)
106 DstPic.Line (X, 0)-(X, PH - 1), RGB(255, 255, 255)
108 DstPic.Refresh
110 DoEvents
        '<EhFooter>
        Exit Sub

ReDrawDstPicVert_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.ReDrawDstPicVert" + " line: " + Str(Erl))

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
'redraw picDst pic preview from pixelsdib, with a horizontal line to show progress
Public Sub ReDrawDstPicHor(Y As Variant)
        '<EhHeader>
        On Error GoTo ReDrawDstPicHor_Err
        '</EhHeader>

100 StretchBlt DstPic.hdc, 0, 0, PW, PH, PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy
102 DstPic.DrawMode = 7        '7=eor,13=normal
104 DstPic.Forecolor = RGB(255, 255, 255)
106 DstPic.Line (0, Y)-(PW - 1, Y), RGB(255, 255, 255)

108 DstPic.Refresh
110 DoEvents

        '<EhFooter>
        Exit Sub

ReDrawDstPicHor_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.ReDrawDstPicHor" + " line: " + Str(Erl))

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
'redraw picDst pic preview from pixelsdib
Public Sub ReDrawDstPic()
        '<EhHeader>
        On Error GoTo ReDrawDstPic_Err
        '</EhHeader>

100 StretchBlt DstPic.hdc, 0, 0, PW, PH, PixelsDib.hdc, 0, 0, PW, PH, vbSrcCopy
102 DstPic.Refresh
104 DoEvents

        '<EhFooter>
        Exit Sub

ReDrawDstPic_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.ReDrawDstPic" + " line: " + Str(Erl))

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



Private Sub sldcontrast_Change()
        '<EhHeader>
        On Error GoTo sldcontrast_Change_Err
        '</EhHeader>
100 Contrast = sldContrast.Value
102 Call PicAdjust
        '<EhFooter>
        Exit Sub

sldcontrast_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldcontrast_Change" + " line: " + Str(Erl))

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
Private Sub sldcontrast_scroll()
        '<EhHeader>
        On Error GoTo sldcontrast_scroll_Err
        '</EhHeader>
100 Call sldcontrast_Change
        '<EhFooter>
        Exit Sub

sldcontrast_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldcontrast_scroll" + " line: " + Str(Erl))

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
Private Sub sldhuex_change()
        '<EhHeader>
        On Error GoTo sldhuex_change_Err
        '</EhHeader>
100 Hue = sldHuex.Value
102 Call PicAdjust
        '<EhFooter>
        Exit Sub

sldhuex_change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldhuex_change" + " line: " + Str(Erl))

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

Private Sub sldhuex_Scroll()
        '<EhHeader>
        On Error GoTo sldhuex_Scroll_Err
        '</EhHeader>
100 Call sldhuex_change
        '<EhFooter>
        Exit Sub

sldhuex_Scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldhuex_Scroll" + " line: " + Str(Erl))

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
Private Sub sldSaturation_change()
        '<EhHeader>
        On Error GoTo sldSaturation_change_Err
        '</EhHeader>
100 Saturation = sldSaturation.Value
102 Call PicAdjust
        '<EhFooter>
        Exit Sub

sldSaturation_change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldSaturation_change" + " line: " + Str(Erl))

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

Private Sub sldSaturation_Scroll()
        '<EhHeader>
        On Error GoTo sldSaturation_Scroll_Err
        '</EhHeader>
100 Call sldSaturation_change
        '<EhFooter>
        Exit Sub

sldSaturation_Scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldSaturation_Scroll" + " line: " + Str(Erl))

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

Private Sub sldBrightness_Change()
        '<EhHeader>
        On Error GoTo sldBrightness_Change_Err
        '</EhHeader>
100 Brightness = sldBrightness.Value
102 Call PicAdjust
        '<EhFooter>
        Exit Sub

sldBrightness_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldBrightness_Change" + " line: " + Str(Erl))

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

Private Sub sldBrightness_scroll()
        '<EhHeader>
        On Error GoTo sldBrightness_scroll_Err
        '</EhHeader>
100 Call sldBrightness_Change
        '<EhFooter>
        Exit Sub

sldBrightness_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.sldBrightness_scroll" + " line: " + Str(Erl))

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




Public Sub PicAdjust()
        '<EhHeader>
        On Error GoTo PicAdjust_Err
        '</EhHeader>
    Dim resp As Long

    
100 Call modGraphical.GPX_Hue(SrcPic.hdc, ZoomWindow.PrevPic.hdc, Hue, resp)
102 Call modGraphical.GPX_Saturation(TmpPic.hdc, SrcPic.hdc, Saturation, resp)
104 Call modGraphical.GPX_Brightness(SrcPic.hdc, TmpPic.hdc, Brightness, resp)
106 Call modGraphical.GPX_Contrast(TmpPic.hdc, SrcPic.hdc, Contrast / 100, Contrast / 100, Contrast / 100, resp)
108 Call modGraphical.GPX_BitBlt(SrcPic.hdc, 0, 0, SrcPic.Width, SrcPic.Height, TmpPic.hdc, 0, 0, vbSrcCopy, resp)
    'Call modGraphical.GPX_Brightness(SrcPic.hdc, TmpPic.hdc, 0, resp)

110 Debug.Print "picadjust debug"
112 Debug.Print "src width " & SrcPic.Width
114 Debug.Print "src height " & SrcPic.Height
116 Debug.Print "tmp width " & TmpPic.Width
118 Debug.Print "tmp height " & TmpPic.Height
    
    'SrcPic.Picture = TmpPic.Image
120 SrcPic.Refresh
    'MsgBox "picadjust is not done"

        '<EhFooter>
        Exit Sub

PicAdjust_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.PicAdjust" + " line: " + Str(Erl))

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

Public Sub ResizePic()
    '1. stretch zoomwindow.LoadedPic to zoomwindow.prevpic
    '2. copies this picture to convertdialog.SrcPic
        '<EhHeader>
        On Error GoTo ResizePic_Err
        '</EhHeader>

    Dim xC As Long
    Dim yC As Long

    Dim Z As Single
    Dim X As Single
    Dim Y As Single

    'clear convert preview pic
100 With ZoomWindow
102     .PrevPic.BackColor = 0
104     .PrevPic.Cls
106     .PrevPic.Refresh
    
108     .PrevPic.Line (0, 0)-(PW, PH), 0, BF
    End With


110 If StretchPic = False Then

112     xC = xD - ((ZoomWindow.LoadedPic.Width * ResizeScaleX / 100) / 2)
114     yC = yD - ((ZoomWindow.LoadedPic.Height * ResizeScaleY / 100) / 2)

116     ZoomWindow.PrevPic.Line (0, 0)-(ZoomWindow.PrevPic.Width, ZoomWindow.PrevPic.Height), 0, BF
118     ZoomWindow.PrevPic.PaintPicture ZoomWindow.LoadedPic.Picture, xC, yC, _
        ZoomWindow.LoadedPic.Width * ResizeScaleX / 100, ZoomWindow.LoadedPic.Height * ResizeScaleY / 100, _
        0, 0, ZoomWindow.LoadedPic.Width, ZoomWindow.LoadedPic.Height, vbSrcCopy
120     ZoomWindow.PrevPic.Refresh
    
        'ConvertDialog.SrcPic.Picture = ZoomWindow.PrevPic.Image
    
    Else

122     If KeepAspect = False Then

124         With ZoomWindow
126         .LoadedPic.ScaleMode = vbPixels

128         .PrevPic.PaintPicture .LoadedPic.Image, 0, 0, _
                                 .PrevPic.Width, .PrevPic.Height, _
                                 0, 0, _
                                 .LoadedPic.Width, .LoadedPic.Height, vbSrcCopy
            End With
        
        Else

130         With ZoomWindow
        
132         .LoadedPic.ScaleMode = vbPixels

134         If .LoadedPic.Width / .LoadedPic.Height < (PW) / (PH) Then

                'magasabb mint szelesebb

136             Z = 1 / (.LoadedPic.Height / PH)
138             X = ((PW - 1) - (Z * .LoadedPic.Width)) / 2

140             .PrevPic.PaintPicture .LoadedPic.Image, _
                                     X, 0, _
                                     Z * .LoadedPic.Width, Z * .LoadedPic.Height, _
                                     0, 0, _
                                     .LoadedPic.Width, .LoadedPic.Height, vbSrcCopy

            Else

142             Z = 1 / (.LoadedPic.Width / PW)
144             Y = ((PH - 1) - (Z * .LoadedPic.Height)) / 2

146             .PrevPic.PaintPicture .LoadedPic.Image, _
                                     0, Y, _
                                     Z * .LoadedPic.Width, Z * .LoadedPic.Height, _
                                     0, 0, _
                                     .LoadedPic.Width, .LoadedPic.Height, vbSrcCopy
        
            End If
            End With
        
        End If

        'ConvertDialog.SrcPic.Picture = ZoomWindow.PrevPic.Image
    End If


        '<EhFooter>
        Exit Sub

ResizePic_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.ResizePic" + " line: " + Str(Erl))

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




'---------------------------------------------------------------
'
'
' BrightnessLadder Code
'
'
'
'-------------------------------------------------------


Private Sub InitLadder()
        '<EhHeader>
        On Error GoTo InitLadder_Err
        '</EhHeader>
    Dim X As Long

100 For X = 0 To 31
102     MyPalette2.PaletteRGB(X) = PaletteRGB(BrLadderTab(X))
104 Next X

106 EnableEntrys
108 HScroll1.Value = BrLadderMax
110 MyPalette2.BoxCount = BrLadderMax - 1
112 MyPalette2.InitSurface
114 SetCap

        '<EhFooter>
        Exit Sub

InitLadder_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.InitLadder" + " line: " + Str(Erl))

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
100 If GradientList.ListIndex >= 0 Then
102 Call DeletePreset
    End If
        '<EhFooter>
        Exit Sub

DeleteSelected_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.DeleteSelected_Click" + " line: " + Str(Erl))

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

Private Sub EnableEntrys()
        '<EhHeader>
        On Error GoTo EnableEntrys_Err
        '</EhHeader>

100 MyPalette2.BoxCount = BrLadderMax - 1

        '<EhFooter>
        Exit Sub

EnableEntrys_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.EnableEntrys" + " line: " + Str(Erl))

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

Private Sub SetCap()
        '<EhHeader>
        On Error GoTo SetCap_Err
        '</EhHeader>

100 Label7.Caption = "Number of Entrys: " + Str(BrLadderMax)

        '<EhFooter>
        Exit Sub

SetCap_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SetCap" + " line: " + Str(Erl))

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
Private Sub HScroll1_Change()
        '<EhHeader>
        On Error GoTo HScroll1_Change_Err
        '</EhHeader>

100 BrLadderMax = HScroll1.Value
102 EnableEntrys
104 SetCap

        '<EhFooter>
        Exit Sub

HScroll1_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.HScroll1_Change" + " line: " + Str(Erl))

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
Private Sub HScroll1_scroll()
        '<EhHeader>
        On Error GoTo HScroll1_scroll_Err
        '</EhHeader>

100 HScroll1_Change

        '<EhFooter>
        Exit Sub

HScroll1_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.HScroll1_scroll" + " line: " + Str(Erl))

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



Private Sub MyPalette2_BoxClicked(Button As Integer, Index As Byte)
        '<EhHeader>
        On Error GoTo MyPalette2_BoxClicked_Err
        '</EhHeader>

100 ActiveEntry = Index
102 MyPalette2.PaletteRGB(ActiveEntry) = PaletteRGB(ActiveIndex)
104 MyPalette2.InitSurface
106 BrLadderTab(ActiveEntry) = ActiveIndex

        '<EhFooter>
        Exit Sub

MyPalette2_BoxClicked_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.MyPalette2_BoxClicked" + " line: " + Str(Erl))

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

100 ActiveIndex = Index
102 MyPalette2.PaletteRGB(ActiveEntry) = PaletteRGB(ActiveIndex)
104 MyPalette2.InitSurface
106 BrLadderTab(ActiveEntry) = ActiveIndex
108 Label20.BackColor = PaletteRGB(ActiveIndex)

        '<EhFooter>
        Exit Sub

MyPalette1_BoxClicked_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.MyPalette1_BoxClicked" + " line: " + Str(Erl))

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

Public Sub BrLadderLoadPreset(ByVal Index As Long)
        '<EhHeader>
        On Error GoTo BrLadderLoadPreset_Err
        '</EhHeader>
    Dim PresetCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long
    Dim Name As String

100 With m_cIni
    
102     .Path = App.Path & "\Ladders.ini"
104     .Section = "main"
106     .Key = "presetcount": .Default = 0: PresetCount = .Value
    
108     If PresetCount <> 0 Then
110             .Key = "Name" & Format(Index): .Default = "": Name = .Value
112             If Name <> "" Then
114                 .Section = Name
116                 .Key = "colorcount": .Default = 0: ColorCount = .Value
118                 If ColorCount <> 0 Then
120                     For Y = 1 To ColorCount
122                         .Key = "colorindex" & Format(Y): .Default = 0: BrLadderTab(Y) = .Value
124                     Next Y
126                     BrLadderMax = ColorCount
                    End If
                End If
        End If

    End With

128 Frame8.Caption = "Selected: " + Name

        '<EhFooter>
        Exit Sub

BrLadderLoadPreset_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.BrLadderLoadPreset" + " line: " + Str(Erl))

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

Private Sub LoadPresetList()
        '<EhHeader>
        On Error GoTo LoadPresetList_Err
        '</EhHeader>

    Dim PresetCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long
    Dim Name As String
    Dim ChipStr As String

100 GradientList.Clear

102 With m_cIni
    
104     .Path = App.Path & "\Ladders.ini"
106     .Section = "main"
108     .Key = "presetcount": .Default = 0: PresetCount = .Value
    
110     If PresetCount <> 0 Then
112         Y = 0
114         For X = 0 To PresetCount
116             .Key = "Name" & Format(X): .Default = "": Name = .Value
            
118             .Section = Name
120             .Key = "type": .Default = "error"
            
122             Select Case .Value
            
                    Case 0
124                     ChipStr = " (VICII)"
126                 Case 1
128                     ChipStr = " (TED)"
130                 Case 2
132                     ChipStr = " (VDC)"
                End Select
            
            
                'If .Value = ChipType Then
134                 GradientList.AddItem Name & ChipStr
136                 GradientList.ItemData(Y) = X
138                 Y = Y + 1
                'End If
140             .Section = "main"
142          Next X
        End If

    End With

        '<EhFooter>
        Exit Sub

LoadPresetList_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.LoadPresetList" + " line: " + Str(Erl))

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

    Dim PresetCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long


100 If Name = "overwriteoldindex" Then


102 With m_cIni
    
104     .Path = App.Path & "\Ladders.ini"
106     .Section = "main"
    
108     .Key = "Name" & Format(GradientList.ListIndex):    Name = .Value
110     .Section = Name
112     .DeleteSection
114     .Key = "colorcount": .Value = BrLadderMax
116     .Key = "type": .Value = ChipType
118     For Y = 0 To BrLadderMax - 1
120         .Key = "colorindex" & Format(Y): .Value = BrLadderTab(Y)
122     Next Y

    End With


    Else


124 With m_cIni
    
126     .Path = App.Path & "\Ladders.ini"
128     .Section = "main"
        .Key = "presetcount":
130     .Default = 0
132     PresetCount = .Value
134     For X = 0 To PresetCount - 1
136         .Key = "Name" & Format(X)
138         If StrComp(.Value, Name, vbTextCompare) = 0 Then
140             MsgBox "A preset with this name does already exist", vbExclamation, "Warning"
142             Name = "alreadyexisted"
                Exit Sub
            End If
144     Next X
        
146     PresetCount = PresetCount + 1
148     .Key = "Name" & Format(PresetCount):  .Value = Name
150     .Key = "presetcount"
152     .Value = PresetCount
        
154     .Section = Name
156     .DeleteSection
158     .Key = "colorcount": .Value = BrLadderMax
160     .Key = "type": .Value = ChipType
162     For Y = 0 To BrLadderMax - 1
164         .Key = "colorindex" & Format(Y): .Value = BrLadderTab(Y)
166     Next Y

    End With


    End If

        '<EhFooter>
        Exit Sub

SavePreset_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SavePreset" + " line: " + Str(Erl))

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

    Dim PresetCount As Long
    Dim ColorCount As Long
    Dim X As Long
    Dim Y As Long
    Dim Name As String
    Dim Temp As String
100 With m_cIni
    
102     .Path = App.Path & "\Ladders.ini"
104     .Section = "main"
106     .Key = "presetcount": .Default = 0: PresetCount = .Value
    
108     If PresetCount <> 0 Then
110         .Key = "Name" & Format(GradientList.ListIndex):    Name = .Value
112         .Section = Name
114         .DeleteSection
116         .Section = "main"
118         .Key = "Name" & Format(GradientList.ListIndex)
120         .DeleteKey
122         PresetCount = PresetCount - 1
        
124         Y = GradientList.ListIndex + 1
126         For X = Y To PresetCount
128             .Key = "Name" & Format(X)
130             Temp = .Value
132             .DeleteKey
134             .Key = "Name" & Format(X - 1)
136             .Value = Temp
138         Next X
140         .Section = "main"
142         .Key = "presetcount"
144         .Value = PresetCount
        End If

    End With

146 Call LoadPresetList
        '<EhFooter>
        Exit Sub

DeletePreset_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.DeletePreset" + " line: " + Str(Erl))

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

Private Sub GradientList_Click()
        '<EhHeader>
        On Error GoTo GradientList_Click_Err
        '</EhHeader>

100 Gradient_Selected(ChipType) = GradientList.ListIndex

102 Call BrLadderLoadPreset(GradientList.ListIndex)
104 Call InitLadder

        '<EhFooter>
        Exit Sub

GradientList_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.GradientList_Click" + " line: " + Str(Erl))

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



106 If NewName <> "alreadyexisted" Then
108     GradientList.AddItem NewName
110     GradientList.ListIndex = GradientList.ListCount - 1
112     GradientList.ItemData(GradientList.ListIndex) = GradientList.ListIndex
114     Call SavePreset(NewName)
    End If

        '<EhFooter>
        Exit Sub

NewPreset_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.NewPreset_Click" + " line: " + Str(Erl))

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

100 If GradientList.ListIndex >= 0 Then
102     Call SavePreset("overwriteoldindex")
    Else
104     MsgBox "You must select a preset you whish to overwrite", vbExclamation
    End If
        '<EhFooter>
        Exit Sub

SaveOverSelected_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.SaveOverSelected_Click" + " line: " + Str(Erl))

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

Private Sub txtCTabFnames_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo txtCTabFnames_Change_Err
        '</EhHeader>

100 ColorTable_Filename(Index) = txtCTabFnames(Index)

        '<EhFooter>
        Exit Sub

txtCTabFnames_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.ConvertDialog.txtCTabFnames_Change" + " line: " + Str(Erl))

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
