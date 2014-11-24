VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DDA53BD0-2CD0-11D4-8ED4-00E07D815373}#1.0#0"; "MBMouse.ocx"
Begin VB.MDIForm MainWin 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Project One"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11370
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   MousePointer    =   1  'Arrow
   ScrollBars      =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   11100
      ScaleHeight     =   7260
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   360
      Width           =   270
      Begin VB.VScrollBar VScroll1 
         Height          =   7215
         LargeChange     =   32
         Left            =   0
         Max             =   200
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   11370
      TabIndex        =   1
      Top             =   7620
      Width           =   11370
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   80
         Left            =   0
         Max             =   319
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   10816
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7875
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":08B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0E02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Freehand"
            Object.ToolTipText     =   "Freehand"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fill"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a"
                  Text            =   "Stric"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b"
                  Text            =   "Compensating"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Brush"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "c"
                  Text            =   "Setup Brush"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   4
            Style           =   1
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   0
         Width           =   2175
      End
   End
   Begin MBMouseHelper.MouseHelper MouseHelper1 
      Left            =   1560
      Top             =   840
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadPicture 
         Caption         =   "Import Picture"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExportPicture 
         Caption         =   "Export Picture"
      End
      Begin VB.Menu hihihz 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavePicture 
         Caption         =   "Save Picture"
         Shortcut        =   ^S
      End
      Begin VB.Menu bahaha 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuLoadKoala 
         Caption         =   "Load Koala"
      End
      Begin VB.Menu mnuSaveKoala 
         Caption         =   "Save Koala"
      End
      Begin VB.Menu mnuSaveDrazlace 
         Caption         =   "Save Drazlace"
      End
      Begin VB.Menu mnuLoadDrazlace 
         Caption         =   "Load Drazlace"
      End
      Begin VB.Menu mnuSaveFli 
         Caption         =   "Save Fli"
      End
      Begin VB.Menu mnuLoadFli 
         Caption         =   "Load Fli"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomLoad 
         Caption         =   "Custom Load"
      End
      Begin VB.Menu mnuCustomSave 
         Caption         =   "Custom Save"
      End
      Begin VB.Menu seperetaror 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadOwn 
         Caption         =   "Load P1 Format"
      End
      Begin VB.Menu mnuSaveOwn 
         Caption         =   "Save P1 format"
      End
      Begin VB.Menu anothaseperatho 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuClipboard 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu seperatahrota 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuCustomScreenMode 
         Caption         =   "Define Custom GfxMode"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "Palette Setup"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFillmode 
         Caption         =   "Fillmode"
         Begin VB.Menu mnuStrictFill 
            Caption         =   "strict"
         End
         Begin VB.Menu mnuCompensatingFill 
            Caption         =   "compensating"
         End
         Begin VB.Menu uhuh 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDitherfill 
            Caption         =   "Ditherfill"
         End
      End
      Begin VB.Menu mnu4thReplace 
         Caption         =   "4th Color Replaces"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuAspect 
         Caption         =   "Aspect Ratio"
      End
   End
   Begin VB.Menu mnuGfxMode 
      Caption         =   "&GfxMode"
      Begin VB.Menu mnuGfxModeCustom 
         Caption         =   "Custom Mode"
      End
      Begin VB.Menu mnuGfxModeHires 
         Caption         =   "Hires"
      End
      Begin VB.Menu mnuGfxModeKoala 
         Caption         =   "Koala"
      End
      Begin VB.Menu mnuGfxModeDrazlace 
         Caption         =   "Drazlace"
      End
      Begin VB.Menu mnuGfxModeAfli 
         Caption         =   "Afli"
      End
      Begin VB.Menu mnuGfxModeFli 
         Caption         =   "Fli"
      End
      Begin VB.Menu mnuGfxModeIFli 
         Caption         =   "IFli"
      End
      Begin VB.Menu mnuGfxModeDrazlaceSpec 
         Caption         =   "Drazlace (2 screenmem)"
      End
      Begin VB.Menu mnuGfxModeUnrestrictedHires 
         Caption         =   "Unrestricted Hires"
      End
      Begin VB.Menu mnuGfxModeUnrestrictedMulti 
         Caption         =   "Unrestricted Multi"
      End
   End
   Begin VB.Menu mnuBrush 
      Caption         =   "Brush"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom In Edit Window"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom Out Edit Window"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnufake 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomInPrev 
         Caption         =   "Zoom In Preview Window"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuZoomOutPrev 
         Caption         =   "Zoom Out Preview Window"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuyetanothersepara 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreviewWindow 
         Caption         =   "Preview Window"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPaletteWindow 
         Caption         =   "Palette Window"
         Shortcut        =   ^L
      End
      Begin VB.Menu yetanotherseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInterlaceEmu 
         Caption         =   "Toggle Interlace Emulation"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuGridOptions 
         Caption         =   "Grid Setup"
         Shortcut        =   ^G
      End
      Begin VB.Menu anotherseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRedraw 
         Caption         =   "Debug (redraw from emulated mem)"
         Index           =   0
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuDebugChar 
         Caption         =   "Debug Char"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "MainWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim actButton As Integer
Dim actShift As Integer
Dim OldX As Single
Dim yOld As Single

 ' Visual Basic code for implementing HTML Help 1.1

'*****
' Declare the following two constants
' as PUBLIC
Private Const HH_HELP_CONTEXT = &HF            ' display mapped numeric
Private Const HH_TP_HELP_WM_HELP = &H11       ' text popup help, same as
                                             ' WinHelp HELP_WM_HELP

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_HELP_FINDER = &H0            ' WinHelp equivalent
Private Const HH_DISPLAY_TOC = &H1            ' WinHelp equivalent
Private Const HH_DISPLAY_INDEX = &H2         ' WinHelp equivalent
Private Const HH_DISPLAY_SEARCH = &H3        ' not currently implemented
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_ENUM_INFO_TYPE = &H7        ' Get Info type name, call
                                             ' repeatedly to enumerate,
                                             ' -1 at end
Private Const HH_SET_INFO_TYPE = &H8         ' Add Info type to filter.
Private Const HH_SYNC = &H9
Private Const HH_ADD_NAV_UI = &HA             ' not currently implemented
Private Const HH_ADD_BUTTON = &HB             ' not currently implemented
Private Const HH_GETBROWSER_APP = &HC        ' not currently implemented
Private Const HH_KEYWORD_LOOKUP = &HD
Private Const HH_DISPLAY_TEXT_POPUP = &HE    ' display string resource id
                                             ' or text in a popup window
                                             ' value in dwData
Private Const HH_TP_HELP_CONTEXTMENU = &H10  ' text popup help, same as
                                             ' WinHelp HELP_CONTEXTMENU
Private Const HH_CLOSE_ALL = &H12             ' close all windows opened
                                             ' directly or indirectly by
                                             ' the caller
Private Const HH_ALINK_LOOKUP = &H13         ' ALink version of
                                             ' HH_KEYWORD_LOOKUP
Private Const HH_GET_LAST_ERROR = &H14       ' not currently implemented
Private Const HH_ENUM_CATEGORY = &H15        ' Get category name, call
                                             ' repeatedly to enumerate,
                                             ' -1 at end
Private Const HH_ENUM_CATEGORY_IT = &H16     ' Get category info type
                                             ' members, call repeatedly to
                                             ' enumerate, -1 at end
Private Const HH_RESET_IT_FILTER = &H17      ' Clear the info type filter
                                             ' of all info types.
Private Const HH_SET_INCLUSIVE_FILTER = &H18 ' set inclusive filtering
                                             ' method for untyped topics
                                             ' to be included in display
Private Const HH_SET_EXCLUSIVE_FILTER = &H19  ' set method for untyped
                                             ' topics to be excluded from
                                             ' the display
Private Const HH_SET_GUID = &H1A              ' For Microsoft Installer --
                                             ' dwData is a pointer to the
                                             ' GUID string
Private Const HH_INTERNAL = &HFF              ' Used internally.

' Button IDs

Private Const IDTB_EXPAND = 200
Private Const IDTB_CONTRACT = 201
Private Const IDTB_STOP = 202
Private Const IDTB_REFRESH = 203
Private Const IDTB_BACK = 204
Private Const IDTB_HOME = 205
Private Const IDTB_SYNC = 206
Private Const IDTB_PRINT = 207
Private Const IDTB_OPTIONS = 208
Private Const IDTB_FORWARD = 209
Private Const IDTB_NOTES = 210                ' not implemented
Private Const IDTB_BROWSE_FWD = 211
Private Const IDTB_BROWSE_BACK = 212
Private Const IDTB_CONTENTS = 213             ' not implemented
Private Const IDTB_INDEX = 214                ' not implemented
Private Const IDTB_SEARCH = 215               ' not implemented
Private Const IDTB_HISTORY = 216              ' not implemented
Private Const IDTB_BOOKMARKS = 217            ' not implemented
Private Const IDTB_JUMP1 = 218
Private Const IDTB_JUMP2 = 219
Private Const IDTB_CUSTOMIZE = 221
Private Const IDTB_ZOOM = 222
Private Const IDTB_TOC_NEXT = 223
Private Const IDTB_TOC_PREV = 224

'Type RECT
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type

Private Type tagHHN_NOTIFY
  hdr As Variant
  pszUrl As String                            ' Multi-byte, null-terminated string
End Type

Private Type tagHH_POPUP
  cbStruct As Integer                         ' sizeof this structure
  hInst As Variant                            ' instance handle for string resource
  idString As Variant                         ' string resource id, or text id if pszFile
                                             ' is specified in HtmlHelp call
  pszText As String                           ' used if idString is zero
  pt As Integer                               ' top center of popup window
  clrForeground As ColorConstants             ' use -1 for default
  clrBackground As ColorConstants             ' use -1 for default
  rcMargins As RECT                           ' amount of space between edges of window and
                                             ' text, -1 for each member to ignore
  pszFont As String                           ' facename, point size, char set, BOLD ITALIC
                                             ' UNDERLINE
End Type

Private Type tagHH_AKLINK
  cbStruct As Integer                         ' sizeof this structure
  fReserved As Boolean                        ' must be FALSE (really!)
  pszKeywords As String                       ' semi-colon separated keywords
  pszUrl As String                            ' URL to jump to if no keywords found (may be
                                             ' NULL)
  pszMsgText As String                        ' Message text to display in MessageBox if
                                             ' pszUrl
                                             ' is NULL and no keyword match
  pszMsgTitle As String                       ' Message text to display in MessageBox if
                                             ' pszUrl is NULL and no keyword match
  pszWindow As String                         ' Window to display URL in
  fIndexOnFail As Boolean                     ' Displays index if keyword lookup fails.
End Type

Private Enum NavigationTypes
  HHWIN_NAVTYPE_TOC
  HHWIN_NAVTYPE_INDEX
  HHWIN_NAVTYPE_SEARCH
  HHWIN_NAVTYPE_BOOKMARKS
  HHWIN_NAVTYPE_HISTORY ' not implemented
End Enum

Private Enum IT
  IT_INCLUSIVE
  IT_EXCLUSIVE
  IT_HIDDEN
End Enum

Private Type tagHH_ENUM_IT
  cbStruct As Integer                         ' size of this structure
  iType As Integer                            ' the type of the information type i.e.
                                             ' Inclusive, Exclusive, or Hidden
  pszCatName As String                        ' Set to the name of the Category to
                                             ' enumerate the info types in a category;
                                             ' else NULL
  pszITName As String                         ' volitile pointer to the name of the
                                             ' infotype. Allocated by call. Caller
                                             ' responsible for freeing
  pszITDescription As String                  ' volitile pointer to the description of the
                                             ' infotype.
End Type

Private Type tagHH_ENUM_CAT
  cbStruct As Integer                         ' size of this structure
  pszCatName As String                        ' volitile pointer to the category name
  pszCatDescription As String                 ' volitile pointer to the category
                                             ' description
End Type

Private Type tagHH_SET_INFOTYPE
  cbStruct As Integer                         ' the size of this structure
  pszCatName As String                        ' the name of the category, if any, the
                                             ' InfoType is a member of.
  pszInfoTypeName As String                   ' the name of the info type to add to the
                                             ' filter
End Type

Private Enum NavTabs
  HHWIN_NAVTAB_TOP
  HHWIN_NAVTAB_LEFT
  HHWIN_NAVTAB_BOTTOM
End Enum

Private Const HH_MAX_TABS = 19 ' maximum number of tabs
Private Enum Tabs
  HH_TAB_CONTENTS
  HH_TAB_INDEX
  HH_TAB_SEARCH
  HH_TAB_BOOKMARKS
  HH_TAB_HISTORY
End Enum

' HH_DISPLAY_SEARCH Command Related Structures and Constants

Private Const HH_FTS_DEFAULT_PROXIMITY = (-1)

Private Type tagHH_FTS_QUERY
  cbStruct As Integer                         ' Sizeof structure in bytes.
  fUniCodeStrings As Boolean                  ' TRUE if all strings are unicode.
  pszSearchQuery As String                    ' String containing the search query.
  iProximity As Long                          ' Word proximity.
  fStemmedSearch As Boolean                   ' TRUE for StemmedSearch only.
  fTitleOnly As Boolean                       ' TRUE for Title search only.
  fExecute As Boolean                         ' TRUE to initiate the search.
  pszWindow As String                         ' Window to display in
End Type

' HH_WINTYPE Structure

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_SHOW = 5

Private Type HH_WINTYPE
  cbStruct As Integer                         ' IN: size of this structure including all
                                             ' Information Types
  fUniCodeStrings As Boolean                  ' IN/OUT: TRUE if all strings are in UNICODE
  pszType As String                           ' IN/OUT: Name of a type of window
  fsValidMembers As Variant                   ' IN: Bit flag of valid members
                                             ' (HHWIN_PARAM_)
  fsWinProperties As Variant                  ' IN/OUT: Properties/attributes of the window
                                             ' (HHWIN_)
  pszCaption As String                        ' IN/OUT: Window title
  dwStyles As Variant                         ' IN/OUT: Window styles
  dwExStyles As Variant                       ' IN/OUT: Extended Window styles
  rcWindowPos As RECT                         ' IN: Starting position, OUT: current
                                             ' position
  nShowState As Integer                       ' IN: show state (e.g., SW_SHOW)
  hwndHelp As Variant                         ' OUT: window handle
  hwndCaller As Variant                       ' OUT: who called this window
                                             ' The following members are only valid if
                                             ' HHWIN_PROP_TRI_PANE is set
  hwndToolBar As Variant                      ' OUT: toolbar window in tri-pane window
  hwndNavigation As Variant                   ' OUT: navigation window in tri-pane window
  hwndHTML As Variant                         ' OUT: window displaying HTML in tri-pane
                                             ' window
  iNavWidth As Integer                        ' IN/OUT: width of navigation window
  rcHTML As RECT                              ' OUT: HTML window coordinates
  pszToc As String                            ' IN: Location of the table of contents file
  pszIndex As String                           ' IN: Location of the index file
  pszFile As String                           ' IN: Default location of the html file
  pszHome As String                           ' IN/OUT: html file to display when Home
                                             ' button is clicked
  fsToolBarFlags As Variant                   ' IN: flags controling the appearance of the
                                             ' toolbar
  fNotExpanded As Boolean                     ' IN: TRUE/FALSE to contract or expand, OUT:
                                             ' current state
  curNavType As Integer                       ' IN/OUT: UI to display in the navigational
                                             ' pane
  tabpos As Integer                           ' IN/OUT: HHWIN_NAVTAB_TOP, HHWIN_NAVTAB_LEFT,
                                             ' or HHWIN_NAVTAB_BOTTOM
  idNotify As Integer                         ' IN: ID to use for WM_NOTIFY messages
  tabOrder(HH_MAX_TABS + 1) As Byte           ' IN/OUT: tab order: Contents, Index,
                                             ' Search, History, Favorites, Reserved 1-5,
                                             ' Custom tabs
  cHistory As Integer                         ' IN/OUT: number of history items to keep
                                             ' (default is 30)
  pszJump1 As String                          ' Text for HHWIN_BUTTON_JUMP1
  pszJump2 As String                          ' Text for HHWIN_BUTTON_JUMP2
  pszUrlJump1 As String                       ' URL for HHWIN_BUTTON_JUMP1
  pszUrlJump2 As String                       ' URL for HHWIN_BUTTON_JUMP2
  rcMinSize As RECT                           ' Minimum size for window (ignored in version
                                             ' 1 of the Workshop)
  cbInfoTypes As Integer                      ' size of paInfoTypes;
End Type

Private Enum Actions
  HHACT_TAB_CONTENTS
  HHACT_TAB_INDEX
  HHACT_TAB_SEARCH
  HHACT_TAB_HISTORY
  HHACT_TAB_FAVORITES
  HHACT_EXPAND
  HHACT_CONTRACT
  HHACT_BACK
  HHACT_FORWARD
  HHACT_STOP
  HHACT_REFRESH
  HHACT_HOME
  HHACT_SYNC
  HHACT_OPTIONS
  HHACT_PRINT
  HHACT_HIGHLIGHT
  HHACT_CUSTOMIZE
  HHACT_JUMP1
  HHACT_JUMP2
  HHACT_ZOOM
  HHACT_TOC_NEXT
  HHACT_TOC_PREV
  HHACT_NOTES
  HHACT_LAST_ENUM
End Enum

Private Type tagHHNTRACK
  hdr As Variant
  pszCurUrl As String                         ' Multi-byte, null-terminated string
  idAction As Integer                         ' HHACT_ value
  phhWinType As HH_WINTYPE                    ' Current window type structure
End Type

Private Type HH_IDPAIR
  dwControlId As Long
  dwTopicId As Long
End Type
Private Declare Function htmlhelp Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hWnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long



Public Function SetHTMLHelpStrings() As String
        '// this presumes the help file is in the same directory as your app, and Main is the name of the window
        '<EhHeader>
        On Error GoTo SetHTMLHelpStrings_Err
        '</EhHeader>
100     SetHTMLHelpStrings = App.Path & "\ProjectOne.chm" '& ">Main"
                                                            ' >main should be there but caused errror....
        '<EhFooter>
        Exit Function

SetHTMLHelpStrings_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.SetHTMLHelpStrings" + " line: " + Str(Erl))

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

Public Sub HTMLHelpContents()
      ' Force the Help window to display
      ' the Contents file (*.hhc) in the left pane
        '<EhHeader>
        On Error GoTo HTMLHelpContents_Err
        '</EhHeader>
100   htmlhelp hWnd, SetHTMLHelpStrings(), HH_DISPLAY_TOC, 0

        '<EhFooter>
        Exit Sub

HTMLHelpContents_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.HTMLHelpContents" + " line: " + Str(Erl))

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

Public Sub HTMLHelpIndex()
        '<EhHeader>
        On Error GoTo HTMLHelpIndex_Err
        '</EhHeader>

      ' Force the Help window to display the Index file
      ' (*.hhk) in the left pane
100   htmlhelp hWnd, SetHTMLHelpStrings(), HH_DISPLAY_INDEX, 0

        '<EhFooter>
        Exit Sub

HTMLHelpIndex_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.HTMLHelpIndex" + " line: " + Str(Erl))

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


Private Sub mnu4thReplace_Click()
        '<EhHeader>
        On Error GoTo mnu4thReplace_Click_Err
        '</EhHeader>
100 mnu4thReplace.Checked = Not mnu4thReplace.Checked
        '<EhFooter>
        Exit Sub

mnu4thReplace_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnu4thReplace_Click" + " line: " + Str(Erl))

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

Private Sub mnuAspect_Click()
        '<EhHeader>
        On Error GoTo mnuAspect_Click_Err
        '</EhHeader>
100 Aspect.Show vbModal
        '<EhFooter>
        Exit Sub

mnuAspect_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuAspect_Click" + " line: " + Str(Erl))

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

Private Sub mnuCompensatingFill_Click()
        '<EhHeader>
        On Error GoTo mnuCompensatingFill_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuCompensatingFill_Click
        '<EhFooter>
        Exit Sub

mnuCompensatingFill_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuCompensatingFill_Click" + " line: " + Str(Erl))

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



Private Sub mnuCopy_Click()
        '<EhHeader>
        On Error GoTo mnuCopy_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuCopy_Click
        '<EhFooter>
        Exit Sub

mnuCopy_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuCopy_Click" + " line: " + Str(Erl))

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

Private Sub mnuCustomLoad_Click()
        '<EhHeader>
        On Error GoTo mnuCustomLoad_Click_Err
        '</EhHeader>

100 If GetLoadName("(*.kla)|*.kla|(*.*)|*.*") = True Then

102 adr_LoadorSave = "load"
104 MemMap.Show vbModal

    End If

        '<EhFooter>
        Exit Sub

mnuCustomLoad_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuCustomLoad_Click" + " line: " + Str(Erl))

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

Private Sub mnuCustomSave_Click()
        '<EhHeader>
        On Error GoTo mnuCustomSave_Click_Err
        '</EhHeader>

100 If GetSaveName("(*.kla)|*.kla|(*.*)|*.*") = False Then Exit Sub
102 Call ZoomWindow.WriteUndo
104 adr_LoadorSave = "save"
106 MemMap.Show vbModal

        '<EhFooter>
        Exit Sub

mnuCustomSave_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuCustomSave_Click" + " line: " + Str(Erl))

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

Private Sub mnuCustomScreenMode_Click()
        '<EhHeader>
        On Error GoTo mnuCustomScreenMode_Click_Err
        '</EhHeader>
100 Custom.Show vbModal
        '<EhFooter>
        Exit Sub

mnuCustomScreenMode_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuCustomScreenMode_Click" + " line: " + Str(Erl))

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

Private Sub mnuDebugChar_Click()
        '<EhHeader>
        On Error GoTo mnuDebugChar_Click_Err
        '</EhHeader>
100 Call ZoomWindow.debugger(Ax, Ay)
        '<EhFooter>
        Exit Sub

mnuDebugChar_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuDebugChar_Click" + " line: " + Str(Erl))

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

Private Sub mnuExit_Click()
        '<EhHeader>
        On Error GoTo mnuExit_Click_Err
        '</EhHeader>

    'Call ZoomWindow.form_unload
100 Unload Me

        '<EhFooter>
        Exit Sub

mnuExit_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuExit_Click" + " line: " + Str(Erl))

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

Private Sub mnuGfxModeCustom_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeCustom_Click_Err
        '</EhHeader>

100 If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102 Custom.Show vbModal
104 Call ZoomWindow.mnuGfxModeCustom_Click
106 Call ZoomWindow.ModeChangeReset
    End If
        '<EhFooter>
        Exit Sub

mnuGfxModeCustom_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeCustom_Click" + " line: " + Str(Erl))

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

Private Sub mnuGfxModeDrazlace_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeDrazlace_Click_Err
        '</EhHeader>
100 If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102 Call ZoomWindow.mnuGfxModeDrazlace_Click
104 Call ZoomWindow.ModeChangeReset
    End If
        '<EhFooter>
        Exit Sub

mnuGfxModeDrazlace_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeDrazlace_Click" + " line: " + Str(Erl))

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

Private Sub mnuGfxModeDrazlaceSpec_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeDrazlaceSpec_Click_Err
        '</EhHeader>
100 If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102 Call ZoomWindow.mnuGfxModeDrazlaceSpec_Click
104 Call ZoomWindow.ModeChangeReset
    End If
        '<EhFooter>
        Exit Sub

mnuGfxModeDrazlaceSpec_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeDrazlaceSpec_Click" + " line: " + Str(Erl))

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

Private Sub mnuGfxModeFli_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeFli_Click_Err
        '</EhHeader>
100 If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102 Call ZoomWindow.mnuGfxModeFli_Click
104 Call ZoomWindow.ModeChangeReset
    End If
        '<EhFooter>
        Exit Sub

mnuGfxModeFli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeFli_Click" + " line: " + Str(Erl))

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
Private Sub mnuGfxModeAfli_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeAfli_Click_Err
        '</EhHeader>

100 If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102     Call ZoomWindow.mnuGfxModeAfli_Click
104     Call ZoomWindow.ModeChangeReset
    End If

        '<EhFooter>
        Exit Sub

mnuGfxModeAfli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeAfli_Click" + " line: " + Str(Erl))

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
Private Sub mnuGfxModeHires_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeHires_Click_Err
        '</EhHeader>

100 If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102     Call ZoomWindow.mnuGfxModeHires_Click
104     Call ZoomWindow.ModeChangeReset
    End If

        '<EhFooter>
        Exit Sub

mnuGfxModeHires_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeHires_Click" + " line: " + Str(Erl))

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

Private Sub mnuGfxModeIFli_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeIFli_Click_Err
        '</EhHeader>
100 If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102 Call ZoomWindow.mnuGfxModeIFli_Click
104 Call ZoomWindow.ModeChangeReset
    End If
        '<EhFooter>
        Exit Sub

mnuGfxModeIFli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeIFli_Click" + " line: " + Str(Erl))

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

Private Sub mnuGfxModeKoala_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeKoala_Click_Err
        '</EhHeader>

100     If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102         Call ZoomWindow.mnuGfxModeKoala_Click
104         Call ZoomWindow.ModeChangeReset
        End If
    
        '<EhFooter>
        Exit Sub

mnuGfxModeKoala_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeKoala_Click" + " line: " + Str(Erl))

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




Private Sub mnuGfxModeUnrestrictedHires_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeUnrestrictedHires_Click_Err
        '</EhHeader>

100     If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102         Call ZoomWindow.mnuGfxModeUnrestrictedHires_Click
104         Call ZoomWindow.ModeChangeReset
        End If
    
        '<EhFooter>
        Exit Sub

mnuGfxModeUnrestrictedHires_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeUnrestrictedHires_Click" + " line: " + Str(Erl))

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
Private Sub mnuGfxModeUnrestrictedMulti_Click()
        '<EhHeader>
        On Error GoTo mnuGfxModeUnrestrictedMulti_Click_Err
        '</EhHeader>

100     If MsgBox("This operation can not be undone." & vbCrLf & " Do you want to continue?", vbYesNo + vbExclamation, "Warning") = vbYes Then
102         Call ZoomWindow.mnuGfxModeUnrestrictedMulti_Click
104         Call ZoomWindow.ModeChangeReset
        End If
    
        '<EhFooter>
        Exit Sub

mnuGfxModeUnrestrictedMulti_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGfxModeUnrestrictedMulti_Click" + " line: " + Str(Erl))

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
Private Sub mnuGridOptions_Click()
        '<EhHeader>
        On Error GoTo mnuGridOptions_Click_Err
        '</EhHeader>

100 Load GridOptions
102 GridOptions.Show vbModal
104 Call ZoomWindow.ZoomWinReDraw(ZoomWinLeft, ZoomWinTop)

        '<EhFooter>
        Exit Sub

mnuGridOptions_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuGridOptions_Click" + " line: " + Str(Erl))

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

Private Sub mnuHelp_Click()
        '<EhHeader>
        On Error GoTo mnuHelp_Click_Err
        '</EhHeader>

100 HTMLHelpContents
    'HTMLHelpIndex
        '<EhFooter>
        Exit Sub

mnuHelp_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuHelp_Click" + " line: " + Str(Erl))

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

Public Sub mnuInterlaceEmu_Click()
        '<EhHeader>
        On Error GoTo mnuInterlaceEmu_Click_Err
        '</EhHeader>

100 If ResoDiv = 1 And BmpBanks = 1 Then

102     If mnuInterlaceEmu.Checked = True Then
104         mnuInterlaceEmu.Checked = False
106         PrevWin.Timer1.Enabled = False
108         Call PrevWin.ReDraw
        Else
110         mnuInterlaceEmu.Checked = True
112         Call ZoomWindow.lacesetup
114         PrevWin.Timer1.Enabled = True
        End If

    End If

        '<EhFooter>
        Exit Sub

mnuInterlaceEmu_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuInterlaceEmu_Click" + " line: " + Str(Erl))

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

Private Sub mnuLoadFli_Click()
        '<EhHeader>
        On Error GoTo mnuLoadFli_Click_Err
        '</EhHeader>
100 Call LoadFli
        '<EhFooter>
        Exit Sub

mnuLoadFli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuLoadFli_Click" + " line: " + Str(Erl))

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

Private Sub mnuLoadOwn_Click()
        '<EhHeader>
        On Error GoTo mnuLoadOwn_Click_Err
        '</EhHeader>
100 Call ZoomWindow.LoadOwnFormat
        '<EhFooter>
        Exit Sub

mnuLoadOwn_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuLoadOwn_Click" + " line: " + Str(Erl))

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

Private Sub mnuLoadPicture_Click()
        '<EhHeader>
        On Error GoTo mnuLoadPicture_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuLoadPicture_Click
        '<EhFooter>
        Exit Sub

mnuLoadPicture_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuLoadPicture_Click" + " line: " + Str(Erl))

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

Private Sub mnuPalette_Click()
        '<EhHeader>
        On Error GoTo mnuPalette_Click_Err
        '</EhHeader>
100 PalSelect.Show vbModal
102 Call PaletteInit
        '<EhFooter>
        Exit Sub

mnuPalette_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuPalette_Click" + " line: " + Str(Erl))

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

Private Sub mnuPaletteWindow_Click()
        '<EhHeader>
        On Error GoTo mnuPaletteWindow_Click_Err
        '</EhHeader>
100 Palett.Visible = Not Palett.Visible
102 mnuPaletteWindow.Checked = Palett.Visible
        '<EhFooter>
        Exit Sub

mnuPaletteWindow_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuPaletteWindow_Click" + " line: " + Str(Erl))

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




Private Sub mnuPaste_Click()
        '<EhHeader>
        On Error GoTo mnuPaste_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuPaste_Click
        '<EhFooter>
        Exit Sub

mnuPaste_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuPaste_Click" + " line: " + Str(Erl))

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

Private Sub mnuPreviewWindow_Click()
        '<EhHeader>
        On Error GoTo mnuPreviewWindow_Click_Err
        '</EhHeader>
100 PrevWin.Visible = Not PrevWin.Visible
102 mnuPreviewWindow.Checked = PrevWin.Visible
        '<EhFooter>
        Exit Sub

mnuPreviewWindow_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuPreviewWindow_Click" + " line: " + Str(Erl))

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

Private Sub mnuRedraw_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo mnuRedraw_Click_Err
        '</EhHeader>
100 Call DrawPicFromMem
        '<EhFooter>
        Exit Sub

mnuRedraw_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuRedraw_Click" + " line: " + Str(Erl))

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

Private Sub mnuSaveFli_Click()
        '<EhHeader>
        On Error GoTo mnuSaveFli_Click_Err
        '</EhHeader>
100 Call SaveFli
        '<EhFooter>
        Exit Sub

mnuSaveFli_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuSaveFli_Click" + " line: " + Str(Erl))

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

Private Sub mnuSaveOwn_Click()
        '<EhHeader>
        On Error GoTo mnuSaveOwn_Click_Err
        '</EhHeader>
100 Call ZoomWindow.SaveOwnFormat
        '<EhFooter>
        Exit Sub

mnuSaveOwn_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuSaveOwn_Click" + " line: " + Str(Erl))

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

Private Sub mnuExportPicture_Click()
        '<EhHeader>
        On Error GoTo mnuExportPicture_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuExportPicture_Click
        '<EhFooter>
        Exit Sub

mnuExportPicture_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuExportPicture_Click" + " line: " + Str(Erl))

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

Private Sub mnuSavePicture_Click()
        '<EhHeader>
        On Error GoTo mnuSavePicture_Click_Err
        '</EhHeader>

100 Select Case GfxMode

        Case "koala"
102         Call SaveKoala
104     Case "Hires"
106         MsgBox "Fast save is not implemented for HIRES, use Custom Save"
108     Case "Drazlace"
110         Call SaveDrazlace
112     Case "Afli"
114         MsgBox "Fast save is not implemented for AFLI, use Custom Save"
116     Case "Ifli"
118         MsgBox "Fast save is not implemented for IFLI, use Custom Save"
120     Case "fli"
122         Call SaveFli
124     Case "Drazlace Special"
126         MsgBox "Fast save is not implemented for Drazlace Special, use Custom Save"
128     Case "Unrestricted"
130         MsgBox "Fast save is not  implemented for Unrestricted mode, use Piture Export"
    End Select

        '<EhFooter>
        Exit Sub

mnuSavePicture_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuSavePicture_Click" + " line: " + Str(Erl))

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

Private Sub mnustrictfill_Click()
        '<EhHeader>
        On Error GoTo mnustrictfill_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnustrictfill_Click
        '<EhFooter>
        Exit Sub

mnustrictfill_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnustrictfill_Click" + " line: " + Str(Erl))

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



Private Sub mnuZoomIn_Click()
        '<EhHeader>
        On Error GoTo mnuZoomIn_Click_Err
        '</EhHeader>
100 Call ZoomWindow.ZoomWinZoomIn
        '<EhFooter>
        Exit Sub

mnuZoomIn_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuZoomIn_Click" + " line: " + Str(Erl))

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

Private Sub mnuZoomInPrev_Click()
        '<EhHeader>
        On Error GoTo mnuZoomInPrev_Click_Err
        '</EhHeader>
100 Call PrevWin.ZoomInPrev
        '<EhFooter>
        Exit Sub

mnuZoomInPrev_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuZoomInPrev_Click" + " line: " + Str(Erl))

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

Private Sub mnuZoomOut_Click()
        '<EhHeader>
        On Error GoTo mnuZoomOut_Click_Err
        '</EhHeader>
100 Call ZoomWindow.ZoomWinZoomOut
        '<EhFooter>
        Exit Sub

mnuZoomOut_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuZoomOut_Click" + " line: " + Str(Erl))

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

Private Sub mnuZoomOutPrev_Click()
        '<EhHeader>
        On Error GoTo mnuZoomOutPrev_Click_Err
        '</EhHeader>
100 Call PrevWin.ZoomOutPrev
        '<EhFooter>
        Exit Sub

mnuZoomOutPrev_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuZoomOutPrev_Click" + " line: " + Str(Erl))

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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar1_ButtonClick_Err
        '</EhHeader>
    Dim X As Long

100 For X = 1 To 7
102     Toolbar1.Buttons(X).Value = tbrUnpressed
104 Next X

106 Button.Value = tbrPressed

108 MainWin.MousePointer = vbCustom

110 If Button.Index = 1 Then
112     ZoomWindow.ActiveTool = "draw"
114     MainWin.MouseIcon = LoadPicture(App.Path & "\Cursors\pencil.ico")
    End If

116 If Button.Index = 3 Then
118     ZoomWindow.ActiveTool = "fill"
120     MainWin.MouseIcon = LoadPicture(App.Path & "\Cursors\fill.ico")
    End If

122 If Button.Index = 5 Then
124     ZoomWindow.ActiveTool = "brush"
126     MainWin.MouseIcon = LoadPicture(App.Path & "\Cursors\brush.ico")
    End If

128 If Button.Index = 7 Then
130     ZoomWindow.ActiveTool = "copy"
132     MainWin.MouseIcon = LoadPicture(App.Path & "\Cursors\cross.ico")
    End If


        '<EhFooter>
        Exit Sub

Toolbar1_ButtonClick_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.Toolbar1_ButtonClick" + " line: " + Str(Erl))

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
        '<EhHeader>
        On Error GoTo form_keydown_Err
        '</EhHeader>
100 Call ZoomWindow.Shared_keydown(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

form_keydown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.form_keydown" + " line: " + Str(Erl))

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
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.form_keyup" + " line: " + Str(Erl))

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
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
        '<EhHeader>
        On Error GoTo Toolbar1_ButtonMenuClick_Err
        '</EhHeader>

100 Call ResetButtonMenuCaption


102 Select Case ButtonMenu.Key
    Case "a"
104     Toolbar1.Buttons(3).ButtonMenus(1).Text = "x Strict"
106     FillMode = "strict"
108 Case "b"
110     Toolbar1.Buttons(3).ButtonMenus(2).Text = "x Compensating"
112     FillMode = "compensating"
114 Case "c"
116     BrushDialog.Show vbModal
118     Call BrushPreCalc
    End Select

        '<EhFooter>
        Exit Sub

Toolbar1_ButtonMenuClick_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.Toolbar1_ButtonMenuClick" + " line: " + Str(Erl))

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
Public Sub ResetButtonMenuCaption()
        '<EhHeader>
        On Error GoTo ResetButtonMenuCaption_Err
        '</EhHeader>

100 Toolbar1.Buttons(3).ButtonMenus(1).Text = "  Strict"
102 Toolbar1.Buttons(3).ButtonMenus(2).Text = "  Compensating"

        '<EhFooter>
        Exit Sub

ResetButtonMenuCaption_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.ResetButtonMenuCaption" + " line: " + Str(Erl))

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
Private Sub mnuSaveKoala_Click()
        '<EhHeader>
        On Error GoTo mnuSaveKoala_Click_Err
        '</EhHeader>
100 If GetSaveName("(*.kla)|*.kla|(*.*)|*.*") = False Then Exit Sub
102 Call SaveKoala
        '<EhFooter>
        Exit Sub

mnuSaveKoala_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuSaveKoala_Click" + " line: " + Str(Erl))

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

Private Sub mnuLoadKoala_click()
        '<EhHeader>
        On Error GoTo mnuLoadKoala_click_Err
        '</EhHeader>
100 Call LoadKoala
        '<EhFooter>
        Exit Sub

mnuLoadKoala_click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuLoadKoala_click" + " line: " + Str(Erl))

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

Private Sub mnusavedrazlace_click()
        '<EhHeader>
        On Error GoTo mnusavedrazlace_click_Err
        '</EhHeader>
100 If GetSaveName("(*.drl)|*.drl|(*.*)|*.*") = False Then Exit Sub
102 Call SaveDrazlace
        '<EhFooter>
        Exit Sub

mnusavedrazlace_click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnusavedrazlace_click" + " line: " + Str(Erl))

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

Private Sub mnuLoadDrazlace_click()
        '<EhHeader>
        On Error GoTo mnuLoadDrazlace_click_Err
        '</EhHeader>
100 Call LoadDrazlace
        '<EhFooter>
        Exit Sub

mnuLoadDrazlace_click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuLoadDrazlace_click" + " line: " + Str(Erl))

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

Private Sub mnuUndo_Click()
        '<EhHeader>
        On Error GoTo mnuUndo_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuUndo_Click
        '<EhFooter>
        Exit Sub

mnuUndo_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuUndo_Click" + " line: " + Str(Erl))

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
Private Sub mnuRedo_Click()
        '<EhHeader>
        On Error GoTo mnuRedo_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuRedo_Click
        '<EhFooter>
        Exit Sub

mnuRedo_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuRedo_Click" + " line: " + Str(Erl))

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
Private Sub mnuBrush_Click()
        '<EhHeader>
        On Error GoTo mnuBrush_Click_Err
        '</EhHeader>
100 Call ZoomWindow.mnuBrush_Click
        '<EhFooter>
        Exit Sub

mnuBrush_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuBrush_Click" + " line: " + Str(Erl))

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

Private Sub mnuDitherfill_Click()
        '<EhHeader>
        On Error GoTo mnuDitherfill_Click_Err
        '</EhHeader>

100 Call ZoomWindow.mnuDitherfill_Click

        '<EhFooter>
        Exit Sub

mnuDitherfill_Click_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mnuDitherfill_Click" + " line: " + Str(Erl))

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

Public Sub invalidate()
        '<EhHeader>
        On Error GoTo invalidate_Err
        '</EhHeader>
    Dim client_rect As RECT
    Dim client_hwnd As Long
    ' Invalidate the picture.
100 client_hwnd = FindWindowEx(Me.hWnd, 0, "MDIClient", vbNullChar)
102 GetClientRect client_hwnd, client_rect
104 InvalidateRect client_hwnd, client_rect, True
        '<EhFooter>
        Exit Sub

invalidate_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.invalidate" + " line: " + Str(Erl))

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

Public Sub mdiform_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo mdiform_Mousedown_Err
        '</EhHeader>

100 actButton = Button
102 actShift = Shift
104 OldX = X
106 yOld = Y
    'MainWin.Text1.Text = "!"
108 ZoomWindow.EventSource = "zoompic"
110 Call ZoomWindow.Shared_MouseDown(Button, Shift, X, Y)

        '<EhFooter>
        Exit Sub

mdiform_Mousedown_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mdiform_Mousedown" + " line: " + Str(Erl))

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

Public Sub mdiform_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo mdiform_Mousemove_Err
        '</EhHeader>

    'actButton = Button
100 actShift = Shift
102 OldX = X
104 yOld = Y

106 MouseOver = "zoompic"
108 ZoomWindow.EventSource = "zoompic"
110 Call ZoomWindow.ZoomPic_MouseMove(Button, Shift, X, Y)

        '<EhFooter>
        Exit Sub

mdiform_Mousemove_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mdiform_Mousemove" + " line: " + Str(Erl))

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

Public Sub mdiform_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo mdiform_Mouseup_Err
        '</EhHeader>

100 actButton = Button
102 actShift = Shift
104 OldX = X
106 yOld = Y

108 ZoomWindow.EventSource = "zoompic"
110 Call ZoomWindow.Shared_MouseUp(Button, Shift, X, Y)

        '<EhFooter>
        Exit Sub

mdiform_Mouseup_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mdiform_Mouseup" + " line: " + Str(Erl))

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

'very fast mouseclicks will be interpreted as doubleclicks
'to not miss a 2nd very close click we need this hack
Public Sub mdiform_dblclick()
        '<EhHeader>
        On Error GoTo mdiform_dblclick_Err
        '</EhHeader>

100 ZoomWindow.EventSource = "zoompic"
102 Call ZoomWindow.Shared_MouseDown(actButton, actShift, OldX, yOld)
104 Call ZoomWindow.Shared_MouseUp(actButton, actShift, OldX, yOld)

        '<EhFooter>
        Exit Sub

mdiform_dblclick_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mdiform_dblclick" + " line: " + Str(Erl))

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

Private Sub mdiform_resize()
        '<EhHeader>
        On Error GoTo mdiform_resize_Err
        '</EhHeader>

100 Call ZoomWindow.ZoomResize

        '<EhFooter>
        Exit Sub

mdiform_resize_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.mdiform_resize" + " line: " + Str(Erl))

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

Public Sub VScroll1_scroll()
        '<EhHeader>
        On Error GoTo VScroll1_scroll_Err
        '</EhHeader>
100 Call ZoomWindow.VScroll1_scroll
        '<EhFooter>
        Exit Sub

VScroll1_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.VScroll1_scroll" + " line: " + Str(Erl))

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
Public Sub HScroll1_scroll()
        '<EhHeader>
        On Error GoTo HScroll1_scroll_Err
        '</EhHeader>
100 Call ZoomWindow.HScroll1_scroll
        '<EhFooter>
        Exit Sub

HScroll1_scroll_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.HScroll1_scroll" + " line: " + Str(Erl))

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
Public Sub VScroll1_Change()
        '<EhHeader>
        On Error GoTo VScroll1_Change_Err
        '</EhHeader>
100 Call ZoomWindow.VScroll1_Change
        '<EhFooter>
        Exit Sub

VScroll1_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.VScroll1_Change" + " line: " + Str(Erl))

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
Public Sub HScroll1_Change()
        '<EhHeader>
        On Error GoTo HScroll1_Change_Err
        '</EhHeader>
100 Call ZoomWindow.HScroll1_Change
        '<EhFooter>
        Exit Sub

HScroll1_Change_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.HScroll1_Change" + " line: " + Str(Erl))

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

Public Sub MouseHelper1_MouseWheel(Ctrl As Variant, Direction As MBMouseHelper.mbDirectionConstants, Button As Long, Shift As Long, Cancel As Boolean)
        '<EhHeader>
        On Error GoTo MouseHelper1_MouseWheel_Err
        '</EhHeader>

100 Call ZoomWindow.MouseHelper1_MouseWheel(Ctrl, Direction, Button, Shift, Cancel)

        '<EhFooter>
        Exit Sub

MouseHelper1_MouseWheel_Err:
    Select Case MsgBox(Error(VBA.Err.Number), vbCritical + vbAbortRetryIgnore, "Error Number" + Str(VBA.Err.Number) + " at " + "Project_One.MainWin.MouseHelper1_MouseWheel" + " line: " + Str(Erl))

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


