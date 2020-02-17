VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmpaint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   Caption         =   "Advanced Paint Brush"
   ClientHeight    =   5895
   ClientLeft      =   2505
   ClientTop       =   1590
   ClientWidth     =   6210
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00400000&
   Icon            =   "frmPaint.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmPaint.frx":0442
   Palette         =   "frmPaint.frx":0594
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5895
   ScaleWidth      =   6210
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "QUIT  "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7440
      Width           =   1935
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6150
      TabIndex        =   30
      Top             =   5520
      Width           =   6210
   End
   Begin VB.CommandButton cmdImi 
      Height          =   375
      Left            =   4680
      Picture         =   "frmPaint.frx":226A
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Image Information"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdHelp 
      Height          =   375
      Left            =   5760
      Picture         =   "frmPaint.frx":2364
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Help"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdText 
      Height          =   375
      Left            =   5040
      Picture         =   "frmPaint.frx":2896
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Text"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdZoomin 
      Height          =   405
      Left            =   3960
      Picture         =   "frmPaint.frx":2998
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Zooms in the picture"
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton cmdZoomout 
      Height          =   405
      Left            =   3600
      Picture         =   "frmPaint.frx":2A82
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Zooms the picture out "
      Top             =   120
      Width           =   345
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   600
      TabIndex        =   14
      Top             =   840
      Width           =   1455
      Begin VB.Frame fraPaint 
         Height          =   735
         Left            =   240
         TabIndex        =   24
         Top             =   0
         Width           =   615
         Begin VB.Image imgPaint 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            Picture         =   "frmPaint.frx":2B6C
            ToolTipText     =   "Draw with a pen"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraText 
         Height          =   735
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   2880
         Width           =   615
         Begin VB.Image imgairbrush2 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   120
            Picture         =   "frmPaint.frx":31EE
            Stretch         =   -1  'True
            ToolTipText     =   "Spray brush"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   615
         Begin VB.Image imgairbrush 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   120
            Picture         =   "frmPaint.frx":A2B8
            Stretch         =   -1  'True
            ToolTipText     =   "Thick line Air brush"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraFilledCircle 
         Height          =   735
         Left            =   840
         TabIndex        =   21
         Top             =   2160
         Width           =   615
         Begin VB.Image imgFilledCircle 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            Picture         =   "frmPaint.frx":11382
            ToolTipText     =   "Draw a Filled Circle"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraCircle 
         Height          =   735
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   615
         Begin VB.Image imgCircle 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            Picture         =   "frmPaint.frx":11A84
            ToolTipText     =   "Draw a Circle"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraFilledSquare 
         Height          =   735
         Left            =   840
         TabIndex        =   19
         Top             =   1440
         Width           =   615
         Begin VB.Image imgFilledSquare 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            Picture         =   "frmPaint.frx":12186
            ToolTipText     =   "Draw a Filled Rectangle "
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraSquare 
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   615
         Begin VB.Image imgSquare 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            Picture         =   "frmPaint.frx":12808
            ToolTipText     =   "Draw a Rectangle "
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraRegion 
         Height          =   735
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   615
         Begin VB.Image imgRegion 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   120
            Picture         =   "frmPaint.frx":12E8A
            ToolTipText     =   "Pick Color"
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.Frame fraFill 
         Height          =   735
         Left            =   840
         TabIndex        =   16
         Top             =   0
         Width           =   615
         Begin VB.Image imgFill 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            Picture         =   "frmPaint.frx":1340C
            ToolTipText     =   "Fill the area "
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraErase 
         Height          =   735
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   615
         Begin VB.Image imgErase 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            Picture         =   "frmPaint.frx":13A8E
            ToolTipText     =   "Erase any part of the picture"
            Top             =   240
            Width           =   390
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   4560
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   6135
      Left            =   2280
      MousePointer    =   99  'Custom
      ScaleHeight     =   6075
      ScaleWidth      =   8715
      TabIndex        =   12
      ToolTipText     =   "Picture Board"
      Top             =   840
      Width           =   8775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   6960
      Width           =   8775
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   6135
      Left            =   11040
      TabIndex        =   10
      Top             =   840
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      Left            =   240
      Max             =   20
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPaste 
      Height          =   405
      Left            =   2880
      Picture         =   "frmPaint.frx":14190
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Paste"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   405
      Left            =   2520
      Picture         =   "frmPaint.frx":146C2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Copy"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdCut 
      BackColor       =   &H8000000A&
      Height          =   405
      Left            =   2160
      Picture         =   "frmPaint.frx":14BF4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cut "
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   405
      Left            =   1440
      Picture         =   "frmPaint.frx":15126
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print"
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton cmdSave 
      Height          =   405
      Left            =   1080
      Picture         =   "frmPaint.frx":15658
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save to disk "
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   405
      Left            =   720
      Picture         =   "frmPaint.frx":1575A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open file"
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton cmdNew 
      Height          =   405
      Left            =   360
      Picture         =   "frmPaint.frx":1585C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "New file "
      Top             =   120
      Width           =   345
   End
   Begin VB.Timer tmrCursor 
      Interval        =   1
      Left            =   2640
      Top             =   5400
   End
   Begin VB.PictureBox pCol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   240
      MouseIcon       =   "frmPaint.frx":15EC6
      MousePointer    =   99  'Custom
      Picture         =   "frmPaint.frx":161D0
      ScaleHeight     =   1110
      ScaleWidth      =   2145
      TabIndex        =   0
      ToolTipText     =   "Color Palette"
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Image arrow 
      Height          =   480
      Left            =   0
      Picture         =   "frmPaint.frx":1DEF2
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cross 
      Height          =   480
      Left            =   480
      Picture         =   "frmPaint.frx":1E044
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image spray 
      Height          =   480
      Left            =   120
      Picture         =   "frmPaint.frx":1E34E
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Eraser 
      Height          =   480
      Left            =   120
      Picture         =   "frmPaint.frx":1E658
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblPenSize 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblXY 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      ToolTipText     =   "Shows X and Y coordinates"
      Top             =   7200
      Width           =   2925
   End
   Begin VB.Image target 
      Height          =   480
      Left            =   3000
      Picture         =   "frmPaint.frx":1E7AA
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bucket 
      Height          =   480
      Left            =   2400
      Picture         =   "frmPaint.frx":1EAB4
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image curpencil 
      Height          =   480
      Left            =   2880
      Picture         =   "frmPaint.frx":1EC06
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1080
      Top             =   4680
      Width           =   615
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFilesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnufiles 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRU 
         Caption         =   "MRU"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuEditClrclip 
         Caption         =   "Clear Clipboard"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsPaint 
         Caption         =   "Paint"
      End
      Begin VB.Menu mnuToolsErase 
         Caption         =   "Erase"
      End
      Begin VB.Menu mnutsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSpray 
         Caption         =   "Spray"
         Begin VB.Menu mnuToolsAB 
            Caption         =   "Air Brush 1"
         End
         Begin VB.Menu mnuToolsAB2 
            Caption         =   "Air Brush 2"
         End
      End
      Begin VB.Menu mnuToolsFill 
         Caption         =   "Fill"
      End
      Begin VB.Menu mnutsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsCircle 
         Caption         =   "Circle"
      End
      Begin VB.Menu mnuToolsFC 
         Caption         =   "Filled Circle"
      End
      Begin VB.Menu mnutsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSquare 
         Caption         =   "Rectangle"
      End
      Begin VB.Menu mnuToolsFR 
         Caption         =   "Filled Rectangle"
      End
      Begin VB.Menu mnutsep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsEyedrop 
         Caption         =   "Eye Dropper"
      End
      Begin VB.Menu mnutsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsGrad 
         Caption         =   "Gradient"
      End
      Begin VB.Menu mnuToolsCirgra 
         Caption         =   "Circular Gradient"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      Begin VB.Menu mnuImageFlipv 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu mnuImagefiph 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu mnuImagesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageinfo 
         Caption         =   "Image Information"
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "&Filter"
      Begin VB.Menu mnuFilterBlur 
         Caption         =   "Blur"
      End
      Begin VB.Menu mnufsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilterlighten 
         Caption         =   "Lighten"
      End
      Begin VB.Menu mnuFilterDarken 
         Caption         =   "Darken"
      End
   End
   Begin VB.Menu mnueffects 
      Caption         =   "Effect&s"
      Begin VB.Menu mnuEffectsStretch 
         Caption         =   "Stretch"
      End
      Begin VB.Menu mnuesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEffectsInvert 
         Caption         =   "Invert"
      End
      Begin VB.Menu mnuesep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEffectsEmb 
         Caption         =   "Emboss"
      End
      Begin VB.Menu mnuEffectsEng 
         Caption         =   "Engrave"
      End
      Begin VB.Menu mnuesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEffectsSweep 
         Caption         =   "Sweeping"
      End
   End
   Begin VB.Menu mnucolor 
      Caption         =   "&Color"
      Begin VB.Menu mnuColorCB 
         Caption         =   "Color Box"
      End
      Begin VB.Menu mnucolorsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorGS 
         Caption         =   "Gray Scale"
      End
      Begin VB.Menu mnuColorBW 
         Caption         =   "Black & White"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelphelp 
         Caption         =   "Help topics"
      End
      Begin VB.Menu mnuhsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmpaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SRCCOPY = &HCC0020
Const Pi = 3.14159265359
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Private Declare Function StretchBlt% Lib "gdi32" (ByVal hdc%, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal xSrc%, ByVal ySrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)

Private sTypes(4) As String
Private Type pointapi
   x As Double
   y As Double
End Type
Const HELP_CONTEXT = &H1
Const HELP_QUIT = &H2
Const HELP_INDEX = &H3
Const HELP_CONTENTS = &H3&
Const HELP_HELPONHELP = &H4
Const HELP_SETINDEX = &H5
Const HELP_SETCONTENTS = &H5&
Const HELP_CONTEXTPOPUP = &H8&
Const HELP_FORCEFILE = &H9&
Const HELP_KEY = &H101
Const HELP_COMMAND = &H102&
Const HELP_PARTIALKEY = &H105&
Const HELP_MULTIKEY = &H201&
Const HELP_SETWINPOS = &H203&

Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Option Explicit

Dim pressed As Boolean     'the variable that i use to tell if the mouse is being held down
Dim colpressed As Boolean  'same as above but for the color chooser
Dim filltool As Boolean    'Variable that tells if the fill tool has been selected
Dim drawtool As Boolean    'variable that tells if the pen tool has been selected
Dim circletool As Boolean  'circle tool variable
Dim rectool As Boolean
Dim linetool As Boolean
Dim fillrectool As Boolean
Dim fillcircle As Boolean
Dim zoomtool As Boolean
Dim spraytool As Boolean
Dim erasetool As Boolean
Dim spraytool2 As Boolean
Dim texttool As Boolean
Dim shape As Boolean
Dim whatl As Variant
Dim whatw As Variant
Dim lenx As Variant
Dim leny As Variant
Dim whatradius As Variant  'used for circle tool
Dim eyedroptool As Boolean 'var for eye dropper
Dim circgradient As Boolean 'var for circualar gradient
Dim radius As Integer
Dim l As Integer
Dim w As Integer
Dim ctx, cty, ctx1, cty1, ctx2, cty2 As Long

'Dim WithEvents PicEx As cPictureEx
Dim picex As PictureBox
Dim about
Dim onlynumbers            'to make sure radius is an integer
Dim point1 As pointapi
Dim point2 As pointapi
Dim g1                     'these 3 are for the gradient tool
Dim g2
Dim g3
Dim cgformat
Dim gformat
Dim gdirection
Dim fdirection As Integer
Dim fformat As Integer
Dim lr1                    'gradient (d)irection,number
Dim lr2
Dim lr3
Dim lr4
Dim ud1
Dim ud2
Dim ud3
Dim ud4
Dim cg1                    'circular gradient variables
Dim cg2
Dim cg3
Dim cg4
Dim Index
Dim index2
Dim index3
Dim index4
Dim i As Integer
Dim c As Integer
Dim bytRed As Integer
Dim bytGreen As Integer
Dim bytBlue As Integer
Dim bytAverage As Integer


Const intUpperBoundX = 600
Const intUpperBoundY = 600
Dim pixels(1 To intUpperBoundX, 1 To intUpperBoundY) As Long

Dim tempbuffer As PictureBox


Dim AllowUndoFlag As Boolean
Dim SaveSmallFlag As Boolean
Dim mresult
Dim mfilespec As String
Dim gcancel As Boolean
' Common Dialog control
Dim gcdg As Object

Dim intXOffset, intYOffset As Integer






Private Sub cmdExit_Click()
Unload Me          'unloads the program
End Sub

Private Sub cmdCopy_Click()
mnuEditCopy_Click

End Sub

Private Sub cmdCut_Click()
mnuEditCut_Click

End Sub

Private Sub cmdHelp_Click()
mnuhelphelp_Click

End Sub

Private Sub cmdImi_Click()
mnuImageinfo_Click
End Sub

Private Sub cmdNew_Click()
mnuFileNew_Click

End Sub

Private Sub cmdOpen_Click()
mnuFileOpen_Click
End Sub

Private Sub cmdPaste_Click()
mnuEditPaste_Click

End Sub

Private Sub cmdPrint_Click()
mnuFilePrint_Click

End Sub

Private Sub cmdSave_Click()
mnuFileSave_Click

End Sub

Private Sub cmdText_Click()
texttool = True
frmText.Show
End Sub

Private Sub cmdZoomin_Click()
picBoard.Picture = picBoard.Image
picBoard.PaintPicture picBoard, 0, 0, picBoard.Width / 1.25, picBoard.Height / 1.25, 0, 0, picBoard.Width, picBoard.Height, vbSrcCopy
picBoard.Width = picBoard.Width / 1.25
picBoard.Height = picBoard.Height / 1.25

End Sub

Private Sub cmdZoomout_Click()
picBoard.Picture = picBoard.Image
picBoard.Width = picBoard.Width * 1.25
picBoard.Height = picBoard.Height * 1.25
picBoard.PaintPicture picBoard, 0, 0, picBoard.Width * 1.25, picBoard.Height * 1.25, 0, 0, picBoard.Width, picBoard.Height, vbSrcCopy



End Sub






Private Sub Command1_Click()
 Dim tmp

        tmp = MsgBox("Save current icon?", vbYesNoCancel + vbQuestion)

        If tmp = vbCancel Then
            Exit Sub
        ElseIf tmp = vbYes Then
            mnuFileSave_Click
        End If

    mnuFileExit_Click

End Sub

Private Sub HScroll1_Change()
picBoard.DataChanged = HScroll1.Value
End Sub

Private Sub Form_Load()
  sTypes(1) = "GIF"
    sTypes(2) = "JPEG"
    sTypes(3) = "PNG"
    sTypes(4) = "BMP"
    'Set PicEx = New cPictureEx
VScroll1.Value = 1
picBoard.DrawWidth = 1
picBoard.MouseIcon = curpencil
pCol.MouseIcon = target
'drawtool = True
Me.ScaleMode = 3
    picBoard.ScaleMode = 3
    
    
    picBoard.AutoRedraw = True
    'Picture1.Picture = Picture1.Image
picBoard.ScaleHeight = 500
picBoard.ScaleWidth = 500

VScroll2.Min = 1
VScroll2.Max = 10000
VScroll2.LargeChange = 1000
VScroll2.SmallChange = 10
HScroll1.Min = 1
HScroll1.Max = 10000
HScroll1.LargeChange = 1000
HScroll1.SmallChange = 10

Dim filename As String
filename = GetSetting(App.Title, "Settings", "Doc1")
If filename <> "" Then
Load mnuMRU(1)
mnuMRU(1).Caption = filename
mnuMRU(1).Visible = True
End If

End Sub

Private Sub imgairbrush_Click()
imgPaint.Appearance = 0
imgFill.Appearance = 0
imgErase.Appearance = 0
imgSquare.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgCircle.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgairbrush.Appearance = 1
mnuToolsAB_Click

End Sub

Private Sub imgairbrush2_Click()
imgPaint.Appearance = 0
imgFill.Appearance = 0
imgErase.Appearance = 0
imgSquare.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgCircle.Appearance = 0
imgairbrush.Appearance = 0
imgRegion.Appearance = 0
imgairbrush2.Appearance = 1
mnuToolsAB2_Click

End Sub

Private Sub imgCircle_Click()
imgPaint.Appearance = 0
imgFill.Appearance = 0
imgErase.Appearance = 0
imgSquare.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgCircle.Appearance = 1
mnuToolsCircle_Click
End Sub

Private Sub imgErase_Click()
imgPaint.Appearance = 0
imgCircle.Appearance = 0
imgSquare.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgFill.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgErase.Appearance = 1
mnuToolsErase_Click

End Sub

Private Sub imgFill_Click()
imgPaint.Appearance = 0
imgCircle.Appearance = 0
imgSquare.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgErase.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgFill.Appearance = 1
mnuToolsFill_Click


End Sub

Private Sub imgFilledCircle_Click()
imgFill.Appearance = 0
imgCircle.Appearance = 0
imgSquare.Appearance = 0
imgErase.Appearance = 0
imgPaint.Appearance = 0
imgFilledSquare.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgFilledCircle.Appearance = 1
mnuToolsFC_Click

End Sub

Private Sub imgFilledSquare_Click()
imgFill.Appearance = 0
imgCircle.Appearance = 0
imgSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgPaint.Appearance = 0
imgErase.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgFilledSquare.Appearance = 1
mnuToolsFR_Click


End Sub

Private Sub imgPaint_Click()
imgFill.Appearance = 0
imgErase.Appearance = 0
imgCircle.Appearance = 0
imgSquare.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgPaint.Appearance = 1

lblPenSize.Visible = True

VScroll1.Visible = True
mnuToolsPaint_Click
End Sub

Private Sub imgRegion_Click()
imgFill.Appearance = 0
imgCircle.Appearance = 0
imgPaint.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgErase.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0

imgSquare.Appearance = 0
imgRegion.Appearance = 1
mnuToolsEyedrop_Click

End Sub

Private Sub imgSquare_Click()
imgFill.Appearance = 0
imgCircle.Appearance = 0
imgPaint.Appearance = 0
imgFilledSquare.Appearance = 0
imgFilledCircle.Appearance = 0
imgErase.Appearance = 0
imgairbrush.Appearance = 0
imgairbrush2.Appearance = 0
imgRegion.Appearance = 0
imgSquare.Appearance = 1

mnuToolsSquare_Click

End Sub

Private Sub imgText_Click()
texttool = True
frmText.Show
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuColorBW_Click()
Me.MousePointer = vbHourglass
 Dim x, y As Integer
    Dim bytRed, bytGreen, bytBlue, bytBlack, bytWhite, bytBW As Integer
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            bytRed = pixels(x, y) And &HFF
            bytGreen = ((pixels(x, y) And &HFF00) / &H100) Mod &H100
            bytBlue = ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100
            
            bytBlack = RGB(0, 0, 0)
            bytWhite = (bytRed + bytGreen + bytBlue) * 2
            bytBW = (bytBlack + bytWhite) / 2
            
            pixels(x, y) = RGB(bytBW, bytBW, bytBW)
         Next y
    Next x
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            picBoard.PSet (x, y), pixels(x, y)
        Next y
    Next x
   Me.MousePointer = vbDefault
    
End Sub

Private Sub mnuColorCB_Click()
On Error GoTo cancel

CommonDialog1.ShowColor
Shape1.FillColor = CommonDialog1.Color

cancel:
End Sub

Private Sub mnuEditClear_Click()
Set picBoard.Picture = Nothing

picBoard.Cls
picBoard.BackColor = &HFFFFFF

End Sub
Private Sub mnuEditClrclip_Click()
Clipboard.Clear
End Sub
Private Sub mnuEditCopy_Click()

Clipboard.Clear
   
      Clipboard.SetData picBoard.Picture
  
End Sub



Private Sub mnuEditCut_Click()

Clipboard.Clear
 
    Clipboard.SetData picBoard.Picture
mnuEditClear_Click




End Sub

Private Sub mnuEditPaste_Click()
      picBoard.Picture = Clipboard.GetData
  End Sub






Private Sub mnuEffectsEmb_Click()
Me.MousePointer = vbHourglass
On Error Resume Next

Const intubx = 600
Const intuby = 600
Dim pixels(1 To intubx, 1 To intuby) As Long

Dim x, y As Integer

    For x = 1 To intubx
        For y = 1 To intuby
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = intubx To 2 Step -1
        For y = intuby To 2 Step -1
            bytRed = ((pixels(x - 1, y - 1) And &HFF) - (pixels(x, y) And &HFF)) + 128
            bytGreen = (((pixels(x - 1, y - 1) And &HFF00) / &H100) Mod &H100 - ((pixels(x, y) And &HFF00) / &H100) Mod &H100) + 128
            bytBlue = (((pixels(x - 1, y - 1) And &HFF0000) / &H1000) Mod &H100 - ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100) + 128
            
            bytAverage = (bytRed + bytGreen + bytBlue) / 3
            pixels(x, y) = RGB(bytAverage, bytAverage, bytAverage)
         Next y
    Next x
    picBoard.Cls
    
    For x = 1 To intubx
        For y = 1 To intuby
            picBoard.PSet (x + 2, y + 2), pixels(x, y)
        Next y
    Next x
    Me.MousePointer = vbDefault
    
End Sub

Private Sub mnuEffectsEng_Click()
Me.MousePointer = vbHourglass
On Error Resume Next

Const intubx = 600
Const intuby = 600
Dim pixels(1 To intubx, 1 To intuby) As Long

Dim x, y As Integer
    
    For x = 1 To intubx
        For y = 1 To intuby
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = 2 To intubx - 1
        For y = 2 To intuby - 1
            bytRed = ((pixels(x + 1, y + 1) And &HFF) - (pixels(x, y) And &HFF)) + 128
            bytGreen = (((pixels(x + 1, y + 1) And &HFF00) / &H100) Mod &H100 - ((pixels(x, y) And &HFF00) / &H100) Mod &H100) + 128
            bytBlue = (((pixels(x + 1, y + 1) And &HFF0000) / &H10000) Mod &H100 - ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100) + 128
            
           bytAverage = (bytRed + bytGreen + bytBlue) / 3
            pixels(x, y) = RGB(bytAverage, bytAverage, bytAverage)
       
         Next y
    Next x
    
    'picBoard.Cls
    
    For x = 1 To intubx
        For y = 1 To intuby
            picBoard.PSet (x - 3, y - 3), pixels(x, y)
        Next y
    Next x
    Me.MousePointer = vbDefault
End Sub


Private Sub mnuEffectsInvert_Click()
Me.MousePointer = vbHourglass
  Dim x, y As Integer
    Dim bytRed, bytGreen, bytBlue, bytA, bytB, bytC, bytAverage As Integer
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            bytRed = pixels(x, y) And &HFF
            bytGreen = ((pixels(x, y) And &HFF00) / &H100) Mod &H100
            bytBlue = ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100
            bytA = bytGreen + 50
            bytB = bytRed / 2
            bytC = bytRed / 2
           ' bytAverage = (bytRed + bytGreen + bytBlue) / 3
            pixels(x, y) = RGB(bytA, bytB, bytC)
         Next y
    Next x
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            picBoard.PSet (x, y), pixels(x, y)
        Next y
    Next x
    Me.MousePointer = vbDefault
    
End Sub



Private Sub mnuEffectsStretch_Click()
picBoard.PaintPicture picBoard.Picture, 0, 0, picBoard.ScaleWidth, picBoard.ScaleHeight

End Sub

Private Sub mnuEffectsSweep_Click()
  Me.MousePointer = vbHourglass
   Dim x, y As Integer
    Dim bytRed, bytGreen, bytBlue As Byte
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = intUpperBoundX - 1 To 1 Step -1
        For y = intUpperBoundY - 1 To 1 Step -1
            bytRed = Abs((pixels(x + 1, y + 1) And &HFF) + (pixels(x, y) And &HFF)) / 2
            bytGreen = Abs(((pixels(x + 1, y + 1) And &HFF00) / &H100) Mod &H100 + ((pixels(x, y) And &HFF00) / &H100) Mod &H100) / 2
            bytBlue = Abs(((pixels(x + 1, y + 1) And &HFF0000) / &H10000) Mod &H100 + ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100) / 2
            pixels(x, y) = RGB(bytRed, bytGreen, bytBlue)
         Next y
    Next x
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            picBoard.PSet (x - 3, y - 3), pixels(x, y)
        Next y
    Next x
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuFilterBlur_Click()
Me.MousePointer = vbHourglass
 Dim x, y As Integer
    Dim bytRed, bytGreen, bytBlue As Byte
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = 1 To intUpperBoundX - 1
        For y = 1 To intUpperBoundY
            bytRed = Abs((pixels(x + 1, y) And &HFF) + (pixels(x, y) And &HFF)) / 2
            bytGreen = Abs(((pixels(x + 1, y) And &HFF00) / &H100) Mod &H100 + ((pixels(x, y) And &HFF00) / &H100) Mod &H100) / 2
            bytBlue = Abs(((pixels(x + 1, y) And &HFF0000) / &H10000) Mod &H100 + ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100) / 2
            pixels(x, y) = RGB(bytRed, bytGreen, bytBlue)
         Next y
    Next x
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            picBoard.PSet (x, y), pixels(x, y)
        Next y
    Next x
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuColorGS_Click()
Me.MousePointer = vbHourglass
  Dim x, y As Integer
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            bytRed = pixels(x, y) And &HFF
            bytGreen = ((pixels(x, y) And &HFF00) / &H100) Mod &H100
            bytBlue = ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100
            
            bytAverage = (bytRed + bytGreen + bytBlue) / 3
            pixels(x, y) = RGB(bytAverage, bytAverage, bytAverage)
         Next y
    Next x
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            picBoard.PSet (x, y), pixels(x, y)
        Next y
    Next x
    Me.MousePointer = vbDefault
    
End Sub

Private Sub mnuFileExit_Click()
Unload Me

End Sub

Private Sub mnuFileNew_Click()
        Dim tmp
        tmp = MsgBox("Save current icon?", vbYesNoCancel + vbQuestion)
        If tmp = vbCancel Then
            Exit Sub
        ElseIf tmp = vbYes Then
            mnuFileSave_Click
        End If
mnuEditClear_Click

End Sub

Private Sub mnuFileOpen_Click()
Dim TheFile As String
On Error GoTo cancel
CommonDialog1.CancelError = True
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Picture Files (*.jpg,*.gif,*.bmp)|*.jpg;*.gif;*.bmp"
CommonDialog1.ShowOpen
TheFile = CommonDialog1.filename
picBoard.Picture = LoadPicture(TheFile)
With CommonDialog1

If GetSetting(App.Title, "Settings", "Doc1") = "" Then
Load mnuMRU(1)
End If
mnuMRU(1).Caption = .filename
mnuMRU(1).Visible = True
SaveSetting App.Title, "Settings", "Doc1", .filename
End With

'tempbuffer.Picture = picBoard.Picture

'picBoard.PaintPicture tempbuffer.Picture, 0, 0, picBoard.Width, picBoard.Height, 0, 0, tempbuffer.Width, tempbuffer.Height, vbSrcCopy

picBoard.ScaleWidth = HScroll1.Value
picBoard.ScaleHeight = VScroll2.Value
 
 ' If Picture.Width > picBoard.ScaleWidth Then
  '   HScroll1.Enabled = True
   ' HScroll1.Max = Picture.Width - picBoard.ScaleWidth
    '    HScroll1.LargeChange = Int(HScroll1.Max \ 25) + 1
     '   HScroll1.SmallChange = Int(HScroll1.Max \ 200) + 1
 ' Else
  '  HScroll1.Enabled = False
   ' End If
    
    'If Picture.Height > picBoard.ScaleHeight Then '
     '   VScroll2.Enabled = True
      '  VScroll2.Max = Picture.Height - picBoard.ScaleHeight
       ' VScroll2.LargeChange = Int(VScroll1.Max \ 25) + 1
       ' VScroll2.SmallChange = Int(VScroll1.Max \ 200) + 1
   ' Else
    '   VScroll2.Enabled = False
    'End If
cancel:
End Sub

Private Sub mnuFilePrint_Click()
frmpaint.PrintForm
End Sub

Private Sub mnuFileSave_Click()

Dim TheFile, tmp
On Error GoTo cancel
CommonDialog1.InitDir = App.Path
CommonDialog1.DefaultExt = ".bmp"

CommonDialog1.Filter = "Bitmaps(*.bmp)|*.bmp"

CommonDialog1.ShowSave
TheFile = CommonDialog1.filename

SavePicture picBoard.Image, TheFile



cancel:

End Sub

Private Sub mnuFilterDarken_Click()
MousePointer = vbHourglass

  Dim x, y As Integer
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            bytRed = pixels(x, y) And &HFF
            bytGreen = ((pixels(x, y) And &HFF00) / &H100) Mod &H100
            bytBlue = ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100
            
            bytRed = bytRed - 100
            If bytRed < 50 Then bytRed = 50
            bytGreen = bytGreen - 100
            If bytGreen < 50 Then bytGreen = 50
            bytBlue = bytBlue - 100
            If bytBlue < 50 Then bytBlue = 50
            
            pixels(x, y) = RGB(bytRed, bytGreen, bytBlue)
         Next y
    Next x
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            picBoard.PSet (x, y), pixels(x, y)
        Next y
    Next x
    MousePointer = vbDefault
End Sub

Private Sub mnuFilterlighten_Click()
MousePointer = vbHourglass
Dim x, y As Integer
    Dim bytRed, bytGreen, bytBlue, bytAverage As Integer
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            pixels(x, y) = picBoard.Point(x, y)
        Next y
    Next x
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            bytRed = pixels(x, y) And &HFF
            bytGreen = ((pixels(x, y) And &HFF00) / &H100) Mod &H100
            bytBlue = ((pixels(x, y) And &HFF0000) / &H10000) Mod &H100
            
            bytRed = bytRed + 100
            If bytRed > 255 Then bytRed = 255
            bytGreen = bytGreen + 100
            If bytGreen > 255 Then bytGreen = 255
            bytBlue = bytBlue + 100
            If bytBlue > 255 Then bytBlue = 255
            
            pixels(x, y) = RGB(bytRed, bytGreen, bytBlue)
         Next y
    Next x
    
    
    For x = 1 To intUpperBoundX
        For y = 1 To intUpperBoundY
            picBoard.PSet (x, y), pixels(x, y)
        Next y
    Next x
    
    MousePointer = vbDefault
End Sub

Private Sub mnuhelphelp_Click()
Dim retval
retval = WinHelp(frmpaint.hwnd, "d:\divya1\6thproject\advpb\painthelp.hlp", HELP_CONTENTS, CLng(0))

End Sub

Private Sub mnuImagefiph_Click()
picBoard.PaintPicture picBoard.Picture, picBoard.ScaleWidth, 0, -1 * picBoard.ScaleWidth, picBoard.ScaleHeight

    
End Sub

Private Sub mnuImageFlipv_Click()
picBoard.PaintPicture picBoard.Picture, 0, picBoard.ScaleHeight, picBoard.ScaleWidth, -1 * picBoard.ScaleHeight


End Sub

Private Sub mnuImageinfo_Click()
 Dim ii As New CImageinfo
    Dim msg As String
If picBoard.Picture Then

 ii.ReadImageInfo (CommonDialog1.filename)
    msg = msg & "FileName: " & CommonDialog1.filename & vbCrLf
    If ii.ImageType Then
        msg = msg & "Width: " & ii.Width & vbCrLf
        msg = msg & "Height: " & ii.Height & vbCrLf
        msg = msg & "Bits per pixel: " & ii.Depth & vbCrLf
        msg = msg & "Type: " & sTypes(ii.ImageType)
    Else
        msg = msg & "Unknown image type"
    End If
    
    MsgBox msg, vbInformation Or vbOKOnly, "Image Information"
Else
MsgBox ("Open a picture file")
End If

End Sub

Private Sub mnuToolsAB_Click()
drawtool = False
erasetool = False
spraytool2 = False

filltool = False
rectool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = False
spraytool = True

End Sub

Private Sub mnuToolsAB2_Click()
drawtool = False
erasetool = False

filltool = False
rectool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = False
spraytool = False
spraytool2 = True

End Sub

Private Sub mnuToolsCircle_Click()

drawtool = False
erasetool = False
spraytool = False
spraytool2 = False

filltool = False
rectool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = True
GetRadius:
whatradius = InputBox("Enter the radius for the circle in pixels:", "Paint")
If IsNumeric(whatradius) Then
    radius = Val(whatradius)
Else
    onlynumbers = MsgBox("You have to enter a number!", vbCritical, "Paint")
GoTo GetRadius
End If

End Sub

Private Sub mnuToolsCirgra_Click()
drawtool = False
erasetool = False

filltool = False
rectool = False
fillrectool = False
fillcircle = False
circletool = False
spraytool = False
spraytool2 = False

circgradient = True

VScroll1.Value = 11
End Sub

Private Sub mnuToolsErase_Click()
drawtool = False

filltool = False
rectool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = False
spraytool = False
spraytool2 = False
eyedroptool = False
erasetool = True
VScroll1.Value = 10
End Sub

Private Sub mnuToolsEyedrop_Click()
drawtool = False
erasetool = False

filltool = False
rectool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = False
spraytool = False
spraytool2 = False
eyedroptool = True


End Sub

Private Sub mnuToolsFC_Click()
drawtool = False
erasetool = False

filltool = False
rectool = False
fillrectool = False

circgradient = False
circletool = False
spraytool = False
spraytool2 = False
eyedroptool = False

fillcircle = True
GetRadius:
whatradius = InputBox("Enter the radius for the circle in pixels:", "Paint")
If IsNumeric(whatradius) Then
    radius = Val(whatradius)
Else
    onlynumbers = MsgBox("You have to enter a number!", vbCritical, "Paint")
GoTo GetRadius
End If
End Sub

Private Sub mnuToolsFill_Click()


drawtool = False
erasetool = False


rectool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = False
spraytool = False
spraytool2 = False
eyedroptool = False

filltool = True
'picBoard.MouseIcon = bucket
End Sub

Private Sub mnuToolsFR_Click()
drawtool = False
erasetool = False

filltool = False
rectool = False

fillcircle = False
circgradient = False
circletool = False
spraytool = False
spraytool2 = False
eyedroptool = False


fillrectool = True
Getendpoints:

whatl = InputBox("Enter the length of the rectangle in pixels:", "Paint")
whatw = InputBox("Enter the width of the rectangle in pixels:", "Paint")

If IsNumeric(whatl) And IsNumeric(whatw) Then
    l = Val(whatl)
    w = Val(whatw)
Else
    onlynumbers = MsgBox("You have to enter a number!", vbCritical, "Paint")
GoTo Getendpoints
End If
End Sub

Private Sub mnuToolsGrad_Click()
Call gradientmaker
End Sub

Private Sub mnuToolsPaint_Click()
picBoard.MouseIcon = curpencil

erasetool = False

filltool = False
rectool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = False
spraytool = False
spraytool2 = False
eyedroptool = False
drawtool = True


End Sub

Private Sub mnuToolsSquare_Click()

drawtool = False
erasetool = False

filltool = False
fillrectool = False
fillcircle = False
circgradient = False
circletool = False
spraytool = False
spraytool2 = False
eyedroptool = False

rectool = True
Getendpoints:

whatl = InputBox("Enter the length of the rectangle in pixels:", "Paint")
whatw = InputBox("Enter the width of the rectangle in pixels:", "Paint")

If IsNumeric(whatl) And IsNumeric(whatw) Then
    l = Val(whatl)
    w = Val(whatw)
Else
    onlynumbers = MsgBox("You have to enter a number!", vbCritical, "Paint")
GoTo Getendpoints
End If

End Sub

Private Sub pCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
colpressed = True                          'this function tells the computer to
Shape1.FillColor = pCol.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = pCol.Point(x, y)
End Sub

Private Sub pCol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If colpressed Then                         'same as above but this is here
Shape1.FillColor = pCol.Point(x, y)        'so you don't have to keep clicking to
picBoard.ForeColor = pCol.Point(x, y)      'change the color...you can just drag
End If
End Sub

Private Sub pCol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
colpressed = False          'stops the selecting of the color when the user 'unclicks'
End Sub



Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
pressed = True
point1.x = x
point1.y = y

If drawtool = True Then                    'draws a point where the user clicks on picBoard
    picBoard.ForeColor = Shape1.FillColor
    picBoard.Line (x, y)-(x, y)
End If


If filltool = True Then                    'if the fill tool is selected it fills picBoard with a custom color
    picBoard.BackColor = Shape1.FillColor
End If
If erasetool = True Then
    picBoard.ForeColor = &HFFFFFF
    picBoard.Line (x, y)-(x, y)
End If

If spraytool = True Then
With frmpaint.picBoard
.DrawWidth = VScroll1.Value
End With
picBoard.PSet (x + 2, y + 1)
picBoard.PSet (x + 2, y + 3)
picBoard.PSet (x - 1, y + 2)
picBoard.PSet (x - 3, y + 2)
picBoard.PSet (x - 2, y - 1)
picBoard.PSet (x - 2, y - 3)
picBoard.PSet (x + 1, y - 2)
picBoard.PSet (x + 3, y - 2)
picBoard.PSet (x, y)
End If

If spraytool2 = True Then
With frmpaint.picBoard
.DrawWidth = VScroll1.Value
End With
Randomize Timer
ctx = x + (1 + (Rnd * 16))
cty = y + (1 + (Rnd * 10))
ctx1 = x + (1 + (Rnd * 6))
cty1 = y + (1 + (Rnd * 7))
ctx2 = x + (1 + (Rnd * 5))
cty2 = y + (1 + (Rnd * 4))



picBoard.PSet (ctx, cty)
picBoard.PSet (ctx1, cty1)
picBoard.PSet (ctx2, cty2)

End If
If circletool = True Then
picBoard.ForeColor = Shape1.FillColor
    picBoard.Circle (point1.x, point1.y), radius
End If
If fillcircle = True Then
picBoard.ForeColor = Shape1.FillColor
    picBoard.FillColor = Shape1.FillColor
    picBoard.FillStyle = vbFSSolid
    picBoard.Circle (point1.x, point1.y), radius
End If

If fillrectool = True Then
picBoard.ForeColor = Shape1.FillColor
    picBoard.Line (point1.x, point1.y)-(w, l), Shape1.FillColor, BF
End If

If rectool = True Then
picBoard.ForeColor = Shape1.FillColor
    picBoard.Line (point1.x, point1.y)-(w, l), , B
End If

If eyedroptool = True Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = picBoard.Point(x, y)
End If



If pressed And circgradient Then
On Error GoTo ErrHandler
cgformat = InputBox("How do you want to format your gradient? (RGB),  1: ##I, 2: #I#, 3: I##, 4:III (black to white)", "Paint", "1,2,3 or 4")
If cgformat = "1" Then GoTo cg1
If cgformat = "2" Then GoTo cg2
If cgformat = "3" Then GoTo cg3
If cgformat = "4" Then GoTo cg4
cg1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")

For Index = 1 To 400 Step 1
    picBoard.Circle (x, y), Index, RGB(g1, g2, Index)
pressed = False
circgradient = False
Exit Sub

cg2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")


For index2 = 1 To 400 Step 1
    picBoard.Circle (x, y), index2, RGB(g1, index2, g2)
Next index2

pressed = False
circgradient = False
Exit Sub
cg3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ")

For index3 = 1 To 400 Step 2
    picBoard.Circle (x, y), index3, RGB(index3, g1, g2)
Next index3

pressed = False
circgradient = False
Exit Sub

cg4:
For index4 = 1 To 400 Step 2
    picBoard.Circle (x, y), index4, RGB(index4, index4, index4)
Next index4

pressed = False
circgradient = False
'drawtool = True
Me.MousePointer = vbDefault

Exit Sub

ErrHandler:
    MsgBox ("An error has occured")
    Exit Sub
Next
End If

End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If pressed And drawtool Then
picBoard.ForeColor = Shape1.FillColor
point2 = point1
point1.x = x
point1.y = y
picBoard.Line (point1.x, point1.y)-(point2.x, point2.y)         'if the mouse is dragged...the line is continued
End If

If pressed And erasetool Then
picBoard.ForeColor = &HFFFFFF

point2 = point1
point1.x = x
point1.y = y
picBoard.Line (point1.x, point1.y)-(point2.x, point2.y)
End If

'If pressed And shape Then

 'On Error Resume Next ' just in case
  '      If (x > 0) And (y > 0) And (x < picBoard.ScaleWidth) And (y < picBoard.ScaleWidth) Then
   '         Shape2.Left = x - (Shape2.Width / 2)
    '        Shape2.Top = y - (Shape2.Height / 2)  ' grab the center of the box
     '       DoEvents
           ' Call picBoard.PaintPicture(picBoard, 0, 0, picBoard.Width, picBoard.Height, Shape2.Left, Shape2.Top, Shape2.Width, Shape2.Height)
            'Clipboard.SetData picBoard.Picture
            'mnuEditCopy_Click
    '          picboard2.Visible = True
   '
   'Call picboard2.PaintPicture(picBoard, 1, 1, picboard2.Width, picboard2.Height, Shape2.Left, Shape2.Top, Shape2.Width, Shape2.Height)
       
    '        DoEvents
     '   End If
'End If
If pressed And spraytool Then
With frmpaint.picBoard
.DrawWidth = VScroll1.Value
.ForeColor = Shape1.FillColor
End With
picBoard.PSet (x + 2, y + 1)
picBoard.PSet (x + 2, y + 3)
picBoard.PSet (x - 1, y + 2)
picBoard.PSet (x - 3, y + 2)
picBoard.PSet (x - 2, y - 1)
picBoard.PSet (x - 2, y - 3)
picBoard.PSet (x + 1, y - 2)
picBoard.PSet (x + 3, y - 2)
picBoard.PSet (x, y)
End If

If pressed And spraytool2 Then
With frmpaint.picBoard
.DrawWidth = VScroll1.Value
.ForeColor = Shape1.FillColor
End With
Randomize Timer
ctx = x + (1 + (Rnd * 16))
cty = y + (1 + (Rnd * 10))
ctx1 = x + (1 + (Rnd * 6))
cty1 = y + (1 + (Rnd * 7))
ctx2 = x + (1 + (Rnd * 5))
cty2 = y + (1 + (Rnd * 4))

picBoard.PSet (ctx, cty)
picBoard.PSet (ctx1, cty1)
picBoard.PSet (ctx2, cty2)
End If

If pressed And eyedroptool Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = picBoard.Point(x, y)
End If

lblXY.Caption = "  X=" & CStr(x) & "  Y=" & CStr(y)

Me.MousePointer = vbDefault
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
pressed = False                    'stops drawing the line

Me.MousePointer = vbDefault
End Sub


Private Sub tmrCursor_Timer()
If drawtool = True Or rectool = True Or fillrectool = True Then
    picBoard.MouseIcon = curpencil

ElseIf filltool = True Then
    picBoard.MouseIcon = bucket

ElseIf circletool = True Or fillcircle = True Then
    picBoard.MouseIcon = target

ElseIf erasetool = True Then
    picBoard.MouseIcon = Eraser

ElseIf spraytool = True Then
    picBoard.MouseIcon = spray
ElseIf spraytool2 = True Then
    picBoard.MouseIcon = spray
ElseIf eyedroptool = True Then
    picBoard.MouseIcon = target
ElseIf shape = True Then
    picBoard.MouseIcon = cross
Else
picBoard.MouseIcon = arrow
'Me.MousePointer = vbDefault

End If




End Sub

Private Sub gradientmaker()
On Error GoTo ErrHandler

gdirection = InputBox("What direction do you want the gradient to fade 1: LEFT-RIGHT  2: UP-DOWN ?", "Paint", "1 or 2")
gformat = InputBox("How do you want to format your gradient? (RGB),  1: ##I, 2: #I#, 3: I##, 4:III (black to white)", "Paint", "1,2,3 or 4")

If gdirection = "1" And gformat = "1" Then GoTo lr1
If gdirection = "1" And gformat = "2" Then GoTo lr2
If gdirection = "1" And gformat = "3" Then GoTo lr3
If gdirection = "1" And gformat = "4" Then GoTo lr4

If gdirection = "2" And gformat = "1" Then GoTo ud1
If gdirection = "2" And gformat = "2" Then GoTo ud2
If gdirection = "2" And gformat = "3" Then GoTo ud3
If gdirection = "2" And gformat = "4" Then GoTo ud4

lr1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")
For i = 1 To 255
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(g1, g2, i)
Next i

Exit Sub
lr2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")

For i = 1 To 255
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(g1, i, g2)
Next i

Exit Sub

lr3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ")

For i = 1 To 255
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(i, g1, g2)
Next i
Exit Sub

lr4:
For i = 1 To 255
    picBoard.Line (i, picBoard.ScaleHeight)-(i, 0), RGB(i, i, i)
Next i
Exit Sub

ud1:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")

For i = 1 To 255
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(g1, g2, i)
Next i
Exit Sub

ud2:
g1 = InputBox("Enter a number 0-255 (This is the value for RED) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")

For i = 1 To 255
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(g1, i, g2)
Next i
Exit Sub

ud3:
g1 = InputBox("Enter a number 0-255 (This is the value for GREEN) ", "Paint", "0")
g2 = InputBox("Enter a number 0-255 (This is the value for BLUE) ", "Paint", "0")

For i = 1 To 255
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(i, g1, g2)
Next i
Exit Sub

ud4:
For i = 1 To 255
    picBoard.Line (picBoard.ScaleHeight, i)-(0, i), RGB(i, i, i)
Next i
Exit Sub

ErrHandler:
    MsgBox ("An error has occured")
End Sub

Private Sub VScroll1_Change()
On Error Resume Next

If VScroll1.Value < 1 Then
MsgBox ("Select a proper width of the pen")
End If



lblPenSize.Caption = VScroll1.Value          'sets the pen size according
picBoard.DrawWidth = VScroll1.Value

End Sub

Private Sub VScroll2_Change()
On Error Resume Next

picBoard.DataChanged = VScroll2.Value

End Sub
