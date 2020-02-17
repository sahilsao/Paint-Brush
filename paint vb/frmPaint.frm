VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmpaint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paint Brush Program by Smita Yadav"
   ClientHeight    =   5475
   ClientLeft      =   540
   ClientTop       =   780
   ClientWidth     =   10380
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   Icon            =   "frmPaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPaint.frx":030A
   Palette         =   "frmPaint.frx":045C
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5475
   ScaleWidth      =   10380
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstTools 
      Height          =   1230
      Left            =   7800
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   5115
      ScaleWidth      =   7515
      TabIndex        =   9
      Top             =   120
      Width           =   7575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   3
      Left            =   7800
      Max             =   25
      Min             =   2
      TabIndex        =   7
      Top             =   600
      Value           =   3
      Width           =   1935
   End
   Begin VB.Timer tmrCursor 
      Interval        =   1
      Left            =   480
      Top             =   5520
   End
   Begin VB.PictureBox pCol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   7800
      MouseIcon       =   "frmPaint.frx":2132
      MousePointer    =   99  'Custom
      Picture         =   "frmPaint.frx":243C
      ScaleHeight     =   1110
      ScaleWidth      =   2145
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image target 
      Height          =   480
      Left            =   2040
      Picture         =   "frmPaint.frx":A15E
      Top             =   5520
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample"
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to choose pen/fill color"
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image bucket 
      Height          =   480
      Left            =   1320
      Picture         =   "frmPaint.frx":A468
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image curpencil 
      Height          =   480
      Left            =   960
      Picture         =   "frmPaint.frx":A5BA
      Top             =   5400
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7800
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pen Size"
      Height          =   255
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblPenSize 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Left            =   8640
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   ""
   End
End
Attribute VB_Name = "frmpaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type pointapi
   x As Double
   y As Double
End Type

Option Explicit

Dim pressed As Boolean     'the variable that i use to tell if the mouse is being held down
Dim colpressed As Boolean  'same as above but for the color chooser
Dim filltool As Boolean    'Variable that tells if the fill tool has been selected
Dim drawtool As Boolean    'variable that tells if the pen tool has been selected
Dim circletool As Boolean  'circle tool variable
Dim whatradius As Variant  'used for circle tool
Dim eyedroptool As Boolean 'var for eye dropper
Dim radius As Integer
Dim about
Dim onlynumbers            'to make sure radius is an integer
Dim point1 As pointapi
Dim point2 As pointapi

Dim index
Dim index2
Dim index3
Dim index4
Dim i As Integer
Dim c As Integer







Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me
End Sub


Private Sub cmdExit_Click()
Unload Me          'unloads the program
End Sub



Private Sub mnuFileSave_Click()
Dim TheFile, tmp
On Error GoTo cancel
CommonDialog1.InitDir = App.Path
CommonDialog1.DefaultExt = ".bmp"

CommonDialog1.Filter = "Bitmaps(*.bmp)|*.bmp"

CommonDialog1.ShowSave
TheFile = CommonDialog1.FileName

SavePicture picBoard.Image, TheFile



cancel:
End Sub


Private Sub Form_Load()
HScroll1.Value = 2
picBoard.DrawWidth = 2
picBoard.MouseIcon = curpencil
pCol.MouseIcon = target
drawtool = True
lstTools.AddItem ("Pen")
lstTools.AddItem ("Circle")
lstTools.AddItem ("Paint Bucket")
lstTools.AddItem ("Eye Dropper")
lstTools.AddItem ("Clear")
picBoard.ScaleHeight = 255
picBoard.ScaleWidth = 255
End Sub

Private Sub HScroll1_Change()
lblPenSize.Caption = HScroll1.Value          'sets the pen size according
picBoard.DrawWidth = HScroll1.Value          'to the value of the scroll bar
End Sub

Private Sub lstTools_Click()
If lstTools.Text = "Pen" Then
drawtool = True    'tells the computer that the
filltool = False   'pen has been selected
circletool = False
eyedroptool = False

End If

If lstTools.Text = "Circle" Then

On Error Resume Next
drawtool = False
filltool = False
circletool = True
eyedroptool = False
GetRadius:
whatradius = InputBox("Enter the radius for the circle in pixels:", "Paint")
If IsNumeric(whatradius) Or radius = "" Then
radius = Val(whatradius)
Else
onlynumbers = MsgBox("You have to enter a number!", vbCritical, "Paint")
GoTo GetRadius
End If

End If
If lstTools.Text = "Paint Bucket" Then
filltool = True                 'tells the computer the
drawtool = False                'fill tool has been selected
circletool = False
eyedroptool = False

End If
If lstTools.Text = "Clear" Then
picBoard.BackColor = &HFFFFFF
End If
If lstTools.Text = "Eye Dropper" Then
eyedroptool = True
drawtool = False                'tells the computer that the
filltool = False                'eye dropper has been selected
circletool = False

picBoard.MouseIcon = target
End If
End Sub


Private Sub mnuAbout_Click()
about = MsgBox("Program by Smita Yadav", vbInformation, "About")
End Sub


Private Sub mnuExit_Click()

Unload Me
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
If filltool = True Then                    'if the fill tool is selected it fills picBoard with a custom color
picBoard.BackColor = Shape1.FillColor
End If
If drawtool = True Then                    'draws a point where the user clicks on picBoard
picBoard.Line (x, y)-(x, y)
End If
If circletool = True Then
picBoard.Circle (point1.x, point1.y), radius
End If
If eyedroptool = True Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = picBoard.Point(x, y)
End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If pressed And drawtool Then
point2 = point1
point1.x = x
point1.y = y
picBoard.Line (point1.x, point1.y)-(point2.x, point2.y)         'if the mouse is dragged...the line is continued
End If
If pressed And eyedroptool Then
On Error Resume Next
Shape1.FillColor = picBoard.Point(x, y)        'set the colors for shape one and the pen
picBoard.ForeColor = picBoard.Point(x, y)
End If
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
pressed = False                    'stops drawing the line
End Sub

Private Sub tmrCursor_Timer()
If drawtool = True Then              'this function sets the cursor for
picBoard.MouseIcon = curpencil       'picBoard according to what tool is
End If                               'selected
If filltool = True Then
picBoard.MouseIcon = bucket
End If
If circletool = True Then
picBoard.MouseIcon = target
End If
End Sub


