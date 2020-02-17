VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4875
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4800
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3364.812
   ScaleMode       =   0  'User
   ScaleWidth      =   4507.448
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Label Label3 
      Caption         =   "Charotar Institute of Technology"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Comments? Suggestions? or if you want this programs source code then e-mail me:     agrawaldivya1@yahoo.com"
      ForeColor       =   &H00008080&
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "VIth C.E."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5224.884
      Y1              =   2484.784
      Y2              =   2484.784
   End
   Begin VB.Label lblDescription 
      Caption         =   "Divya  Agrawal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   2085
   End
   Begin VB.Label lblTitle 
      Caption         =   "Advanced Paint Brush"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   480
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5210.798
      Y1              =   2484.784
      Y2              =   2484.784
   End
   Begin VB.Label lblVersion 
      Caption         =   "Developed by:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   1725
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
Unload frmAbout
frmpaint.Show
End Sub
