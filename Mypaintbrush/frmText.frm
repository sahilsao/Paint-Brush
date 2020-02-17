VERSION 5.00
Begin VB.Form frmText 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the text in the text box:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Unload Me
frmpaint.Show
End Sub

Private Sub OKButton_Click()
Unload Me
frmpaint.Show

End Sub

Private Sub Text1_Change()
Dim txt As String
txt = Text1.Text
frmpaint.picBoard.CurrentX = frmpaint.picBoard.ScaleWidth / 6
frmpaint.picBoard.CurrentY = frmpaint.picBoard.ScaleHeight / 6
frmpaint.picBoard.Print txt

End Sub
