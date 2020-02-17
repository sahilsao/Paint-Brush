VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImageinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BUFFERSIZE As Long = 65535

Public Enum eImageType
    itUNKNOWN = 0
    itGIF = 1
    itJPEG = 2
    itPNG = 3
    itBMP = 4
End Enum

Private m_Width As Long
Private m_Height As Long
Private m_Depth As Byte
Private m_ImageType As eImageType


Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get Depth() As Byte
    Depth = m_Depth
End Property

Public Property Get ImageType() As eImageType
    ImageType = m_ImageType
End Property

Public Sub ReadImageInfo(sFileName As String)
    Dim bBuf(BUFFERSIZE) As Byte
    Dim iFN As Integer
    
    m_Width = 0
    m_Height = 0
    m_Depth = 0
    m_ImageType = itUNKNOWN
    
    iFN = FreeFile
    Open sFileName For Binary As iFN
    Get #iFN, 1, bBuf()
    Close iFN
    
    If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
    
        m_ImageType = itPNG
        
        Select Case bBuf(25)
            Case 0
                m_Depth = bBuf(24)
                
            Case 2
                m_Depth = bBuf(24) * 3
                
            Case 3
                m_Depth = 8
                
            Case 4
                m_Depth = bBuf(24) * 2
                
            Case 6
                m_Depth = bBuf(24) * 4
                
            Case Else
                m_ImageType = itUNKNOWN
                
        End Select
        
        If m_ImageType Then
            m_Width = Mult(bBuf(19), bBuf(18))
            
            m_Height = Mult(bBuf(23), bBuf(22))
        End If
        
    End If
    
    If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
        
        m_ImageType = itGIF
        m_Width = Mult(bBuf(6), bBuf(7))
        m_Height = Mult(bBuf(8), bBuf(9))
        
        m_Depth = (bBuf(10) And 7) + 1
    End If
    
    If bBuf(0) = 66 And bBuf(1) = 77 Then
'bmp file
        m_ImageType = itBMP
        
        m_Width = Mult(bBuf(18), bBuf(19))
        
        m_Height = Mult(bBuf(22), bBuf(23))
        
        m_Depth = bBuf(28)
    End If

    If m_ImageType = itUNKNOWN Then
    ' if it is a JPEG file
        Dim lPos As Long
        
        Do
            If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 And bBuf(lPos + 2) = &HFF) _
            Or (lPos >= BUFFERSIZE - 10) Then Exit Do
            
            lPos = lPos + 1
            
        Loop
        
        lPos = lPos + 2
        If lPos >= BUFFERSIZE - 10 Then Exit Sub
        
        
        Do
        
            Do
                If bBuf(lPos) = &HFF And bBuf(lPos + 1) <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= BUFFERSIZE - 10 Then Exit Sub
            Loop
            
            lPos = lPos + 1
            
            Select Case bBuf(lPos)
                Case &HC0 To &HC3, &HC5 To &HC7, &HC9 To &HCB, &HCD To &HCF
                    Exit Do
            End Select
            
            lPos = lPos + Mult(bBuf(lPos + 2), bBuf(lPos + 1))
            
            If lPos >= BUFFERSIZE - 10 Then Exit Sub
            
        Loop
        
        
        m_ImageType = itJPEG
        
        m_Height = Mult(bBuf(lPos + 5), bBuf(lPos + 4))
        
        m_Width = Mult(bBuf(lPos + 7), bBuf(lPos + 6))
        
        m_Depth = bBuf(lPos + 8) * 8
        
    End If
    
End Sub
Private Function Mult(lsb As Byte, msb As Byte) As Long
    Mult = lsb + (msb * CLng(256))
End Function
