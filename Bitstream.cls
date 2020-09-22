VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bitstream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' partially translated from libmad (GPL)

Private Const CHAR_BIT  As Long = 8

Private m_btBuffer()    As Byte

Private m_lngBufferSize As Long

Private m_lngBytePos    As Long
Private m_lngBitPos     As Long

Private m_blnGotBuffer  As Boolean

Private m_lngPower2(31) As Long

Public Sub SetData(btData() As Byte, ByVal lngOffset As Long, ByVal lngLen As Long)
    Dim i   As Long
    
    ReDim m_btBuffer(lngLen - 1) As Byte

    For i = lngOffset To lngOffset + lngLen - 1
        m_btBuffer(i - lngOffset) = btData(i)
    Next
    
    m_lngBufferSize = lngLen
    m_lngBytePos = 0
    m_lngBitPos = CHAR_BIT
    
    m_blnGotBuffer = True
End Sub

Public Sub SkipBits(ByVal lngBits As Long)
    m_lngBytePos = m_lngBytePos + (lngBits \ CHAR_BIT)
    m_lngBitPos = m_lngBitPos - (lngBits Mod CHAR_BIT)
End Sub

Public Function ReadBits(ByVal lngBits As Long) As Long
    Dim lngValue    As Long
    
    If lngBits < m_lngBitPos Then
        lngValue = shr(m_btBuffer(m_lngBytePos) And (shl(1, m_lngBitPos) - 1), m_lngBitPos - lngBits)
        m_lngBitPos = m_lngBitPos - lngBits
    Else
        lngValue = m_btBuffer(m_lngBytePos) And (shl(1, m_lngBitPos) - 1)
        lngBits = lngBits - m_lngBitPos
        m_lngBitPos = CHAR_BIT
        m_lngBytePos = m_lngBytePos + 1
        
        Do While lngBits >= CHAR_BIT
            lngValue = shl(lngValue, CHAR_BIT) Or m_btBuffer(m_lngBytePos)
            m_lngBytePos = m_lngBytePos + 1
            lngBits = lngBits - CHAR_BIT
        Loop
        
        If lngBits > 0 Then
            lngValue = shr(lngValue, lngBits) Or (shr(m_btBuffer(m_lngBytePos), CHAR_BIT - lngBits))
            m_lngBitPos = m_lngBitPos - lngBits
        End If
    End If
    
    ReadBits = lngValue
End Function

Private Function shl(ByVal lngValue As Long, ByVal bits As Byte) As Long
    shl = lngValue * m_lngPower2(bits)
End Function

Private Function shr(ByVal lngValue As Long, ByVal bits As Byte) As Long
    shr = lngValue \ m_lngPower2(bits)
End Function

Private Sub Class_Initialize()
    Dim i   As Long
    
    For i = 0 To 30
        m_lngPower2(i) = 2 ^ i
    Next
    m_lngPower2(31) = &H80000000
End Sub
