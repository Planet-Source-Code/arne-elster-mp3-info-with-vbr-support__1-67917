VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MP3Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' MP3 info class by rm
'
' ID3v2 size calculation by Mathias Kunter (www.magic-tagger.com)

Private Declare Function CreateFile Lib "kernel32.dll" _
Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
    lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long _
) As Long

Private Declare Function ReadFile Lib "Kernel32" ( _
    ByVal hFile As Long, lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function SetFilePointer Lib "Kernel32" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long _
) As Long

Private Declare Function GetFileSize Lib "Kernel32" ( _
    ByVal hFile As Long, ByVal lpFileSizeHigh As Long _
) As Long

Private Declare Function CloseHandle Lib "Kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Enum FILE_OPEN_METHOD
    CREATE_NEW = 1
    CREATE_ALWAYS = 2
    OPEN_EXISTING = 3
    OPEN_ALWAYS = 4
End Enum

Private Enum FILE_SHARE_RIGHTS
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum

Private Enum FILE_ACCESS_RIGHTS
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
End Enum

Private Enum SEEK_METHOD
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2
End Enum

Private Const INVALID_HANDLE    As Long = -1

Private Declare Sub CpyMem Lib "Kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cb As Long _
)

Private Declare Sub ZeroMem Lib "Kernel32" Alias "RtlZeroMemory" ( _
    pDst As Any, ByVal cb As Long _
)

Private Const SMALL_BUFFER      As Long = 2048

Private Const CHAR_I            As Long = 73
Private Const CHAR_D            As Long = 68
Private Const CHAR_3            As Long = 51

Private Enum MpegID
    MPEG_V25 = 0
    MPEG_RSVD = 1
    MPEG_V2 = 2
    MPEG_V1 = 3
End Enum

Private Enum MpegLayer
    MPEG_LAYER_RSVD = 0
    MPEG_LAYER_III = 1
    MPEG_LAYER_II = 2
    MPEG_LAYER_I = 3
End Enum

Private Enum MpegChannelMode
    MPEG_CH_STEREO = 0
    MPEG_CH_JOINT_STEREO = 1
    MPEG_CH_DUAL_CHANNEL = 2
    MPEG_CH_SINGLE_CHANNEL = 3
End Enum

Private Enum MpegEmphasis
    MPEG_EMPHASIS_NONE = 0
    MPEG_EMPHASIS_5015 = 1
    MPEG_EMPHASIS_RSVD = 2
    MPEG_EMPHASIS_CCIT = 3
End Enum

Private Const MPEG_SAMP_RSVD    As Long = 3
Private Const MPEG_BITR_RSVD    As Long = 15

Private Type ID3v2TagHeader
    Identifier(2)           As Byte
    Version(1)              As Byte
    Flags                   As Byte
    size(3)                 As Byte
End Type

Private Enum ID3v2Flags
    ID3_Unsync = (2 ^ 7)
    ID3_ExtHdr = (2 ^ 6)
    ID3_Experm = (2 ^ 5)
    ID3_Footer = (2 ^ 4)
End Enum

Private Enum VBRMode
    VBRNone
    VBRXing
    VBRVBRI
End Enum

Private Enum XingFlags
    XING_Frames = &H1
    XING_Bytes = &H2
    XING_TOC = &H4
    XING_Scale = &H8
End Enum

Private Type XingHeader
    Flags                   As XingFlags
    Bytes                   As Long
    Frames                  As Long
    vbrscale                As Long
    Toc(99)                 As Byte
End Type

Private Type VBRIHeader
    VbriVersion             As Long
    VbriDelay               As Single
    VbriQuality             As Long
    VbriStreamBytes         As Long
    VbriStreamFrames        As Long
    VbriTableSize           As Long
    VbriTableScale          As Long
    VbriEntryBytes          As Long
    VbriEntryFrames         As Long
    VbriTable()             As Long
End Type

Private Type FrameHeader
    Raw(3)                  As Byte
    Channels                As MpegChannelMode
    Layer                   As MpegLayer
    MpegVersion             As MpegID
    SamplerateIndex         As Long
    Bitrate                 As Long
    Copyrighted             As Boolean
    Original                As Boolean
    Private                 As Boolean
    Protected               As Boolean
    Padding                 As Boolean
    Emphasis                As MpegEmphasis
    ModeExtension           As Long
End Type

Private Type FormatedMPEGInfo
    Channels                As Long
    Samplerate              As Long
    Bitrate                 As Long
    Bytes                   As Long
    Frames                  As Long
    Layer                   As Long
    MpegVersion             As Single
    Copyrighted             As Boolean
    Original                As Boolean
    Private                 As Boolean
    Protected               As Boolean
    Padding                 As Boolean
    Emphasis                As Boolean
    Duration                As Double
    GotXing                 As Boolean
    GotVBRI                 As Boolean
    InfoHeaderValid         As Boolean
    VBRMode                 As VBRMode
    XingHdr                 As XingHeader
    InfoHdr                 As XingHeader
    VBRIHdr                 As VBRIHeader
End Type

Private m_hFile             As Long
Private m_lngFileSize       As Long
Private m_lngFirstFrame     As Long
Private m_lngFirstFrameSize As Long
Private m_udtInfo           As FormatedMPEGInfo

Private m_lngID3v2State     As Long
Private m_lngFHState        As Long

Public Property Get FirstFrame() As Long
    FirstFrame = m_lngFirstFrame
End Property

Public Property Get Duration() As Double
    Duration = m_udtInfo.Duration
End Property

Public Property Get Channels() As Long
    Channels = m_udtInfo.Channels
End Property

Public Property Get Samplerate() As Long
    Samplerate = m_udtInfo.Samplerate
End Property

Public Property Get Bitrate() As Long
    Bitrate = m_udtInfo.Bitrate
End Property

Public Property Get FileSize() As Long
    FileSize = m_lngFileSize
End Property

Public Property Get MpegBytes() As Long
    MpegBytes = m_udtInfo.Bytes
End Property

Public Property Get Frames() As Long
    Frames = m_udtInfo.Frames
End Property

Public Property Get Layer() As Long
    Layer = m_udtInfo.Layer
End Property

Public Property Get MpegVersion() As Single
    MpegVersion = m_udtInfo.MpegVersion
End Property

Public Property Get CopyrightBit() As Boolean
    CopyrightBit = m_udtInfo.Copyrighted
End Property

Public Property Get OriginalBit() As Boolean
    OriginalBit = m_udtInfo.Original
End Property

Public Property Get PrivateBit() As Boolean
    PrivateBit = m_udtInfo.Private
End Property

Public Property Get Protected() As Boolean
    Protected = m_udtInfo.Protected
End Property

Public Property Get Emphasis() As Boolean
    Emphasis = m_udtInfo.Emphasis
End Property

Public Property Get IsVBR() As Boolean
    IsVBR = m_udtInfo.VBRMode <> VBRNone
End Property

Public Function ReadMP3Info(ByVal strPath As String) As Boolean
    Dim lngFirstFrame       As Long
    Dim lngFirstFrameSize   As Long
    
    ZeroMem m_udtInfo, Len(m_udtInfo)
    m_lngFirstFrame = 0
    m_lngFirstFrameSize = 0
    m_lngID3v2State = 0
    m_lngFHState = 0
    m_lngFileSize = 0
    
    m_hFile = CreateFile(strPath, GENERIC_READ, _
                         FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                         ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If m_hFile <> INVALID_HANDLE Then
        m_lngFileSize = SetFilePointer(m_hFile, 0, 0, FILE_END)
        
        If FindFirstFrame(lngFirstFrame, lngFirstFrameSize) Then
            m_lngFirstFrame = lngFirstFrame
            m_lngFirstFrameSize = lngFirstFrameSize
            
            m_udtInfo = ReadMpegInfo(m_lngFirstFrame, m_lngFirstFrameSize, m_lngFileSize)
            
            ReadMP3Info = True
        End If
    End If
    
    CloseHandle m_hFile
    m_hFile = 0
End Function

Private Function ReadMpegInfo(ByVal lngOffset As Long, ByVal lngSize As Long, ByVal lngFileSize As Long) As FormatedMPEGInfo
    Dim btFrame()       As Byte
    Dim lngOldPos       As Long
    Dim lngSideInfoSize As Long
    Dim lngID3v1Size    As Long
    Dim lngRead         As Long
    Dim udtFrame        As FrameHeader
    Dim udtInfo         As FormatedMPEGInfo
    
    udtFrame = ReadFrameHeader(lngOffset)
    lngSideInfoSize = MpegSideInfoSize(udtFrame)
    
    ' Frame Header information
    With udtInfo
        .Bitrate = MpegBitrate(udtFrame)
        .Channels = IIf(udtFrame.Channels = MPEG_CH_SINGLE_CHANNEL, 1, 2)
        .Copyrighted = udtFrame.Copyrighted
        .Layer = MpegLayer(udtFrame)
        .Emphasis = udtFrame.Emphasis <> MPEG_EMPHASIS_NONE
        .MpegVersion = MpegVersionID(udtFrame)
        .Original = udtFrame.Original
        .Padding = udtFrame.Padding
        .Private = udtFrame.Private
        .Protected = Not udtFrame.Protected
        .Samplerate = MpegSamplerate(udtFrame)
        .VBRMode = VBRNone
        .Bytes = lngFileSize - lngOffset
        If FindID3v1(0, lngID3v1Size, lngFileSize) Then
            .Bytes = .Bytes - lngID3v1Size
        End If
    End With
    
    ReDim btFrame(lngSize - 1) As Byte
    
    ' read complete frame
    lngOldPos = SetFilePointer(m_hFile, 0, 0, FILE_CURRENT)
    SetFilePointer m_hFile, lngOffset, 0, FILE_BEGIN
    ReadFile m_hFile, btFrame(0), lngSize, lngRead, ByVal 0&
    SetFilePointer m_hFile, lngOldPos, 0, FILE_BEGIN
    
    ' VBR info
    Select Case BytesToAnsi(btFrame, lngSideInfoSize + 4, 4)
        Case "Xing":
            If ExtractXingHeader(btFrame, lngSideInfoSize + 4, udtInfo.XingHdr) Then udtInfo.VBRMode = VBRXing
        Case "Info":
            If ExtractXingHeader(btFrame, lngSideInfoSize + 4, udtInfo.InfoHdr) Then udtInfo.InfoHeaderValid = True
        Case Else:
            If BytesToAnsi(btFrame, 36, 4) = "VBRI" Then
                If ExtractVBRIHeader(btFrame, 36, udtInfo.VBRIHdr) Then udtInfo.VBRMode = VBRVBRI
            End If
    End Select
    
    Select Case udtInfo.VBRMode
        Case VBRXing:
            If (udtInfo.XingHdr.Flags And XING_Bytes) Then udtInfo.Bytes = udtInfo.XingHdr.Bytes
        Case VBRVBRI:
            If (udtInfo.VBRIHdr.VbriStreamBytes > 0) Then udtInfo.Bytes = udtInfo.VBRIHdr.VbriStreamBytes
        Case VBRNone:
            If udtInfo.InfoHeaderValid And (udtInfo.InfoHdr.Flags And XING_Bytes) Then udtInfo.Bytes = udtInfo.InfoHdr.Bytes
    End Select
    
    udtInfo.Frames = udtInfo.Bytes / MpegFrameSize(udtFrame)
    
    Select Case udtInfo.VBRMode
        Case VBRXing:
            If (udtInfo.XingHdr.Flags And XING_Frames) Then udtInfo.Frames = udtInfo.XingHdr.Frames
        Case VBRVBRI:
            If (udtInfo.VBRIHdr.VbriStreamFrames > 0) Then udtInfo.Frames = udtInfo.VBRIHdr.VbriStreamFrames
        Case VBRNone:
            If udtInfo.InfoHeaderValid And (udtInfo.InfoHdr.Flags And XING_Frames) Then udtInfo.Frames = udtInfo.InfoHdr.Frames
    End Select
    
    With udtInfo
        .Duration = .Frames * MpegSamplesPerFrame(udtFrame) / .Samplerate * 1000
        If udtInfo.VBRMode <> VBRNone Then .Bitrate = .Bytes * 8 / .Duration * 1000
    End With
    
    ReadMpegInfo = udtInfo
End Function

Private Function FindFirstFrame(lngOffset As Long, lngSize As Long) As Boolean
    Dim btBuffer()          As Byte
    Dim lngRead             As Long
    Dim lngID3v2Size        As Long
    Dim lngBufferStart      As Long
    Dim i                   As Long
    Dim udtFrameHeader      As FrameHeader
    Dim udtNextFrameHeader  As FrameHeader
    Dim blnGotHeader        As Boolean

    ReDim btBuffer(SMALL_BUFFER - 1) As Byte
    lngOffset = -1
    lngSize = -1

    SetFilePointer m_hFile, 0, 0, FILE_BEGIN

    Do
        ReadFile m_hFile, btBuffer(0), SMALL_BUFFER, lngRead, ByVal 0&
        
        If lngRead > 0 Then
            For i = 0 To SMALL_BUFFER - 1
                If IsID3v2(btBuffer(i)) Then
                    ' skip ID3v2 if existent, it could hold some bytes looking like a MP3 frame
                    lngID3v2Size = GetID3v2Len(lngBufferStart + i - 2)
                    lngBufferStart = lngBufferStart + lngID3v2Size
                    SetFilePointer m_hFile, lngBufferStart, 0, FILE_BEGIN
                    lngBufferStart = lngBufferStart - SMALL_BUFFER
                    Exit For
                ElseIf IsFrameHeader(btBuffer(i)) Then
                    ' seems to be a frame header, read and validate
                    udtFrameHeader = ReadFrameHeader(lngBufferStart + i - 1)
                    If IsValidFrameHeader(udtFrameHeader) Then
                        ' probably is one, to make sure validate the next frame, too
                        udtNextFrameHeader = ReadFrameHeader(lngBufferStart + i + MpegFrameSize(udtFrameHeader) - 1)
                        If IsValidFrameHeader(udtNextFrameHeader) Then
                            ' got the frame header
                            lngOffset = lngBufferStart + i - 1
                            lngSize = MpegFrameSize(udtFrameHeader)
                            blnGotHeader = True
                            Exit For
                        End If
                    End If
                End If
            Next
            
            If Not blnGotHeader Then
                If lngBufferStart - lngID3v2Size > 10& * 1024& Then
                    Exit Do
                End If
            End If
        End If
        
        lngBufferStart = lngBufferStart + SMALL_BUFFER
    Loop While lngRead > 0 And Not blnGotHeader
    
    FindFirstFrame = blnGotHeader
End Function

Private Function GetID3v2Len(ByVal lngOffset As Long) As Long
    Dim udtHdr      As ID3v2TagHeader
    Dim lngOldPos   As Long
    Dim lngRead     As Long
    Dim lngSize     As Long
    
    lngOldPos = SetFilePointer(m_hFile, 0, 0, FILE_CURRENT)
    SetFilePointer m_hFile, lngOffset, 0, FILE_BEGIN
    
    ReadFile m_hFile, udtHdr, Len(udtHdr), lngRead, ByVal 0&
    
    SetFilePointer m_hFile, lngOldPos, 0, FILE_BEGIN
    
    'The size stored in the header excludes itself, and excludes the footer (if present).
    lngSize = Data2Long(VarPtr(udtHdr.size(0)), True) + Len(udtHdr)

    'v 2.4 (or later?) flags: %abcd0000 abc = ignored, d = footer present
    If udtHdr.Version(0) >= 4 Then
        If udtHdr.Flags And ID3_Footer Then
            'Add the size of the footer (which is the same size than the header) to the existing size.
            lngSize = lngSize + Len(udtHdr)
        End If
    End If
    
    GetID3v2Len = lngSize
End Function

Private Function ExtractVBRIHeader( _
    ByRef Frame() As Byte, _
    ByVal Position As Long, _
    ByRef header As VBRIHeader _
) As Boolean

    Dim TableLength As Long
    Dim i           As Long

    Position = Position + 4
    
    header.VbriVersion = VBRIExtract(Frame, Position, 2)
    header.VbriDelay = VBRIExtract(Frame, Position, 2)
    header.VbriQuality = VBRIExtract(Frame, Position, 2)
    header.VbriStreamBytes = VBRIExtract(Frame, Position, 4)
    header.VbriStreamFrames = VBRIExtract(Frame, Position, 4)
    header.VbriTableSize = VBRIExtract(Frame, Position, 2)
    header.VbriTableScale = VBRIExtract(Frame, Position, 2)
    header.VbriEntryBytes = VBRIExtract(Frame, Position, 2)
    header.VbriEntryFrames = VBRIExtract(Frame, Position, 2)
    
    ReDim header.VbriTable(header.VbriTableSize + 1) As Long
    
    For i = 0 To header.VbriTableSize
        header.VbriTable(i) = VBRIExtract(Frame, Position, header.VbriEntryBytes) * header.VbriTableScale
    Next
    
    ExtractVBRIHeader = True
End Function

Private Function ExtractXingHeader( _
    ByRef Frame() As Byte, _
    ByVal Position As Long, _
    ByRef header As XingHeader _
) As Boolean

    Dim foff  As Long

    header.Flags = XingExtractI4(Frame, Position + 4)
    
    If (header.Flags And XING_Frames) Then
        header.Frames = XingExtractI4(Frame, Position + 8 + foff)
        foff = foff + 4
    End If
    
    If (header.Flags And XING_Bytes) Then
        header.Bytes = XingExtractI4(Frame, Position + 8 + foff)
        foff = foff + 4
    End If
    
    If (header.Flags And XING_TOC) Then
        CpyMem header.Toc(0), Frame(Position + 8 + foff), 100
        foff = foff + 100
    End If
    
    If (header.Flags And XING_Scale) Then
        header.vbrscale = XingExtractI4(Frame, Position + 8 + foff)
        foff = foff + 4
    End If
    
    ExtractXingHeader = True
End Function

Private Function VBRIExtract(buf() As Byte, ByRef Position As Long, ByVal Bytes As Long) As Long
    Dim i As Long, b As Long, number As Long
    
    For i = 0 To Bytes - 1
        b = Bytes - 1 - i
        number = number Or ((buf(Position + i) And &HFF) * (2 ^ (8 * b)))
    Next
    
    Position = Position + Bytes
    VBRIExtract = number
End Function

Private Function XingExtractI2(buf() As Byte, ByVal Position As Long) As Integer
    Dim x       As Integer
    Dim Bt(1)   As Byte
    
    Bt(0) = buf(Position + 1)
    Bt(1) = buf(Position + 0)
    
    CpyMem x, Bt(0), 2
    XingExtractI2 = x
End Function

Private Function XingExtractI4(buf() As Byte, ByVal Position As Long) As Long
    Dim x       As Long
    Dim Bt(3)   As Byte

    Bt(0) = buf(Position + 3)
    Bt(1) = buf(Position + 2)
    Bt(2) = buf(Position + 1)
    Bt(3) = buf(Position + 0)
    
    CpyMem x, Bt(0), 4
    XingExtractI4 = x
End Function

Private Function IsValidFrameHeader(udtHdr As FrameHeader) As Boolean
    Dim blnRet  As Boolean
    
    blnRet = True
    
    If udtHdr.Raw(0) <> &HFF Then blnRet = False
    If (udtHdr.Raw(1) And &HE0) <> &HE0 Then blnRet = False
    If udtHdr.MpegVersion = MPEG_RSVD Then blnRet = False
    If udtHdr.Layer = MPEG_LAYER_RSVD Then blnRet = False
    If udtHdr.Bitrate = MPEG_BITR_RSVD Then blnRet = False
    If udtHdr.Bitrate = 0 Then blnRet = False
    If udtHdr.SamplerateIndex = MPEG_SAMP_RSVD Then blnRet = False
    
    If blnRet Then
        If udtHdr.MpegVersion = MPEG_V1 Then
            If udtHdr.Layer = MPEG_LAYER_II Then
                Select Case MpegBitrate(udtHdr) \ 1000
                    Case 32, 48, 56, 80:
                        blnRet = udtHdr.Channels = MPEG_CH_SINGLE_CHANNEL
                    Case 224, 256, 320, 384:
                        blnRet = udtHdr.Channels <> MPEG_CH_SINGLE_CHANNEL
                End Select
            End If
        End If
    End If
    
    IsValidFrameHeader = blnRet
End Function

Private Function ReadFrameHeader(ByVal lngOffset As Long) As FrameHeader
    Dim lngOldPos   As Long
    Dim lngRead     As Long
    Dim btHdr(3)    As Byte
    Dim clsBits     As Bitstream
    
    lngOldPos = SetFilePointer(m_hFile, 0, 0, FILE_CURRENT)
    SetFilePointer m_hFile, lngOffset, 0, FILE_BEGIN
    
    ReadFile m_hFile, btHdr(0), 4, lngRead, ByVal 0&
    Set clsBits = New Bitstream
    clsBits.SetData btHdr, 0, 4
    
    SetFilePointer m_hFile, lngOldPos, 0, FILE_BEGIN
    
    With ReadFrameHeader
        clsBits.SkipBits 11
        .Raw(0) = btHdr(0): .Raw(1) = btHdr(1)
        .Raw(2) = btHdr(2): .Raw(3) = btHdr(3)
        .MpegVersion = clsBits.ReadBits(2)
        .Layer = clsBits.ReadBits(2)
        .Protected = clsBits.ReadBits(1)
        .Bitrate = clsBits.ReadBits(4)
        .SamplerateIndex = clsBits.ReadBits(2)
        .Padding = clsBits.ReadBits(1)
        .Private = clsBits.ReadBits(1)
        .Channels = clsBits.ReadBits(2)
        .ModeExtension = clsBits.ReadBits(2)
        .Copyrighted = clsBits.ReadBits(1)
        .Original = clsBits.ReadBits(1)
        .Emphasis = clsBits.ReadBits(2)
    End With
End Function

Private Function IsFrameHeader(ByVal btByte As Byte) As Boolean
    Select Case m_lngFHState
        Case 0:
            If btByte = &HFF Then
                m_lngFHState = 1
            Else
                m_lngFHState = 0
            End If
        Case 1:
            If btByte And &HE0 = &HE0 Then
                m_lngFHState = 2
            Else
                m_lngFHState = 0
            End If
        Case 2
            If btByte = &HFF Then
                m_lngFHState = 1
            Else
                m_lngFHState = 0
            End If
    End Select
    
    IsFrameHeader = m_lngFHState = 2
End Function

Private Function IsID3v2(ByVal btByte As Byte) As Boolean
    Select Case m_lngID3v2State
        Case 0:
            Select Case btByte
                Case CHAR_I: m_lngID3v2State = 1
                Case Else:   m_lngID3v2State = 0
            End Select
        Case 1:
            Select Case btByte
                Case CHAR_D: m_lngID3v2State = 2
                Case Else:   m_lngID3v2State = 0
            End Select
        Case 2:
            Select Case btByte
                Case CHAR_3: m_lngID3v2State = 3
                Case Else:   m_lngID3v2State = 0
            End Select
        Case 3:
            m_lngID3v2State = 0
    End Select
    
    IsID3v2 = m_lngID3v2State = 3
End Function

Private Property Get MpegSideInfoSize(udtFrameHeader As FrameHeader) As Long
    Select Case udtFrameHeader.MpegVersion
        Case MPEG_V1:
            Select Case udtFrameHeader.Channels
                Case MPEG_CH_SINGLE_CHANNEL:    MpegSideInfoSize = 17
                Case Else:                      MpegSideInfoSize = 32
            End Select
        Case MPEG_V2, MPEG_V25:
            Select Case udtFrameHeader.Channels
                Case MPEG_CH_SINGLE_CHANNEL:    MpegSideInfoSize = 9
                Case Else:                      MpegSideInfoSize = 17
            End Select
    End Select
End Property

Private Property Get MpegFrameSize(udtFrameHeader As FrameHeader) As Long
    Dim lngBitrate      As Long
    Dim lngSamplerate   As Long
    
    lngBitrate = MpegBitrate(udtFrameHeader)
    lngSamplerate = MpegSamplerate(udtFrameHeader)
    
    Select Case udtFrameHeader.MpegVersion
        Case MPEG_V1:
            Select Case udtFrameHeader.Layer
                Case MPEG_LAYER_I:
                    MpegFrameSize = Fix((12 * lngBitrate / lngSamplerate) + Abs(udtFrameHeader.Padding)) * 4
                Case MPEG_LAYER_II, MPEG_LAYER_III:
                    MpegFrameSize = Fix((144 * lngBitrate / lngSamplerate) + Abs(udtFrameHeader.Padding))
            End Select
        Case MPEG_V2, MPEG_V25:
            Select Case udtFrameHeader.Layer
                Case MPEG_LAYER_I:
                    MpegFrameSize = Fix((12 * lngBitrate / lngSamplerate) + Abs(udtFrameHeader.Padding)) * 4
                Case MPEG_LAYER_II:
                    MpegFrameSize = Fix((144 * lngBitrate / lngSamplerate) + Abs(udtFrameHeader.Padding))
                Case MPEG_LAYER_III:
                    MpegFrameSize = Fix((72 * lngBitrate / lngSamplerate) + Abs(udtFrameHeader.Padding))
            End Select
    End Select
End Property

Private Property Get MpegSamplesPerFrame(udtFrameHeader As FrameHeader) As Long
    Select Case udtFrameHeader.MpegVersion
        Case MPEG_V1:
            Select Case udtFrameHeader.Layer
                Case MPEG_LAYER_I:      MpegSamplesPerFrame = 384
                Case MPEG_LAYER_II:     MpegSamplesPerFrame = 1152
                Case MPEG_LAYER_III:    MpegSamplesPerFrame = 1152
            End Select
        Case MPEG_V2, MPEG_V25:
            Select Case udtFrameHeader.Layer
                Case MPEG_LAYER_I:      MpegSamplesPerFrame = 384
                Case MPEG_LAYER_II:     MpegSamplesPerFrame = 1152
                Case MPEG_LAYER_III:    MpegSamplesPerFrame = 576
            End Select
    End Select
End Property

Private Property Get MpegBitrate(udtFrameHeader As FrameHeader) As Long
    Select Case udtFrameHeader.MpegVersion
        Case MPEG_V1:
            Select Case udtFrameHeader.Layer
                Case MPEG_LAYER_I:
                    Select Case udtFrameHeader.Bitrate
                        Case 1: MpegBitrate = 32000
                        Case 2: MpegBitrate = 64000
                        Case 3: MpegBitrate = 96000
                        Case 4: MpegBitrate = 128000
                        Case 5: MpegBitrate = 160000
                        Case 6: MpegBitrate = 192000
                        Case 7: MpegBitrate = 224000
                        Case 8: MpegBitrate = 256000
                        Case 9: MpegBitrate = 288000
                        Case 10: MpegBitrate = 320000
                        Case 11: MpegBitrate = 352000
                        Case 12: MpegBitrate = 384000
                        Case 13: MpegBitrate = 416000
                        Case 14: MpegBitrate = 448000
                    End Select
                    
                Case MPEG_LAYER_II:
                    Select Case udtFrameHeader.Bitrate
                        Case 1: MpegBitrate = 32000
                        Case 2: MpegBitrate = 48000
                        Case 3: MpegBitrate = 56000
                        Case 4: MpegBitrate = 64000
                        Case 5: MpegBitrate = 80000
                        Case 6: MpegBitrate = 96000
                        Case 7: MpegBitrate = 112000
                        Case 8: MpegBitrate = 128000
                        Case 9: MpegBitrate = 160000
                        Case 10: MpegBitrate = 192000
                        Case 11: MpegBitrate = 224000
                        Case 12: MpegBitrate = 256000
                        Case 13: MpegBitrate = 320000
                        Case 14: MpegBitrate = 384000
                    End Select
                    
                Case MPEG_LAYER_III:
                    Select Case udtFrameHeader.Bitrate
                        Case 1: MpegBitrate = 32000
                        Case 2: MpegBitrate = 40000
                        Case 3: MpegBitrate = 48000
                        Case 4: MpegBitrate = 56000
                        Case 5: MpegBitrate = 64000
                        Case 6: MpegBitrate = 80000
                        Case 7: MpegBitrate = 96000
                        Case 8: MpegBitrate = 112000
                        Case 9: MpegBitrate = 128000
                        Case 10: MpegBitrate = 160000
                        Case 11: MpegBitrate = 192000
                        Case 12: MpegBitrate = 224000
                        Case 13: MpegBitrate = 256000
                        Case 14: MpegBitrate = 320000
                    End Select
            End Select
            
        Case MPEG_V2, MPEG_V25:
            Select Case udtFrameHeader.Layer
                Case MPEG_LAYER_I:
                    Select Case udtFrameHeader.Bitrate
                        Case 1: MpegBitrate = 32000
                        Case 2: MpegBitrate = 48000
                        Case 3: MpegBitrate = 56000
                        Case 4: MpegBitrate = 64000
                        Case 5: MpegBitrate = 80000
                        Case 6: MpegBitrate = 96000
                        Case 7: MpegBitrate = 112000
                        Case 8: MpegBitrate = 128000
                        Case 9: MpegBitrate = 144000
                        Case 10: MpegBitrate = 160000
                        Case 11: MpegBitrate = 176000
                        Case 12: MpegBitrate = 192000
                        Case 13: MpegBitrate = 224000
                        Case 14: MpegBitrate = 256000
                    End Select
                    
                Case MPEG_LAYER_II, MPEG_LAYER_III:
                    Select Case udtFrameHeader.Bitrate
                        Case 1: MpegBitrate = 8000
                        Case 2: MpegBitrate = 16000
                        Case 3: MpegBitrate = 24000
                        Case 4: MpegBitrate = 32000
                        Case 5: MpegBitrate = 40000
                        Case 6: MpegBitrate = 48000
                        Case 7: MpegBitrate = 56000
                        Case 8: MpegBitrate = 64000
                        Case 9: MpegBitrate = 80000
                        Case 10: MpegBitrate = 96000
                        Case 11: MpegBitrate = 112000
                        Case 12: MpegBitrate = 128000
                        Case 13: MpegBitrate = 144000
                        Case 14: MpegBitrate = 160000
                    End Select
                    
            End Select
    End Select
End Property

Private Property Get MpegVersionID(udtFrameHeader As FrameHeader) As Single
    Select Case udtFrameHeader.MpegVersion
        Case MPEG_V1:   MpegVersionID = 1
        Case MPEG_V2:   MpegVersionID = 2
        Case MPEG_V25:  MpegVersionID = 2.5
    End Select
End Property

Private Property Get MpegSamplerate(udtFrameHeader As FrameHeader) As Long
    Select Case udtFrameHeader.MpegVersion
        Case MPEG_V1:
            Select Case udtFrameHeader.SamplerateIndex
                Case 0: MpegSamplerate = 44100
                Case 1: MpegSamplerate = 48000
                Case 2: MpegSamplerate = 32000
            End Select
        Case MPEG_V2:
            Select Case udtFrameHeader.SamplerateIndex
                Case 0: MpegSamplerate = 22050
                Case 1: MpegSamplerate = 24000
                Case 2: MpegSamplerate = 16000
            End Select
        Case MPEG_V25:
            Select Case udtFrameHeader.SamplerateIndex
                Case 0: MpegSamplerate = 11025
                Case 1: MpegSamplerate = 12000
                Case 2: MpegSamplerate = 8000
            End Select
    End Select
End Property

Private Property Get MpegLayer(udtFrameHeader As FrameHeader) As Long
    Select Case udtFrameHeader.Layer
        Case MPEG_LAYER_I:      MpegLayer = 1
        Case MPEG_LAYER_II:     MpegLayer = 2
        Case MPEG_LAYER_III:    MpegLayer = 3
    End Select
End Property

Private Function Data2Long(ByVal pData As Long, ByVal bSynchSafe As Boolean) As Long
    Dim i As Integer, Data(3) As Byte

    CpyMem Data(0), ByVal pData, 4

    'Avoid converting wrong synchsafe integers. If bit 7 of any byte is set, it is not synchsafe.
    'However, we can't detect wrong coded values which have bit 7 zeroed.
    For i = 0 To 3
        If Data(i) And &H80& Then bSynchSafe = False
    Next i

    'Perform left-shifts, done by multiplication with the hex values of 2^n. Finally, bit-or the values.
    If bSynchSafe Then
        Data2Long = (Data(0) * &H200000) Or (Data(1) * &H4000&) Or (Data(2) * &H80&) Or Data(3)
    Else
        Data2Long = (Data(0) * &H1000000) Or (Data(1) * &H10000) Or (Data(2) * &H100&) Or Data(3)
    End If
End Function

Private Function FindID3v1(ByRef Position As Long, ByRef size As Long, ByVal lngFileLen As Long) As Boolean
    Dim lngOldFPos      As Long
    Dim lngRead         As Long
    Dim btBuffer(127)   As Byte
    
    lngOldFPos = SetFilePointer(m_hFile, 0, 0, FILE_CURRENT)
    
    SetFilePointer m_hFile, lngFileLen - 128, 0, FILE_BEGIN
    ReadFile m_hFile, btBuffer(0), UBound(btBuffer) + 1, lngRead, ByVal 0&
    
    If BytesToAnsi(btBuffer, 0, 3) = "TAG" Then
        Position = lngFileLen - 128
        size = 128
        FindID3v1 = True
    End If
    
    SetFilePointer m_hFile, lngOldFPos, 0, FILE_BEGIN
End Function

Private Function BytesToAnsi(buf() As Byte, offset As Long, Length As Long) As String
    Dim strOut  As String
    Dim i       As Long
    
    For i = offset To offset + Length - 1
        strOut = strOut & Chr(buf(i))
    Next
    
    BytesToAnsi = strOut
End Function
