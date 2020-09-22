VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MP3 Info"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3150
      Width           =   2940
   End
   Begin MSComDlg.CommonDialog dlgMP3 
      Left            =   4050
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "MP3/2 (*.mp3;*.mp2)|*.mp3;*.mp2|All files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "open MP3..."
      Height          =   350
      Left            =   3150
      TabIndex        =   1
      Top             =   3150
      Width           =   1365
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   75
      Width           =   4365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsMP3Info    As MP3Info

Private Sub cmdOpen_Click()
    With dlgMP3
        .FileName = vbNullString
        .ShowOpen
    End With
    
    txtInfo.Text = ""
    
    If dlgMP3.FileName <> vbNullString Then
        txtFile.Text = dlgMP3.FileName
        If Not m_clsMP3Info.ReadMP3Info(dlgMP3.FileName) Then
            MsgBox "Not a valid MP3 file!", vbExclamation
        Else
            ShowMP3Info
        End If
    End If
End Sub

Private Sub ShowMP3Info()
    AppendLine "Size: " & m_clsMP3Info.FileSize & " bytes"
    AppendLine "Actual Audio Data: " & m_clsMP3Info.MpegBytes & " bytes"
    AppendLine "Header found at: " & m_clsMP3Info.FirstFrame & " bytes"
    AppendLine "Length: " & FormatSeconds(Fix(m_clsMP3Info.Duration / 1000)) & " min (" & Fix(m_clsMP3Info.Duration / 1000) & " seconds)"
    AppendLine "MPEG " & m_clsMP3Info.MpegVersion & " layer " & m_clsMP3Info.Layer
    AppendLine CLng(m_clsMP3Info.Bitrate / 1000) & "kbit" & IIf(m_clsMP3Info.IsVBR, " (VBR)", "") & ", " & m_clsMP3Info.Frames & " frames"
    AppendLine m_clsMP3Info.Samplerate & "Hz, " & IIf(m_clsMP3Info.Channels = 1, "1 channel", "2 channels")
    AppendLine "CRCs: " & IIf(m_clsMP3Info.Protected, "yes", "no")
    AppendLine "Copyrighted: " & IIf(m_clsMP3Info.CopyrightBit, "yes", "no")
    AppendLine "Original: " & IIf(m_clsMP3Info.OriginalBit, "yes", "no")
    AppendLine "Protected: " & IIf(m_clsMP3Info.Protected, "yes", "no")
    AppendLine "Emphasis: " & IIf(m_clsMP3Info.Emphasis, "yes", "no")
End Sub

Private Sub AppendLine(ByVal strText As String)
    txtInfo.Text = txtInfo.Text & strText & vbCrLf
End Sub

Private Function FormatSeconds(ByVal secs As Long) As String
    FormatSeconds = (secs \ 60) & ":" & Format(secs Mod 60, "00")
End Function

Private Sub Form_Load()
    Set m_clsMP3Info = New MP3Info
End Sub
