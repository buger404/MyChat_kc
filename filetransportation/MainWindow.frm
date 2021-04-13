VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyChat - 文件传输工具"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog dialog 
      Left            =   5640
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sock 
      Index           =   0
      Left            =   5640
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame ConfirmFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8E8E8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   6375
      Begin VB.Label ProF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008FDF54&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label OKBtn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008FDF54&
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label ProB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   5895
      End
   End
   Begin VB.Label Content 
      BackStyle       =   0  'Transparent
      Caption         =   "正文"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   5490
   End
   Begin VB.Label Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标题"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temple As String
Dim diCount As Long, dCount As Long, pCount As Long, file As String
Dim data() As Byte, Fsize As Long, ip As String, port As Long, Dsize As Long
Public Sub SetPro(num As Single)
    ProB.Visible = True
    ProF.Visible = True
    ProF.Width = num * ProB.Width
End Sub

Private Sub Form_Load()
    Dim t() As String, f() As String
    t = Split(Command, " ")
    '-o filepath downloadpersoncount port
    Me.Show
    If t(0) = "-o" Then
        file = Base64DecodeString(t(1))
        Title.Caption = "正在准备文件传输"
        Content.Caption = "请稍等片刻。"
        OKBtn.Visible = False
        DoEvents
        Dim b(1023) As Byte
        ReDim data(FileLen(file) - 1)
        Open file For Binary As #1
        For i = 0 To UBound(data) Step 1024
            Get #1, , b
            CopyMemory data(i), b(0), 1024
            If i Mod 102400 = 0 Then
                SetPro i / UBound(data)
                DoEvents
            End If
        Next
        Close #1
        DoEvents
        Title.Caption = "文件传输已开放"
        ProB.Visible = False: ProF.Visible = False
        f = Split(file, "\")
        pCount = Val(t(2))
        port = Val(t(3))
        temple = "文件 '" & f(UBound(f)) & "' " & vbCrLf & "正在下载：{diCount}/" & pCount & vbCrLf & "下载完毕：{dCount}/" & pCount
        Content.Caption = Replace(Replace(temple, "{diCount}", diCount), "{dCount}", dCount)
        DoEvents
        sock(0).LocalPort = port
        sock(0).Listen
        DoEvents
    End If
    '-d filename filesize user ip port
    If t(0) = "-d" Then
        file = Base64DecodeString(t(1))
        Fsize = Val(t(2))
        Title.Caption = Base64DecodeString(t(3)) & "请求向你发送文件"
        Content.Caption = "文件名：" & file & vbCrLf & "文件大小：" & Int(Fsize / 1024 / 1024 * 1000) / 1000 & "MB"
        DoEvents
        ip = t(4): port = Val(t(5))
    End If
End Sub

Private Sub OKBtn_Click()
    If OKBtn.Caption = "打开" Then
        ShellExecuteA 0, "open", "c:\windows\explorer.exe", "/select,""" & dialog.FileName & """", "", SW_SHOW
        Exit Sub
    End If
    dialog.FileName = file
    dialog.ShowSave
    If dialog.FileName = "" Then Exit Sub
    Open dialog.FileName For Binary As #1
    OKBtn.Visible = False
    sock(0).Connect ip, port
End Sub

Private Sub sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Load sock(sock.UBound + 1)
    sock(sock.UBound).Accept requestID
    sock(0).Close
    sock(0).Listen
    diCount = diCount + 1
    Content.Caption = Replace(Replace(temple, "{diCount}", diCount), "{dCount}", dCount)
    sock(sock.UBound).SendData data
    DoEvents
End Sub

Private Sub sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim b() As Byte
    sock(Index).GetData b
    Put #1, , b
    Dsize = Dsize + UBound(b) + 1
    SetPro Dsize / Fsize
    DoEvents
    If Dsize >= Fsize Then
        Close #1
        Content.Caption = "传输已完成。"
        ProB.Visible = False
        ProF.Visible = False
        OKBtn.Visible = True
        OKBtn.Caption = "打开"
        sock(Index).Close
    End If
End Sub

Private Sub sock_SendComplete(Index As Integer)
    dCount = dCount + 1
    Content.Caption = Replace(Replace(temple, "{diCount}", diCount), "{dCount}", dCount)
End Sub
