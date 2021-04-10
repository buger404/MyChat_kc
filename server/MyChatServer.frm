VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "服务端"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   Icon            =   "MyChatServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   3660
      Width           =   6300
   End
   Begin VB.CommandButton OCR 
      Caption         =   "图片转文字"
      Height          =   615
      Left            =   6480
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Audio 
      BackColor       =   &H00FFFFFF&
      Caption         =   "语音识别粘贴"
      Height          =   615
      Left            =   6480
      TabIndex        =   11
      ToolTipText     =   "就你能说话"
      Top             =   1920
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "禁言"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      ToolTipText     =   "就你能说话"
      Top             =   1440
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "保存记录"
      Height          =   375
      Left            =   6480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      ToolTipText     =   "保存在D盘"
      Top             =   960
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "发送"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "MyChatServer.frx":1BCC2
      Top             =   720
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   270
      Left            =   5565
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3615
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSWinsockLib.Winsock lis 
      Left            =   6960
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   6480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "快快把IP跟Port告诉你的小伙伴吧！"
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "号客户机的连接"
      Height          =   255
      Left            =   4230
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "断开与"
      Height          =   270
      Left            =   3015
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim State As Boolean, pop As Single
Dim grpExistId As Integer
Dim pypid
Dim g As String, q As Single, m As Single
Dim MainPage As MainPage, IPPage As IPPage

Public Sub Command1_Click()
    Dim C As Single
    C = Val(Text2.Text)
    
    If C > Winsock.ubound Then
        MsgBox ("没有此用户")
    Else
        If Winsock(C).State = 7 Then
            Winsock(C).close
            MsgBox ("已断开")
        End If
    End If

    pop = pop - 1
    Me.Caption = lis.LocalIP & " - " & "已连接" & pop & "人"
    Text2.Text = ""
End Sub
'===============================================================================================================
'Emerald框架部分
Private Sub InitEmeraldFramework()
    '启动Emerald
    StartEmerald Me.hwnd, 1100, 600, False
    'ScaleGame Screen.Width / Screen.TwipsPerPixelX / 1280, ScaleDefault
    '创建字体渲染
    Set EF = New GFont
    EF.MakeFont "微软雅黑"
    '实例化页面管理器核心
    Set ECore = New GMan
    '实例化页面控制器
    Set MainPage = New MainPage
    Set IPPage = New IPPage
    '显示
    DrawTimer.Enabled = True
    ECore.ActivePage = "IPPage"
End Sub
Private Sub UnloadEmeraldFramework()
    DrawTimer.Enabled = False
    EndEmerald
End Sub
Private Sub DrawTimer_Timer()
    '更新画面
    ECore.Display
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    If Mouse.State = 0 Then
        UpdateMouse x, y, 0, button
    Else
        Mouse.x = x: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'===============================================================================================================
Public Sub Command2_Click()

    If Text4.Text = "" Then VBA.Beep: Exit Sub
    
    Dim S As Single
    S = 1
    Do While (S <= Winsock.ubound)
        If Winsock(S).State = 7 Then
            Winsock(S).SendData "msg;" + "groupid;" + "主机" + ";id;" + Base64EncodeString(Text4.Text) + ";"
            DoEvents
        End If
        S = S + 1
    Loop
    
    Text3.Text = "我：" & Text4.Text & vbCrLf & Text3.Text
    Text4.Text = ""

End Sub

Public Sub Command3_Click()
    Text3.Text = ""
    Text4.Text = ""
End Sub

Public Sub Command4_Click()
    Open App.path & "\" & "服务端消息记录" & Str(q) & ".txt" For Output As #1
    Print #1, Text3.Text
    Close #1
    q = q + 1
End Sub

Public Sub Command5_Click()
    If State = False Then
        State = True
        Command5.Caption = "解除禁言"
        Dim S As Single
        g = "服务器开启了禁言"
        S = 1
        Do While (S <= Winsock.ubound)
            If Winsock(S).State = 7 Then
                Winsock(S).SendData g
                DoEvents
            End If
            S = S + 1
        Loop
        Text3.Text = g & vbCrLf & Text3.Text
    Else
        State = False
        Command5.Caption = "禁言"
        g = "服务器关闭了禁言"
        S = 1
        Do While (S <= Winsock.ubound)
            If Winsock(S).State = 7 Then
                Winsock(S).SendData g
                DoEvents
            End If
            S = S + 1
        Loop
        Text3.Text = g & vbCrLf & Text3.Text
    End If
End Sub

Public Sub Audio_Click()
    ShellEx "python """ & App.path & "\" & "server.py"" -y "
    If Dir("audio_text.txt") = "" Then Shell "python """ & App.path & "\" & "server.py"" -y "
    
    Dim strfile As String
    strfile = "audio_text.txt"
    Open strfile For Input As #1
        Text4.Text = StrConv(InputB(FileLen(strfile), 1), vbUnicode)
    Close #1
    Kill "audio.wav"
    Kill "audio_text.txt"
    Kill "audio.pcm"
End Sub



Public Sub OCR_Click()
    ShellEx "python """ & App.path & "\" & "server.py"" -t "
    If Dir("OCR_text.txt") = "" Then Shell "python """ & App.path & "\" & "server.py"" -t "
    
    Dim strfile As String
    strfile = "OCR_text.txt"
    Open strfile For Input As #1
        Text4.Text = StrConv(InputB(FileLen("OCR_text.txt"), 1), vbUnicode)
    Close #1
    Kill "OCR_text.txt"
End Sub

Private Sub Form_Load()
    ReDim groups(0)
    '测试用
    AddGroup 1, 1, True, "测试讨论组1"
    AddGroup 2, 1, False, "未加入测试"
    AddGroup 3, 1, True, "测试讨论组2"
    AddGroup 4, 1, True, "testtest"
    AddGroup 5, 1, True, "hash"
    AddMessage 1, 1, "测试组员", "我发送了一条消息啦啦啦啦啦啦！"
    AddMessage 1, 2, "测试组员", "我还能换行" & vbCrLf & "乌拉乌拉"
    AddMessage 1, -1, "系统消息", "您被禁言，才怪。"
    AddMessage 1, -2, "老师", "不要乱发消息"
    userId = -2
    
    Call InitEmeraldFramework
    Set Shadow = New aShadow
    With Shadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 8
            .Transparency = 20
        End If
    End With
    
    Text3.Visible = False: Text4.Visible = False
    Me.Show
    
    pypid = Shell("python """ & App.path & "\" & "server.py"" -o " & lis.LocalIP, 6)
    
<<<<<<< HEAD
    Call AddGroup(0, -2, True, "公共")
    grpExistId = 0
    
    
    Text3.Visible = True: Text4.Visible = True
=======
    Text3.Visible = False: Text4.Visible = True
>>>>>>> 3b07fd90bb91919c2b4047f89cff85163e999376
    Command5.Enabled = False
    State = False
    m = 1
    
    lis.LocalPort = 2001
    lis.Listen
    
    '精准控制坐标
    Text3.Move 300 + 0, 60, Me.ScaleWidth - 300, Me.ScaleHeight - 60 - 120
    Text4.Move 300 + 50, Me.ScaleHeight - 80 + 25, Me.ScaleWidth - 245 - 300, 80 - 50
    
    Me.Caption = lis.LocalIP & " - " & "已连接" & pop & "人"
    
    ECore.NewTransform , , "MainPage"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Server.Command4_Click
    TerminateProcess OpenProcess(PROCESS_TERMINATE, 0, pypid), 0
    
    Set Shadow = Nothing
    Call UnloadEmeraldFramework
End Sub

Private Sub lis_ConnectionRequest(ByVal requestID As Long)
    Load Winsock(m)
    Command5.Enabled = True
    
    pop = Winsock.ubound
    
    If Winsock(m).State = sckClosed Then
        Winsock(m).Accept requestID
    End If
    
    Call SetJoinState(0, True)
    
    Me.Caption = lis.LocalIP & " - " & "已连接" & pop & "人"
    
    m = m + 1
End Sub

Private Sub Text4_Change()
    '自动调整文本框大小
    Dim Line As Long
    Line = UBound(Split(Text4.Text, vbCrLf)) + 1
    If Line <= 0 Then Line = 1
    Dim Border As Integer, Height As Long
    Border = IIf(Line > 1, 1, 0)
    Height = Line * 30
    If Height > Me.ScaleHeight - 120 Then Height = Me.ScaleHeight - 120 '防止过多行溢出
    If Text4.BorderStyle <> Border Then Text4.BorderStyle = Border
    If Text4.Height <> Height Then
        Text4.Height = Height
        Text4.Top = Me.ScaleHeight - 80 + 25 - Height + 30
    End If
End Sub

Private Sub Winsock_Close(index As Integer)
    pop = pop - 1
    Me.Caption = lis.LocalIP & " - " & "已连接" & pop & "人"
End Sub

Private Sub Winsock_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If State = True Then Exit Sub
    
    Dim strSplit
    Dim id As Integer
    Dim MsgType As String
    Dim grpId As String
    Dim name As String
    Dim MsgContent As String
    Dim strData As String
    Winsock(index).GetData strData
    
    Dim S As Single
    S = 1
    Do While (S <= Winsock.ubound)
        If Winsock(S).State = 7 Then
            Winsock(S).SendData strData
            DoEvents
        End If
        S = S + 1
    Loop
    
    strSplit = Split(strData, ";")
    id = index
    MsgType = strSplit(0)
<<<<<<< HEAD

    
    Select Case MsgType
    Case "msg"
    name = strSplit(2)
    grpId = strSplit(1)
    MsgContent = strSplit(4)
    MsgContent = Base64DecodeString(MsgContent)
    Text3.Text = name + ":" + MsgContent + "   #" + Str(id) + "#" + Str(grpId) + "#" + vbCrLf + Text3.Text
    Case "picmsg"
    Case "addgroup"
    Case "okgroup"
    Case "creategroup"
    
    Dim grpCreateName As String
    
    End Select
=======
    name = strSplit(2)
    MsgContent = strSplit(4)
    MsgContent = Base64DecodeString(MsgContent)
>>>>>>> 3b07fd90bb91919c2b4047f89cff85163e999376
    
End Sub
