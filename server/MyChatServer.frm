VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin MSComDlg.CommonDialog trans 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "任何文件|*.*"
   End
   Begin MSScriptControlCtl.ScriptControl vbs 
      Index           =   0
      Left            =   6600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
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
Dim pypid
Dim grpid As Integer
Dim g As String, q As Single, m As Single
Dim IPPage As IPPage

Public Sub Command1_Click()
    Dim C As Single
    C = Val(Text2.Text)
    
    If C > Winsock.UBound Then
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
    DoEvents
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

Public Sub SendMsg()

    If Text4.Text = "" Then VBA.Beep: Exit Sub
    
    Dim S As Single
    S = 1
    Do While (S <= Winsock.UBound)
        If Winsock(S).State = 7 Then
            'MsgBox Str(MainPage.selectIndex)
            Winsock(S).SendData "msg;" + str(groups(MainPage.selectIndex).id) + ";" + Base64EncodeString(userName) + ";" + str(userId) + ";" + Base64EncodeString(Text4.Text) + ";" + vbCrLf
            DoEvents
        End If
        S = S + 1
    Loop
    
    Call AddMessage(groups(MainPage.selectIndex).id, userId, userName, Text4.Text)
    
    Text4.Text = ""

End Sub



Public Sub Command3_Click()
    Text3.Text = ""
    Text4.Text = ""
End Sub

Public Sub Command4_Click()
    Open App.path & "\" & "服务端消息记录" & str(q) & ".txt" For Output As #1
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
        Do While (S <= Winsock.UBound)
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
        Do While (S <= Winsock.UBound)
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
    LoadBlackList
    ReDim groups(0): ReDim bans(0)
    Set Robots = New RobotCore
    AddGroup 1, -2, True, "大厅", "老师"
    userId = -2: userName = "老师"
    
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
    Call AddMessage(1, -1, "系统消息", "服务端ip：" + lis.LocalIP)
    Text3.Visible = False: Text4.Visible = True
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
    
    Dim ttt As String, code As String
    Open App.path & "\robots\random.vbs" For Input As #1
    Do While Not EOF(1)
        Line Input #1, ttt
        code = code & ttt & vbCrLf
    Loop
    Close #1
    
    'vbs.AddCode code
    'vbs.Run "Process", "hhhhh"
    'MsgBox vbs.Eval("Guidence")
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
    
    pop = Winsock.UBound
    
    If Winsock(m).State = sckClosed Then
        Winsock(m).Accept requestID
        Winsock(m).SendData "identify;" & m & vbCrLf
        For i = 1 To UBound(groups)
            Winsock(m).SendData "addgroup;" & groups(i).leader & ";" & Base64EncodeString(groups(i).Name) & ";" & Base64EncodeString(groups(i).LeaderName) & ";" & groups(i).id & vbCrLf
            For j = 1 To UBound(groups(i).members)
                Winsock(m).SendData "addmember;" & Base64EncodeString(groups(i).members(j).Name) & ";" & groups(i).members(j).id & ";" & groups(i).id & vbCrLf
            Next
        Next
        Winsock(m).SendData "black;" & GetBlackString & vbCrLf
        DoEvents
    End If
    
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
    'Call AddMessage(1, -1, "系统消息", "")
    DeleteMember index, 1
    For Each w In Winsock
        If w.State = 7 Then w.SendData "deletemember;" & index & ";" & 1 & vbCrLf
    Next
End Sub

Private Sub Winsock_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If State = True Then Exit Sub
    
    Dim strSplit() As String
    Dim id As Integer
    Dim MsgType As String
    Dim strData As String
    Winsock(index).GetData strData
    
    Dim S As Single
  '  S = 1
   ' Do While (S <= Winsock.ubound)
   '     If Winsock(S).State = 7 Then
   '         Winsock(S).SendData strData
   '         DoEvents
   '     End If
   '     S = S + 1
   ' Loop
    
    Dim cmds() As String
    cmds = Split(strData, vbCrLf)
    
    For k = 0 To UBound(cmds) - 1
        strSplit = Split(cmds(k), ";")
        id = index
        MsgType = strSplit(0)
    
        Select Case MsgType
        Case "getId"
            Winsock(id).SendData "getId;" + str(id) + ";" + vbCrLf
        Case "filerecv"
            ShellExecuteA 0, "open", App.path & "\FileTransportation.exe", "-d;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & ";" & strSplit(4) & ";" & strSplit(5), "", SW_SHOW
        Case "filesend"
            If Val(strSplit(6)) = 0 Then
                ShellExecuteA 0, "open", App.path & "\FileTransportation.exe", "-d;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & ";" & strSplit(4) & ";" & strSplit(5), "", SW_SHOW
                For j = 1 To UBound(groups)
                    If groups(j).id = Val(strSplit(7)) Then
                        For i = 1 To UBound(groups(j).members)
                            If Winsock(groups(j).members(i).id).State = 7 And groups(j).members(i).id <> index Then
                                Winsock(groups(j).members(i).id).SendData "filerecv;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & ";" & strSplit(4) & ";" & strSplit(5) & vbCrLf
                            End If
                            DoEvents
                        Next
                    End If
                Next
            ElseIf Val(strSplit(6)) = -2 Then
                ShellExecuteA 0, "open", App.path & "\FileTransportation.exe", "-d;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & ";" & strSplit(4) & ";" & strSplit(5), "", SW_SHOW
            Else
                If Winsock(Val(strSplit(6))).State = 7 Then
                    Winsock(Val(strSplit(6))).SendData "filerecv;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & ";" & strSplit(4) & ";" & strSplit(5) & vbCrLf
                End If
            End If
        Case "msg"
            Dim Name As String
            Dim MsgContent As String
            Name = Base64DecodeString(strSplit(2))
            grpid = strSplit(1)
            MsgContent = strSplit(4)
            MsgContent = Base64DecodeString(MsgContent)
            Call AddMessage(Int(grpid), Val(strSplit(3)), Name, MsgContent)
            
            strData = MsgType + ";" + str(grpid) + ";" + Base64EncodeString(Name) + ";" + strSplit(3) + ";" + Base64EncodeString(MsgContent) & vbCrLf
    
            S = 1
            Do While (S <= Winsock.UBound)
                If Winsock(S).State = 7 And S <> id Then
                    Winsock(S).SendData strData
                    DoEvents
                End If
                S = S + 1
            Loop
        Case "picmsg"
        
        Case "addgrouprequest"
            If groups(Val(strSplit(3))).leader = -2 Then
                If Val(strSplit(3)) = 1 Then GoTo skipNotify
                Me.SetFocus
                If MsgBox(Base64DecodeString(strSplit(1)) & "(#" & Val(strSplit(2)) & ") 申请加入组“" & groups(Val(strSplit(3))).Name & "”，是否同意？", 48 + vbYesNo) = vbYes Then
skipNotify:
                    AddMember Base64DecodeString(strSplit(1)), Val(strSplit(2)), Val(strSplit(3))
                    For Each w In Winsock
                        If w.State = 7 Then w.SendData "addmember;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & vbCrLf
                        DoEvents
                    Next
                End If
            Else
                Winsock(groups(Val(strSplit(3))).leader).SendData "grouprequest;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & vbCrLf
            End If
        Case "broadcast"
            Dim newcmd As String
            For i = 1 To UBound(strSplit)
                newcmd = newcmd & strSplit(i) & IIf(i < UBound(strSplit), ";", "")
            Next
            For Each w In Winsock
                If w.State = 7 And w.index <> index Then w.SendData newcmd & vbCrLf
                DoEvents
            Next
        Case "addmember"
            AddMember Base64DecodeString(strSplit(1)), Val(strSplit(2)), Val(strSplit(3))
            AddMessage Val(strSplit(3)), -1, "系统消息", Base64DecodeString(strSplit(1)) & "加入了本讨论组"
        Case "creategroup"
            ProcessCreateGroup strSplit, id
        Case "quitgroup"
            If index = groups(Val(strSplit(1))).leader Then
                '解散处理
                DeleteGroup Val(strSplit(1))
                For Each w In Winsock
                    If w.State = 7 Then w.SendData "deletegroup;" & Val(strSplit(1)) & vbCrLf
                    DoEvents
                Next
            Else
                '退群处理
                DeleteMember index, Val(strSplit(1))
                For Each w In Winsock
                    If w.State = 7 Then w.SendData "deletemember;" & Val(strSplit(1)) & ";" & index & vbCrLf
                    DoEvents
                Next
            End If
        Case "addban"
            ProcessBan Val(strSplit(1)), Val(strSplit(3)), Val(strSplit(2))
        End Select
    Next
End Sub
Public Sub ProcessCreateGroup(arg() As String, id As Integer)
    Dim gid As Integer
    Call AddGroup(groups(UBound(groups)).id + 1, id, True, Base64DecodeString(arg(1)), Base64DecodeString(arg(2)))
    gid = groups(UBound(groups)).id
    AddMember groups(UBound(groups)).LeaderName, groups(UBound(groups)).id, UBound(groups)
    Dim w As Winsock
    For Each w In Winsock
        If w.State = 7 Then w.SendData "addgroup;" & id & ";" & arg(1) & ";" & arg(2) & ";" & gid & vbCrLf & "addmember;" & arg(2) & ";" & id & ";" & gid & vbCrLf
        DoEvents
    Next
End Sub
Public Sub ProcessBan(group As Integer, id As Integer, Duration As Long)
    Dim bname As String
    bname = Robots.GetMemberName(group, Robots.GetMemberIndex(group, id))
    bname = bname & "(#" & id & ")"
    AddMessage group, -1, "系统消息", bname & "被禁言" & Round(Duration / 60) & "分钟"
    For Each w In Winsock
        If w.State = 7 Then
            If w.index = id Then w.SendData "addban;" & group & ";" & Duration & vbCrLf
            w.SendData "msg;" & group & ";" & Base64EncodeString("系统消息") & ";-1;" & Base64EncodeString(bname & "被禁言" & Round(Duration / 60) & "分钟") & vbCrLf
        End If
        DoEvents
    Next
End Sub
