VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim ProTime As Long, mode As Integer
'======================================

Public Sub Update()
    Page.Clear '清空画布
    '标题栏
    Dim m As MButtonState
    Page.Paint 0, 0, 0, GW, 61, argb(255, 89, 89, 89)
    m = CheckMouse(0, 0, GW - 120, 61)
    If m = mMouseDown Then
        '发送标题栏移动消息
        SendMessageA GHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        ReleaseCapture
    End If
    
    If Server.lis.State = 0 Then Page.Writes Server.Caption & "（连接未启动）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
    If Server.lis.State = 1 Then Page.Writes Server.Caption & "（连接已准备）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Server.lis.State = 2 Then Page.Writes Server.Caption & "（等待连接）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Server.lis.State = 3 Then Page.Writes Server.Caption & "（连接被挂起）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
    If Server.lis.State = 4 Then Page.Writes Server.Caption & "（正在解析域名）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Server.lis.State = 5 Then Page.Writes Server.Caption & "（已识别服务器主机）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Server.lis.State = 6 Then Page.Writes Server.Caption & "（正在连接服务端）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Server.lis.State = 7 Then Page.Writes Server.Caption & "（已连接服务端）", 20, 15, 20, argb(255, 232, 232, 232), style:=FontStyleBold
    If Server.lis.State = 8 Then Page.Writes Server.Caption & "（连接被关闭）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
    If Server.lis.State = 9 Then Page.Writes Server.Caption & "（连接服务端失败）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold

    m = CheckMouse(GW - 45, 15, 30, 30)
    Page.Paint 1, GW - 45, 15, 30, 30, argb(IIf(m, 255, 100), 255, 83, 57)
    Page.Writes "×", GW - 41, 16, 18, argb(255, 255, 255, 255), style:=FontStyleBold
    If m = mMouseUp Then Unload Server: End
    m = CheckMouse(GW - 85, 15, 30, 30)
    Page.Paint 1, GW - 85, 15, 30, 30, argb(IIf(m, 255, 100), 128, 128, 128)
    Page.Writes "-", GW - 77, 16, 18, argb(255, 255, 255, 255), style:=FontStyleBold
    If m = mMouseUp Then Server.WindowState = 1
    
    '计算动画百分比
    Dim pro As Single, dpro As Single
    pro = (GetTickCount - ProTime) / 300
    If pro > 1 Then pro = 1
    If mode = 1 Then pro = 1 - pro
    pro = Cubic(pro, 0, 1, 1, 1)
    dpro = 1 - pro
    
    '底部发送框
    Dim TextV As Boolean
    TextV = (pro = 1)
    If Server.Text4.Visible <> TextV Then Server.Text4.Visible = TextV

    Page.Paint 0, 0, GH - 82, GW, 83, argb(255, 232, 232, 232)
    
    If pro = 0 Then GoTo SkipTab1
    
    Page.Paint 2, 33, GH - (82 - 18) * pro + 40 * dpro, GW - 250 + 70, 42, argb(255 * pro, 255, 255, 255), Radius:=42
    Page.DrawImage "send.png", 630, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "send.png", 630, GH - 40 * pro + 40 * dpro, Alpha:=IIf(Server.Command2.Enabled, IIf(m, 1, 0.8), 0.3) * pro, pos:=posOnCenter
    ECore.TipMsg "发送消息", 0
    If m = mMouseDown And Mouse.button = 1 And Server.Text4.Visible Then Server.Command2_Click: Mouse.button = 0

    
    Page.DrawImage "clear.png", 690, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "clear.png", 690, GH - 40 * pro + 40 * dpro, Alpha:=1 * pro, pos:=posOnCenter
    ECore.TipMsg "清空", 0
    If m = mMouseDown And Mouse.button = 1 Then
        Server.Command3_Click
        Mouse.button = 0
    End If

SkipTab1:
    If pro = 1 Then GoTo SkipTab2
    
    Page.DrawImage "disconnect.png", 50, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "disconnect.png", 50, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "断开客户端连接", 0
    If m = mMouseDown And Mouse.button = 1 Then Server.Text2.Text = InputBox("将断开的客户端的序号", "断开客户端连接"): Server.Command1_Click: Mouse.button = 0
    Page.DrawImage "ban.png", 110, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "ban.png", 110, GH - 40 * dpro + 40 * pro, Alpha:=IIf(Server.Command5.Enabled, 1, 0.3) * dpro, pos:=posOnCenter
    ECore.TipMsg "客户端禁言", 0
    If m = mMouseDown And Mouse.button = 1 And Server.Command5.Enabled Then Server.Command5_Click: Mouse.button = 0
    
    Page.DrawImage "image.png", 692, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "image.png", 692, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "图片文字识别", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Server.OCR_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    Page.DrawImage "voice.png", 632, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "voice.png", 632, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "语音转文本", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Server.Audio_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0

    Page.DrawImage "save.png", 572, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "save.png", 572, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "保存记录", 0
    If m = mMouseDown And Mouse.button = 1 Then
        Server.Command4_Click
        Mouse.button = 0
    End If

    
SkipTab2:
    
    Page.DrawImage "tools.png", 752, GH - (40 - 6), Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "tools.png", 752, GH - (40 - 6), Alpha:=IIf(m, 1, 0.8), pos:=posOnCenter
    ECore.TipMsg "工具菜单", 0
    If m = mMouseDown And Mouse.button = 1 Then ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, Msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入资源
    Page.Res.NewImages App.path & "\assets", 0.5
    '创建页面
    ECore.Add Page, "MainPage"
End Sub
