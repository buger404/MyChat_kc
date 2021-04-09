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
    
    If Client.Winsock1.State = 0 Then Page.Writes Client.Caption & "（连接未启动）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
    If Client.Winsock1.State = 1 Then Page.Writes Client.Caption & "（连接已准备）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Client.Winsock1.State = 2 Then Page.Writes Client.Caption & "（等待连接）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Client.Winsock1.State = 3 Then Page.Writes Client.Caption & "（连接被挂起）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
    If Client.Winsock1.State = 4 Then Page.Writes Client.Caption & "（正在解析域名）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Client.Winsock1.State = 5 Then Page.Writes Client.Caption & "（已识别服务器主机）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Client.Winsock1.State = 6 Then Page.Writes Client.Caption & "（正在连接服务端）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
    If Client.Winsock1.State = 7 Then Page.Writes Client.Caption & "（已连接服务端）", 20, 15, 20, argb(255, 232, 232, 232), style:=FontStyleBold
    If Client.Winsock1.State = 8 Then Page.Writes Client.Caption & "（连接被关闭）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
    If Client.Winsock1.State = 9 Then Page.Writes Client.Caption & "（连接服务端失败）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold

    m = CheckMouse(GW - 45, 15, 30, 30)
    Page.Paint 1, GW - 45, 15, 30, 30, argb(IIf(m, 255, 100), 255, 83, 57)
    Page.Writes "×", GW - 41, 16, 18, argb(255, 255, 255, 255), style:=FontStyleBold
    If m = mMouseUp Then Unload Client: End
    m = CheckMouse(GW - 85, 15, 30, 30)
    Page.Paint 1, GW - 85, 15, 30, 30, argb(IIf(m, 255, 100), 128, 128, 128)
    Page.Writes "-", GW - 77, 16, 18, argb(255, 255, 255, 255), style:=FontStyleBold
    If m = mMouseUp Then Client.WindowState = 1
    
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
    If Not (Client.Text1.Visible Or Client.Picture2.Visible) Then TextV = False
    If Client.Text2.Visible <> TextV Then Client.Text2.Visible = TextV

    Page.Paint 0, 0, GH - 82, GW, 83, argb(255, 232, 232, 232)
    
    If pro = 0 Then GoTo SkipTab1
    
    If Client.Text1.Visible Or Client.Picture2.Visible Then
        Page.Paint 2, 33, GH - (82 - 18) * pro + 40 * dpro, GW - 250 + 70, 42, argb(255 * pro, 255, 255, 255), Radius:=42
        Page.DrawImage "send.png", 630, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "send.png", 630, GH - 40 * pro + 40 * dpro, Alpha:=IIf(Client.Command2.Enabled, IIf(m, 1, 0.8), 0.3) * pro, pos:=posOnCenter
        ECore.TipMsg "发送消息", 0
        If m = mMouseDown And Mouse.button = 1 And Client.Text1.Visible Then Client.SendMsg: Mouse.button = 0
    End If
    
    Dim ClearEn As Boolean
    ClearEn = ClearEn Or (Client.Command3.Enabled And Client.Text1.Text <> "" And Client.Text1.Visible)
    ClearEn = ClearEn Or (Client.Picture2.Visible)
    ClearEn = ClearEn Or (Client.Picture1.Visible)
    ClearEn = ClearEn Or (Client.Text5.Text <> "" And Client.Text5.Visible)
    
    Page.DrawImage "clear.png", 690, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "clear.png", 690, GH - 40 * pro + 40 * dpro, Alpha:=IIf(ClearEn, IIf(m, 1, 0.8), 0.3) * pro, pos:=posOnCenter
    ECore.TipMsg "清空", 0
    If m = mMouseDown And Mouse.button = 1 Then
        If Client.Text1.Visible Then Client.Command3_Click
        If Client.Text5.Visible Then Client.Text5.Text = ""
        If Client.Picture2.Visible Then Client.Picture2.Cls
        If Client.Picture1.Visible Then Client.Command4_Click
        Mouse.button = 0
    End If
    
    If Client.Picture1.Visible Then
        Dim ct(3) As Byte
        CopyMemory ct(0), Client.LBtnColor, 4
        Page.Paint 1, 50, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, ct(0), ct(1), ct(2)), pos:=posOnCenter
        If CheckMouse2 = mMouseUp Then Client.ColorPad.Color = Client.LBtnColor: Client.ColorPad.ShowColor: Client.LBtnColor = Client.ColorPad.Color
        Page.Paint 1, 50, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, 0, 0, 0), pos:=posOnCenter, size:=2, style:=1
        CopyMemory ct(0), Client.RBtnColor, 4
        Page.Paint 1, 110, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, ct(0), ct(1), ct(2)), pos:=posOnCenter
        If CheckMouse2 = mMouseUp Then Client.ColorPad.Color = Client.RBtnColor: Client.ColorPad.ShowColor: Client.RBtnColor = Client.ColorPad.Color
        Page.Paint 1, 110, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, 0, 0, 0), pos:=posOnCenter, size:=2, style:=1
        Dim pSize As Single
        pSize = Client.Picture1.DrawWidth / 20
        Page.ShowSliderBar pSize, "slider.png", 162, GH - 49 * pro
        If Int(pSize * 20) < 1 Then pSize = 1 / 20
        If Int(pSize * 20) <> Client.Picture1.DrawWidth Then Client.Picture1.DrawWidth = Int(pSize * 20)
        Page.Writes "笔刷大小：" & Client.Picture1.DrawWidth, 360, GH - 33 * pro, 18, argb(255 * pro, 128, 128, 128), align:=StringAlignmentCenter
        
        Page.DrawImage "drawing" & IIf(Client.DotMode, "d", "") & ".png", 632, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "drawing" & IIf(Client.DotMode, "d", "") & ".png", 632, GH - 40 * pro + 40 * dpro, Alpha:=1 * pro, pos:=posOnCenter
        ECore.TipMsg IIf(Not Client.DotMode, "连线模式", "画点模式"), 0
        If m = mMouseDown And Mouse.button = 1 Then Client.DotMode = Not Client.DotMode: Mouse.button = 0
    End If
    
    If Client.Text5.Visible Then
        Page.Writes "总字数 " & Len(Client.Text5.Text) & ", 行数 " & (UBound(Split(Client.Text5.Text, vbCrLf)) + 1), 20, GH - 71 * pro, 20, argb(255 * pro, 32, 32, 32)
        Page.Writes "光标选取位置 " & Client.Text5.SelStart & ", 选取字数 " & Client.Text5.SelLength, 20, GH - 46 * pro, 20, argb(255 * pro, 96, 96, 96)
    End If

SkipTab1:
    If pro = 1 Then GoTo SkipTab2

    Page.DrawImage "chat" & IIf(Not Client.Text1.Visible, "d", "") & ".png", 50, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "chat" & IIf(Not Client.Text1.Visible, "d", "") & ".png", 50, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "聊天模式", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.Option3_Click: Mouse.button = 0
    Page.DrawImage "notepad" & IIf(Not Client.Text5.Visible, "d", "") & ".png", 110, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "notepad" & IIf(Not Client.Text5.Visible, "d", "") & ".png", 110, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "记事本", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.Option2_Click: Mouse.button = 0
    Page.DrawImage "drawing" & IIf(Not Client.Picture1.Visible, "d", "") & ".png", 170, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "drawing" & IIf(Not Client.Picture1.Visible, "d", "") & ".png", 170, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "画板", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.Option1_Click: Mouse.button = 0
    Page.DrawImage "dj" & IIf(Not Client.Picture2.Visible, "d", "") & ".png", 230, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "dj" & IIf(Not Client.Picture2.Visible, "d", "") & ".png", 230, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "DJ Mode", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.Option4_Click: Mouse.button = 0
    
    Page.DrawImage "image.png", 692, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "image.png", 692, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    
    ECore.TipMsg "图片文字识别", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Client.OCR_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    Page.DrawImage "voice.png", 632, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "voice.png", 632, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    
    ECore.TipMsg "语音转文本", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Client.Audio_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    If Client.Picture2.Visible = False Then
        Page.DrawImage "save.png", 572, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "save.png", 572, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
        ECore.TipMsg "保存文件", 0
        If m = mMouseDown And Mouse.button = 1 Then
            If Client.Text1.Visible Then Client.save3_Click
            If Client.Text5.Visible Then Client.save2_Click
            If Client.Picture1.Visible Then Client.saveDrawing
            Mouse.button = 0
        End If
    End If
    
SkipTab2:

    Page.DrawImage "tools.png", 752, GH - (40 - 6), Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "tools.png", 752, GH - (40 - 6), Alpha:=IIf(m, 1, 0.8), pos:=posOnCenter
    ECore.TipMsg "工具菜单", 0
    If m = mMouseDown And Mouse.button = 1 Then
        ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    End If
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
