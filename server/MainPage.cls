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
    Dim selectIndex As Integer
    Dim scroll As ScrollArea, scrollY As Long, preY As Long, scrBottom As Boolean
'======================================

Public Sub Update()
redraw:
    Page.Clear '清空画布
    '标题栏
    Dim m As MButtonState
    Page.DrawImage "back.jpg", 0, 0
    Page.Paint 0, 0, 0, 301, GH, argb(100, 0, 0, 0)
    Page.Paint 0, 300, 0, GW - 300, GH, argb(230, 255, 255, 255)
    'Page.Paint 0, 0, 0, GW, 61, argb(255, 89, 89, 89)
    m = CheckMouse(300, 0, GW - 300 - 120, 61)
    If m = mMouseDown Then
        '发送标题栏移动消息
        SendMessageA GHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        ReleaseCapture
    End If
    
    For i = 1 To UBound(groups)
        m = CheckMouse(0, (i - 1) * 70 + 60, 301, 71)
        Page.Paint 0, 0, (i - 1) * 70 + 60, 301, 71, argb(IIf(selectIndex = i, 60, IIf(m > 0, 30, 0)), 0, 0, 0)
        Page.Writes groups(i).Name, 30, (i - 1) * 70 + 60 + 5, 20, argb(255, 255, 255, 255), style:=FontStyleBold
        Dim tip As String
        tip = groups(i).Msg(UBound(groups(i).Msg)).Name & ":" & groups(i).Msg(UBound(groups(i).Msg)).Content
        If Len(tip) > 12 Then tip = Left(tip, 10) & "..."
        Page.Writes tip, 30, (i - 1) * 70 + 60 + 33, 18, argb(200, 255, 255, 255)
        If m = 3 Then selectIndex = i: scrBottom = True
        If groups(i).unreadTick > 0 Then
            Page.Paint 2, 230, (i - 1) * 70 + 60 + 20, 50, 30, argb(255, 251, 77, 63), 100
            Page.Writes IIf(groups(i).unreadTick = 100, "99+", groups(i).unreadTick), 230 + 25, (i - 1) * 70 + 60 + 21, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        End If
        If selectIndex = i Then groups(i).unreadTick = 0
    Next
    
    Page.Writes "MyChat · 教师端", 20, 15, 20, argb(160, 255, 255, 255)
    Page.Writes groups(selectIndex).Name, 20 + 300, 15, 20, argb(255, 0, 0, 0), style:=FontStyleBold
    
    StartScrollArea Page, scroll
        Page.Clear argb(0, 0, 0, 0)
        Dim drx As Long, dry As Long, tw As Long, th As Long, ory As Long
        dry = 20 - scrollY: ory = dry
        For i = 1 To UBound(groups(selectIndex).Msg)
            With groups(selectIndex).Msg(i)
                drx = IIf(.id = userId, scroll.Width - 20, 20)

                Page.Writes .Name & IIf(id > 0, "(#" & .id & ")", "") & "   " & format(.time, "hh:MM:ss"), drx, dry, 18, IIf(.id > 0, argb(255, 168, 168, 168), argb(255, 0, 176, 240)), align:=IIf(.id = userId, StringAlignmentFar, StringAlignmentNear)

                tw = EF.GetWidth(Page.GG, .Content, 18, StringAlignmentNear, FontStyleRegular)
                th = EF.GetHeight(Page.GG, .Content, 18, StringAlignmentNear, FontStyleRegular)
                If dry >= -(th + 40 + 20) And dry <= scroll.Height Then
                    If .id = -1 Then
                        Page.Paint 2, scroll.Width / 2, dry + 30, tw + 40, 40, argb(255, 212, 212, 212), 100, pos:=posOnCenter
                        Page.Writes .Content, scroll.Width / 2, dry + 20 - 4, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter
                        
                    Else
                        Page.Writes .Name & IIf(id > 0, "(#" & .id & ")", "") & "   " & format(.time, "hh:MM:ss"), drx, dry, 18, IIf(.id > 0 And .id <> userId, argb(255, 168, 168, 168), argb(255, 84, 223, 143)), align:=IIf(.id = userId, StringAlignmentFar, StringAlignmentNear)
                        If .id = userId Then drx = drx - tw - 20
                        If tw > scroll.Width - 60 Then
                            th = th + (tw - scroll.Width + 60) / 18 / 0.75 - 3
                            tw = scroll.Width - 60
                        End If
                        Page.Paint 2, drx, dry + 25, tw + 20, th + 20, IIf(.id = userId, argb(255, 84, 223, 143), argb(255, 232, 232, 232)), 16
                        If CheckMouse(drx + 300, dry + 60, tw + 20, th + 20) > 0 And Mouse.button = 2 Then
                            MenuWindow.id = .id: Mouse.button = 0
                            Server.PopupMenu MenuWindow.msgMenu
                        End If
                        Page.Writes .Content, drx + 10, dry + 25 + 10, 18, IIf(.id = userId, argb(255, 255, 255, 255), argb(255, 0, 0, 0)), tw + 1, th
                        
                    End If
                End If
                If .id = -1 Then
                    dry = dry + th + 20 + 20
                Else
                    dry = dry + th + 40 + 20
                End If
            End With
        Next
        Dim maxScr As Long
        If dry - ory > scroll.Height - 30 Then
            maxScr = dry - ory - scroll.Height + 30
        Else
            maxScr = 0
        End If
        If scrBottom Then
            scrBottom = False
            scrollY = maxScr
            GoTo redraw
        End If
        If maxScr = 0 Then maxScr = 1
        dry = 20 + (scroll.Height - 20 - 40) * (scrollY / maxScr)
        Page.Paint 2, scroll.Width - 15, dry, 8, 40, argb(255, 128, 128, 128), 10
        m = CheckMouse(scroll.Width - 15 + 300, 20 + 60, 8, scroll.Height - 40)
        If (m > 0 And Mouse.button = 1) Or (Mouse.button = 1 And preY <> 0) Then
            If preY = 0 Then preY = Mouse.y - dry - 60
            dry = Mouse.y - 60 - preY * 2
            If dry < 20 Then dry = 0
            If dry > scroll.Height - 20 - 40 Then dry = scroll.Height - 20 - 40
            scrollY = dry / (scroll.Height - 20 - 40) * maxScr
        ElseIf preY <> 0 Then
            preY = 0
        End If
        If scrollY > maxScr Then scrollY = maxScr: GoTo redraw
        If scrollY < 0 Then scrollY = 0: GoTo redraw
        
    EndScrollArea Page, 300, 60, 0, 0

    'If Server.lis.State = 0 Then Page.Writes Server.Caption & "（连接未启动）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
   ' If Server.lis.State = 1 Then Page.Writes Server.Caption & "（连接已准备）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
   ' If Server.lis.State = 2 Then Page.Writes Server.Caption & "（等待连接）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
   ' If Server.lis.State = 3 Then Page.Writes Server.Caption & "（连接被挂起）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
  '  If Server.lis.State = 4 Then Page.Writes Server.Caption & "（正在解析域名）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
  '  If Server.lis.State = 5 Then Page.Writes Server.Caption & "（已识别服务器主机）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
   ' If Server.lis.State = 6 Then Page.Writes Server.Caption & "（正在连接服务端）", 20, 15, 20, argb(255, 255, 170, 49), style:=FontStyleBold
   ' If Server.lis.State = 7 Then Page.Writes Server.Caption & "（已连接服务端）", 20, 15, 20, argb(255, 232, 232, 232), style:=FontStyleBold
   ' If Server.lis.State = 8 Then Page.Writes Server.Caption & "（连接被关闭）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold
  '  If Server.lis.State = 9 Then Page.Writes Server.Caption & "（连接服务端失败）", 20, 15, 20, argb(255, 255, 83, 57), style:=FontStyleBold

    m = CheckMouse(GW - 45, 15, 30, 30)
    Page.Paint 1, GW - 45, 15, 30, 30, argb(IIf(m, 255, 100), 180, 180, 180)
    Page.Writes "×", GW - 41, 16, 18, argb(255, 0, 0, 0), style:=FontStyleBold
    If m = mMouseUp Then Unload Server: End
    m = CheckMouse(GW - 85, 15, 30, 30)
    Page.Paint 1, GW - 85, 15, 30, 30, argb(IIf(m, 255, 100), 180, 180, 180)
    Page.Writes "-", GW - 77, 16, 18, argb(255, 0, 0, 0), style:=FontStyleBold
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

    Page.Paint 0, 300, GH - 82, GW, 83, argb(255, 232, 232, 232)
    
    If pro = 0 Then GoTo SkipTab1
    
    Page.Paint 2, 300 + 33, GH - (82 - 18) * pro + 40 * dpro, GW - 250 - 300 + 70, 42, argb(255 * pro, 255, 255, 255), Radius:=42
    Page.DrawImage "send.png", 300 + 630, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "send.png", 300 + 630, GH - 40 * pro + 40 * dpro, Alpha:=IIf(Server.Command2.Enabled, IIf(m, 1, 0.8), 0.3) * pro, pos:=posOnCenter
    ECore.TipMsg "发送消息", 0

    'If m = mMouseDown And Mouse.button = 1 And Server.Text4.Visible Then Call Server.SendMsg: Mouse.button = 0


    If m = mMouseDown And Mouse.button = 1 And Server.Text4.Visible Then
        Server.SendMsg
        Mouse.button = 0
        'AddMessage groups(selectIndex).id, userId, userName, Server.Text4.Text
        scrBottom = True
    End If

    
    Page.DrawImage "clear.png", 300 + 690, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "clear.png", 300 + 690, GH - 40 * pro + 40 * dpro, Alpha:=1 * pro, pos:=posOnCenter
    ECore.TipMsg "清空", 0
    If m = mMouseDown And Mouse.button = 1 Then
        Server.Command3_Click
        Mouse.button = 0
    End If

SkipTab1:
    If pro = 1 Then GoTo SkipTab2
    
    Page.DrawImage "disconnect.png", 300 + 50, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "disconnect.png", 300 + 50, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "断开客户端连接", 0
    If m = mMouseDown And Mouse.button = 1 Then Server.Text2.Text = InputBox("将断开的客户端的序号", "断开客户端连接"): Server.Command1_Click: Mouse.button = 0
    
    Page.DrawImage "disconnect.png", 300 + 110, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "disconnect.png", 300 + 110, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "创建组", 0
    If m = mMouseDown And Mouse.button = 1 Then Call Server.createGrp: Mouse.button = 0
    
    Page.DrawImage "ban.png", 300 + 170, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "ban.png", 300 + 170, GH - 40 * dpro + 40 * pro, Alpha:=IIf(Server.Command5.Enabled, 1, 0.3) * dpro, pos:=posOnCenter
    ECore.TipMsg "客户端禁言", 0
    If m = mMouseDown And Mouse.button = 1 And Server.Command5.Enabled Then Server.Command5_Click: Mouse.button = 0
    
    Page.DrawImage "image.png", 300 + 692, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "image.png", 300 + 692, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "图片文字识别", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Server.OCR_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    Page.DrawImage "voice.png", 300 + 632, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "voice.png", 300 + 632, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "语音转文本", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Server.Audio_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0

    Page.DrawImage "save.png", 300 + 572, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "save.png", 300 + 572, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "保存记录", 0
    If m = mMouseDown And Mouse.button = 1 Then
        Server.Command4_Click
        Mouse.button = 0
    End If

    
SkipTab2:
    
    Page.DrawImage "tools.png", 300 + 752, GH - (40 - 6), Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "tools.png", 300 + 752, GH - (40 - 6), Alpha:=IIf(m, 1, 0.8), pos:=posOnCenter
    ECore.TipMsg "工具菜单", 0
    If m = mMouseDown And Mouse.button = 1 Then ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
    scrollY = scrollY - Direction * Depth * 100
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
    Page.Res.newImage App.path & "\assets\back.jpg", GW, GH
    Page.Res.ApplyBlurEffect "back.jpg", 20, 0
    Page.Res.NewImages App.path & "\assets", 0.5
    scroll = CreateScrollArea(798, 418)
    selectIndex = 2
    '创建页面
    ECore.Add Page, "MainPage"
End Sub
