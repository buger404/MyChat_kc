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
    Public selectIndex As Integer
    Dim scroll As ScrollArea, scrollY As Long, preY As Long, scrBottom As Boolean, sideY As Long
'======================================

Public Sub Update()
ReDraw:
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
    
    If UBound(groups) = 0 Then Exit Sub
    If selectIndex > UBound(groups) Then selectIndex = UBound(groups)
    If Not groups(selectIndex).isJoin Then mode = 1
    
    For i = 1 To UBound(groups)
        m = CheckMouse(0, (i - 1) * 70 + 60 - sideY, 301, 71)
        Page.Paint 0, 0, (i - 1) * 70 + 60 - sideY, 301, 71, argb(IIf(selectIndex = i, 60, IIf(m > 0, 30, 0)), 0, 0, 0)
        Page.Writes groups(i).Name, 30, (i - 1) * 70 + 60 + 5 - sideY, 20, argb(255, 255, 255, 255), style:=FontStyleBold
        Dim tip As String
        tip = groups(i).Msg(UBound(groups(i).Msg)).Name & ":" & groups(i).Msg(UBound(groups(i).Msg)).Content
        If Len(tip) > 12 Then tip = Left(tip, 10) & "..."
        If tip = ":" Then tip = "(暂无消息)"
        If Not groups(i).isJoin Then tip = "（尚未加入）"
        Page.Writes tip, 30, (i - 1) * 70 + 60 + 33 - sideY, 18, argb(200, 255, 255, 255)
        If m = 3 Then
            If Mouse.button = 1 Then
                selectIndex = i: scrBottom = True: mode = 0
            ElseIf Mouse.button = 2 And i > 1 Then
                MenuWindow.groupId = i
                Mouse.button = 0
                MenuWindow.quitGroup.Caption = IIf(groups(i).leader = userId, "解散", "退出")
                Client.PopupMenu MenuWindow.groupMenu
            End If
        End If
        If (Not groups(i).isJoin) And groups(i).unreadTick > 0 Then groups(i).unreadTick = 0
        If groups(i).unreadTick > 0 Then
            Page.Paint 2, 230, (i - 1) * 70 + 60 + 20 - sideY, 50, 30, argb(255, 251, 77, 63), 100
            Page.Writes IIf(groups(i).unreadTick = 100, "99+", groups(i).unreadTick), 230 + 25, (i - 1) * 70 + 60 + 21 - sideY, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        End If
        If selectIndex = i And groups(i).unreadTick > 0 Then groups(i).unreadTick = 0: scrBottom = True
    Next

    Page.Writes "MyChat · 学生端", 20, 15, 20, argb(160, 255, 255, 255)
    Dim maxScr As Long
    If (i - 1) * 70 + 60 > GH - 30 Then
        maxScr = (i - 1) * 70 + 60 - GH + 30
    Else
        maxScr = 0
    End If
    If sideY > maxScr Then sideY = maxScr: GoTo ReDraw
    If sideY < 0 Then sideY = 0: GoTo ReDraw

    Page.Writes groups(selectIndex).Name & "(" & UBound(groups(selectIndex).members) & ")", 20 + 300, 15, 20, argb(255, 0, 0, 0), style:=FontStyleBold
    
    Dim disbledToolbar As Boolean
    If Not groups(selectIndex).isJoin Then
        Page.Paint 2, 300 + (GW - 300) / 2, GH / 2, 450, 250, argb(120, 0, 0, 0), 20, pos:=posOnCenter
        Page.Writes "您尚未加入该讨论组", 300 + (GW - 300) / 2, GH / 2 - 70, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        Page.Writes "组长：" & groups(selectIndex).LeaderName, 300 + (GW - 300) / 2, GH / 2 - 30, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        If Page.ShowColorButton(2, 300 + (GW - 300) / 2 - 100, GH - 250, 200, 40, "申请加入", argb(255, 255, 255, 255), argb(220, 84, 223, 143), argb(255, 84, 223, 143), 99, size:=20) = mMouseUp Then
            ' 申请加入的代码放置此处
            Client.Winsock1.SendData "addgrouprequest;" & Base64EncodeString(userName) & ";" & userId & ";" & groups(selectIndex).id & vbCrLf
        End If
        disbledToolbar = True
        GoTo nomsg
    End If
    
    StartScrollArea Page, scroll
        Page.Clear argb(0, 0, 0, 0)
        Dim drx As Long, dry As Long, tw As Long, th As Long, ory As Long
        dry = 20 - scrollY: ory = dry
        For i = 1 To UBound(groups(selectIndex).Msg)
            With groups(selectIndex).Msg(i)
                drx = IIf(.id = userId, scroll.Width - 20, 20)
                
                tw = EF.GetWidth(Page.GG, .Content, 18, StringAlignmentNear, FontStyleRegular)
                th = EF.GetHeight(Page.GG, .Content, 18, StringAlignmentNear, FontStyleRegular)
                If dry >= -(th + 40 + 20) And dry <= scroll.Height Then
                    If .id = -1 Then
                        Page.Paint 2, scroll.Width / 2, dry + 30, tw + 40, 40, argb(255, 212, 212, 212), 100, pos:=posOnCenter
                        Page.Writes .Content, scroll.Width / 2, dry + 20 - 4, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter
                        
                    Else
                        Page.Writes .Name & IIf(id > 0, "(#" & .id & ")", "") & "   " & format(.time, "hh:MM:ss"), drx, dry, 18, IIf(.id > 0, argb(255, 168, 168, 168), argb(255, 0, 176, 240)), align:=IIf(.id = userId, StringAlignmentFar, StringAlignmentNear)
                        Page.Writes .Name & IIf(id > 0, "(#" & .id & ")", "") & "   " & format(.time, "hh:MM:ss"), drx, dry, 18, IIf(.id > 0 And .id <> userId, argb(255, 168, 168, 168), argb(255, 84, 223, 143)), align:=IIf(.id = userId, StringAlignmentFar, StringAlignmentNear)
                        If .id = userId Then drx = drx - tw - 20
                        If tw > scroll.Width - 60 Then
                            th = th + (tw - scroll.Width + 60) / 18 / 0.75 - 3
                            tw = scroll.Width - 60
                        End If
                        Page.Paint 2, drx, dry + 25, tw + 20, th + 20, IIf(.id = userId, argb(255, 84, 223, 143), argb(255, 232, 232, 232)), 16
                        If CheckMouse(drx + 300, dry + 60 + 25, tw + 20, th + 20) > 0 And Mouse.button = 2 Then
                            MenuWindow.id = .id: selectMsg = groups(selectIndex).Msg(i): Mouse.button = 0
                            MenuWindow.banMsg.Visible = (.id <> userId And (userId = groups(selectIndex).leader Or userId = -2))
                            MenuWindow.kickGroup.Visible = (.id <> userId And (userId = groups(selectIndex).leader Or userId = -2))
                            Client.PopupMenu MenuWindow.msgMenu
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
        If dry - ory > scroll.Height - 30 Then
            maxScr = dry - ory - scroll.Height + 30
        Else
            maxScr = 0
        End If
        If scrBottom Then
            scrBottom = False
            scrollY = maxScr
            GoTo ReDraw
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
        If scrollY > maxScr Then scrollY = maxScr: GoTo ReDraw
        If scrollY < 0 Then scrollY = 0: GoTo ReDraw
        
    EndScrollArea Page, 300, 60, 0, 0

    For i = 1 To UBound(bans)
        Dim span As Date, tickspan As Long
        If i > UBound(bans) Then Exit For
        If bans(i).groupId = groups(selectIndex).id And bans(i).id = userId Then
            span = Now - bans(i).StartTime
            tickspan = Minute(span) * 60 + Second(span) + Hour(span) * 60 * 60
            If tickspan < bans(i).Duration Then
                mode = 1
                Page.Paint 2, 300 + (GW - 300) / 2, GH / 2, 450, 250, argb(100, 0, 0, 0), 20, pos:=posOnCenter
                Page.Writes "您正被禁言，无法发言", 300 + (GW - 300) / 2, GH / 2 - 70, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
                Page.Writes Int((bans(i).Duration - tickspan) / 60) & ":" & format((bans(i).Duration - tickspan) Mod 60, "00"), 300 + (GW - 300) / 2, GH / 2 - 0, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
                disbledToolbar = True
                Exit For
            Else
                bans(i) = bans(UBound(bans))
                ReDim Preserve bans(UBound(bans) - 1)
                i = i - 1
            End If
        End If
    Next
nomsg:
    
    m = CheckMouse(GW - 45, 15, 30, 30)
    Page.Paint 1, GW - 45, 15, 30, 30, argb(IIf(m, 255, 100), 180, 180, 180)
    Page.Writes "×", GW - 41, 16, 18, argb(255, 0, 0, 0), style:=FontStyleBold
    If m = mMouseUp Then Unload Client: End
    m = CheckMouse(GW - 85, 15, 30, 30)
    Page.Paint 1, GW - 85, 15, 30, 30, argb(IIf(m, 255, 100), 180, 180, 180)
    Page.Writes "-", GW - 77, 16, 18, argb(255, 0, 0, 0), style:=FontStyleBold
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
    If Not (Client.Text1.Enabled Or Client.Picture2.Visible) Then TextV = False
    If Client.Text2.Visible <> TextV Then Client.Text2.Visible = TextV

    If Not disbledToolbar Then Page.Paint 0, 300, GH - 82, GW, 83, argb(255, 232, 232, 232)
    
    If pro = 0 Or disbledToolbar Then GoTo SkipTab1
    
    If Client.Text1.Enabled Or Client.Picture2.Visible Then
        Page.Paint 2, 33 + 300, GH - (82 - 18) * pro + 40 * dpro, GW - 250 - 300 + 70, 42, argb(255 * pro, 255, 255, 255), Radius:=42
        Page.DrawImage "send.png", 630 + 300, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "send.png", 630 + 300, GH - 40 * pro + 40 * dpro, Alpha:=IIf(Client.Command2.Enabled, IIf(m, 1, 0.8), 0.3) * pro, pos:=posOnCenter
        ECore.TipMsg "发送消息", 0
        If m = mMouseDown And Mouse.button = 1 And Client.Text1.Enabled Then Client.SendMsg: Mouse.button = 0
    End If
    
    Dim ClearEn As Boolean
    ClearEn = ClearEn Or (Client.Command3.Enabled And Client.Text1.Text <> "" And Client.Text1.Enabled)
    ClearEn = ClearEn Or (Client.Picture2.Visible)
    ClearEn = ClearEn Or (Client.Picture1.Visible)
    ClearEn = ClearEn Or (Client.Text5.Text <> "" And Client.Text5.Visible)
    
    Page.DrawImage "clear.png", 690 + 300, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "clear.png", 690 + 300, GH - 40 * pro + 40 * dpro, Alpha:=IIf(ClearEn, IIf(m, 1, 0.8), 0.3) * pro, pos:=posOnCenter
    ECore.TipMsg "清空", 0
    If m = mMouseDown And Mouse.button = 1 Then
        If Client.Text1.Enabled Then Client.Command3_Click
        If Client.Text5.Visible Then Client.Text5.Text = ""
        If Client.Picture2.Visible Then Client.Picture2.Cls
        If Client.Picture1.Visible Then Client.Command4_Click
        Mouse.button = 0
    End If
    
    If Client.Picture1.Visible Then
        Dim ct(3) As Byte
        CopyMemory ct(0), Client.LBtnColor, 4
        Page.Paint 1, 50 + 300, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, ct(0), ct(1), ct(2)), pos:=posOnCenter
        If CheckMouse2 = mMouseUp Then Client.ColorPad.Color = Client.LBtnColor: Client.ColorPad.ShowColor: Client.LBtnColor = Client.ColorPad.Color
        Page.Paint 1, 50 + 300, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, 0, 0, 0), pos:=posOnCenter, size:=2, style:=1
        CopyMemory ct(0), Client.RBtnColor, 4
        Page.Paint 1, 110 + 300, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, ct(0), ct(1), ct(2)), pos:=posOnCenter
        If CheckMouse2 = mMouseUp Then Client.ColorPad.Color = Client.RBtnColor: Client.ColorPad.ShowColor: Client.RBtnColor = Client.ColorPad.Color
        Page.Paint 1, 110 + 300, GH - 42 * pro + 40 * dpro, 46, 46, argb(255, 0, 0, 0), pos:=posOnCenter, size:=2, style:=1
        Dim pSize As Single
        pSize = Client.Picture1.DrawWidth / 20
        Page.ShowSliderBar pSize, "slider.png", 162 + 300, GH - 49 * pro
        If Int(pSize * 20) < 1 Then pSize = 1 / 20
        If Int(pSize * 20) <> Client.Picture1.DrawWidth Then Client.Picture1.DrawWidth = Int(pSize * 20)
        Page.Writes "笔刷大小：" & Client.Picture1.DrawWidth, 360 + 300, GH - 33 * pro, 18, argb(255 * pro, 128, 128, 128), align:=StringAlignmentCenter
        
        Page.DrawImage "drawing" & IIf(Client.DotMode, "d", "") & ".png", 632 + 300, GH - 40 * pro + 40 * dpro, Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "drawing" & IIf(Client.DotMode, "d", "") & ".png", 632 + 300, GH - 40 * pro + 40 * dpro, Alpha:=1 * pro, pos:=posOnCenter
        ECore.TipMsg IIf(Not Client.DotMode, "连线模式", "画点模式"), 0
        If m = mMouseDown And Mouse.button = 1 Then Client.DotMode = Not Client.DotMode: Mouse.button = 0
    End If
    
    If Client.Text5.Visible Then
        Page.Writes "总字数 " & Len(Client.Text5.Text) & ", 行数 " & (UBound(Split(Client.Text5.Text, vbCrLf)) + 1), 20 + 300, GH - 71 * pro, 20, argb(255 * pro, 32, 32, 32)
        Page.Writes "光标选取位置 " & Client.Text5.SelStart & ", 选取字数 " & Client.Text5.SelLength, 20 + 300, GH - 46 * pro, 20, argb(255 * pro, 96, 96, 96)
    End If

SkipTab1:
    If pro = 1 Or disbledToolbar Then GoTo SkipTab2

    Page.DrawImage "chat" & IIf(Client.Text5.Visible Or Client.Picture1.Visible Or Client.Picture2.Visible, "d", "") & ".png", 50 + 300, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "chat" & IIf(Client.Text5.Visible Or Client.Picture1.Visible Or Client.Picture2.Visible, "d", "") & ".png", 50 + 300, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "聊天模式", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.Option3_Click: Mouse.button = 0
    Page.DrawImage "notepad" & IIf(Not Client.Text5.Visible, "d", "") & ".png", 110 + 300, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "notepad" & IIf(Not Client.Text5.Visible, "d", "") & ".png", 110 + 300, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "记事本", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.Option2_Click: Mouse.button = 0
    Page.DrawImage "drawing" & IIf(Not Client.Picture1.Visible, "d", "") & ".png", 170 + 300, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "drawing" & IIf(Not Client.Picture1.Visible, "d", "") & ".png", 170 + 300, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "画板", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.Option1_Click: Mouse.button = 0
    Page.DrawImage "dj" & IIf(Not Client.Picture2.Visible, "d", "") & ".png", 230 + 300, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "dj" & IIf(Not Client.Picture2.Visible, "d", "") & ".png", 230 + 300, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "DJ Mode", 0
    If m = mMouseDown And Mouse.button = 1 Then Client.fileServer: Mouse.button = 0
    
    
    Page.DrawImage "image.png", 692 + 300, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "image.png", 692 + 300, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    
    ECore.TipMsg "图片文字识别", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Client.OCR_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    Page.DrawImage "voice.png", 632 + 300, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "voice.png", 632 + 300, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    
    ECore.TipMsg "语音转文本", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Client.Audio_Click: ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    If Client.Picture2.Visible = False Then
        Page.DrawImage "save.png", 572 + 300, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "save.png", 572 + 300, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
        ECore.TipMsg "保存文件", 0
        If m = mMouseDown And Mouse.button = 1 Then
            If Client.Text1.Enabled Then Client.save3_Click
            If Client.Text5.Visible Then Client.save2_Click
            If Client.Picture1.Visible Then Client.saveDrawing
            Mouse.button = 0
        End If
    End If
    
SkipTab2:

    If Not disbledToolbar Then
        Page.DrawImage "tools.png", 300 + 752, GH - (40 - 6), Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "tools.png", 300 + 752, GH - (40 - 6), Alpha:=IIf(m, 1, 0.8), pos:=posOnCenter
        ECore.TipMsg "工具菜单", 0
        If m = mMouseDown And Mouse.button = 1 Then ProTime = GetTickCount: mode = IIf(mode = 0, 1, 0): Mouse.button = 0
    End If
    
    Page.DrawImage "back.jpg", 0, 0, 0, 0, 300, 60, 1
    Page.Paint 0, 0, 0, 301, 61, argb(100, 0, 0, 0)
    Page.Writes "MyChat · 学生端", 20, 15, 20, argb(160, 255, 255, 255)
    m = CheckMouse(255, 15, 30, 30)
    Page.Paint 1, 255, 15, 30, 30, argb(IIf(m, 255, 100), 180, 180, 180)
    Page.Writes "+", 260, 16, 18, argb(255, 0, 0, 0), style:=FontStyleBold
    If m = mMouseUp Then
        Dim grName As String
reinput:
        grName = InputBox("给新建的组起个名字吧~")
        If grName = "" Then MsgBox "名字不能为空。", 48: GoTo reinput
        Client.Winsock1.SendData "creategroup;" & Base64EncodeString(grName) & ";" & Base64EncodeString(userName) & vbCrLf
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
    If Mouse.x < 300 Then
        sideY = sideY - Direction * Depth * 100
    Else
        scrollY = scrollY - Direction * Depth * 100
    End If
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
    selectIndex = 1
    '创建页面
    ECore.Add Page, "MainPage"
End Sub
