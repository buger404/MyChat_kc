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
    Dim ProTime As Long, Mode As Integer
    Public selectIndex As Integer
    Dim scroll As ScrollArea, scrollY As Long, preY As Long, scrBottom As Boolean, sideY As Long
    Dim willPop As Object, bIndex As Integer, bTick As Long, mIndex As Integer
'======================================
Public Sub TransFile(id As Integer)
    Dim port As Long, sfile As String
    port = 1000 + Int(Rnd * 9000)
    Server.trans.ShowOpen
    sfile = Server.trans.filename
    If sfile <> "" Then
        Dim f As String
        f = Split(sfile, "\")(UBound(Split(sfile, "\")))
        If id = 0 Then
            For i = 1 To UBound(groups)
                If groups(i).id = group Then
                    ShellExecuteA 0, "open", App.path & "\FileTransportation.exe", "-o;" & Base64EncodeString(sfile) & ";" & str(UBound(groups(i).members)) & ";" & str(port), "", SW_SHOW
                    Exit For
                End If
            Next
        Else
            ShellExecuteA 0, "open", App.path & "\FileTransportation.exe", "-o;" & Base64EncodeString(sfile) & ";1;" & str(port), "", SW_SHOW
        End If
        If id = 0 Then
            For Each w In Server.Winsock
                If w.State = 7 Then w.SendData "filerecv;" & Base64EncodeString(f) & ";" & FileLen(sfile) & ";" & Base64EncodeString(userName) & ";" & Server.lis.LocalIP & ";" & port & vbCrLf
                DoEvents
            Next
        Else
            If Server.Winsock(id).State = 7 Then Server.Winsock(id).SendData "filerecv;" & Base64EncodeString(f) & ";" & FileLen(sfile) & ";" & Base64EncodeString(userName) & ";" & Server.lis.LocalIP & ";" & port & vbCrLf
        End If
    End If
End Sub
Public Sub Update()
    If Not willPop Is Nothing Then
        Server.PopupMenu willPop
        Set willPop = Nothing
        Mouse.button = 0
    End If
    
ReDraw:
    Page.Clear '清空画布
    '标题栏
    Dim m As MButtonState
    Page.DrawImage "back.jpg", 0, 0
    Page.Paint 0, 0, 0, 301, GH, argb(100, 0, 0, 0)
    Page.Paint 0, 300, 0, GW - 300, GH, argb(230, 255, 255, 255)
    'Page.Paint 0, 0, 0, GW, 61, argb(255, 89, 89, 89)
    m = CheckMouse(300, 0, GW - 300 - 250, 61)
    If m = mMouseDown Then
        '发送标题栏移动消息
        SendMessageA GHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
        ReleaseCapture
    End If
    
    If selectIndex > UBound(groups) Then selectIndex = UBound(groups)
    If Not groups(selectIndex).isJoin Then Mode = 1
    
    For i = 1 To UBound(groups)
        m = CheckMouse(0, (i - 1) * 70 + 60 - sideY, 301, 71)
        Page.Paint 0, 0, (i - 1) * 70 + 60 - sideY, 301, 71, argb(IIf(selectIndex = i, 60, IIf(m > 0, 30, 0)), 0, 0, 0)
        Page.Writes groups(i).Name, 30, (i - 1) * 70 + 60 + 5 - sideY, 20, argb(255, 255, 255, 255), style:=FontStyleBold
        Dim tip As String
        tip = groups(i).Msg(UBound(groups(i).Msg)).Name & ":" & groups(i).Msg(UBound(groups(i).Msg)).Content
        If Len(tip) > 12 Then tip = Left(tip, 10) & "..."
        If tip = ":" Then tip = "(暂无消息)"
        Page.Writes tip, 30, (i - 1) * 70 + 60 + 33 - sideY, 18, argb(200, 255, 255, 255)
        If m = 3 Then
            If Mouse.button = 1 Then
                selectIndex = i: scrBottom = True: Mode = 0
            ElseIf Mouse.button = 2 And i <> 1 Then
                MenuWindow.groupid = i: Mouse.button = 0
                Server.PopupMenu MenuWindow.groupMenu
            End If
        End If
        If groups(i).unreadTick > 0 Then
            Page.Paint 2, 230, (i - 1) * 70 + 60 + 20 - sideY, 50, 30, argb(255, 251, 77, 63), 100
            Page.Writes IIf(groups(i).unreadTick = 100, "99+", groups(i).unreadTick), 230 + 25, (i - 1) * 70 + 60 + 21 - sideY, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        End If
        If selectIndex = i And groups(i).unreadTick > 0 Then groups(i).unreadTick = 0: scrBottom = True
    Next

    If realSize <> 0 Then
        ReDim Preserve groups(realSize)
        realSize = 0
    End If

    Dim maxScr As Long
    If (i - 1) * 70 + 60 > GH - 30 Then
        maxScr = (i - 1) * 70 + 60 - GH + 30
    Else
        maxScr = 0
    End If
    If sideY > maxScr Then sideY = maxScr: GoTo ReDraw
    If sideY < 0 Then sideY = 0: GoTo ReDraw

    If selectIndex > UBound(groups) Then selectIndex = UBound(groups)
    
    Dim disbledToolbar As Boolean
    If Not groups(selectIndex).isJoin Then
        Page.Paint 2, 300 + (GW - 300) / 2, GH / 2, 450, 250, argb(120, 0, 0, 0), 20, pos:=posOnCenter
        Page.Writes "您尚未加入该讨论组", 300 + (GW - 300) / 2, GH / 2 - 70, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        Page.Writes "组长：" & groups(selectIndex).LeaderName, 300 + (GW - 300) / 2, GH / 2 - 30, 20, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        If Page.ShowColorButton(2, 300 + (GW - 300) / 2 - 100, GH - 250, 200, 40, "申请加入", argb(255, 255, 255, 255), argb(220, 84, 223, 143), argb(255, 84, 223, 143), 99, size:=20) = mMouseUp Then
            ' 申请加入的代码放置此处
        End If
        disbledToolbar = True
        GoTo nomsg
    End If

    If Mode = 3 Then
        Page.Writes "窗口黑名单", 20 + 300, 15, 20, argb(255, 0, 0, 0), style:=FontStyleBold
        If IsKeyPress(VK_CONTROL) And IsKeyPress(VK_MENU) And IsKeyPress(VK_Q) Then
            Dim hwnd As Long, Title As String * 255, Class As String * 255, imName As String * 255, imName2 As String, imN() As String
            hwnd = GetForegroundWindow
            GetWindowTextA hwnd, Title, 255
            GetClassNameA hwnd, Class, 255
            imName = GetProcessPath(hwnd)
            imN = Split(imName, "\")
            imName2 = LCase(imN(UBound(imN)))
            Title = LCase(Title): Class = LCase(Class)
            Dim tit As String, cl As String, im As String
            tit = InputBox("要添加的窗口的标题（可以使用通配符）", "窗口黑名单", UnSpace(Title))
            cl = InputBox("要添加的窗口的类名（可以使用通配符）", "窗口黑名单", UnSpace(Class))
            im = InputBox("要添加的窗口的镜像名（可以使用通配符）", "窗口黑名单", Replace(imName2, Chr(32), ""))
            AddBlack tit, cl, im
        End If
        Page.Writes "切换到要添加黑名单的窗口，然后按下Ctrl+Alt+Q", 20 + 300, 80, 18, argb(255, 0, 0, 0), style:=FontStyleUnderline
        For i = bIndex * 10 + 1 To bIndex * 10 + 10
            If i > UBound(nbf.item) Or i < 0 Then Exit For
            Page.Paint 0, 20 + 300, 120 + (i - bIndex * 10) * 30, 750, 31, argb(IIf(i Mod 2 = 0, 255, 0), 232, 232, 232)
            Page.Writes i & ". " & nbf.item(i).Title & ";" & nbf.item(i).Class & ";" & nbf.item(i).image, 20 + 300, 120 + (i - bIndex * 10) * 30, 18, argb(255, 0, 0, 0)
            If Page.ShowColorButton(0, 20 + 950, 120 + (i - bIndex * 10) * 30, 100, 31, "移除", argb(255, 255, 255, 255), argb(200, 239, 78, 47), argb(255, 239, 78, 47), size:=18) = mMouseUp Then
                Mouse.button = 0
                For j = i To UBound(nbf.item) - 1
                    nbf.item(j) = nbf.item(j + 1)
                Next
                ReDim Preserve nbf.item(UBound(nbf.item) - 1)
            End If
        Next
        Dim maxP As Integer
        maxP = Int(UBound(nbf.item) / 10) + IIf(UBound(nbf.item) - Int(UBound(nbf.item) / 10) * 10 <> 0, 1, 0)
        Page.Writes "第" & bIndex + 1 & "页，共" & maxP & "页", 300 + scroll.Width / 2, GH - 95, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter
        If Page.ShowColorButton(1, 20 + 350, GH - 100, 30, 30, "<", argb(255, 255, 255, 255), argb(200, 239, 78, 47), argb(255, 239, 78, 47), size:=18) = mMouseUp Then
            bIndex = bIndex - 1
        End If
        If Page.ShowColorButton(1, 20 + 950, GH - 100, 30, 30, ">", argb(255, 255, 255, 255), argb(200, 239, 78, 47), argb(255, 239, 78, 47), size:=18) = mMouseUp Then
            bIndex = bIndex + 1
        End If
        If Page.ShowColorButton(0, 20 + 350, GH - 50, 100, 30, "确定", argb(255, 255, 255, 255), argb(200, 239, 78, 47), argb(255, 239, 78, 47), size:=18) = mMouseUp Then
            Call BlackList.Dump
            Mode = 0
        End If
        If Page.ShowColorButton(0, 20 + 900, GH - 50, 100, 30, "取消", argb(255, 255, 255, 255), argb(200, 239, 78, 47), argb(255, 239, 78, 47), size:=18) = mMouseUp Then
            Mode = 0
        End If
        If bIndex < 0 Then bIndex = 0
        If bIndex > maxP - 1 Then bIndex = maxP - 1
    Else
        Page.Writes groups(selectIndex).Name & "(" & UBound(groups(selectIndex).members) & "人)", 20 + 300, 15, 20, argb(255, 0, 0, 0), style:=FontStyleBold
    End If
    If Mode = 4 Then
        Page.Writes "该组内的所有人", 60 + 300, 80, 18, argb(255, 0, 0, 0)
        Page.DrawImage "file.png", 20 + 300, 80, Alpha:=1
        ECore.TipMsg "传输文件", 0
        If CheckMouse2 >= mMouseDown And Mouse.button = 1 Then
            Mouse.button = 0
            TransFile 0
        End If
        For i = mIndex * 10 + 1 To mIndex * 10 + 10
            If i > UBound(groups(selectIndex).members) Or i < 0 Then Exit For
            Page.Paint 0, 20 + 300, 120 + (i - mIndex * 10) * 30, 750, 31, argb(IIf(i Mod 2 = 0, 255, 0), 232, 232, 232)
            Page.Writes i & ". " & groups(selectIndex).members(i).Name & "(#" & groups(selectIndex).members(i).id & ")", 100 + 300, 120 + (i - mIndex * 10) * 30, 18, argb(255, 0, 0, 0)
            If userId = -2 Or userId = groups(selectIndex).leader Then
                Page.DrawImage "menu.png", 20 + 300, 120 + (i - mIndex * 10) * 30, Alpha:=1
                ECore.TipMsg "管理", 0
                If CheckMouse2 >= mMouseDown And Mouse.button = 1 Then
                    MenuWindow.id = groups(selectIndex).members(i).id: Mouse.button = 0
                    MenuWindow.banMsg.Visible = True
                    MenuWindow.kickGroup.Visible = True
                    MenuWindow.copyMsg.Visible = False
                    Set willPop = MenuWindow.msgMenu
                End If
            End If
            Page.DrawImage "file.png", 60 + 300, 120 + (i - mIndex * 10) * 30, Alpha:=1
            ECore.TipMsg "传输文件", 0
            If CheckMouse2 >= mMouseDown And Mouse.button = 1 Then
                Mouse.button = 0
                TransFile groups(selectIndex).members(i).id
            End If
        Next
        maxP = Int(UBound(groups(selectIndex).members) / 10) + IIf(UBound(groups(selectIndex).members) - Int(UBound(groups(selectIndex).members)) * 10 <> 0, 1, 0)
        Page.Writes "第" & mIndex + 1 & "页，共" & maxP & "页", 300 + scroll.Width / 2, GH - 95, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter
        If Page.ShowColorButton(1, 20 + 350, GH - 100, 30, 30, "<", argb(255, 255, 255, 255), argb(200, 239, 78, 47), argb(255, 239, 78, 47), size:=18) = mMouseUp Then
            mIndex = mIndex - 1
        End If
        If Page.ShowColorButton(1, 20 + 950, GH - 100, 30, 30, ">", argb(255, 255, 255, 255), argb(200, 239, 78, 47), argb(255, 239, 78, 47), size:=18) = mMouseUp Then
            mIndex = mIndex + 1
        End If
        If mIndex < 0 Then mIndex = 0
        If mIndex > maxP - 1 Then mIndex = maxP - 1
    End If
    
    If Mode > 2 Then disbledToolbar = True: GoTo nomsg
    
    StartScrollArea Page, scroll
        Page.Clear argb(0, 0, 0, 0)
        Dim drx As Long, dry As Long, tw As Long, th As Long, ory As Long
        dry = 20 - scrollY: ory = dry
        For i = 1 To UBound(groups(selectIndex).Msg)
            With groups(selectIndex).Msg(i)
                drx = IIf(.id = userId, scroll.Width - 20, 20)
                
                tw = EF.GetWidth(Page.GG, .Content, 18, StringAlignmentNear, FontStyleRegular)
                th = EF.GetHeight(Page.GG, .Content, 18, StringAlignmentNear, FontStyleRegular)
                If tw > scroll.Width - 60 Then
                    th = th + (tw - scroll.Width + 60) / 18 / 0.75 - 3
                    tw = scroll.Width - 60
                End If

                If dry >= -(th + 40 + 20) And dry <= scroll.Height Then
                    If .id = -1 Then
                        Page.Paint 2, scroll.Width / 2, dry + 30, tw + 40, 40, argb(255, 212, 212, 212), 100, pos:=posOnCenter
                        Page.Writes .Content, scroll.Width / 2, dry + 20 - 4, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter
                        
                    Else
                        Page.Writes .Name & IIf(id > 0, "(#" & .id & ")", "") & "   " & format(.time, "hh:MM:ss"), drx, dry, 18, IIf(.id > 0, argb(255, 168, 168, 168), argb(255, 0, 176, 240)), align:=IIf(.id = userId, StringAlignmentFar, StringAlignmentNear)
                        Page.Writes .Name & IIf(id > 0, "(#" & .id & ")", "") & "   " & format(.time, "hh:MM:ss"), drx, dry, 18, IIf(.id > 0 And .id <> userId, argb(255, 168, 168, 168), argb(255, 84, 223, 143)), align:=IIf(.id = userId, StringAlignmentFar, StringAlignmentNear)
                        If .id = userId Then drx = drx - tw - 20
                        Page.Paint 2, drx, dry + 25, tw + 20, th + 20, IIf(.id = userId, argb(255, 84, 223, 143), argb(255, 232, 232, 232)), 16
                        
                        If CheckMouse(drx + 300, dry + 60 + 25, tw + 20, th + 20) > 0 And Mouse.button = 2 Then
                            MenuWindow.id = .id: selectMsg = groups(selectIndex).Msg(i): Mouse.button = 0
                            MenuWindow.copyMsg.Visible = True
                            MenuWindow.banMsg.Visible = (.id <> userId And (userId = groups(selectIndex).leader Or userId = -2))
                            MenuWindow.kickGroup.Visible = (.id <> userId And (userId = groups(selectIndex).leader Or userId = -2))
                            Set willPop = MenuWindow.msgMenu
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
        If bans(i).groupid = groups(selectIndex).id And bans(i).id = userId Then
            span = Now - bans(i).StartTime
            tickspan = Minute(span) * 60 + Second(span) + Hour(span) * 60 * 60
            If tickspan < bans(i).Duration Then
                Mode = 1
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
    Page.DrawImage "robot.png", GW - 135, 13, Alpha:=1
    If CheckMouse2 = mMouseUp Then
        Server.PopupMenu MenuWindow.robotMenu
    End If
    Page.DrawImage "black.png", GW - 175, 14, Alpha:=1
    If CheckMouse2 = mMouseUp Then
        Mode = 3
        nbf = bf
        'Server.PopupMenu MenuWindow.robotMenu
    End If
    Page.DrawImage "setting.png", GW - 215, 14, Alpha:=1
    If CheckMouse2 = mMouseUp Then
        Mode = 4
    End If
    
    '计算动画百分比
    Dim pro As Single, dpro As Single
    pro = (GetTickCount - ProTime) / 300
    If pro > 1 Then pro = 1
    If Mode = 1 Then pro = 1 - pro
    pro = Cubic(pro, 0, 1, 1, 1)
    dpro = 1 - pro
    
    '底部发送框
    Dim TextV As Boolean
    TextV = (pro = 1) And Mode = 0
    If Server.Text4.Visible <> TextV Then Server.Text4.Visible = TextV

    If Not disbledToolbar Then Page.Paint 0, 300, GH - 82, GW, 83, argb(255, 232, 232, 232)
    
    If pro = 0 Or disbledToolbar Or Mode > 2 Then GoTo SkipTab1
    
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
    If pro = 1 Or disbledToolbar Or Mode > 2 Then GoTo SkipTab2
    
    Page.DrawImage "disconnect.png", 300 + 50, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "disconnect.png", 300 + 50, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "断开客户端连接", 0
    If m = mMouseDown And Mouse.button = 1 Then Server.Text2.Text = InputBox("将断开的客户端的序号", "断开客户端连接"): Server.Command1_Click: Mouse.button = 0
    
    Page.DrawImage "ban.png", 300 + 110, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "ban.png", 300 + 110, GH - 40 * dpro + 40 * pro, Alpha:=IIf(Server.Command5.Enabled, 1, 0.3) * dpro, pos:=posOnCenter
    ECore.TipMsg "客户端禁言", 0
    If m = mMouseDown And Mouse.button = 1 And Server.Command5.Enabled Then Server.Command5_Click: Mouse.button = 0
    
    Page.DrawImage "image.png", 300 + 692, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "image.png", 300 + 692, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "图片文字识别", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Server.OCR_Click: ProTime = GetTickCount: Mode = IIf(Mode = 0, 1, 0): Mouse.button = 0
    Page.DrawImage "voice.png", 300 + 632, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "voice.png", 300 + 632, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "语音转文本", 0
    If m = mMouseDown And Mouse.button = 1 Then Mouse.button = 0: Server.Audio_Click: ProTime = GetTickCount: Mode = IIf(Mode = 0, 1, 0): Mouse.button = 0

    Page.DrawImage "save.png", 300 + 572, GH - 40 * dpro + 40 * pro, Alpha:=0, pos:=posOnCenter
    m = CheckMouse2
    Page.DrawImage "save.png", 300 + 572, GH - 40 * dpro + 40 * pro, Alpha:=1 * dpro, pos:=posOnCenter
    ECore.TipMsg "保存记录", 0
    If m = mMouseDown And Mouse.button = 1 Then
        Server.Command4_Click
        Mouse.button = 0
    End If

    
SkipTab2:
    If Not disbledToolbar Then
        Page.DrawImage "tools.png", 300 + 752, GH - (40 - 6), Alpha:=0, pos:=posOnCenter
        m = CheckMouse2
        Page.DrawImage "tools.png", 300 + 752, GH - (40 - 6), Alpha:=IIf(m, 1, 0.8), pos:=posOnCenter
        ECore.TipMsg "工具菜单", 0
        If m = mMouseDown And Mouse.button = 1 Then ProTime = GetTickCount: Mode = IIf(Mode = 0, 1, 0): Mouse.button = 0
    End If
    
    Page.DrawImage "back.jpg", 0, 0, 0, 0, 300, 60, 1
    Page.Paint 0, 0, 0, 301, 61, argb(100, 0, 0, 0)
    Page.Writes "MyChat · 教师端", 20, 15, 20, argb(160, 255, 255, 255)
    m = CheckMouse(255, 15, 30, 30)
    Page.Paint 1, 255, 15, 30, 30, argb(IIf(m, 255, 100), 180, 180, 180)
    Page.Writes "+", 260, 16, 18, argb(255, 0, 0, 0), style:=FontStyleBold
    If m = mMouseUp Then
        Dim grName As String
reinput:
        grName = InputBox("给新建的组起个名字吧~")
        If grName = "" Then MsgBox "名字不能为空。", 48: GoTo reinput
        Dim cmd() As String
        cmd = Split("creategroup;" & Base64EncodeString(grName) & ";" & Base64EncodeString(userName), ";")
        Server.ProcessCreateGroup cmd, userId
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
