VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim IP As String
    Dim Mode As Integer
    Public FinalIP As String
'======================================

Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制i代码请放在此处）
    
    Page.Clear '清空画布

    If Mode = 1 Then
        Page.ShowLoading GW / 2 - 150, GH / 2 - 150, 300, 300, 5, argb(255, 0, 176, 240), argb(255, 255, 51, 102), argb(255, 255, 214, 100)
        'ECore.NewTransform transToLeft, 500, "MainPage"
        Exit Sub
    End If

    m = CheckMouse(GW - 45, 15, 30, 30)
    Page.Paint 1, GW - 45, 15, 30, 30, argb(IIf(m, 255, 100), 255, 83, 57)
    Page.Writes "×", GW - 41, 16, 18, argb(255, 255, 255, 255), style:=FontStyleBold
    If m = mMouseUp Then Unload Client: End
    
    Page.Writes "输入服务端使用的本地IP地址以开始使用", GW / 2, GH / 2 - 80, 24, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    
    Page.ShowEdit IP, 2, GW / 2 - 200, GH / 2 + 10, 400, 50, argb(255, 0, 0, 0), argb(255, 232, 232, 232), argb(255, 222, 222, 222), argb(255, 0, 176, 240), 99, size:=24

    If UBound(Split(IP, ".")) = 3 Then
        If Page.ShowColorButton(2, GW / 2 - 100, GH - 100, 200, 50, "开始使用", argb(255, 255, 255, 255), argb(220, 0, 176, 240), argb(255, 0, 176, 240), 99, size:=24) = mMouseUp Then
            FinalIP = IP
            ECore.NewTransform
            Mode = 1
        End If
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
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
    Set Page = New GPage  '创建页面绘制器
    Page.Create Me  '将绘制器与控制器绑定
    Page.Res.NewImages App.path & "\assets"  '导入资源
    ECore.Add Page, "IPPage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
