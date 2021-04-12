VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RobotCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public currentRobot As String

Private Sub Class_Initialize()
    ReDim machine(0)
    Dim r As String
    r = Dir(App.path & "\robots\")
    Do While r <> ""
        ImportRobot r
        r = Dir()
    Loop
End Sub

Public Sub ImportRobot(file As String)
    Dim code As String, t As String
    Open App.path & "\robots\" & file For Input As #1
    Do While Not EOF(1)
        Line Input #1, t
        code = code & t & vbCrLf
    Loop
    Close #1
    ReDim Preserve machine(UBound(machine) + 1)
    Load Server.vbs(UBound(machine))
    Set machine(UBound(machine)) = Server.vbs(UBound(machine))
    With machine(UBound(machine))
        .AddCode code
        .AddObject "Core", Me, True
    End With
    Load MenuWindow.robotBtn(MenuWindow.robotBtn.UBound + 1)
    With MenuWindow.robotBtn(MenuWindow.robotBtn.UBound)
        .Caption = machine(UBound(machine)).Eval("RobotName")
        .Visible = True
    End With
End Sub

Public Function GetLeader(group)
    For i = 1 To UBound(groups)
        If groups(i).id = group Then
            GetLeader = groups(i).leader
            Exit For
        End If
    Next
End Function
Public Function GetMemberCount(group)
    For i = 1 To UBound(groups)
        If groups(i).id = group Then
            GetMemberCount = UBound(groups(i).members)
            Exit For
        End If
    Next
End Function
Public Function GetMemberId(group, index)
    For i = 1 To UBound(groups)
        If groups(i).id = group Then
            GetMemberId = groups(i).members(index).id
            Exit For
        End If
    Next
End Function
Public Function GetMemberName(group, index)
    For i = 1 To UBound(groups)
        If groups(i).id = group Then
            GetMemberName = groups(i).members(index).Name
            Exit For
        End If
    Next
End Function
Public Function GetMemberIndex(group, id)
    For i = 1 To UBound(groups)
        If groups(i).id = group Then
            For j = 1 To UBound(groups(i).members)
                If groups(i).members(j).id = id Then
                    GetMemberIndex = j
                    Exit Function
                End If
            Next
        End If
    Next
End Function
Public Sub SendMessage(group, content)
    AddMessage Val(group), -4, currentRobot, content
    For Each w In Server.Winsock
        If w.State = 7 Then w.SendData "msg;" & group & ";" & Base64EncodeString(currentRobot) & ";" & -4 & ";" & Base64EncodeString(content) & vbCrLf
    Next
End Sub
