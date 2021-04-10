Attribute VB_Name = "GroupCore"
Public Type Messages
    id As Integer
    Name As String
    Content As String
    time As Date
End Type
Public Type group
    id As Integer
    leader As Integer
    isJoin As Boolean
    Name As String
    Msg() As Messages
<<<<<<< HEAD
=======
    unreadTick As Integer
>>>>>>> 53e2e550e623ad8f9cb82be9bf2b880d4cb697f9
End Type
Public Type MsgBan
    id As Integer
    groupid As Integer
    StartTime As Date
    Duration As Long
End Type
Public userId As Integer, userName As String
Public groups() As group, bans() As MsgBan

Public Sub AddGroup(id As Integer, leader As Integer, isJoin As Boolean, Name As String)
    ReDim Preserve groups(UBound(groups) + 1)
    With groups(UBound(groups))
        .id = id
        .isJoin = isJoin
        .Name = Name
        .leader = leader
        ReDim .Msg(0)
    End With
End Sub
Public Sub DeleteGroup(id As Integer)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            For j = i To UBound(groups) - 1
                groups(j) = groups(j + 1)
            Next
            ReDim Preserve groups(UBound(groups) - 1)
            Exit For
        End If
    Next
End Sub
Public Sub AddMessage(id As Integer, memberid As Integer, Name As String, Content As String)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            ReDim Preserve groups(i).Msg(UBound(groups(i).Msg) + 1)
            With groups(i).Msg(UBound(groups(i).Msg))
                .Content = Content
                .id = memberid
                .Name = Name
                .time = Now
            End With
            groups(i).unreadTick = groups(i).unreadTick + 1
            If groups(i).unreadTick > 100 Then groups(i).unreadTick = 100
            Exit For
        End If
    Next
End Sub
Public Sub SetJoinState(id As Integer, isJoin As Boolean)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            groups(i).isJoin = isJoin
            Exit For
        End If
    Next
End Sub
Public Sub AddBan(id As Integer, group As Integer, Duration As Long)
    ReDim Preserve bans(UBound(bans) + 1)
    With bans(UBound(bans))
        .id = id
        .groupid = group
        .StartTime = Now
        .Duration = Duration
    End With
End Sub
