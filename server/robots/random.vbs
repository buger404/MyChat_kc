RobotName = "摇号机"
Guidence = "发送“/摇号”抽取组内成员"
Sub Process(msg,group,fromid)
	if msg = "/摇号" then
		count = Core.GetMemberCount(group)
		if count = 0 then
			Core.SendMessage group, "没有组员可以抽取！"
		else
			Core.SendMessage group, "恭喜" & Core.GetMemberName(group,int(rnd * count) + 1) & "被抽中！"
		end if
	end if
End Sub