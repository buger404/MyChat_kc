RobotName = "ҡ�Ż�"
Guidence = "���͡�/ҡ�š���ȡ���ڳ�Ա"
Sub Process(msg,group,fromid)
	if msg = "/ҡ��" then
		count = Core.GetMemberCount(group)
		if count = 0 then
			Core.SendMessage group, "û����Ա���Գ�ȡ��"
		else
			Core.SendMessage group, "��ϲ" & Core.GetMemberName(group,int(rnd * count) + 1) & "�����У�"
		end if
	end if
End Sub