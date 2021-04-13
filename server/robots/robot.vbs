RobotName = "测试机器人"
Guidence = "帮助"
Sub Process(msg,group,fromid)
	Core.SendMessage group, "机器人收到了消息：" & msg

End Sub