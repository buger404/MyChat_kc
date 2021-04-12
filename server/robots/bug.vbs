RobotName = "BUG"
Guidence = "·¢ËÍ¡°/BUG¡±´¥·¢bug"
Sub Process(msg,group,fromid)
	if msg = "/BUG" then
		Core.SendMessage 1 / 0
	end if
End Sub