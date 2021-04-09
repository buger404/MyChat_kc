@echo off
set /p msg=请描述本次提交更新的内容
git add .
git commit -m %msg%
git push --set-upstream origin master
echo 如果上传失败，则表示你需要先同步代码
pause