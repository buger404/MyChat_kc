@echo off
set /p msg=�����������ύ���µ�����
git add .
git commit -m %msg%
git push --set-upstream origin master
echo ����ϴ�ʧ�ܣ����ʾ����Ҫ��ͬ������
pause