# coding=utf-8
import socketserver
import base64
import os
import pymysql
import sys
import pyaudio
import wave
import requests 
import cv2
import json
import getopt

class TCPServer(socketserver.BaseRequestHandler):

	def handle(recData):
		#rec = .recv(1024).strip()
		totalData = ""
		while True:
			rec = str(recData.request.recv(1024), "utf-8")
			if rec == '': 
				break
			totalData = totalData + rec
			g = rec.split(":")
			if g[0] == 'num':
				break
			
		rec = totalData.split(":")
		cmd = rec[0]

		if cmd == 'num':
			num = rec[1]
			
			try:

				db = pymysql.connect('localhost','root','root:2020sql','face_db')

			except (Exception) as e:

				print(e)
				print('连接错误')

			cursor = db.cursor()
			sql = "SELECT * FROM face_tbl WHERE SEAT_NUM = %s" % (num)

			try:

				cursor.execute(sql)
				result = cursor.fetchall()

			except (Exception) as e:
				print(e)
				print('向数据库发送请求错误')
			
			try:

				i = result[0]
				stu = i[1]
				face_info = i[2]
				returnData = stu + ';' + face_info	

			except (Exception) as e:

				print(e)
				print('读取返回错误')

			try:

				print("{} wrote:".format(recData.client_address[0]) + str(num))
				recData.request.sendall(bytes(returnData, "utf-8"))

			except (Exception) as e:
				print(e)
				print('发送返回客户机的数据时错误')


			#返回人脸
		elif cmd == 'name':
			#注册人脸
			name = rec[1]
			baseData = rec[2]

			db = pymysql.connect('localhost','root','root:2020sql','face_db')   #后期vb写入sql密码
			cursor = db.cursor()
			sql = "INSERT INTO face_tbl(STU_NAME,BASE_DATA)VALUES ('" + name + "', '" + baseData + "')"
			
			print('注册' + name +'成功！')
			
			try:
				cursor.execute(sql)
				db.commit()
			except:
				db.rollback()
			db.close()
#处理返回信息


def imgRead(winname,path,b):
	#变量winname表示拍摄窗口标题
	#变量path表示希望保存的文件名路径
	#变量b表示是否返回图像
	cap = cv2.VideoCapture(0)
	print("Press (s)tart or (q)uit. ")
	a = True
	while True:
		ret,frame = cap.read()
		cv2.imshow(winname,frame)
		#初始化窗口

		if cv2.waitKey(30) == ord('s'):
			cv2.imwrite(path,frame)
			def get_file(filePath):
				with open(filePath, 'rb') as fp:
					return fp.read()
					#读取文件函数
			img = get_file(path)
			cap.release()
			cv2.destroyAllWindows()
			break

		if cv2.waitKey(30) == ord('q'):
			return a
			break

	if b == True and a == True:
		return img	
#调用摄像头保存图片

def OCR():
	from aip import AipOcr
	APP_ID = '17950303'
	API_KEY = 'QQGYgGXgQuydRyxP2cGMrG8n'
	SECRET_KEY = 'HsAS01NA35b41i7CI3RqQ9p6deX6Qef3'
	client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
	#引入百度ai组件

	image = imgRead('OCR','ocr_img.png',True)
	#调用imgRead函数捕捉图像
	if image != True:
		client.basicGeneral(image);
		req = client.basicGeneral(image);
		os.remove("ocr_img.png")
		all_text = ''
		for i in req['words_result']:
			all_text += i['words']
		print(all_text)
		ocrtext = open('OCR_text.txt','w')
		ocrtext.write('\n'+ all_text)
		ocrtext.close()
	else:
		ocrtext = open('OCR_text.txt','w')
		ocrtext.write('none')
		ocrtext.close()
#图片识别

def audio():
	from aip import AipSpeech
	APP_ID = '18808958'
	API_KEY = 'MEc3KEhONqzPMtf4F2CQ7a1s'
	SECRET_KEY = 'RBhmn9pFDYK8clG7GP8pyG3664vGsCKN'
	client = AipSpeech(APP_ID, API_KEY, SECRET_KEY)
	#引入百度ai组件

	CHUNK = 1024
	FORMAT = pyaudio.paInt16
	CHANNELS = 2
	RATE = 16000
	RECORD_TIME = 5
	p = pyaudio.PyAudio()
	stream = p.open(format=FORMAT,
		channels=CHANNELS,
		rate=RATE,
		input=True,
		frames_per_buffer=CHUNK)
	#录音参数设定	
	
	print("录音五秒....")
	frames = []
	for i in range(0, int(RATE / CHUNK * RECORD_TIME)):
		data = stream.read(CHUNK)
		frames.append(data)
	stream.stop_stream()
	stream.close()
	p.terminate()
	print("停了.")
	#暂存录音
	wav_file = 'audio.wav'
	wf = wave.open(wav_file, 'wb')
	wf.setnchannels(CHANNELS)
	wf.setsampwidth(p.get_sample_size(FORMAT))
	wf.setframerate(RATE)
	wf.writeframes(b''.join(frames))
	wf.close()	
	#储存录音过程

	pcm_file = "%s.pcm" %(wav_file.split(".")[0])
	os.system("ffmpeg -y  -i %s  -acodec pcm_s16le -f s16le -ac 1 -ar 16000 %s"%(wav_file,pcm_file))
	#wav转pcm格式

	with open(pcm_file, 'rb') as fp:
		file = fp.read()
	#读取pcm文件
	res = client.asr(file, 'pcm', 16000, {
		'dev_pid': 1537,
	})
	#使用百度ai组件获得语音转文字结果

	output = res.get("result")[0]
	print(output)
	audiotext = open('audio_text.txt','w')
	audiotext.write(output)
	audiotext.close()
	#输出结果

#语音识别

def removeFun():
	if os.path.exists('/id_info.txt'):
		os.remove('/id_info.txt')
	if os.path.exists('/audio.wav'):
		os.remove('/audio.wav')
	if os.path.exists('/audio.pcm'):
		os.remove('/audio.pcm')	
	if os.path.exists('/face.png'):
		os.remove('/face.png')
	if os.path.exists('/sign_img.png'):
		os.remove('/sign_img.png')
	if os.path.exists('/ocr_img.png'):
		os.remove('/ocr_img.png')

def main(argv):

	try:
		opts, args = getopt.getopt(argv,"o:ty",["ip"])
	except getopt.GetoptError:
		sys.exit(2)
		#定义命令行指令

	for opt, arg in opts:
		if opt in ("-o","--ip"):
			HOST = arg
			PORT = 80
			servername = open('server_info.txt','w')
			servername.write(HOST)
			servername.close()

			try:
				with socketserver.TCPServer((HOST, PORT), TCPServer) as server:
					print("已开启服务器..." + HOST)
					server.serve_forever()
		
			except KeyboardInterrupt:
				print("...已关闭")
		#定义开启服务器
		elif opt in ("-t"):
			removeFun()
			OCR()
			removeFun()
		#定义OCR指令
		elif opt in ("-y"):
			removeFun()
			audio()
			removeFun()
		#定义语音识别指令

#cmd唤醒程序

if __name__ == "__main__":
	main(sys.argv[1:])
#获取cmd指令