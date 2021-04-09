# coding=utf-8
import pyaudio
import wave
import requests 
import cv2
import json
import base64
import os
import socket
import sys
import getopt
#初始化

def imgRead(winname,path,b):
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
			a = False
			return a
			break

	if b == True and a == True:
		return img	
#调用摄像头保存图片

def sendOut(HOST,dataOut,a):
	try:
		with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
			sock.connect((HOST, 80))
			sock.sendall(bytes(dataOut, "utf-8"))
			if a == False:
				sock.close()
			elif a == True:
				totalData = ''
				while True:
					rec = str(sock.recv(10240), "utf-8")
					if not rec:
						break
					totalData = totalData + rec
				sock.close()
				return totalData
	except:
		print("送出错误")

def OCR():
	from aip import AipOcr
	APP_ID = '17950303'
	API_KEY = 'QQGYgGXgQuydRyxP2cGMrG8n'
	SECRET_KEY = 'HsAS01NA35b41i7CI3RqQ9p6deX6Qef3'
	client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
	#引入组件

	image = imgRead('OCR','ocr_img.png',True)
	if image != True:
		try:
			client.basicGeneral(image);
			req = client.basicGeneral(image);
			all_text = ''
			for i in req['words_result']:
				all_text += i['words']
			print(all_text)
			ocrtext = open('OCR_text.txt','w')
			ocrtext.write('\n'+ all_text)
			ocrtext.close()
		except Exception as e:
			ocrtext = open('OCR_text.txt','w')
			ocrtext.write('none')
			ocrtext.close()
			removeFun()
		
#图片识别

def audio():
	from aip import AipSpeech
	APP_ID = '18808958'
	API_KEY = 'MEc3KEhONqzPMtf4F2CQ7a1s'
	SECRET_KEY = 'RBhmn9pFDYK8clG7GP8pyG3664vGsCKN'
	client = AipSpeech(APP_ID, API_KEY, SECRET_KEY)
	#引入组件
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

	print("录音五秒钟嗷...")
	frames = []

	for i in range(0, int(RATE / CHUNK * RECORD_TIME)):
		data = stream.read(CHUNK)
		frames.append(data)

	print("停了.")
	stream.stop_stream()
	stream.close()
	p.terminate()

	wav_file = 'audio.wav'
	wf = wave.open(wav_file, 'wb')
	wf.setnchannels(CHANNELS)
	wf.setsampwidth(p.get_sample_size(FORMAT))
	wf.setframerate(RATE)
	wf.writeframes(b''.join(frames))
	wf.close()	

	pcm_file = "%s.pcm" %(wav_file.split(".")[0])
	os.system("ffmpeg -y  -i %s  -acodec pcm_s16le -f s16le -ac 1 -ar 16000 %s"%(wav_file,pcm_file))

	with open(pcm_file, 'rb') as fp:
		file = fp.read()

	res = client.asr(file, 'pcm', 16000, {
		'dev_pid': 1537,
	})
	try:
		output = res.get("result")[0]
		print(output)
		audiotext = open('audio_text.txt','w')
		audiotext.write(output)
		audiotext.close()
	except Exception as e:
		audiotext = open('audio_text.txt','w')
		audiotext.write("none")
		audiotext.close()
		removeFun()
	
#语音识别

def face(ip):

	host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=CGOxLTWnGMfLyBAxPsGXb7an&client_secret=fIjKS3L6uKSfhZZAdx2NZ6VWiwltdT7Z'
	response = requests.get(host)
	access_token = eval(response.text)['access_token']
	api = "https://aip.baidubce.com/rest/2.0/face/v3/match"+"?access_token="+access_token
	#获取百度云token
	
	HOST = ip
	def jsfunc(FaceId):
		with open('face.png','rb') as f:
			pic1 = base64.b64encode(f.read())
		try:

			send = 'num:' + str(FaceId)
			recvied = ''
			recvied = sendOut(HOST,send,True)
			i = recvied.split(';')
			pic2 = i[1]
			name = i[0]
			params = json.dumps([
				{"image":str(pic1, "utf-8"),"image_type":'BASE64',"face_type":"LIVE"},
				{"image":pic2,"image_type":'BASE64',"face_type":"LIVE"}
			])
			return [params,name]

		except(Exception) as e:

			print(e)
			err = True
			return err
	#写jsonの函数

	imgRead('FaceCheck','face.png',False)
	#保存图片

	faceid = 1
	#number = 3
	state = False

	while state == False:
		print('正在处理请稍候...')
		h = jsfunc(faceid)
		p = h[0]
		print(faceid)

		try:
			content = requests.post(api,p).text
			percentage = eval(content)['result']['score']
			print ("相似度： " + str(percentage))
			if percentage >= 90:
				idcheck = open('id_info.txt','w')
				idcheck.write(h[1])
				idcheck.close()
				state = True
				#break
			else:
				state = False
				faceid = faceid + 1
				
		except (Exception) as e :
			print(e)
			faceid = faceid + 1

	return state
#面部识别

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
	removeFun()

	try:
		opts, args = getopt.getopt(argv,"hoal:s:",["ipl=","ips="])
	except getopt.GetoptError:
		print('test1')
		sys.exit(2)
	for opt, arg in opts:
		if opt == '-h':
			print('-l:登陆 -s：注册 -o OCR -a 语音转文字')
			sys.exit()
		elif opt in ("-l","--ipl"):
			if os.path.exists('/face.png'):
				os.remove('/face.png')
			try:
				ipAdd = arg
				with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
					sock.connect((ipAdd, 80))
					sock.close()
			except(Exception) as e:
				idcheck = open('id_info.txt','w')
				idcheck.write("404")
				removeFun()
			face(ipAdd)
			removeFun()
		elif opt in ("-s","--ips"):
			
			removeFun()
			HOST = arg    
			name = input('输入您的名字：')

			imgRead("sign in","sign_img.png",False)
			with open('sign_img.png','rb') as f:
				img = base64.b64encode(f.read())

			send = 'name:' + name +':' + str(img,"utf-8")
			sendOut(HOST,send,False)
		
			removeFun()
		
		elif opt in ("-o"):
			removeFun()
			OCR()
			removeFun()

		elif opt in ("-a"):
			removeFun()			
			audio()
			removeFun()	


if __name__ == "__main__":
	main(sys.argv[1:])