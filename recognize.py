import speech_recognition as sr
import subprocess
import os
import pyautogui
import pyttsx3
import time
import numpy as np
import pyscreenshot as ImageGrab
import cv2
import pytesseract
import win32gui
#from win32gui import GetWindowText, GetForegroundWindow

# GetWindowText(GetForegroundWindow() # Get windows

screenWidth, screenHeight = pyautogui.size() # Get the size of the primary monitor.

def console(text):
	try:
		print(text.encode('utf-8').decode('cp1251'))
	except:
		print()

# Поиск индекса микрофона
# for index, name in enumerate(sr.Microphone.list_microphone_names()):
# 	print(f'Microphone with name "{name}" found for Microphone(device_index="{index}")')

def record_volume():
	r = sr.Recognizer()
	with sr.Microphone(device_index=1) as source:
		print('Настраиваюсь...')
		r.adjust_for_ambient_noise(source, duration=0.5) # Настройка посторонних шумов
		print('Слушаю...')
		audio = r.listen(source)
	print('Услышал.')
	try:
		query = r.recognize_google(audio, language='ru-RU')
		text = query.lower()
		print(f'Вы сказали: {text}')
		if 'пауза' in text:
			pyautogui.press('k')
			return
		if 'старт' in text:
			pyautogui.press('k')
			return
		if 'вперёд' in text:
			pyautogui.press('right')
			return
		if 'вперёд 2' in text:
			pyautogui.press('right')
			return
		if 'вперёд 3' in text:
			pyautogui.press(['right', 'right'])
			return
		if 'вперёд 4' in text:
			pyautogui.press(['right', 'right', 'right'])
			return
		if 'назад' in text:
			pyautogui.press('left')
			return
		if 'назад 2' in text:
			pyautogui.press('left')
			return
		if 'назад 3' in text:
			pyautogui.press(['left', 'left'])
			return
		if 'назад 4' in text:
			pyautogui.press(['left', 'left', 'left'])
			return
		if 'разверни видео' in text:
			pyautogui.press('f')
			return
		if 'сверни видео' in text:
			pyautogui.press('f')
			return
		if 'выключи звук' in text:
			pyautogui.press('m')
			return
		if 'включи звук' in text:
			pyautogui.press('m')
			return
		if 'включи стим' in text or 'включи steam' in text or 'включить steam' in text:
			subprocess.Popen(r"C:\Program Files (x86)\Steam\steam.exe")
			return
		if 'выключи стим' in text or 'выключи steam' in text or 'выключить steam' in text:
			os.system("TASKKILL /F /IM steam.exe")
			return
		if 'включи контру' in text or 'включи kant.ru' in text or 'включи control' in text:
			# http://programmerbook.ru/html/common-values/url/protocol/steam/
			# steam://rungameid/730
			subprocess.call(r"C:\Program Files (x86)\Steam\Steam.exe -applaunch 730")
			return
		if 'выключи контру' in text or 'выключи kant.ru' in text or 'выключи control' in text:
			os.system("TASKKILL /F /IM csgo.exe")
			return
		if 'разверни контру' in text or 'разверни kant.ru' in text or 'разверни control' in text:
			activate('counter-strike: global offensive')
			return
		if 'разверни браузер' in text:
			activate('chrome')
			return
		if 'включи браузер' in text:
			subprocess.call(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
			return
		if 'начни поиск' in text:
			activate('counter-strike: global offensive')
			time.sleep(0.1)
			pyautogui.moveTo(screenWidth/100*2.5, screenHeight/100*13.5)
			pyautogui.click()
			pyautogui.moveTo(screenWidth/100*85, screenHeight/100*95)
			time.sleep(0.1)
			pyautogui.click()
			# Никита
			# pyautogui.moveTo(725, 764)
			# time.sleep(0.1)
			# pyautogui.click()
			return
		if 'останови поиск' in text:
			activate('counter-strike: global offensive')
			time.sleep(0.1)
			pyautogui.moveTo(screenWidth/100*2.5, screenHeight/100*13.5)
			pyautogui.click()
			pyautogui.moveTo(screenWidth/100*85, screenHeight/100*95)
			time.sleep(0.1)
			pyautogui.click()
			# Никита
			# pyautogui.moveTo(725, 764)
			# time.sleep(0.1)
			# pyautogui.click()
			return
		if 'закрой браузер' in text:
			os.system("TASKKILL /F /IM chrome.exe")
			return
		if 'закрой' in text:
			pyautogui.moveTo(screenWidth-5, 5)
			pyautogui.click()
			return
		if 'выключи компьютер' in text:
			subprocess.call('shutdown /s')
			return
		if 'перезагрузи компьютер' in text:
			subprocess.call('shutdown /r')
			return
		if 'прими игру' in text:
			activate('counter-strike: global offensive')
			time.sleep(0.1)
			pyautogui.moveTo(screenWidth/2, screenHeight/100*55)
			pyautogui.click()
			return
		if 'выключи себя' in text:
			os.system("TASKKILL /F /IM recognize.exe")
			return
		if 'пригласи друзей' in text:
			activate('counter-strike: global offensive')
			time.sleep(0.1)
			pyautogui.moveTo(screenWidth-5, screenHeight-5)
			coords = get_coodrs(1650, 282, screenWidth-5, screenHeight-5)
			invite(coords)
			return
	except Exception as e:
		print(e)
		return

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

def activate(name):
	top_windows = []
	win32gui.EnumWindows(windowEnumerationHandler, top_windows)
	for i in top_windows:
		if name in i[1].lower():
			win32gui.ShowWindow(i[0],5)
			win32gui.SetForegroundWindow(i[0])
			break

def invite(coords):
	for i in coords:
		print(coords)
		pyautogui.moveTo(i[0], i[1])
		pyautogui.click()
		time.sleep(0.1)
		pyautogui.moveTo(i[0]-65, i[1]+20)
		pyautogui.click()
		time.sleep(0.1)

def get_coodrs(x1, y1, screenWidth, screenHeight):
	# pylint: disable=maybe-no-member
	filename = 'Image.png'
	screen = np.array(ImageGrab.grab(bbox=(x1, y1, screenWidth, screenHeight)))
	cv2.imshow('window', cv2.cvtColor(screen, cv2.COLOR_BGR2RGB))
	cv2.imwrite(filename, screen)
	cv2.destroyAllWindows()
	img = cv2.imread('Image.png')
	boxes = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
	result = []
	for i in range(len(boxes['text'])):
		word = boxes['text'][i].lower()
		print(word)
		if 'go' in word or '60' in word or '6o' in word or 'co' in word or 'g0' in word or 'cs' in word:
			result.append([boxes['left'][i] + x1, boxes['top'][i] + y1])
	return result

def say(text):
	tts = pyttsx3.init()
	rate = tts.getProperty('rate') #Скорость произношения
	tts.setProperty('rate', rate-40)

	volume = tts.getProperty('volume') #Громкость голоса
	tts.setProperty('volume', volume+0.9)

	voices = tts.getProperty('voices')

	# Задать голос по умолчанию
	tts.setProperty('voice', 'ru')

	# Попробовать установить предпочтительный голос
	for voice in voices:
		if voice.name == 'Aleksandr':
			tts.setProperty('voice', voice.id)

	tts.say(text)
	tts.runAndWait()

say('я родился')
while True:
	try:
		record_volume()
	except:
		continue