import os
import pyautogui
import cv2
import pytesseract
import time
import pyperclip
import win32api

def join(dire, name):
	return os.path.join(dire,name)

def caltime(fun):
	start = time.time()
	fun()
	return time.time() - start

class Window:
	folder = os.path.dirname(__file__)
	region_fld = join(folder, 'Region')
	screenshot_fld = join(folder, 'Screenshots')
	tesseract_fld = join(folder, 'Tesseract')
	pytesseract.pytesseract.tesseract_cmd = join(tesseract_fld,'tesseract.exe')
	tools = {'Brush':['alt','h','b','enter'], 'Text':['alt','h','t'], 'Select':['alt','h','s','e','r']}


	def __init__(self, ribbon,statusbar):
		self.ribbon = ribbon
		self.current_tool = 'Brush'
		self.tool_region = {}
		self.ribbon_is_open = self.check_ribbon()
		self.reset_ribbon()
		self.statusbar = statusbar
		self.reset_status_bar()



	def reset_status_bar(self):
		if self.statusbar:
			self.open_status_bar()
		else:
			self.close_status_bar()

	def close_status_bar(self):
		time.sleep(1)
		pos = pyautogui.locateOnScreen(join(self.region_fld, 'ZoomMinus.PNG'))
		if pos:
			pyautogui.press(['alt','v','s'])	

	def open_status_bar(self):
		time.sleep(1)
		pos = pyautogui.locateOnScreen(join(self.region_fld, 'ZoomMinus.PNG'))
		print(pos)
		if not pos:	
			pyautogui.press(['alt','v','s'])

	def get_font(self):
		temp = self.current_tool
		self.change_tool('Text')
		time.sleep(1)
		pyautogui.moveTo(self.left, self.top)
		pyautogui.click(clicks=2)
		pyautogui.press(['alt', 't','f','f'])
		pyautogui.hotkey('ctrl', 'c')
		time.sleep(0.5)
		self.police = pyperclip.paste()
		if pyautogui.locateOnScreen(join(self.region_fld, 'Bold.PNG')):
			self.bold = True
		else:
			self.bold = False

		if pyautogui.locateOnScreen(join(self.region_fld, 'Italic.PNG')):
			self.italic = True
		else:
			self.italic = False

		if pyautogui.locateOnScreen(join(self.region_fld, 'Underline.PNG')):
			self.underline = True
		else:
			self.underline = False

		if pyautogui.locateOnScreen(join(self.region_fld, 'Strikethrough.PNG')):
			self.strikethrough = True
		else:
			self.strikethrough = False
		pyautogui.press('enter')
		pyautogui.press(['alt', 't','f','s'])
		pyautogui.hotkey('ctrl','c')
		time.sleep(0.5)
		self.size = int(pyperclip.paste())
		pyautogui.press('enter')
		if temp == 'Text':
			pyautogui.press(['alt','h','t'])
		else:
			self.change_tool(temp)
		
	def check_ribbon(self):
		is_close = pyautogui.locateOnScreen(join(self.region_fld,'RibbonOpen.PNG'))
		if is_close:
			self.tool_region['Ribbon'] = is_close
			return False
		is_open = pyautogui.locateOnScreen(join(self.region_fld, 'RibbonClose.PNG'))
		if is_open:
			self.tool_region['Ribbon'] = is_open
			return True

	def reset_ribbon(self):
		if not self.ribbon:
			self.close_ribbon()

		else:
			self.open_ribbon()

	def close_ribbon(self):
		if self.ribbon_is_open:
			pyautogui.hotkey('ctrl', 'f1')
		self.ribbon_is_open = False

	def open_ribbon(self):
		if not self.ribbon_is_open:
			pyautogui.hotkey('ctrl', 'f1')
		self.ribbon_is_open = True

	def resize(self):
		def read_size(size_region):
			location = join(self.screenshot_fld, 'size.png')
			pyautogui.screenshot(location, region=size_region)
			img = cv2.imread(location)
			txt = pytesseract.image_to_string(img)
			txt = txt.split(' ')
			print((txt[0],txt[-1]))
			x = int(txt[0])
			y = int(txt[-1])
			return x,y
		time.sleep(2) #remove later
		self.open_status_bar()
		time.sleep(0.3)
		pyautogui.moveTo(pyautogui.locateCenterOnScreen(join(self.region_fld, 'ZoomMinus.PNG')))
		for _ in range(3):
			pyautogui.click()

		limitx = pyautogui.locateCenterOnScreen(join(self.region_fld,'LimitX.PNG'))
		limity = pyautogui.locateCenterOnScreen(join(self.region_fld,'LimitY.PNG'))
		while True:
			if limitx:
				pyautogui.click(x=limitx[0], y=limitx[1])
			if limity:
				pyautogui.click(x=limity[0], y=limity[1])
			time.sleep(0.75)
			corner_pos = pyautogui.locateCenterOnScreen(join(self.region_fld, 'Corner.PNG'))
			if corner_pos:
				break

		pyautogui.moveTo(corner_pos)

		


		while True:
			size_start = pyautogui.locateOnScreen(join(self.region_fld, 'SizeStart.PNG'))
			size_end = list(pyautogui.locateAllOnScreen(join(self.region_fld, 'SizeEnd.PNG')))[-1]
			x,y = read_size((size_start.left + size_start.width, size_start.top, size_end.left - size_start.left - size_start.width, size_start.height))
			if x <= 1000 and y <= 500:
				break
			if x > 900:
				pyautogui.drag(-25,0,1,button='left')
			if y > 400:
				pyautogui.drag(0,-25,1,button='left')



		pyautogui.moveTo(pyautogui.locateCenterOnScreen(join(self.region_fld, 'ZoomPlus.PNG')))
		for  i in range(3):
			pyautogui.click()
		pyautogui.moveTo(pyautogui.locateCenterOnScreen(join(self.region_fld, 'Corner.PNG')))

		side = pyautogui.locateOnScreen(join(self.region_fld, 'Top.PNG'))
		pyautogui.drag(side.left - pyautogui.position()[0], 0, 3, button='left')
		while True:
			if pyautogui.locateOnScreen(join(self.region_fld, 'SideScroll.PNG')):
				pyautogui.drag(-1,0,1,button='left')
				break
			pyautogui.drag(1,0,1,button='left')
		bottom = pyautogui.locateOnScreen(join(self.region_fld, 'Bottom.PNG'))
		pyautogui.drag(0,bottom.top - pyautogui.position()[1], 2, button='left')
		self.reset_status_bar()
		while True:
			if pyautogui.locateOnScreen(join(self.region_fld, 'Scroll.PNG')):
				limitx = pyautogui.locateCenterOnScreen(join(self.region_fld,'LimitX.PNG'))
				limity = pyautogui.locateCenterOnScreen(join(self.region_fld,'LimitY.PNG'))
				pyautogui.click(x=limitx[0], y=limitx[1])
				pyautogui.click(x=limity[0], y=limity[1])
				time.sleep(1)
				pyautogui.moveTo(pyautogui.locateCenterOnScreen(join(self.region_fld, 'Corner.PNG')))
				pyautogui.drag(0,-1,1,button='left')
				break
			pyautogui.drag(0,1,1,button='left')

	def get_sizes(self):
		top_left = pyautogui.locateCenterOnScreen(join(self.region_fld, 'TopLeft.PNG'))
		top_right = pyautogui.locateCenterOnScreen(join(self.region_fld, 'TopRight.PNG'))
		bottom_right = pyautogui.locateCenterOnScreen(join(self.region_fld, 'Corner.PNG'))

		self.left, self.top, self.width, self.height = top_left[0], top_left[1], top_right[0] - top_left[0], bottom_right[1] - top_left[1] - 1

	def clear_screen(self):
		pyautogui.hotkey('ctrl', 'a')
		pyautogui.press('delete')
		self.current_tool = 'Select'

	def change_tool(self, tool):
		if self.current_tool != tool:
			pyautogui.press(self.tools[tool])
		self.current_tool = tool

class Line():
	def __init__(self,win,xy,xy2):
		self.start = win.left + xy[0], win.top + xy[1]
		self.end = win.left + xy2[0], win.top + xy2[1]
		self.win = win
	def draw(self):
		self.win.change_tool('Brush')
		pyautogui.moveTo(self.start[0], self.start[1])
		pyautogui.dragTo(self.end[0], self.end[1], button='left')

class Rect():
	def __init__(self,win,xy,xy2):
		self.win = win
		self.pos1 = win.left + xy[0], win.top + xy[1]
		self.pos2 =  win.left + xy2[0]
		self.pos3 =  win.top + xy2[1]
		self.pos4 =  win.left + xy[0]
		self.pos5 =  win.top + xy[1]
		self.left , self.top, self.width, self.height = win.left + xy[0], win.top + xy[1], xy2[0] - xy[0], xy2[1] - xy[1]

	def draw(self):
		self.win.change_tool('Brush')
		pyautogui.moveTo(self.pos1[0], self.pos1[1])
		pyautogui.dragTo(self.pos2, None, button='left')
		pyautogui.dragTo(None, self.pos3, button='left')
		pyautogui.dragTo(self.pos4, None, button='left')
		pyautogui.dragTo(None, self.pos5, button='left')

def draw_circle():
	pass

class Text():
	def __init__(self, win, xy, text, police, size, isbold, isitalic, isunderline, isstrickethrough, length):
		self.win = win
		self.xy = xy
		self.text = text
		self.xy = win.left + xy[0], win.top + xy[1]
		self.police = police
		self.size = size
		self.isbold = isbold
		self.isitalic = isitalic
		self.isunderline = isunderline
		self.isstrickethrough = isstrickethrough
		self.length = length

	def draw(self):
		self.win.change_tool('Text')
		pyautogui.moveTo(self.xy[0], self.xy[1])

		if self.length != 0:
			pyautogui.drag(self.length, 0, 1, button='left')
		else:
			pyautogui.click(clicks=2)
			
		if self.win.police != self.police:
			time.sleep(1)
			pyautogui.press(['alt', 't', 'f','f'])
			pyautogui.write(self.police)
			pyautogui.press('enter')
			self.win.police = self.police
			

		if self.win.size != self.size:
			time.sleep(1)
			pyautogui.press(['alt', 't', 'f','s'])
			pyautogui.write(str(self.size))
			pyautogui.press('enter')
			self.win.size = self.size

		if self.win.bold != self.isbold:
			pyautogui.hotkey('ctrl','b')
			self.win.bold = self.isbold

		if self.win.italic != self.isitalic:
			pyautogui.hotkey('ctrl','i')
			self.win.italic = self.isitalic

		if self.win.underline != self.isunderline:
			pyautogui.hotkey('ctrl','u')
			self.win.underline = self.isunderline

		if self.win.strikethrough != self.isstrickethrough:
			pyautogui.press(['alt','t','f','x'])
			self.win.strikethrough = self.isstrickethrough
		pyautogui.write(self.text)
		pyautogui.press(['alt','h','t'])

class Button():
	def __init__(self, rect, text, func):
		self.rect = rect
		self.text = text
		self.state_left = win32api.GetKeyState(0x01)
		self.func = func

	def draw(self):
		self.rect.draw()
		time.sleep(2)
		self.text.draw()

	def update(self):
		a = win32api.GetKeyState(0x01)
		if a != self.state_left:
			self.state_left = a
			pos = pyautogui.position()
			if pos[0] > self.rect.left and pos[0] < self.rect.left + self.rect.width and pos[1] > self.rect.top and pos[1] < self.rect.top + self.rect.height:
				if a < 0:
					self.func()

class Entry():
	def __init__(self, rect):
		self.rect = rect

	def draw(self):
		self.rect.draw()

	def get_text(self):
		location = join(self.rect.win.screenshot_fld, 'Entry.PNG')
		pyautogui.screenshot(location, region=(self.rect.left, self.rect.top, self.rect.width, self.rect.height))
		img = cv2.imread(location)
		txt = pytesseract.image_to_string(img)
		print(txt)


def hi():
	global w, x
	w.clear_screen()
	del x
	e = Entry(Rect(w,(50,50),(150,150)))
	e.draw()
	input()
	e.get_text()
	w.clear_screen()
	for i in range(0,w.width,30):
		e = Line(w,(i,0), (i, w.height))
		e.draw()
	for i in range(0,w.height, 30):
		e= Line(w,(0,i),(w.width, i))
		e.draw()

if __name__ == '__main__':
	w = Window(True, False)

	#w.resize()
	w.get_sizes()
	w.get_font()
	w.clear_screen()
	time.sleep(1)
	x= Button(Rect(w,(50,50), (150,150)), Text(w,(100,100), 'hi', 'Arial', 30, False,False,False,False, 500), hi)
	x.draw()
	print('starting')
	while True:
		try:
			x.update()
		except:
			break
	#print(caltime(w.clear_screen))


	#l = Line(w, (20,50), (100, 200))
	#r = Rect(w, (20,50), (100, 200))

	#l.draw()
	#r.draw()

	#w.resize()
