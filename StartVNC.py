'''Automatically signs into the VNC server, finds the desired window displayed 
through the server screen, and repositions VNC to show that window only.

After signing in to VNC, screenshots are taken to find the rectangular border 
of the window of interest displayed through the VNC viewer. When that 
rectangle is found, the VNC window is resized to fit only that rectangle of 
interest. The scroll buttons are clicked until the rectangle is centered 
correctly in view. Finally, the VNC window is moved to a better location.
'''

from PIL import Image, ImageFilter
#python -m pip install pillow
import pyscreenshot
#python -m pip install pyscreenshot
import win32gui, win32con, win32api
#python -m pip install pypiwin32
import time
import os
import sys

print("SET VNC TO USE 8-BIT COLOR")

VNC_WINDOW_TITLE = ""
VNC_VIEWER_EXE_PATH = ""
VNC_PASSWORD = "" #typeString() only works for lowercase passwords

if (VNC_WINDOW_TITLE == "") or (VNC_VIEWER_EXE_PATH == ""):
	print("Forgot to set constants in source.")
	sys.exit()

class Rectangle:
	'''Keeps track of rectangle coordinates and dimensions.'''
	def resetWidthHeight(self):
		'''Calculates the with and height attributes.'''
		self.width = self.right - self.left
		self.height = self.bottom - self.top
	
	def __init__(self, LTRBTuple):
		'''Initialization requires a 4-tuple (left, top, right, bottom).'''
		self.left = LTRBTuple[0]
		self.top = LTRBTuple[1]
		self.right = LTRBTuple[2]
		self.bottom = LTRBTuple[3]
		self.resetWidthHeight()
	
	def getLTRBTuple(self):
		'''Collects the left, top, right, and bottom values as a 4-tuple.'''
		return (self.left, self.top, self.right, self.bottom)
	
	def __str__(self):
		'''Hastily converts all attributes into a string.'''
		return str((self.left, self.top, self.right, self.bottom, 
					self.width, self.height))

hwnd = win32gui.FindWindow(None, VNC_WINDOW_TITLE)
if (hwnd != 0):
	print("VNC already open.")
	sys.exit()

os.startfile(VNC_VIEWER_EXE_PATH)
time.sleep(1)

def pressKey(virtualKeyCode):
	'''Presses key down and up.'''
	win32api.keybd_event(virtualKeyCode, 0, 1, 0)
	win32api.keybd_event(virtualKeyCode, 0, 2, 0)

def typeString(text):
	'''Presses alphabetic keys, ignoring capitalization.'''
	text = text.upper()
	for char in text:
		pressKey(ord(char))
		time.sleep(0.1)

def clickMouse():
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0)

pressKey(win32con.VK_RETURN)
time.sleep(1)
typeString(VNC_PASSWORD)
pressKey(win32con.VK_RETURN)
time.sleep(1)

hwnd = win32gui.FindWindow(None, VNC_WINDOW_TITLE)
if (hwnd == 0):
	print("Failed to find VNC window.")
	sys.exit()

#Collect window dimensions
windowRect = Rectangle(win32gui.GetWindowRect(hwnd))
print("Left:", windowRect.left, "Top:", windowRect.top, 
	  "Right:", windowRect.right, "Bottom:", windowRect.bottom)

#Reset windowRect to desired window coordinates
windowRect.left = windowRect.top = 0
windowRect.right = windowRect.width
windowRect.bottom = windowRect.height

win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, windowRect.left, 
					  windowRect.top, windowRect.width, windowRect.height, 0)

#Wait for screen to finish loading
#VNC viewer shows gray (204, 204, 204) while loading
i = None	#Stores the last screenshot of the window
r = g = b = 204
while (r == g == b == 204):
	time.sleep(0.1)
	i = pyscreenshot.grab(windowRect.getLTRBTuple(), False)
	(r,g,b) = i.getpixel((windowRect.width-30, windowRect.height-30))

def grayToBW(color):
	'''Converts grayscale colors to strict black or white.'''
	if color <= 128:
		return 0
	else:
		return 255

#Convert window screenshot to black or white pixels
i = i.convert("L")
i = i.point(grayToBW)

def findRectangleAtCoords(i, col, row, minRectWidth, minRectHeight):
	'''Searches a black/white screenshot for a black rectangle at (col, row).
	
	i	= window screenshot as an Image object, assumed black and white pixels
	col = int, column of the upper-left coordinate of the potential rectangle
	row = int, row of the upper-left coordinate of the potential rectangle
	minRectWidth 	= int, minimum width of a valid rectangle
	minRectHeight 	= int, minimum height of a valid rectangle
	
	Returns a Rectangle object describing the found black rectangle, or None 
	if none was found at position (col, row). A valid rectangle is one which 
	has non-broken lines of black pixels to complete the four borders, and 
	each border is long enough to satisfy minRectWidth and minRectHeight. 
	'''
	#Test if (col, row) looks like a black upper-left corner pixel
	if not ((i.getpixel((col, row)) == 0) 
	   and (i.getpixel((col-1, row)) == 255) 
	   and (i.getpixel((col, row-1)) == 255) 
	   and (i.getpixel((col+1, row)) == 0) 
	   and (i.getpixel((col, row+1)) == 0)):
		return None
	upperLeft = (col, row)
	#Find an upper-right corner, following the black pixels right
	upperRight = None
	tcol = col + 1
	while tcol < width:
		if i.getpixel((tcol, row)) == 255:
			#Right border found
			if tcol-1 != col:
				upperRight = (tcol-1, row)
			tcol = width
		tcol += 1
	if upperRight == None:
		return None
	rectWidth = upperRight[0] - col
	if rectWidth < minRectWidth:
		return None
	#Find a lower-left corner, following the black pixels down
	lowerLeft = None
	trow = row + 1
	while trow < height:
		if i.getpixel((col, trow)) == 255:
			#Bottom border found
			if trow-1 != row:
				lowerLeft = (col, trow-1)
			trow = height
		trow += 1
	if lowerLeft == None:
		return None
	rectHeight = lowerLeft[1] - row
	if rectHeight < minRectHeight:
		return None
	#Verify bottom border is black
	tcol = col + 1
	trow = lowerLeft[1]
	bottomBorderBlack = True
	while tcol <= upperRight[0]:
		if i.getpixel((tcol, trow)) == 255:
			bottomBorderBlack = False
			tcol = upperRight[0] + 1
		tcol += 1
	if not bottomBorderBlack:
		return None
	#Verify right border is black
	tcol = upperRight[0]
	trow = upperRight[1] + 1
	leftBorderBlack = True
	while trow <= lowerLeft[1]:
		if i.getpixel((tcol, trow)) == 255:
			leftBorderBlack = False
			trow = lowerLeft[1] + 1
		trow += 1
	if not leftBorderBlack:
		return None
	return Rectangle((col, row, upperRight[0], lowerLeft[1]))

#Search for the black rectangle that I want to see, testing every pixel
rectangle = None
minRectWidth = 100
minRectHeight = 100

(width, height) = i.size
for row in range(1, height-1):
	for col in range(1, width-1):
		rectangle = findRectangleAtCoords(i, col, row, 
										  minRectWidth, minRectHeight)
		if rectangle != None:
			break
	if rectangle != None:
		break
print("Rectangle:", rectangle)
if rectangle == None:
	sys.exit()

#Now resize and reposition window to show the rectangle
windowRect.right = rectangle.width + 8
windowRect.bottom = rectangle.height + 70
windowRect.resetWidthHeight()

win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, windowRect.left, 
					  windowRect.top, windowRect.width, windowRect.height, 0)

#Horizontal scroll to center the rectangle
lastKnownLeftBorderCol = rectangle.left
xClicksEstimation = int(rectangle.left/3)	#To prevent clicking forever
loopcount = 0
#Right scroll button
win32api.SetCursorPos((windowRect.width - 30, windowRect.height - 12))
leftBorderFound = True
#Click to scroll right until the left border of the rectangle is off the screen
while leftBorderFound and (loopcount < xClicksEstimation + 20):
	clickMouse()
	time.sleep(0.1)
	i = pyscreenshot.grab(windowRect.getLTRBTuple(), False)
	i = i.convert("L")
	i = i.point(grayToBW)
	leftBorderFound = False
	while (not leftBorderFound) and (lastKnownLeftBorderCol > 0):
		leftBorderFound = True
		for trow in range(rectangle.top, rectangle.top + minRectHeight):
			if i.getpixel((lastKnownLeftBorderCol, trow)) == 255:
				leftBorderFound = False
				break
		lastKnownLeftBorderCol -= 1
	loopcount += 1
#Click left once to see the left border again
win32api.SetCursorPos((12, windowRect.height - 12))
time.sleep(0.1)
clickMouse()
time.sleep(0.1)

#Vertical scroll to center rectangle
lastKnownTopBorderRow = rectangle.top
yClicksEstimation = int(rectangle.top/3.3)+1	#To prevent clicking forever
loopcount = 0
#Down scroll button
win32api.SetCursorPos((windowRect.width - 12, windowRect.height - 30))
topBorderFound = True
#Click to scroll down until the top border of the rectangle is off the screen
while topBorderFound and (loopcount < yClicksEstimation + 20):
	clickMouse()
	time.sleep(0.1)
	i = pyscreenshot.grab(windowRect.getLTRBTuple(), False)
	i = i.convert("L")
	i = i.point(grayToBW)
	topBorderFound = False
	while (not topBorderFound) and (lastKnownTopBorderRow > 26):
		topBorderFound = True
		for tcol in range(rectangle.left, rectangle.left + minRectWidth):
			if i.getpixel((tcol, lastKnownTopBorderRow)) == 255:
				topBorderFound = False
				break
		lastKnownTopBorderRow -= 1
	loopcount += 1
#Click up once or twice to see the border again
win32api.SetCursorPos((windowRect.width - 12, 55))
time.sleep(0.1)
clickMouse()
time.sleep(0.1)
clickMouse()
time.sleep(0.1)

#Consider adjusting for screen height later instead of repositioning with 
#absolute coordinates
#screenHeight = win32api.GetSystemMetrics(win32con.SM_CYFULLSCREEN)

win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 1864, 276, 
					  windowRect.width, windowRect.height, 0)
