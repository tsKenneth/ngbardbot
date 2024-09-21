
import cv2
import numpy as np
from datetime import datetime
import win32com.client as comclt
import time
import random
import mss

skillbar_bb = {
	"y": 1025,
	"x": 625,
	"h": 55,
	"w": 670
	}

debuff_bb = {
	"y": 100,
	"x": 800,
	"h": 75,
	"w": 350
}

skill_table = {
	"consonantchain" : { "is_buff": False, "hotkey": "1", "icon":"icons/consonantchain.jpg", "cooldown":0, "duration": 30, "casttime":2 },
	"crashingchords": { "is_buff": False, "hotkey": "2", "icon":"icons/crashingchords.jpg", "cooldown":35, "duration": 0, "casttime":3 },
	"sonicboom": { "is_buff": False, "hotkey": "3", "icon":"icons/sonicboom.jpg", "cooldown":8.5, "duration": 0, "casttime":2 },
	"bellow": { "is_buff": False, "hotkey": "4", "icon":"icons/bellow.jpg", "cooldown":10, "duration": 0, "casttime":2 },
	"euphonicdirge": { "is_buff": False, "hotkey": "5", "icon":"icons/euphonicdirge.jpg", "cooldown":0, "duration": 18, "casttime":3 },
	"subvertedsymphony": { "is_buff": False, "hotkey": "6", "icon":"icons/subvertedsymphony.jpg", "cooldown":0, "duration": 24, "casttime":3 },
	"righteousrhapsody": { "is_buff": False, "hotkey": "7", "icon":"icons/righteousrhapsody.jpg", "cooldown":0, "duration": 45, "casttime":3 },
	"battlehymn": { "is_buff": True, "hotkey": "8", "icon":"icons/battlehymn.jpg", "cooldown":25, "duration": 30, "casttime":3 },
	"melodyofmana": { "is_buff": True, "hotkey": "9", "icon":"icons/melodyofmana.jpg", "cooldown":35, "duration": 45, "casttime":3 },
	"litanyoflife": { "is_buff": True, "hotkey": "0", "icon":"icons/litanyoflife.jpg", "cooldown":35, "duration": 45, "casttime":3 },
	"chromaticsonata":{ "is_buff": True, "hotkey": "-", "icon":"icons/chromaticsonata.jpg", "cooldown":35, "duration": 45, "casttime":3 },
	"militantcadence": { "is_buff": True, "hotkey": "=", "icon":"icons/militantcadence.jpg", "cooldown":60, "duration": 90, "casttime":3 }
}

location_table = {
	"town": {"screen":"screens/town.jpg"}
}

priority_list = ["battlehymn","melodyofmana","litanyoflife","chromaticsonata","militantcadence","righteousrhapsody","sonicboom","subvertedsymphony","bellow","euphonicdirge","consonantchain","crashingchords"]

def check_location(img):
	loc="Dungeon"
	loc_score=0.5

	for k,v in location_table.items():
		screen_img  = cv2.imread(v["screen"])
		screen_img = cv2.cvtColor(screen_img,cv2.COLOR_RGB2GRAY)
		res = cv2.matchTemplate(img, screen_img, cv2.TM_CCOEFF_NORMED)
		min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
		if max_val > loc_score:
			loc=k
			loc_score = max_val

	return loc

def check_available(img):
	img_skillbar = img[skillbar_bb['y']:skillbar_bb['y'] + skillbar_bb['h'],
				   skillbar_bb['x']:skillbar_bb['x'] + skillbar_bb['w']]

	availability_list = {}

	for k,v in skill_table.items():
		skill_icon_img  = cv2.imread(v["icon"])
		skill_icon_img = cv2.cvtColor(skill_icon_img,cv2.COLOR_RGB2GRAY)
		res = cv2.matchTemplate(img_skillbar, skill_icon_img, cv2.TM_CCOEFF_NORMED)
		min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
		availability_list[k] = max_val > 0.97

	return availability_list

def check_rr(img):
	img_debuff = img[debuff_bb['y']:debuff_bb['y'] + debuff_bb['h'],
				   debuff_bb['x']:debuff_bb['x'] + debuff_bb['w']]

	debuff_icon_img  = cv2.imread("icons/righteousrhapsodydebuff.jpg")
	debuff_icon_img = cv2.cvtColor(debuff_icon_img, cv2.COLOR_RGB2GRAY)
	res = cv2.matchTemplate(img_debuff, debuff_icon_img, cv2.TM_CCOEFF_NORMED)
	min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

	return max_val > 0.75

time_casted = { k:None for k in skill_table.keys() }

time.sleep(5)
wsh = comclt.Dispatch("WScript.Shell")
wsh.AppActivate("Nevergrind Online")

delay_list = [0.1, 0.121, 0.114, 0.112, 0.119, 0.106]

while True:
	now = datetime.now()
	delay=random.choice(delay_list)

	# Get Timing
	time_elapsed_since_cast = { k:(None if time_casted[k] is None else (now-time_casted[k]).total_seconds()) for k in time_casted.keys()}
	duration_expired = { k:(True if time_elapsed_since_cast[k] is None else skill_table[k]["duration"] - skill_table[k]["casttime"] - delay - v <= 0) for k,v in time_elapsed_since_cast.items()}
	# skill_ready = { k:(True if time_elapsed_since_cast[k] is None else skill_table[k]["cooldown"] - v <= 0) for k,v in time_elapsed_since_cast.items()}

	# Get raw pixels from the screen, save it to a Numpy array
	sct = mss.mss()
	screenshot = cv2.cvtColor(np.array(sct.grab(sct.monitors[1])), cv2.COLOR_RGB2GRAY)

	# If location is town, sleep and skip
	if check_location(screenshot) == "town":
		print("Player is in town")
		time.sleep(5)
		continue

	skill_ready = check_available(screenshot)
	rr_active = check_rr(screenshot)

	for skill in priority_list:
		if skill == "righteousrhapsody":
			if not rr_active:
				print("Casting " + skill)
				wsh.SendKeys(skill_table[skill]["hotkey"])
				time_casted[skill] = datetime.now()
				time.sleep(skill_table[skill]["casttime"] + delay)
				break
		elif skill_ready[skill] and duration_expired[skill]:
			print("Casting " + skill)
			wsh.SendKeys(skill_table[skill]["hotkey"])
			time_casted[skill] = datetime.now()
			time.sleep(skill_table[skill]["casttime"] + delay)
			break

# sample  = cv2.imread("sample.jpg",  0)
