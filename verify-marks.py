import pandas as pd
import sys
# web driver manager: https://github.com/SergeyPirogov/webdriver_manager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.edge.service import Service as EdgeService
# from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select

cols = {}
subject_ids = {}

class MyArgs:
	def __init__(self, args):
		self.file_name = args[1]
		self.subject = str(args[2]).strip().lower()
		self.info = {}
		if self.subject in subject_ids:
			self.info['subject'] = subject_ids[self.subject]
		else:
			raise Exception("Invalid subject passed as argument")

		try:
			file = open(args[3], "r")
			self.username = file.readline()
			self.password = file.readline()
			file.close()
		except:
			raise Exception("Credential file could not be opened")

		try:
			file = open(args[4], "r")
			lines = file.readlines()
			for line in lines:
				if len(line.strip()) == 0:
					break
				key_value = line.split(":")
				self.info[key_value[0].strip().lower()] = key_value[1].strip()
			file.close()
		except:
			raise Exception("Config file could not be opened")

 
def print_help():
	print("You need to provide the following arguments:")
	print("  - name of excel file")
	print("  - subject to upload")
 
def arg_parser():
	if len(sys.argv) == 2 and sys.argv[1] == '--help':
		print_help()
		return None
	if len(sys.argv) < 5:
		raise Exception("Run with --help")
	return MyArgs(sys.argv)
 
def get_df(args: MyArgs):
	xl = pd.ExcelFile(args.file_name)
	sheet_name = None
	for name in xl.sheet_names:
		if args.subject in name.lower():
			sheet_name = name
			break
	if sheet_name == None:
		raise Exception("Could not find proper sheet for subject")
	return pd.read_excel(args.file_name, sheet_name=sheet_name,header=None, skiprows=5).fillna(0)
 
def setup_driver():
	# driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
	driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
	return driver

def setup_cols():
	cols['english'] = 11
	cols['hindi'] = 11
	cols['mathematics'] = 6
	cols['environment studies'] = 5
	cols['other subjects'] = 4

def setup_subject_ids():
	#english
	subject_ids['english'] = 'English'
	#hindi
	subject_ids['hindi'] = 'Hindi'
	#maths
	subject_ids['mathematics'] = 'Mathematics'
	subject_ids['math'] = 'Mathematics'
	subject_ids['maths'] = 'Mathematics'
	#evs
	subject_ids['environment studies'] = 'Environment Studies'
	subject_ids['evs'] = 'Environment Studies'
	subject_ids['e.v.s.'] = 'Environment Studies'
	subject_ids['evs.'] = 'Environment Studies'
	#other subjects
	subject_ids['other subjects'] = 'Other Subjects'
	subject_ids['other'] = 'Other Subjects'
	subject_ids['others'] = 'Other Subjects'


def close_windows(driver, current_window):
	window_handles = driver.window_handles
	for handle in window_handles:
		if handle.lower() != current_window.lower():
			try:
				driver.switch_to.window(handle)
				driver.close()
			except:
				pass

def find_and_click_in_nav(driver, text: str):
	navigation_tree = WebDriverWait(driver, timeout = 10).until(lambda d: d.find_element(by=By.ID, value = "Treeview1"))
	tree_children = navigation_tree.find_elements(by=By.TAG_NAME, value = 'font')
	to_click = None
	for child in tree_children:
		if text.upper() in child.text.upper():
				to_click = child
				break
	if to_click == None:
		raise Exception(f'Could not locate {text} in navigation pane.')
	to_click.click()

def login(driver, args: MyArgs, user_id:str, pass_id: str):
	username_txt = driver.find_element(by=By.ID, value = user_id)
	username_txt.send_keys(args.username + Keys.TAB)
	password_txt = driver.find_element(by=By.ID, value = pass_id)
	password_txt.send_keys(args.password + Keys.ENTER)

def navigate_to_marks(driver, args: MyArgs):
	driver.get("https://aisg46.amizone.net/Amitranetg46/Webforms/Admin/Login.aspx")
	driver.implicitly_wait(0.5)
	# sign in
	current_window = driver.current_window_handle
	login(driver, args= args, user_id= "SignIn1_username", pass_id= "SignIn1_password")
	
	# navigate to aspect marks
	driver.implicitly_wait(10)
	close_windows(driver, current_window)
	driver.switch_to.window(current_window)
	driver.implicitly_wait(5)
	driver.switch_to.frame("contents")
	find_and_click_in_nav(driver, "Academics")
	find_and_click_in_nav(driver, "PrimaryReportCard")
	find_and_click_in_nav(driver, "EnterStudentAspectMarks")
	driver.switch_to.default_content()
	driver.switch_to.frame("main")

	login(driver, args= args, user_id= "txtuid", pass_id= "txtpassword")
	WebDriverWait(driver, timeout = 10).until(lambda d: d.find_element(by=By.ID, value = "ddlSession"))

def select_value(driver, id: str, value: str = None, index: int = None):
	select = Select(driver.find_element(by=By.ID, value = id))
	if value != None:
		select.select_by_visible_text(value)
	elif index != None:
		select.select_by_index(index)
	driver.implicitly_wait(1)

def make_selections(driver, args: MyArgs):
	select_value(driver, id = "ddlSession", value = args.info['session'])
	select_value(driver, id = "ddlclass", value = args.info['class'])
	select_value(driver, id = "ddlSection", value = args.info['section'])
	select_value(driver, id = "ddlExam", index= args.info['term'])
	select_value(driver, id = "ddlExamMonths", index = args.info['exam set'])
	select_value(driver, id = "ddlSubject", value= args.info['subject'])
	driver.find_element(by = By.ID, value = "btnView").click()
	WebDriverWait(driver, timeout = 10).until(lambda d: d.find_element(by=By.ID, value = "RepDetails_ctl01_lblStudent"))

def verify_marks_for(driver, args: MyArgs, student_row):
	mark_index = 1 + int(args.info['exam set'])
	text_box_str = "RepDetails_ctl{rollno}_datalistdisp_ctl{colno}_txtMarks"
	save_btn_str = "RepDetails_ctl{rollno}_imgbtnAddGrade"
	current_col = 0
	while current_col < cols[args.info['subject'].lower()]:
		text_box_id = text_box_str.format(rollno = str(student_row[0]).zfill(2), colno = str(current_col).zfill(2))
		marks_box = driver.find_element(by = By.ID, value = text_box_id)
		marks_online = float(marks_box.get_attribute('value'))
		if float(str(student_row[mark_index])) != marks_online:
			exam_set = args.info['exam set']
			print(f'- Mismatch for {student_row[1]}: replace {current_col+1} column for set {exam_set} with {marks_online}')
		mark_index = mark_index + 2
		current_col = current_col + 1
	# WebDriverWait(driver, timeout = 10).until(lambda d: d.find_element(by=By.ID, value = "lblMessage"))
	# label = driver.find_element(by=By.ID, value = "lblMessage")
	# label.text
	

def run():
	setup_cols()
	setup_subject_ids()
	myargs = arg_parser()
	if myargs == None:
		return
	df = get_df(myargs)
	driver = setup_driver()
	navigate_to_marks(driver, myargs)
	make_selections(driver, myargs)
	for _, row in df.iterrows():
		verify_marks_for(driver, args = myargs, student_row = row)
	_ = input("type anything to quit: ")
	driver.close()
	driver.quit()
 
run()
