import subprocess
import pygetwindow as gw
import os
import webbrowser
import time
import win32com.client



def get_latest(items):
    max = 0
    latest = ""
    for item in items:
        id = int(item.split('.')[0])
        if id > max:
            max = id
            latest = item.strip()

    return latest


def find_latest_file_or_dir(path, find_flag):
    if find_flag == "dir":
        l_dirs = [d for d in os.listdir(path) if os.path.isdir(path + "\\" + d)]
        latest_item = get_latest(l_dirs)
    elif find_flag == "file":
        l_files = [f for f in os.listdir(path) if os.path.isfile(path + "\\" + f)]
        latest_item = get_latest(l_files)

    return latest_item


## Task 1: Execute mongosh.exe
try:
    subprocess.Popen(['start', 'cmd', '/k', "mongosh.exe"], shell=True)
    time.sleep(2)   # to cater to the time delay needed to open cmd window
    cmd_window = gw.getActiveWindow()
    cmd_window.maximize()

except FileNotFoundError:
    print("mongosh.exe not found. Please check the path or install MongoDB.")


## Task 2: Open the directory for the latest section in File Explorer
dir_path = "D:\\Study\\Data Engineering\\Udemy\\MongoDB - The Complete Developer's Guide 2023\\"
latest_subdir = find_latest_file_or_dir(dir_path, 'dir')
latest_subdir_path = dir_path + latest_subdir

try:
    subprocess.Popen("explorer " + latest_subdir_path)
    time.sleep(2)
    cmd_window = gw.getActiveWindow()
    cmd_window.maximize()

except FileNotFoundError:
    print(f"Directory not found: {latest_subdir_path}")


## Task 3: Open the word document for the latest lecture in MS Word
latest_file = find_latest_file_or_dir(latest_subdir_path, 'file')
latest_file_abs_path = latest_subdir_path + "\\" + latest_file

try:
    word_app = win32com.client.Dispatch("Word.Application")

    # Open the Word document
    doc = word_app.Documents.Open(latest_file_abs_path)

    # Make Word visible (optional)
    word_app.Visible = True

except Exception as e:
    print(f"Error: {e}")


## Task 4: Open the course in Chrome
chrome_url = "https://www.udemy.com/course/mongodb-the-complete-developers-guide/"
webbrowser.open(chrome_url)

# # Check if the specific part of the URL is already open in Chrome
# url_is_open = False
# # for _ in range(3):  # Try a few times (adjust as needed)
# #    webbrowser.open(chrome_url)
# #    time.sleep(2)  # Wait for Chrome to open
# open_chrome_processes = [p.name() for p in psutil.process_iter(['name']) if 'chrome' in p.info['name'].lower()]
# if any(chrome_url in process for process in open_chrome_processes):
#     url_is_open = True



## Task 5: Open the course tracker spreadsheet in MS Excel
course_tracker_path = "D:\\Study\\Data Engineering\\Udemy\\MongoDB - The Complete Developer's Guide 2023\\Course Tracker.xlsx"
try:
    # Open the file with excel
    subprocess.run(["start", "excel", course_tracker_path], shell=True)
    excel_window = gw.getActiveWindow()
    excel_window.maximize()

except Exception as e:
    print(f"Error: {e}")


## Task 6: Open the text file with commands in Notepad++
notepad_plus_plus_path = "C:\\Program Files\\Notepad++\\notepad++.exe"
file_to_open = "D:\\Study\\Data Engineering\\Udemy\\MongoDB - The Complete Developer's Guide 2023\\MongoDB  commands.txt"

try:
    subprocess.Popen([notepad_plus_plus_path, file_to_open])
except FileNotFoundError:
    print("Notepad++ not found. Please check the path to Notepad++ executable.")
except Exception as e:
    print(f"An error occurred: {e}")