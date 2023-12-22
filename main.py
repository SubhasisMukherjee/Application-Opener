import subprocess
import pygetwindow as gw
import os
import webbrowser
import time
import win32com.client



def find_latest_file_or_dir(path):
    max = 0
    latest = ""
    for pth in os.listdir(path):
        id = int(pth.split('.')[0])
        if id > max:
            max = id
            latest = pth

    return latest


# Task 1: Execute mongosh.exe
try:
    subprocess.Popen(['start', 'cmd', '/k', "mongosh.exe"], shell=True)
    time.sleep(2)   # to cater to the time delay needed to open cmd window
    cmd_window = gw.getActiveWindow()
    cmd_window.maximize()
    # # Create STARTUPINFO structure to configure the window size
    # startupinfo = subprocess.STARTUPINFO()
    # startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    # startupinfo.wShowWindow = ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 3)  # SW_MAXIMIZE

except FileNotFoundError:
    print("mongosh.exe not found. Please check the path or install MongoDB.")


# Task 2: Open the latest directory in File Explorer
directory_path = "D:\\Study\\Data Engineering\\Udemy\\MongoDB - The Complete Developer's Guide 2023\\"
latest_subdir = find_latest_file_or_dir(directory_path)
latest_subdirectory_path = directory_path + latest_subdir

try:
    subprocess.Popen("explorer " + latest_subdirectory_path)
    time.sleep(2)
    cmd_window = gw.getActiveWindow()
    cmd_window.maximize()

except FileNotFoundError:
    print(f"Directory not found: {latest_subdirectory_path}")


# Task 3: Open the latest doc inside the directory in MS Word
latest_file = find_latest_file_or_dir(latest_subdirectory_path)
latest_file_abs_path = latest_subdirectory_path + "\\" + latest_file

try:
    word_app = win32com.client.Dispatch("Word.Application")

    # Open the Word document
    doc = word_app.Documents.Open(latest_file_abs_path)

    # Make Word visible (optional)
    word_app.Visible = True


except Exception as e:
    print(f"Error: {e}")


# Task 4: Open a specific URL in Chrome
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

# Task 5: Open an excel spreadsheet in MS Excel
excel_file_path = r"C:\Users\Subhasis\Desktop\Udemy - MongoDB .xlsx"
try:
    # Open the file with excel
    subprocess.run(["start", "excel", excel_file_path], shell=True)
    excel_window = gw.getActiveWindow()
    excel_window.maximize()

except Exception as e:
    print(f"Error: {e}")


# Task 6: Open a text file in Notepad++
notepad_plus_plus_path = r"C:\Program Files\Notepad++\notepad++.exe"
file_to_open = r"C:/Users/Subhasis/Desktop/MongoDB  commands.txt"

try:
    subprocess.Popen([notepad_plus_plus_path, file_to_open])
except FileNotFoundError:
    print("Notepad++ not found. Please check the path to Notepad++ executable.")
except Exception as e:
    print(f"An error occurred: {e}")