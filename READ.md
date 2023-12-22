As part of a course, I had to do all the following on daily basis.

The main_basic.py file contains the below feeatures:
1) Execute mongosh.exe in command prompt
2) Open the directory with the notes for the latest course section in File Explorer
3) Open the notes for the latest lecture of the latest section in MS Word
4) Open the course URL in Chrome
5) Open the course tracker spreadsheet in MS Excel
6) Open the text file with all the commands in Notepad++

I automated the whole process using this script. 
I have created an executable of this script using pyinstaller python library. Just double clicking on the exe file does all the above for me.

pipreqs library has been used to generate the requirements.txt file. More of this library can be found at https://github.com/bndr/pipreqs


Enhancements available in main_advanced.py file::
1. Detecting the current lecture being marked "Complete" in the course tracker and opening a new word document fot the next lecture.
2. Detecting that all the lectures of the current section are marked "Complete" in the course tracker, hence creating a new directory for the next section, and creating a new word document for the first lecture of that section.
3. Opening MS Paint and Snipping Tool.