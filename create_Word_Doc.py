#-- Comments start with a hash sign   version 6/15(b)/2025
# I didn't intend for this Py code to be my first Git-tracked one, but things were such a mess
#  ... that I didn't know what I had in this file
# It turns out to be a test py script for creating a Word Document using the pywin32 MODULE
# which then sends VBA commands to the Word doc to drive it with "wd.Selection.<method>" object calls
# The original code was included in WiseOwl Lesson "17b" which is much more complex
# WiseOwl:  Python Part 17b - VBA using pywin32
# https://www.youtube.com/watch?v=PiHm9k0gd1M&t=1339s&pp=ygVEQWxleCBweSBzY3JpcHQgZm9yIGNyZWF0aW5nIGEgV29yZCBEb2N1bWVudCB1c2luZyB0aGUgcHl3aW4zMiBNT0RVTEU%3D

# More research:
#   PyCharm Version Control (VCS):
#   https://www.google.com/search?q=pycharm+version+control&oq=PyCharm+version+control&sourceid=chrome&ie=UTF-8
#   Max Rohowsky: https://www.youtube.com/watch?v=8ZEssR8VTKo&t=175s

#   Word VBA Selection methods:
#   https://learn.microsoft.com/en-us/office/vba/api/word.selection
#
# Step 1: Need to first pip install pywin32 in the terminal !!!
# In PyCharm use these tabs: View --> Tools Window --> Terminal
import win32com.client      # COM stands for Component Object Model

# In addition to driving the wd.Selection object, I will use pyprclip to transfer clip board contents into the wd.doc
# More specifically, the plan is to have a while loop that keeps pasting into the wd.doc until user says stop

import pyperclip        #<-- this imported module enables copy/paste to clipboard (use pip to install)
import os               #<-- this module will allow us to manipulate the targeted file (fyl)

#  Step 2: Here we "launch" the MS Word application so we can create a blank wd doc
wd= win32com.client.Dispatch("Word.application")    # <-- "Dispatch" means launch

#  Next, we need to make the Launched app visible
#  For more info, see www.iCodeGuru.com/WebServer/Python-Programming-on-Win32/ch2.htm

wd.Visible= True

# Next we add a new "Document" to the Work Space
wd.Documents.Add()

# Here is where we have a while loop for copy/paste operations
# ---------- Colored escape definitions -------------
esc_white: str = '\033[97m'  #<-- this escape sequence will switch print()s to output white letters
esc_yellow: str = '\033[93m'  #<-- this escape sequence should switch print()s to output yellow letters
esc_red: str = '\033[91m'  #<-- this escape sequence is for red. See:
# https://jakob-bagterp.github.io/colorist-for-python/ansi-escape-codes/standard-16-colors/-
# -#foreground-text-and-background-colors  Note background colors can also be set

halted = False
i = 0
while not halted:
    i += 1
    print('Keep copying contents into Clipboard and press just Enter to paste. Type xx to terminate.')
    junk1 = input(f'({i}) Copy {esc_yellow} new content {esc_white} to Clipboard and hit <Enter> with Cursor placed here -->({i})__')
    new_content = pyperclip.paste()  # <--- pull the content string from the clipboard
    if junk1 == 'xx' or junk1 == 'XX': #<-- here we test for the 'xx' terminating text
        halted = True
        break
    else:
        wd.Selection.Paste()
        continue

print(f'Copy-Paste loop has been halted at i= {i}')

# Now we can start dumping the copied content into the new Wd Document
# wd.Selection.TypeText ("Now is the time for all good men to rise and stand ....")

# Finally go the Terminal and run this new "create_Word_Doc.py" program
# Make sure PyCharm's Run pointer is pointing to "Current File" and not to "main"
# And guess what ??? IT WORKED, but only ON THE SECOND RUN  --why?




