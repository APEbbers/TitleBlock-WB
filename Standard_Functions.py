import ctypes  # An included library with Python install.
import tkinter as tk
from tkinter.filedialog import asksaveasfile


#  Message Styles:
#  0 : OK
#  1 : OK | Cancel
#  2 : Abort | Retry | Ignore
#  3 : Yes | No | Cancel
#  4 : Yes | No
#  5 : Retry | Cancel
#  6 : Cancel | Try Again | Continuemport ctypes  # An included library with Python install.
def Mbox(text, title, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


# files must be like:
# files = [('All Files', '*.*'),
#              ('Python Files', '*.py'),
#              ('Text Document', '*.txt')]
def SaveDialog(files):
    # Create the window
    root = tk.Tk()
    # Hide the window
    root.withdraw()

    file = asksaveasfile(filetypes=files, defaultextension=files)
    return file.name


def GetLetterFromNumber(number: int, UCase: bool):
    Alfabet = {"abcdfghijklmnopqrstuvwxyz"}
    Letter = str(Alfabet)[number]

    # If UCase is true, convert to upper case
    if UCase is True:
        Letter = Letter.upper()

    return Letter


def GetNumberFromLetter(Letter):
    Alfabet = {"abcdfghijklmnopqrstuvwxyz"}
    Number = int(str(Alfabet).index(Letter.lower())) - 1

    return Number
