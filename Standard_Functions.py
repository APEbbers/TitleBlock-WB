import ctypes  # An included library with Python install.
import tkinter as tk
from tkinter.filedialog import asksaveasfile
from PySide import QtCore
from PySide import QtGui


#  Message Styles:
#  0 : OK
#  1 : Yes | No
#  2 : Inputbox
def Mbox(text, title="", style=0):
    if style == 0:
        reply = str(QtGui.QMessageBox.information(None, title, text))
        return reply
    if style == 1:
        reply = QtGui.QMessageBox.question(
            None,
            title,
            text,
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
            QtGui.QMessageBox.No,
        )
        if reply == QtGui.QMessageBox.Yes:
            return "yes"
        if reply == QtGui.QMessageBox.No:
            return "no"
    if style == 2:
        reply = QtGui.QInputDialog.getText(None, title, text)
        if reply[1]:
            # user clicked OK
            replyText = reply[0]
        else:
            # user clicked Cancel
            replyText = reply[0]  # which will be "" if they clicked Cancel
        return str(replyText)


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
