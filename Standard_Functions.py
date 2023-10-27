import ctypes  # An included library with Python install.
import tkinter as tk
from tkinter.filedialog import asksaveasfile
from PySide import QtCore
from PySide import QtGui


def Mbox(text, title="", style=0, default="", stringList="[,]"):
    """
    Message Styles:
    0 : OK                          (text, title, style)
    1 : Yes | No                    (text, title, style)
    2 : Inputbox                    (text, title, style, default)
    3 : Inputbox with dropdown      (text, title, style, default, stringlist)
    """
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
        reply = QtGui.QInputDialog.getText(None, title, text, text=default)

        if reply[1]:
            # user clicked OK
            replyText = reply[0]
        else:
            # user clicked Cancel
            replyText = reply[0]  # which will be "" if they clicked Cancel
        return str(replyText)
    if style == 3:
        reply = QtGui.QInputDialog.getItem(None, title, text, stringList, 1, True)

        if reply[1]:
            # user clicked OK
            replyText = reply[0]
        else:
            # user clicked Cancel
            replyText = reply[0]  # which will be "" if they clicked Cancel
        return str(replyText)


def SaveAsDialog(files):
    """
    files must be like:
    files = [
        ('All Files', '*.*'),
        ('Python Files', '*.py'),
        ('Text Document', '*.txt')
    ]
    """

    # Create the window
    root = tk.Tk()
    # Hide the window
    root.withdraw()

    file = asksaveasfile(filetypes=files, defaultextension=files)
    if file is not None:
        return file.name


def GetLetterFromNumber(number: int, UCase: bool = True):
    Alfabet = {"abcdefghijklmnopqrstuvwxyz"}
    Letter = str(Alfabet)[number + 1]

    # If UCase is true, convert to upper case
    if UCase is True:
        Letter = Letter.upper()

    return Letter


def GetNumberFromLetter(Letter):
    Alfabet = {"abcdefghijklmnopqrstuvwxyz"}
    Number = int(str(Alfabet).index(Letter.lower())) - 1

    return Number


def GetA1fromR1C1(input: str) -> str:
    if input[:1] == "'":
        input = input[1:]
    try:
        input = input.upper()
        ColumnPosition = input.find("C")
        RowNumber = int(input[1:(ColumnPosition)])
        ColumnNumber = int(input[(ColumnPosition + 1) :])

        ColumnLetter = GetLetterFromNumber(ColumnNumber)

        return str(ColumnLetter + str(RowNumber))
    except Exception:
        return ""
