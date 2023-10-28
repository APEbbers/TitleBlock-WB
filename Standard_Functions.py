# ***************************************************************************
# *   Copyright (c) 2015 Paul Ebbers paul.ebbers@gmail.com                  *
# *                                                                         *
# *   This file is part of the FreeCAD CAx development system.              *
# *                                                                         *
# *   This program is free software; you can redistribute it and/or modify  *
# *   it under the terms of the GNU Lesser General Public License (LGPL)    *
# *   as published by the Free Software Foundation; either version 2 of     *
# *   the License, or (at your option) any later version.                   *
# *   for detail see the LICENCE text file.                                 *
# *                                                                         *
# *   FreeCAD is distributed in the hope that it will be useful,            *
# *   but WITHOUT ANY WARRANTY; without even the implied warranty of        *
# *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         *
# *   GNU Lesser General Public License for more details.                   *
# *                                                                         *
# *   You should have received a copy of the GNU Library General Public     *
# *   License along with FreeCAD; if not, write to the Free Software        *
# *   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  *
# *   USA                                                                   *
# *                                                                         *
# ***************************************************************************/


def Mbox(text, title="", style=0, default="", stringList="[,]"):
    """
    Message Styles:
    0 : OK                          (text, title, style)
    1 : Yes | No                    (text, title, style)
    2 : Inputbox                    (text, title, style, default)
    3 : Inputbox with dropdown      (text, title, style, default, stringlist)
    """
    from PySide import QtGui

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


def SaveDialog(files, OverWrite: bool = True):
    """
    files must be like:
    files = [
        ('All Files', '*.*'),
        ('Python Files', '*.py'),
        ('Text Document', '*.txt')
    ]

    OverWrite:
    If True, file will be overwritten
    If False, only the path+filename will be returned
    """
    import tkinter as tk
    from tkinter.filedialog import asksaveasfile
    from tkinter.filedialog import askopenfilename

    # Create the window
    root = tk.Tk()
    # Hide the window
    root.withdraw()

    if OverWrite is True:
        file = asksaveasfile(filetypes=files, defaultextension=files)
        if file is not None:
            return file.name
    if OverWrite is False:
        file = askopenfilename(filetypes=files, defaultextension=files)
        if file is not None:
            return file


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


def CheckIfWorkbookExists(FullFileName: str, CreateIfNone: bool = True):
    import os
    from openpyxl import Workbook

    result = False
    try:
        result = os.path.exists(FullFileName)
    except Exception:
        if CreateIfNone is True:
            Filter = [
                ("Excel", "*.xlsx"),
                (
                    "Excel Macro-enabled Workbook",
                    "*.xlsm",
                ),
            ]
            FullFileName = SaveDialog(Filter)
            if FullFileName.strip():
                wb = Workbook(str(FullFileName))
                wb.save(FullFileName)
                wb.close()
                result = True
        if CreateIfNone is False:
            result = False
    return result
