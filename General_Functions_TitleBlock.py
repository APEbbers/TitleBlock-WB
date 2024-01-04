# ***************************************************************************
# *   Copyright (c) 2023 Paul Ebbers paul.ebbers@gmail.com                  *
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

# region imports

import os
import FreeCAD as App
import Standard_Functions_TitleBlock as Standard_Functions

# Get the settings
from Settings import SPREADSHEET_COLUMNFONTSTYLE_UNDERLINE
from Settings import SPREADSHEET_COLUMNFONTSTYLE_ITALIC
from Settings import SPREADSHEET_COLUMNFONTSTYLE_BOLD
from Settings import SPREADSHEET_TABLEFONTSTYLE_UNDERLINE
from Settings import SPREADSHEET_TABLEFONTSTYLE_ITALIC
from Settings import SPREADSHEET_TABLEFONTSTYLE_BOLD
from Settings import SPREADSHEET_TABLEFOREGROUND
from Settings import SPREADSHEET_TABLEBACKGROUND_2
from Settings import SPREADSHEET_TABLEBACKGROUND_1
from Settings import SPREADSHEET_HEADERFONTSTYLE_UNDERLINE
from Settings import SPREADSHEET_HEADERFONTSTYLE_ITALIC
from Settings import SPREADSHEET_HEADERFONTSTYLE_BOLD
from Settings import SPREADSHEET_HEADERFOREGROUND
from Settings import SPREADSHEET_HEADERBACKGROUND
from Settings import DRAW_NO_FiELD
from Settings import USE_FILENAME_DRAW_NO
from Settings import ENABLE_DEBUG
from Settings import DOCINFO_COMMENT
from Settings import DOCINFO_LICENSEURL
from Settings import DOCINFO_LICENSE
from Settings import DOCINFO_COMPANY
from Settings import DOCINFO_LASTMODIFIEDDATE
from Settings import DOCINFO_LASTMODIFIEDBY
from Settings import DOCINFO_CREATEDDATE
from Settings import DOCINFO_CREATEDBY
from Settings import DOCINFO_NAME
from Settings import MAP_NOSHEETS
from Settings import MAP_MASS
from Settings import MAP_ANGLE
from Settings import MAP_LENGTH
from Settings import EXTERNAL_SOURCE_STARTCELL
from Settings import EXTERNAL_SOURCE_SHEET_NAME
from Settings import EXTERNAL_SOURCE_PATH
from Settings import USE_EXTERNAL_SOURCE
from Settings import INCLUDE_NO_SHEETS
from Settings import INCLUDE_MASS
from Settings import INCLUDE_ANGLE
from Settings import INCLUDE_LENGTH
from Settings import AUTOFIT_FACTOR
from Settings import preferences
# endregion

# Define the translation
translate = App.Qt.translate


# Function the return the correct string to use in the FormatTable function
def FontStyle(Bold: bool, Italic: bool, UnderLine: bool) -> str:
    result = ""
    if Bold is True:
        if result == "":
            result = "bold"
        if result != "":
            result = result + "|bold"
    if Italic is True:
        if result == "":
            result = "italic"
        if result != "":
            result = result + "|italic"
    if UnderLine is True:
        if result == "":
            result = "underline"
        if result != "":
            result = result + "|underline"
    return result


# Format the table based on the preferences
def FormatTable(sheet, Endrow):
    # HeaderRange
    RangeAlign1 = "A1:A" + str(Endrow)
    RangeStyle1 = "A1:E1"

    # TableRange
    RangeAlign2 = "B1:E" + str(Endrow)
    RangeStyle2 = "B2:E" + str(Endrow)
    # First column
    RangeStyle3 = "A2:A" + str(Endrow)

    # Font style for the top row
    sheet.setStyle(
        RangeStyle1,
        FontStyle(
            SPREADSHEET_HEADERFONTSTYLE_BOLD,
            SPREADSHEET_HEADERFONTSTYLE_ITALIC,
            SPREADSHEET_HEADERFONTSTYLE_UNDERLINE,
        ),
    )

    # Font style for the first column
    sheet.setStyle(
        RangeStyle3,
        FontStyle(
            SPREADSHEET_COLUMNFONTSTYLE_BOLD,
            SPREADSHEET_COLUMNFONTSTYLE_ITALIC,
            SPREADSHEET_COLUMNFONTSTYLE_UNDERLINE,
        ),
    )

    # Font style for the rest of the table
    sheet.setStyle(
        RangeStyle2,
        FontStyle(
            SPREADSHEET_TABLEFONTSTYLE_BOLD,
            SPREADSHEET_TABLEFONTSTYLE_ITALIC,
            SPREADSHEET_TABLEFONTSTYLE_UNDERLINE,
        ),
    )

    sheet.setBackground(RangeStyle1, SPREADSHEET_HEADERBACKGROUND)
    sheet.setForeground(RangeStyle1, SPREADSHEET_HEADERFOREGROUND)

    # Style the rest of the table
    for i in range(2, int(Endrow + 1), 2):
        RangeStyle3 = f"A{i}:E{i}"
        RangeStyle4 = f"A{i+1}:E{i+1}"
        sheet.setBackground(RangeStyle3, SPREADSHEET_TABLEBACKGROUND_1)
        sheet.setBackground(RangeStyle4, SPREADSHEET_TABLEBACKGROUND_2)
        sheet.setForeground(RangeStyle3, SPREADSHEET_TABLEFOREGROUND)
        sheet.setForeground(RangeStyle4, SPREADSHEET_TABLEFOREGROUND)

    # align the columns
    sheet.setAlignment(RangeAlign1, "left|vcenter")
    sheet.setAlignment(RangeAlign2, "center|vcenter")

    # Set the column width
    for i in range(1, Endrow):
        Standard_Functions.SetColumnWidth_SpreadSheet(
            sheet=sheet, column=f"A{i}", cellValue=sheet.getContents(f"A{i}"), factor=AUTOFIT_FACTOR)
        Standard_Functions.SetColumnWidth_SpreadSheet(
            sheet=sheet, column=f"B{i}", cellValue=sheet.getContents(f"B{i}"), factor=AUTOFIT_FACTOR)
        Standard_Functions.SetColumnWidth_SpreadSheet(
            sheet=sheet, column=f"C{i}", cellValue=sheet.getContents(f"C{i}"), factor=AUTOFIT_FACTOR)
        Standard_Functions.SetColumnWidth_SpreadSheet(
            sheet=sheet, column=f"D{i}", cellValue=sheet.getContents(f"D{i}"), factor=AUTOFIT_FACTOR)
        Standard_Functions.SetColumnWidth_SpreadSheet(
            sheet=sheet, column=f"E{i}", cellValue=sheet.getContents(f"E{i}"), factor=AUTOFIT_FACTOR)
    return
