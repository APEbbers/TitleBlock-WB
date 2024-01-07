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
from Settings import AUTOFIT_FACTOR


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


def FormatTable(
    sheet,
    HeaderRange,
    TableRange,
    FirstColumnRange,
):
    """_summary_

    Args:
        sheet (object): FreeCAD sheet object
        HeaderRange (string): Range for the header.
        TableRange (string): Range for the table
        FirstColumnRange (string): Range for the first column
    """

    # Format the header ------------------------------------------------------------------------------------------------
    # Set the font style for the header
    sheet.setStyle(
        HeaderRange,
        FontStyle(
            SPREADSHEET_HEADERFONTSTYLE_BOLD, SPREADSHEET_HEADERFONTSTYLE_ITALIC,
            SPREADSHEET_HEADERFONTSTYLE_UNDERLINE))  # \bold|italic|underline'
    # Set the colors for the header
    sheet.setBackground(HeaderRange, SPREADSHEET_HEADERBACKGROUND)
    sheet.setForeground(
        HeaderRange, SPREADSHEET_HEADERFOREGROUND
    )  # RGBA
    # ------------------------------------------------------------------------------------------------------------------

    # Format the table -------------------------------------------------------------------------------------------------
    # Get the first column and first row
    TableRangeColumnStart = Standard_Functions.RemoveNumbersFromString(
        TableRange.split(":")[0]
    )
    TableRangeRowStart = int(
        Standard_Functions.RemoveLettersFromString(TableRange.split(":")[0])
    )

    # Get the last column and last row
    TableRangeColumnEnd = Standard_Functions.RemoveNumbersFromString(
        TableRange.split(":")[1]
    )
    TableRangeRowEnd = int(
        Standard_Functions.RemoveLettersFromString(TableRange.split(":")[1])
    )

    # Get the second column
    TableRangeSecondColumn = Standard_Functions.GetLetterFromNumber(
        Standard_Functions.GetNumberFromLetter(
            Standard_Functions.RemoveNumbersFromString(
                TableRange.split(":")[0]
            )
        ) + 1
    )

    print(f"{TableRangeColumnStart}:{TableRangeColumnStart}, {TableRangeColumnEnd}:{TableRangeRowEnd}, {TableRangeSecondColumn}")

    # Calculate the delta between the start and end of the table in vertical direction (Rows).
    DeltaRange = TableRangeRowEnd - TableRangeRowStart + 1
    # Go through the range
    for i in range(1, DeltaRange + 2, 2):
        # Correct the position
        j = i - 1
        # Define the first row
        FirstRow = f"{TableRangeColumnStart}{str(j+TableRangeRowStart)}:{TableRangeColumnEnd}{str(j+TableRangeRowStart)}"
        # Define the second row
        SecondRow = f"{TableRangeColumnStart}{str(j+TableRangeRowStart+1)}:{TableRangeColumnEnd}{str(j+TableRangeRowStart+1)}"

        # if the first and second rows are within the range, set the colors
        if i <= DeltaRange:
            sheet.setBackground(
                FirstRow, SPREADSHEET_TABLEBACKGROUND_1
            )
            sheet.setForeground(
                FirstRow, SPREADSHEET_TABLEFOREGROUND
            )
        if i + 1 <= DeltaRange:
            sheet.setBackground(
                SecondRow, SPREADSHEET_TABLEBACKGROUND_2
            )
            sheet.setForeground(
                SecondRow, SPREADSHEET_TABLEFOREGROUND
            )

    # Set the font style for the table
    sheet.setStyle(
        TableRange,
        FontStyle(
            SPREADSHEET_TABLEFONTSTYLE_BOLD, SPREADSHEET_TABLEFONTSTYLE_ITALIC,
            SPREADSHEET_TABLEFONTSTYLE_UNDERLINE))  # \bold|italic|underline'

    # Set the font style for the first column
    sheet.setStyle(
        FirstColumnRange,
        FontStyle(
            SPREADSHEET_COLUMNFONTSTYLE_BOLD, SPREADSHEET_COLUMNFONTSTYLE_ITALIC, SPREADSHEET_COLUMNFONTSTYLE_UNDERLINE
        )
    )  # \bold|italic|underline'
    # ------------------------------------------------------------------------------------------------------------------

    # Align the table and headers --------------------------------------------------------------------------------------
    # align the columns
    sheet.setAlignment(
        f"{TableRangeColumnStart}{TableRangeRowStart}:{TableRangeColumnStart}{TableRangeRowEnd}", "left|vcenter"
    )
    sheet.setAlignment(
        f"{TableRangeSecondColumn}{TableRangeRowStart}:{TableRangeColumnEnd}{TableRangeRowEnd}", "center|vcenter"
    )

    # Set the column width
    for i in range(TableRangeRowStart, TableRangeRowEnd):
        for j in range(Standard_Functions.GetNumberFromLetter(TableRangeColumnStart),
                       Standard_Functions.GetNumberFromLetter(TableRangeColumnEnd)):
            Standard_Functions.SetColumnWidth_SpreadSheet(
                sheet=sheet, column=f"{FirstColumnRange}{j}", cellValue=sheet.getContents(
                    f"{Standard_Functions.GetLetterFromNumber(j)}{i}"),
                factor=AUTOFIT_FACTOR)
    # ------------------------------------------------------------------------------------------------------------------
    return sheet
