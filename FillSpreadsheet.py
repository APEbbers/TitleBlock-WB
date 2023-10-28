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


# This macro creates an Spreadsheet named "TitleBlock", reads editable texts from the first page
# and fills the spreadsheet.
# The data in spreadsheet can be changed by the user as long the data in column A remains unchanged.
# There will be three colums created in the spreadsheet:
#  - Property Name           -> the name of the editable text item
#  - Property Value          -> the value of the editable text item
#  - Increase, Yes or No?    -> do you want to increase the  value by 1 for the next sheet.
# there a multiple senarios for using these macro's:
#  - When changing the template for an different size, you can quickly refill the titleblock.
#    Of course, the editable text in the new template must be equal.
#  - Link the date to model properties like parameters. Basicly all what is achievable with the Spreadsheet workbench
#  - Import data from excel or libre office
#  - Etc.

import os
import FreeCAD as App
import Standard_Functions

# Get the settings
from Settings import INCLUDE_LENGTH
from Settings import INCLUDE_ANGLE
from Settings import INCLUDE_MASS
from Settings import INCLUDE_NO_SHEETS
from Settings import USE_EXTERNAL_SOURCE
from Settings import EXTERNAL_SOURCE_PATH
from Settings import EXTERNAL_SOURCE_SHEET_NAME
from Settings import EXTERNAL_SOURCE_STARTCELL
from Settings import MAP_LENGTH
from Settings import MAP_ANGLE
from Settings import MAP_MASS
from Settings import MAP_NOSHEETS
from Settings import ENABLE_DEBUG
from Settings import USE_FILENAME_DRAW_NO
from Settings import DRAW_NO_FiELD

# If no start cell is defined. the start cell will be "A1"
if len(EXTERNAL_SOURCE_STARTCELL) == 0:
    EXTERNAL_SOURCE_STARTCELL = "A1"

# When you copy/paste from a cell in the spreadsheet workbench, the char ' is automaticly added
# at the begining of the text.
# To avoid annociance, if this char is detected, it is removed from the setting.
if str(MAP_LENGTH).startswith("'"):
    MAP_LENGTH = str(MAP_LENGTH)[1:]
if str(MAP_ANGLE).startswith("'"):
    MAP_ANGLE = str(MAP_ANGLE)[1:]
if str(MAP_MASS).startswith("'"):
    MAP_MASS = str(MAP_MASS)[1:]
if str(MAP_NOSHEETS).startswith("'"):
    MAP_NOSHEETS = str(MAP_NOSHEETS)[1:]
if str(DRAW_NO_FiELD).startswith("'"):
    DRAW_NO_FiELD = str(DRAW_NO_FiELD)[1:]


def AddExtraData(sheet, StartRow):
    # The following system values can be added. You can use these for you specific templates.
    # Use the "bind function" of the spreadsheet workbench to create the correct entry for your template.:
    #   1. add the correct Property name
    #   2. bind the cell to the value of your specific template property.
    #   3. Mark the cell in column C if this value needs to be increased per page

    # If the debug mode is active, show which property is includex.difference(y)
    if ENABLE_DEBUG is True:
        if INCLUDE_LENGTH is True:
            print("Length unit is included: " + str(INCLUDE_LENGTH))
        if INCLUDE_ANGLE is True:
            print("Angle unit is included: " + str(INCLUDE_ANGLE))
        if INCLUDE_MASS is True:
            print("Mass unit is included: " + str(INCLUDE_MASS))
        if INCLUDE_NO_SHEETS is True:
            print("Number of pages is included: " + str(INCLUDE_NO_SHEETS))

    # get units scheme
    SchemeNumber = App.Units.getSchema()

    # Add the length units of your FreeCAD application
    if INCLUDE_LENGTH is True:
        sheet.set("A" + str(StartRow + 1), "Length_Units")
        sheet.set(
            "B" + str(StartRow + 1),
            str(
                App.Units.schemaTranslate(
                    App.Units.Quantity(50, App.Units.Length), SchemeNumber
                )
            )
            .split()[1]
            .replace("'", "")
            .replace(",", ""),
        )
        StartRow = StartRow + 1

    # Add the angular units of your FreeCAD application
    if INCLUDE_ANGLE is True:
        sheet.set("A" + str(StartRow + 1), "Angle_Units")
        sheet.set(
            "B" + str(StartRow + 1),
            str(
                App.Units.schemaTranslate(
                    App.Units.Quantity(50, App.Units.Angle), SchemeNumber
                )
            )
            .split()[1]
            .replace("'", "")
            .replace(",", ""),
        )
        StartRow = StartRow + 1

    # Add the mass units of your FreeCAD application
    if INCLUDE_MASS is True:
        sheet.set("A" + str(StartRow + 1), "Mass_Units")
        sheet.set(
            "B" + str(StartRow + 1),
            str(
                App.Units.schemaTranslate(
                    App.Units.Quantity(50, App.Units.Mass), SchemeNumber
                )
            )
            .split()[1]
            .replace("'", "")
            .replace(",", ""),
        )
        print(
            App.Units.schemaTranslate(
                App.Units.Quantity(50, App.Units.Mass), SchemeNumber
            )
        )
        StartRow = StartRow + 1

    # Add the total number of sheets. You can use this for your title block
    if INCLUDE_NO_SHEETS is True:
        # Get all the pages
        pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
        # Create the Property and add the value
        sheet.set("A" + str(StartRow + 1), "Number of sheets")
        sheet.set("B" + str(StartRow + 1), str(len(pages)))
    return


# Map data from the system and/or document to the spreadsheet
def MapData(sheet):
    # The following system values can be mapped. The values will be automaticly filled in defined properties

    # Get the filename
    filename = os.path.basename(App.ActiveDocument.FileName).split(".")[0]

    # if the debug mode is on, show what is mapped to which property
    if ENABLE_DEBUG is True:
        if str(MAP_LENGTH).strip():
            print("Length unit is mapped to: " + str(MAP_LENGTH))
        if str(MAP_ANGLE).strip():
            print("Angle unit is mapped to: " + str(MAP_ANGLE))
        if str(MAP_MASS).strip():
            print("Mass unit is mapped to: " + str(MAP_MASS))
        if str(MAP_NOSHEETS).strip():
            print("the number of pages is mapped to: " + str(MAP_NOSHEETS))
        if USE_FILENAME_DRAW_NO is True:
            if str(DRAW_NO_FiELD).strip():
                print(
                    "The filename ("
                    + str(filename)
                    + ") is mapped to: "
                    + str(DRAW_NO_FiELD)
                )

    # get units scheme
    SchemeNumber = App.Units.getSchema()

    # Go through the column A in the spreadsheet and find the properties.
    for RowNum in range(1000):
        # Start with x+1 first, to make sure that x is at least 1.
        RowNum = RowNum + 2

        # Map length units of your FreeCAD application
        if str(MAP_LENGTH).strip():
            # If the cell in column A is equal to MAP_LENGTH, add the value in column B
            if str(sheet.get("A" + str(RowNum))) == MAP_LENGTH:
                sheet.set(
                    "B" + str(RowNum),
                    str(
                        App.Units.schemaTranslate(
                            App.Units.Quantity(50, App.Units.Length), SchemeNumber
                        )
                    )
                    .split()[1]
                    .replace("'", "")
                    .replace(",", ""),
                )

        # Map angle units of your FreeCAD application
        if str(MAP_ANGLE).strip():
            # If the cell in column A is equal to MAP_ANGLE, add the value in column B
            if str(sheet.get("A" + str(RowNum))) == MAP_ANGLE:
                sheet.set(
                    "B" + str(RowNum),
                    str(
                        App.Units.schemaTranslate(
                            App.Units.Quantity(50, App.Units.Angle), SchemeNumber
                        )
                    )
                    .split()[1]
                    .replace("'", "")
                    .replace(",", ""),
                )

        # Map mass units of your FreeCAD application
        if str(MAP_MASS).strip():
            # If the cell in column A is equal to MAP_MASS, add the value in column B
            if str(sheet.get("A" + str(RowNum))) == MAP_MASS:
                sheet.set(
                    "B" + str(RowNum),
                    str(
                        App.Units.schemaTranslate(
                            App.Units.Quantity(50, App.Units.Mass), SchemeNumber
                        )
                    )
                    .split()[1]
                    .replace("'", "")
                    .replace(",", ""),
                )

        # Map the number of pages
        if str(MAP_NOSHEETS).strip():
            # Get all the pages
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
            # If the cell in column A is equal to MAP_NOSHEETS, add the value in column B
            if str(sheet.get("A" + str(RowNum))) == MAP_NOSHEETS:
                sheet.set("B" + str(RowNum), str(len(pages)))

        # Map the filename
        if USE_FILENAME_DRAW_NO is True:
            if str(DRAW_NO_FiELD).strip():
                # If the cell in column A is equal to DRAW_NO_FiELD, add the value in column B
                if str(sheet.get("A" + str(RowNum))) == DRAW_NO_FiELD:
                    print("I got here! " + filename)
                    sheet.set("B" + str(RowNum), filename)

        # Check if the next row exits. If not this is the end of all the available values.
        try:
            sheet.get("A" + str(RowNum + 1))
        except Exception:
            # print("end of range")
            return


# Fill the spreadsheet with all the date from the titleblock
def FillSheet():
    try:
        # get the fist page
        page = App.ActiveDocument.Page
        # get the editable texts
        texts = page.Template.EditableTexts
        # get the spreadsheet "TitleBlock"
        sheet = App.ActiveDocument.getObject("TitleBlock")

        # Debug mode is active, show all editable text in the page
        if ENABLE_DEBUG is True:
            print("the following editable text are present in your page:")
            for EditableText in texts.items():
                print(EditableText)

        # set the headers in the spreadsheet
        sheet.set("A1", "Property Name")
        sheet.set("B1", "Property Value")
        sheet.set("C1", "Increase value")
        sheet.set("D1", "Multiplier")
        sheet.set("E1", "Remarks")

        # set the start value for the start row.
        # (x=0, the spreadsheet whill be populated from the first row. the headers will be overwritten)
        StartRow = 1
        for key, value in texts.items():
            # Increase StartRow by one, to fill the next row
            StartRow = StartRow + 1
            # Fill the property name
            sheet.set("A" + str(StartRow), "{0}".format(key, value))
            # Fill the property value
            sheet.set("B" + str(StartRow), "{1}".format(key, value))
            # If there is no value yet, the increase function will be set empty by default.
            try:
                str(sheet.get("C" + str(StartRow)))
            except Exception:
                sheet.set("C" + str(StartRow), "")

        # Finally recompute the spreadsheet
        sheet.recompute()

        # Run the def to map system data
        MapData(sheet)

        # Run the def to add extra system data
        AddExtraData(sheet, StartRow)

        # Finally recompute the spreadsheet
        sheet.recompute()
    except Exception as e:
        Text = "TitleBlock Workbench: an error occurred!!"
        if ENABLE_DEBUG is True:
            Text = (
                "TitleBlock Workbench: an error occurred!!\n"
                + "See the report view for details"
            )
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
        if ENABLE_DEBUG is True:
            raise (e)


# Import data from a (central) excel workbook
def ImportDataExcel():
    from openpyxl import load_workbook

    # if debug mode is enabled, show the external file including path.
    if ENABLE_DEBUG is True:
        print(str(EXTERNAL_SOURCE_PATH))

    try:
        # Check if it is allowed to use an external source and if so, continue
        if USE_EXTERNAL_SOURCE is True:
            # try to open the source. if not show an messagebox and if debug mode is enabled, show the exeption as well
            try:
                wb = load_workbook(str(EXTERNAL_SOURCE_PATH), data_only=True)
                ws = wb[str(EXTERNAL_SOURCE_SHEET_NAME)]
            except Exception as e:
                if ENABLE_DEBUG is True:
                    raise (e)
                Standard_Functions.Mbox(
                    "an problem occured while openening the excel file!\nDo you have it open in an another application?"
                )
                return

            # get the spreadsheet "TitleBlock"
            sheet = App.ActiveDocument.getObject("TitleBlock")

            # Get the startcolumn and the other three columns from there
            StartCell = EXTERNAL_SOURCE_STARTCELL
            if (Standard_Functions.GetA1fromR1C1(StartCell)).strip():
                StartCell = Standard_Functions.GetA1fromR1C1(StartCell)
            StartColumn = StartCell[:1]
            # If debug mode is on, show the start colum and its number
            if ENABLE_DEBUG is True:
                print("Start column is: " + str(StartColumn))
                print(
                    "Column number is: "
                    + str(Standard_Functions.GetNumberFromLetter(StartColumn))
                )
            Column2 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 2), True
            )
            Column3 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 3), True
            )
            Column4 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 4), True
            )
            Column5 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 5), True
            )

            # Get the start row
            StartRow = EXTERNAL_SOURCE_STARTCELL[1:2]
            # if debug mode is on, show your start row
            if ENABLE_DEBUG is True:
                print("the start row is: " + str(StartRow))

            # import the headers from the excelsheet into the spreadsheet
            sheet.set("A1", str(ws[str(StartColumn) + str(StartRow)].value))
            sheet.set("B1", str(ws[str(Column2) + str(StartRow)].value))
            sheet.set("C1", str(ws[str(Column3) + str(StartRow)].value))
            sheet.set("D1", str(ws[str(Column4) + str(StartRow)].value))
            sheet.set("E1", str(ws[str(Column4) + str(StartRow)].value))

            # Go through the excel until the cell in the first column is empty.
            for i in range(1000):
                # Define the start row. This is the Header row +1 + i as counter
                RowNumber = int(StartRow) + i + 1

                # check if you reached the end of the data.
                if ws[str(StartColumn) + str(RowNumber)].value is None:
                    break

                # Get the number of row difference between the start row in the excelsheet
                # and the first row in the spreadsheet.
                # This to start at second row in the spreadsheet. (under the headers)
                Delta = int(StartRow) - 1

                # Fill the property name
                sheet.set(
                    str("A" + str(RowNumber - Delta)),
                    str(ws[str(StartColumn) + str(RowNumber)].value),
                )
                # Fill the property value
                if ws[Column2 + str(RowNumber)].value is not None:
                    sheet.set(
                        str("B" + str(RowNumber - Delta)),
                        str(ws[Column2 + str(RowNumber)].value),
                    )
                # Fill the value for auto increasement(yes or no)
                if ws[Column3 + str(RowNumber)].value is not None:
                    sheet.set(
                        str("C" + str(RowNumber - Delta)),
                        str(ws[Column3 + str(RowNumber)].value),
                    )
                # Fill the multipliers
                if ws[Column4 + str(RowNumber)].value is not None:
                    sheet.set(
                        str("D" + str(RowNumber - Delta)),
                        str(ws[Column4 + str(RowNumber)].value),
                    )
                # Fill the remarks
                if ws[Column5 + str(RowNumber)].value is not None:
                    sheet.set(
                        str("E" + str(RowNumber - Delta)),
                        str(ws[Column4 + str(RowNumber)].value),
                    )

            # Finally recompute the spreadsheet
            sheet.recompute()

            # Run the def to add extra system data.
            MapData(sheet)

            # Run the def to add extra system data. This is the final value of "RowNumber" minus the "StartRow".
            AddExtraData(sheet, RowNumber - int(StartRow))

            # Finally recompute the spreadsheet
            sheet.recompute()

        else:
            Standard_Functions.Mbox("External source is not enabled!", "", 0)
    except Exception as e:
        Text = "TitleBlock Workbench: an error occurred!!"
        if ENABLE_DEBUG is True:
            Text = (
                "TitleBlock Workbench: an error occurred!!\n"
                + "See the report view for details"
            )
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
        if ENABLE_DEBUG is True:
            raise (e)


def Start(command):
    # Find or Create spreadsheet
    try:
        # check if there is already an spreadsheet called "TitleBlock"
        sheet = App.ActiveDocument.getObject("TitleBlock")

        # if the debug mode is on, report presense of titleblock spreadsheet
        if ENABLE_DEBUG is True:
            print("TitleBlock already present")

        # Proceed with the macro.
        if command == "FillSpreadsheet":
            FillSheet()
        if command == "ImportExcel":
            ImportDataExcel()
    except Exception:
        # if there is not yet an spreadsheet called "TitleBlock", create one
        sheet = App.ActiveDocument.addObject("Spreadsheet::Sheet", "TitleBlock")

        # if the debug mode is on, report creation of titleblock spreadsheet
        if ENABLE_DEBUG is True:
            print("TitleBlock created")

        # set the label to "TitleBlock"
        sheet.Label = "TitleBlock"

        # Proceed with the macro.
        if command == "FillSpreadsheet":
            FillSheet()
        if command == "ImportExcel":
            ImportDataExcel()
