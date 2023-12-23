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
import Standard_Functions_TitleBlock as Standard_Functions

# Get the preferences
from Settings import preferences

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
from Settings import DOCINFO_NAME
from Settings import DOCINFO_CREATEDBY
from Settings import DOCINFO_CREATEDDATE
from Settings import DOCINFO_LASTMODIFIEDBY
from Settings import DOCINFO_LASTMODIFIEDDATE
from Settings import DOCINFO_COMPANY
from Settings import DOCINFO_LICENSE
from Settings import DOCINFO_LICENSEURL
from Settings import DOCINFO_COMMENT
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
if str(DOCINFO_NAME).startswith("'"):
    DOCINFO_NAME = str(DOCINFO_NAME)[1:]
if str(DOCINFO_CREATEDBY).startswith("'"):
    DOCINFO_CREATEDBY = str(DOCINFO_CREATEDBY)[1:]
if str(DOCINFO_CREATEDDATE).startswith("'"):
    DOCINFO_CREATEDDATE = str(DOCINFO_CREATEDDATE)[1:]
if str(DOCINFO_LASTMODIFIEDBY).startswith("'"):
    DOCINFO_LASTMODIFIEDBY = str(DOCINFO_LASTMODIFIEDBY)[1:]
if str(DOCINFO_LASTMODIFIEDDATE).startswith("'"):
    DOCINFO_LASTMODIFIEDDATE = str(DOCINFO_LASTMODIFIEDDATE)[1:]
if str(DOCINFO_COMPANY).startswith("'"):
    DOCINFO_COMPANY = str(DOCINFO_COMPANY)[1:]
if str(DOCINFO_LICENSE).startswith("'"):
    DOCINFO_LICENSE = str(DOCINFO_LICENSE)[1:]
if str(DOCINFO_LICENSEURL).startswith("'"):
    DOCINFO_LICENSEURL = str(DOCINFO_LICENSEURL)[1:]
if str(DOCINFO_COMMENT).startswith("'"):
    DOCINFO_COMMENT = str(DOCINFO_COMMENT)[1:]

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
            Standard_Functions.Print("Length unit is included: " + str(INCLUDE_LENGTH), "Log")
        if INCLUDE_ANGLE is True:
            Standard_Functions.Print("Angle unit is included: " + str(INCLUDE_ANGLE), "Log")
        if INCLUDE_MASS is True:
            Standard_Functions.Printt("Mass unit is included: " + str(INCLUDE_MASS), "Log")
        if INCLUDE_NO_SHEETS is True:
            Standard_Functions.Print("Number of pages is included: " + str(INCLUDE_NO_SHEETS), "Log")

    # get units scheme
    SchemeNumber = App.Units.getSchema()

    # Add the length units of your FreeCAD application
    if INCLUDE_LENGTH is True:
        sheet.set("A" + str(StartRow + 1), "Length_Units")
        sheet.set(
            "B" + str(StartRow + 1),
            str(App.Units.schemaTranslate(App.Units.Quantity(50, App.Units.Length), SchemeNumber))
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
            str(App.Units.schemaTranslate(App.Units.Quantity(50, App.Units.Angle), SchemeNumber))
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
            str(App.Units.schemaTranslate(App.Units.Quantity(50, App.Units.Mass), SchemeNumber))
            .split()[1]
            .replace("'", "")
            .replace(",", ""),
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
def MapData(sheet, MapSpecific: int = 0):
    """
    0:  Map all\n
    1:  Map length unit only\n
    2:  Map angle unit only\n
    3:  Map mass unit only\n
    4:  Map No of sheets only\n
    5:  Map filename only\n
    """
    # Get the filename
    filename = os.path.basename(App.ActiveDocument.FileName).split(".")[0]

    # if the debug mode is on, show what is mapped to which property
    if ENABLE_DEBUG is True:
        if str(MAP_LENGTH).strip():
            Standard_Functions.Print("Length unit is mapped to: " + str(MAP_LENGTH), "Log")
        if str(MAP_ANGLE).strip():
            Standard_Functions.Print("Angle unit is mapped to: " + str(MAP_ANGLE), "Log")
        if str(MAP_MASS).strip():
            Standard_Functions.Print("Mass unit is mapped to: " + str(MAP_MASS), "Log")
        if str(MAP_NOSHEETS).strip():
            Standard_Functions.Print("the number of pages is mapped to: " + str(MAP_NOSHEETS), "Log")
        if USE_FILENAME_DRAW_NO is True:
            Standard_Functions.Print("The filename (" + str(filename) + ") is mapped to: " + str(DRAW_NO_FiELD))
        return

    # get units scheme
    SchemeNumber = App.Units.getSchema()

    # Go through the column A in the spreadsheet and find the properties.
    for RowNum in range(1000):
        # Start with x+1 first, to make sure that x is at least 1.
        RowNum = RowNum + 2

        # Map length units of your FreeCAD application
        # Map only as requested
        if MapSpecific == 0 or MapSpecific == 1:
            if str(MAP_LENGTH).strip():
                # If the cell in column A is equal to MAP_LENGTH, add the value in column B
                if str(sheet.getContents("A" + str(RowNum))) == MAP_LENGTH:
                    sheet.set(
                        "B" + str(RowNum),
                        str(App.Units.schemaTranslate(App.Units.Quantity(50, App.Units.Length), SchemeNumber))
                        .split()[1]
                        .replace("'", "")
                        .replace(",", ""),
                    )

        # Map angle units of your FreeCAD application
        # Map only as requested
        if MapSpecific == 0 or MapSpecific == 2:
            if str(MAP_ANGLE).strip():
                # If the cell in column A is equal to MAP_ANGLE, add the value in column B
                if str(sheet.getContents("A" + str(RowNum))) == MAP_ANGLE:
                    sheet.set(
                        "B" + str(RowNum),
                        str(App.Units.schemaTranslate(App.Units.Quantity(50, App.Units.Angle), SchemeNumber))
                        .split()[1]
                        .replace("'", "")
                        .replace(",", ""),
                    )

        # Map mass units of your FreeCAD application
        # Map only as requested
        if MapSpecific == 0 or MapSpecific == 3:
            if str(MAP_MASS).strip():
                # If the cell in column A is equal to MAP_MASS, add the value in column B
                if str(sheet.getContents("A" + str(RowNum))) == MAP_MASS:
                    sheet.set(
                        "B" + str(RowNum),
                        str(App.Units.schemaTranslate(App.Units.Quantity(50, App.Units.Mass), SchemeNumber))
                        .split()[1]
                        .replace("'", "")
                        .replace(",", ""),
                    )

        # Map the number of pages
        # Map only as requested
        if MapSpecific == 0 or MapSpecific == 4:
            if str(MAP_NOSHEETS).strip():
                # Get all the pages
                pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
                # If the cell in column A is equal to MAP_NOSHEETS, add the value in column B
                if str(sheet.getContents("A" + str(RowNum))) == MAP_NOSHEETS:
                    sheet.set("B" + str(RowNum), str(len(pages)))

        # Map the filename
        # Map only as requested
        if MapSpecific == 0 or MapSpecific == 5:
            if USE_FILENAME_DRAW_NO is True:
                if str(DRAW_NO_FiELD).strip():
                    # If the cell in column A is equal to DRAW_NO_FiELD, add the value in column B
                    if str(sheet.getContents("A" + str(RowNum))) == DRAW_NO_FiELD:
                        sheet.set("B" + str(RowNum), filename)

        # Check if the next row exits. If not this is the end of all the available values.
        try:
            sheet.getContents("A" + str(RowNum + 1))
        except Exception:
            # print("end of range")
            return


# Map document information
def MapDocInfo(sheet):
    doc = App.ActiveDocument

    # if the debug mode is on, show what is mapped to which property
    if ENABLE_DEBUG is True:
        if str(DOCINFO_NAME).strip():
            Standard_Functions.Print("Document name is mapped to:  " + str(DOCINFO_NAME), "Log")
        if str(DOCINFO_CREATEDBY).strip():
            Standard_Functions.Print("Created by value is mapped to:  " + str(DOCINFO_CREATEDBY), "Log")
        if str(DOCINFO_CREATEDDATE).strip():
            Standard_Functions.Print("Created date is mapped to:  " + str(DOCINFO_CREATEDDATE), "Log")
        if str(DOCINFO_LASTMODIFIEDBY).strip():
            Standard_Functions.Print("Last modified by value is mapped to:  " + str(DOCINFO_LASTMODIFIEDBY), "Log")
        if str(DOCINFO_LASTMODIFIEDDATE).strip():
            Standard_Functions.Print("Last modified date is mapped to:  " + str(DOCINFO_LASTMODIFIEDDATE), "Log")
        if str(DOCINFO_COMPANY).strip():
            Standard_Functions.Print("Company name is mapped to:  " + str(DOCINFO_COMPANY), "Log")
        if str(DOCINFO_LICENSE).strip():
            Standard_Functions.Print("License name is mapped to:  " + str(DOCINFO_LICENSE), "Log")
        if str(DOCINFO_LICENSEURL).strip():
            Standard_Functions.Print("License link is mapped to:  " + str(DOCINFO_LICENSEURL), "Log")
        if str(DOCINFO_COMMENT).strip():
            Standard_Functions.Print("Comment is mapped to:  " + str(DOCINFO_COMMENT), "Log")

    # Go through the column A in the spreadsheet and find the properties.
    for RowNum in range(1000):
        # Start with x+1 first, to make sure that x is at least 1.
        RowNum = RowNum + 2

        # DOCINFO_NAME
        if str(DOCINFO_NAME).strip():
            # If the cell in column A is equal to # DOCINFO_NAME, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_NAME:
                sheet.set("B" + str(RowNum), doc.Name)

        # DOCINFO_CREATEDBY
        if str(DOCINFO_CREATEDBY).strip():
            # If the cell in column A is equal to # DOCINFO_CREATEDBY, add the value in column B
            print(f"{DOCINFO_CREATEDBY}, {str(sheet.getContents('A' + str(RowNum)))}")
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_CREATEDBY:
                print(doc.CreatedBy)
                sheet.set("B" + str(RowNum), doc.CreatedBy)

        # DOCINFO_CREATEDDATE
        if str(DOCINFO_CREATEDDATE).strip():
            # If the cell in column A is equal to # DOCINFO_CREATEDDATE, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_CREATEDDATE:
                CreationDate = doc.CreationDate.split("T")[0]
                sheet.set("B" + str(RowNum), CreationDate)

        # DOCINFO_LASTMODIFIEDBY
        if str(DOCINFO_LASTMODIFIEDBY).strip():
            # If the cell in column A is equal to # DOCINFO_LASTMODIFIEDBY, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_LASTMODIFIEDBY:
                sheet.set("B" + str(RowNum), doc.LastModifiedBy)

        # DOCINFO_LASTMODIFIEDDATE
        if str(DOCINFO_LASTMODIFIEDDATE).strip():
            # If the cell in column A is equal to # DOCINFO_LASTMODIFIEDDATE, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_LASTMODIFIEDDATE:
                sheet.set("B" + str(RowNum), doc.LastModifiedDate)

        # DOCINFO_COMPANY
        if str(DOCINFO_COMPANY).strip():
            # If the cell in column A is equal to # DOCINFO_COMPANY, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_COMPANY:
                sheet.set("B" + str(RowNum), doc.Company)

        # DOCINFO_LICENSE
        if str(DOCINFO_LICENSE).strip():
            # If the cell in column A is equal to # DOCINFO_LICENSE, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_LICENSE:
                sheet.set("B" + str(RowNum), doc.License)

        # DOCINFO_LICENSEURL
        if str(DOCINFO_LICENSEURL).strip():
            # If the cell in column A is equal to # DOCINFO_LICENSEURL, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_LICENSEURL:
                sheet.set("B" + str(RowNum), doc.LicenseURL)

        # DOCINFO_COMMENT
        if str(DOCINFO_COMMENT).strip():
            # If the cell in column A is equal to # DOCINFO_COMMENT, add the value in column B
            if str(sheet.getContents("A" + str(RowNum))) == DOCINFO_COMMENT:
                sheet.set("B" + str(RowNum), doc.Comment)
    return


def FormatTable(sheet, Endrow):
    RangeAlign1 = "A1:A" + str(Endrow)
    RangeAlign2 = "B1:E" + str(Endrow)
    RangeStyle1 = "A1:E1"
    RangeStyle2 = "A2:A" + str(Endrow)

    # Style the top row
    sheet.setStyle(RangeStyle1, "bold")  # \bold|italic|underline'
    sheet.setStyle(RangeStyle2, "bold")  # \bold|italic|underline'
    sheet.setBackground(RangeStyle1, Standard_Functions.ColorConvertor([255, 128, 0]))
    # sheet.setBackground(RangeStyle2, Standard_Functions.ColorConvertor([0, 153, 0]))
    sheet.setForeground(RangeStyle1, Standard_Functions.ColorConvertor([0, 0, 0]))  # RGBA

    # Style the rest of the table

    for i in range(2, int(Endrow + 1), 2):
        RangeStyle3 = f"A{i}:E{i}"
        RangeStyle4 = f"A{i+1}:E{i+1}"
        sheet.setBackground(RangeStyle3, Standard_Functions.ColorConvertor([128, 128, 128]))
        sheet.setBackground(RangeStyle4, Standard_Functions.ColorConvertor([64, 64, 64]))
        sheet.setForeground(RangeStyle3, Standard_Functions.ColorConvertor([0, 0, 0]))
        sheet.setForeground(RangeStyle4, Standard_Functions.ColorConvertor([0, 0, 0]))

    # align the columns
    sheet.setAlignment(RangeAlign1, "left|vcenter")
    sheet.setAlignment(RangeAlign2, "center|vcenter")

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
            Standard_Functions.Print("the following editable text are present in your page:", "Log")
            for EditableText in texts.items():
                Standard_Functions.Print(str(EditableText), "Log")

        # set the headers in the spreadsheet
        sheet.set("A1", "Property Name")
        sheet.set("B1", "Property Value")
        sheet.set("C1", "Increase value")
        sheet.set("D1", "Factor")
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
                str(sheet.getContents("C" + str(StartRow)))
            except Exception:
                sheet.set("C" + str(StartRow), "")

        # Finally recompute the spreadsheet
        sheet.recompute()

        # Run the def to map system data
        MapData(sheet=sheet)

        # Run the def to map document information
        MapDocInfo(sheet=sheet)

        # Run the def to add extra system data
        AddExtraData(sheet, StartRow)

        # Format the spreadsheet
        extraRows = 0
        if INCLUDE_LENGTH is True:
            extraRows = extraRows + 1
        if INCLUDE_ANGLE is True:
            extraRows = extraRows + 1
        if INCLUDE_MASS is True:
            extraRows = extraRows + 1
        if INCLUDE_NO_SHEETS is True:
            extraRows = extraRows + 1
        FormatTable(sheet=sheet, Endrow=StartRow + extraRows)

        # Finally recompute the spreadsheet
        sheet.recompute()
        App.ActiveDocument.recompute()
    # except TypeError:
    #     pass
    except Exception as e:
        Text = "TitleBlock Workbench: an error occurred!!\n"
        if ENABLE_DEBUG is True:
            Text = "TitleBlock Workbench: an error occurred!!\n" + "See the report view for details"
            raise e
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
    return


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
                if EXTERNAL_SOURCE_SHEET_NAME == "":
                    # Set the sheetname with a inputbox
                    Worksheets_List = [i for i in wb.sheetnames if i != "Settings"]
                    Input_SheetName = str(
                        Standard_Functions.Mbox(
                            text="Please enter the name of the worksheet",
                            title="",
                            style=3,
                            default="TitleBlockData",
                            stringList=Worksheets_List,
                        )
                    )
                    # if the user canceled, exit this function.
                    if not Input_SheetName.strip():
                        return
                    ws = wb[str(Input_SheetName)]
                if EXTERNAL_SOURCE_SHEET_NAME != "":
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
            if EXTERNAL_SOURCE_SHEET_NAME == "":
                # Set EXTERNAL_SOURCE_SHEET_NAME to the chosen sheetname
                preferences.SetString("SheetName", Input_SheetName)
                ws = wb[Input_SheetName]

                # Set the startcell with an inputbox
                StartCell = str(
                    Standard_Functions.Mbox(
                        "Please enter the name of the cell.\n"
                        + "Enter a single cell like 'A1', 'B2', etc. Other notations will be ignored!",
                        style=2,
                        default="A1",
                    )
                )
                if not StartCell.strip():
                    StartCell = "A1"

                # Set SHEETNAME_STARTCELL to the chosen sheetname
                preferences.SetString("StartCell", StartCell)

            if (Standard_Functions.GetA1fromR1C1(StartCell)).strip():
                StartCell = Standard_Functions.GetA1fromR1C1(StartCell)
            StartColumn = StartCell[:1]
            # If debug mode is on, show the start colum and its number
            if ENABLE_DEBUG is True:
                Standard_Functions.Print("Start column is: " + str(StartColumn), "Log")
                Standard_Functions.Print(
                    "Column number is: " + str(Standard_Functions.GetNumberFromLetter(StartColumn)), "Log"
                )
            Column2 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 1), True
            )
            Column3 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 2), True
            )
            Column4 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 3), True
            )
            Column5 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 4), True
            )

            # Get the start row
            StartRow = EXTERNAL_SOURCE_STARTCELL[1:2]
            # if debug mode is on, show your start row
            if ENABLE_DEBUG is True:
                Standard_Functions.Print("the start row is: " + str(StartRow), "Log")

            # import the headers from the excelsheet into the spreadsheet
            sheet.set("A1", str(ws[str(StartColumn) + str(StartRow)].value))
            sheet.set("B1", str(ws[str(Column2) + str(StartRow)].value))
            sheet.set("C1", str(ws[str(Column3) + str(StartRow)].value))
            sheet.set("D1", str(ws[str(Column4) + str(StartRow)].value))
            sheet.set("E1", str(ws[str(Column5) + str(StartRow)].value))

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
            MapData(sheet=sheet)

            # Run the def to add document information
            MapDocInfo(sheet=sheet)

            # Run the def to add extra system data. This is the final value of "RowNumber" minus the "StartRow".
            AddExtraData(sheet, RowNumber - int(StartRow))

            # Format the spreadsheet
            extraRows = 0
            if INCLUDE_LENGTH is True:
                extraRows = extraRows + 1
            if INCLUDE_ANGLE is True:
                extraRows = extraRows + 1
            if INCLUDE_MASS is True:
                extraRows = extraRows + 1
            if INCLUDE_NO_SHEETS is True:
                extraRows = extraRows + 1
            FormatTable(sheet=sheet, Endrow=RowNumber + extraRows - 1)

            # Finally recompute the spreadsheet
            sheet.recompute()
            App.ActiveDocument.recompute()

        else:
            Standard_Functions.Mbox("External source is not enabled!", "", 0)
    except Exception as e:
        Text = "TitleBlock Workbench: an error occurred!!\n"
        if ENABLE_DEBUG is True:
            Text = "TitleBlock Workbench: an error occurred!!\n" + "See the report view for details"
            raise e
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
    return


def Start(command):
    try:
        sheet = sheet = App.ActiveDocument.getObject("TitleBlock")
        # check if the result is not empty
        if sheet is not None:
            # Proceed with the macro.
            if command == "FillSpreadsheet":
                FillSheet()
            if command == "ImportExcel":
                ImportDataExcel()

            # if the debug mode is on, report presense of titleblock spreadsheet
            if ENABLE_DEBUG is True:
                Standard_Functions.Print("TitleBlock already present", "Log")

            return
        # if the result is empty, create a new titleblock spreadsheet
        if sheet is None:
            sheet = App.ActiveDocument.addObject("Spreadsheet::Sheet", "TitleBlock")

            # Proceed with the macro.
            if command == "FillSpreadsheet":
                FillSheet()
            if command == "ImportExcel":
                ImportDataExcel()
            # if the debug mode is on, report creation of titleblock spreadsheet
            if ENABLE_DEBUG is True:
                Standard_Functions.Print("TitleBlock created", "Log")

            return
    except Exception as e:
        if ENABLE_DEBUG is True:
            raise e
    return
