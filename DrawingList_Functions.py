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

# Additional functions to map properties based on values in a drawing list

# region imports
import os
import FreeCAD as App
import Standard_Functions_TitleBlock as Standard_Functions
import TableFormat_Functions

# Get the settings
from Settings import ENABLE_DEBUG
from Settings import USE_SIMPLE_LIST
from Settings import USE_EXTERNAL_SOURCE_SIMPLE_LIST
from Settings import EXTERNAL_FILE_SIMPLE_LIST
from Settings import SHEETNAME_SIMPLE_LIST
from Settings import STARTCELL_SIMPLE_LIST
from Settings import PROPERTY_NAME_SIMPLE_LIST
from Settings import PROPERTY_NAME_TITLEBLOCK_SIMPLE_LIST
from Settings import USE_PAGE_NAMES_SIMPLE_LIST
from Settings import AUTOFILL_TITLEBLOCK_SIMPLE_LIST
from Settings import USE_ADVANCED_LIST
from Settings import EXTERNAL_FILE_ADVANCED_LIST
from Settings import SHEETNAME_ADVANCED_LIST
from Settings import STARTCELL_ADVANCED_LIST
from Settings import PROPERTY_NAME_ADVANCED_LIST
from Settings import PROPERTY_NAME_TITLEBLOCK_ADVANCED_LIST
from Settings import SORTING_PREFIX_ADVANCED_LIST
from Settings import USE_EXTERNAL_SOURCE_ADVANCED_LIST
from Settings import AUTOFILL_TITLEBLOCK_ADVANCED_LIST
from Settings import preferences

# Define the translation
translate = App.Qt.translate

# endregion


def MapSimpleDrawingList(sheet):
    # Check if it is allowed to use an external source and if so, continue
    if USE_SIMPLE_LIST is True and USE_EXTERNAL_SOURCE_SIMPLE_LIST is False:
        try:
            # get the spreadsheet "TitleBlock"
            if sheet is None:
                sheet = App.ActiveDocument.getObject("TitleBlock")
            # Get the Drawing list.
            DrawingSheet = App.ActiveDocument.getObject(SHEETNAME_SIMPLE_LIST)
            if DrawingSheet is None:
                SheetName = Standard_Functions.Mbox(text="DrawingList", title="Title block workbench", style=20)
                if SheetName != "":
                    DrawingSheet = SheetName
                if SheetName == "":
                    return

            # Get the startcolumn and the other three columns from there
            StartCellExt = EXTERNAL_FILE_SIMPLE_LIST

            if (Standard_Functions.GetA1fromR1C1(StartCellExt)).strip():
                StartCellExt = Standard_Functions.GetA1fromR1C1(StartCellExt)
            StartColumnExt = Standard_Functions.RemoveNumbersFromString(StartCellExt)
            # If debug mode is on, show the start colum and its number
            if ENABLE_DEBUG is True:
                Standard_Functions.Print(
                    translate(
                        "TitleBlock Workbench",
                        "Start column for the drawing list is: " + str(StartColumnExt),
                    ),
                    "Log",
                )
                Standard_Functions.Print(
                    translate(
                        "TitleBlock Workbench",
                        "Column number for the drawing list is: "
                        + str(Standard_Functions.GetNumberFromLetter(StartColumnExt)),
                    ),
                    "Log",
                )
            Column2 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumnExt) + 1), True
            )

            # Get the start row
            StartRow = Standard_Functions.RemoveLettersFromString(
                EXTERNAL_FILE_ADVANCED_LIST
            )
            # if debug mode is on, show your start row
            if ENABLE_DEBUG is True:
                Text = translate(
                    "TitleBlock Workbench", "the start row for the drawing list is: " + str(StartRow)
                )
                Standard_Functions.Print(Text, "Log")

            # Define place holders for the property name and value in the excel list.
            PropertyNameExt = ""
            PropertyValueExt = ""

            # Get the pages in the document
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")

            # Go through the excel list and collect the property value based on the property name searched for.
            for i in range(1000):
                # Define the start row. This is the Header row + 1 + i as counter
                RowNumber = int(StartRow) + i + 1

                # Get the property name in the excel list. If it starts with "'", remove it
                PropertyNameExt = str(DrawingSheet.getContents(f"{StartColumnExt}{RowNumber}"))
                if PropertyNameExt[:1] == "'":
                    PropertyNameExt = PropertyNameExt[1:]

                # Get the property value in the excel list. If it starts with "'", remove it
                PropertyValueExt = str(DrawingSheet.getContents(f"{Column2}{RowNumber}"))
                if PropertyValueExt[:1] == "'":
                    PropertyValueExt = PropertyValueExt[1:]

                # If page names are not to be mapped, go here
                if USE_PAGE_NAMES_SIMPLE_LIST is False:
                    # If the property name in the drawing list matches the property name that
                    # must be search for in the title block, continue.
                    if PropertyNameExt == PROPERTY_NAME_SIMPLE_LIST:
                        # Go through all the pages
                        for page in pages:
                            # Get the editable texts
                            texts = page.Template.EditableTexts
                            # Get the property name in the titleblock spreadsheet and fill it with the property value from the drawing list
                            texts[PROPERTY_NAME_TITLEBLOCK_SIMPLE_LIST] = PropertyValueExt
                        return

                # If page names are to be mapped, go here
                if USE_PAGE_NAMES_SIMPLE_LIST is True:
                    # Go through the pages
                    for page in pages:
                        # If the Property name in the drawing list matches the page label:
                        # Fill in the desired editable text with the property value from the drawing list
                        if PropertyNameExt == page.Label:
                            # Get the editable texts
                            texts = page.Template.EditableTexts
                            texts[PROPERTY_NAME_TITLEBLOCK_SIMPLE_LIST] = PropertyValueExt
                            return

                # If the property name in the drawing list is empty, it is the end of the list.
                # Exit the function
                if PropertyNameExt == "":
                    break

            # Recomute the document
            App.ActiveDocument.recompute(None, True, True)
        except Exception as e:
            Text = translate(
                "TitleBlock Workbench", "TitleBlock Workbench: an error occurred!!\n"
            )
            if ENABLE_DEBUG is True:
                Text = translate(
                    "TitleBlock Workbench",
                    "TitleBlock Workbench: an error occurred!!\n"
                    + "See the report view for details",
                )
                raise e
            Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)


def MapSimpleDrawingList_Excel(sheet):
    from openpyxl import load_workbook

    # Check if it is allowed to use an external source and if so, continue
    if USE_SIMPLE_LIST is True and USE_EXTERNAL_SOURCE_SIMPLE_LIST is True:
        # if debug mode is enabled, show the external file including path.
        if ENABLE_DEBUG is True:
            Text = translate("TitleBlock Workbench", f"The drawing list is: {EXTERNAL_FILE_SIMPLE_LIST}")
            Standard_Functions.Print(Text, "Log")

        # try to open the source. if not show an messagebox and if debug mode is enabled, show the exeption as well
        try:
            wb = ""
            if EXTERNAL_FILE_SIMPLE_LIST.lower().endswith(".fcstd"):
                Filter = [
                    ("Excel", "*.xlsx"),
                ]
                FileName = Standard_Functions.GetFileDialog(files=Filter, SaveAs=False)
                if FileName != "":
                    wb = load_workbook(FileName, read_only=True, data_only=True)
                if FileName == "":
                    return
            else:
                wb = load_workbook(
                    str(EXTERNAL_FILE_SIMPLE_LIST), read_only=True, data_only=True
                )
            if SHEETNAME_SIMPLE_LIST == "":
                # Set the sheetname with a inputbox
                Worksheets_List = [i for i in wb.sheetnames if i != "Settings"]
                Text = translate(
                    "TitleBlock Workbench", "Please enter the name of the worksheet"
                )
                Input_SheetName = str(
                    Standard_Functions.Mbox(
                        text=Text,
                        title="TitleBlock Workbench",
                        style=3,
                        default="TitleBlockData",
                        stringList=Worksheets_List,
                    )
                )
                # if the user canceled, exit this function.
                if not Input_SheetName.strip():
                    return
                ws = wb[str(Input_SheetName)]
            if SHEETNAME_SIMPLE_LIST != "":
                ws = wb[str(SHEETNAME_SIMPLE_LIST)]
        except IOError:
            Standard_Functions.Mbox(
                "Permission error!!\nDo you have the file open?",
                "Titleblock Workbench",
                0,
                IconType="Critical",
            )
            return
        except Exception as e:
            if ENABLE_DEBUG is True:
                raise (e)
            Text = translate(
                "TitleBlock Workbench",
                "an problem occured while openening the excel file!\nDo you have it open in an another application",
            )
            Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
            return

        try:
            # get the spreadsheet "TitleBlock"
            if sheet is None:
                sheet = App.ActiveDocument.getObject("TitleBlock")

            # Get the startcolumn and the other three columns from there
            StartCell = STARTCELL_SIMPLE_LIST
            if SHEETNAME_SIMPLE_LIST == "":
                # Set EXTERNAL_SOURCE_SHEET_NAME to the chosen sheetname
                preferences.SetString("SheetName_SimpleList", Input_SheetName)
                ws = wb[Input_SheetName]

                # Set the startcell with an inputbox
                Text = translate(
                    "TitleBlock Workbench",
                    "Please enter the name of the cell.\n"
                    + "Enter a single cell like 'A1', 'B2', etc. Other notations will be ignored!",
                )
                StartCell = str(
                    Standard_Functions.Mbox(
                        text=Text,
                        title="TitleBlock Workbench",
                        style=20,
                        default="A1",
                    )
                )
                if not StartCell.strip():
                    StartCell = "A1"

                # Set SHEETNAME_STARTCELL to the chosen sheetname
                preferences.SetString("StartCell_SimpleList", StartCell)

            if (Standard_Functions.GetA1fromR1C1(StartCell)).strip():
                StartCell = Standard_Functions.GetA1fromR1C1(StartCell)
            StartColumn = Standard_Functions.RemoveNumbersFromString(StartCell)
            # If debug mode is on, show the start colum and its number
            if ENABLE_DEBUG is True:
                Standard_Functions.Print(
                    translate(
                        "TitleBlock Workbench",
                        "Start column for the drawing list is: " + str(StartColumn),
                        "Log",
                    )
                )
                Standard_Functions.Print(
                    translate(
                        "TitleBlock Workbench",
                        "Column number for the drawing list is: "
                        + str(Standard_Functions.GetNumberFromLetter(StartColumn)),
                    ),
                    "Log",
                )
            Column2 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumn) + 1), True
            )

            # Get the start row
            StartRow = StartRow = Standard_Functions.RemoveLettersFromString(STARTCELL_SIMPLE_LIST)
            # if debug mode is on, show your start row
            if ENABLE_DEBUG is True:
                Text = translate(
                    "TitleBlock Workbench", "the start row for the drawing list is: " + str(StartRow)
                )
                Standard_Functions.Print(Text, "Log")

            # Define place holders for the property name and value in the excel list.
            PropertyNameExcel = ""
            PropertyValueExcel = ""

            # Get the pages in the document
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")

            # Go through the excel list and collect the property value based on the property name searched for.
            for i in range(1000):
                # Define the start row. This is the Header row + 1 + i as counter
                RowNumber = int(StartRow) + i + 1

                # Get the property name in the excel list. If it starts with "'", remove it
                PropertyNameExcel = ws[f"{StartColumn}{RowNumber}"].value
                if PropertyNameExcel[:1] == "'":
                    PropertyNameExcel = PropertyNameExcel[1:]

                # Get the property value in the excel list. If it starts with "'", remove it
                PropertyValueExcel = ws[f"{Column2}{RowNumber}"].value
                if PropertyValueExcel[:1] == "'":
                    PropertyValueExcel = PropertyValueExcel[1:]

                # If page names are not to be mapped, go here
                if USE_PAGE_NAMES_SIMPLE_LIST is False:
                    # If the property name in the drawing list matches the property name that
                    # must be search for in the title block, continue.
                    if PropertyNameExcel == PROPERTY_NAME_SIMPLE_LIST:
                        # Go through all the pages
                        for page in pages:
                            # Get the editable texts
                            texts = page.Template.EditableTexts
                            # Get the property name in the titleblock spreadsheet and fill it with the property value from the drawing list
                            texts[PROPERTY_NAME_TITLEBLOCK_SIMPLE_LIST] = PropertyValueExcel
                        return

                # If page names are to be mapped, go here
                if USE_PAGE_NAMES_SIMPLE_LIST is True:
                    # Go through the pages
                    for page in pages:
                        # If the Property name in the drawing list matches the page label:
                        # Fill in the desired editable text with the property value from the drawing list
                        if PropertyNameExcel == page.Label:
                            # Get the editable texts
                            texts = page.Template.EditableTexts
                            texts[PROPERTY_NAME_TITLEBLOCK_SIMPLE_LIST] = PropertyValueExcel
                            return

                # If the property name in the drawing list is empty, it is the end of the list.
                # Exit the function
                if PropertyNameExcel == "":
                    break

            # Recomute the document
            App.ActiveDocument.recompute(None, True, True)

        except Exception as e:
            Text = translate(
                "TitleBlock Workbench", "TitleBlock Workbench: an error occurred!!\n"
            )
            if ENABLE_DEBUG is True:
                Text = translate(
                    "TitleBlock Workbench",
                    "TitleBlock Workbench: an error occurred!!\n"
                    + "See the report view for details",
                )
                raise e
            Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
        finally:
            # Close the excel workbook
            wb.close()
    else:
        Text = translate("TitleBlock Workbench", "Use of a simple drawing list is not enabled!")
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
    return


def MapSimpleDrawingList_FreeCAD(sheet):
    # Check if it is allowed to use an external source and if so, continue
    if USE_SIMPLE_LIST is True and USE_EXTERNAL_SOURCE_SIMPLE_LIST is True:
        # if debug mode is enabled, show the external file including path.
        if ENABLE_DEBUG is True:
            Text = translate("TitleBlock Workbench", f"The drawing list is: {EXTERNAL_FILE_SIMPLE_LIST}")
            Standard_Functions.Print(Text, "Log")

        # Get the name of the external source
        Input_SheetName = SHEETNAME_SIMPLE_LIST
        # get the spreadsheet "TitleBlock"
        if sheet is None:
            sheet = App.ActiveDocument.getObject("TitleBlock")
        # Clear the sheet
        sheet.clearAll()
        # Define the External sheet and document
        ExtSheet = None
        # ff = None

        # try to open the source. if not show an messagebox and if debug mode is enabled, show the exeption as well
        try:
            if EXTERNAL_FILE_SIMPLE_LIST.lower().endswith(".xlsx"):
                Filter = [
                    ("FreeCAD", "*.FCStd"),
                ]
                FileName = Standard_Functions.GetFileDialog(
                    files=Filter, SaveAs=False
                )
                if FileName != "":
                    ff = App.openDocument(FileName, True)
                if FileName == "":
                    return
            else:
                ff = App.openDocument(EXTERNAL_FILE_SIMPLE_LIST, True)
            if EXTERNAL_FILE_SIMPLE_LIST == "":
                # Set the sheetname with a inputbox
                Spreadsheet_List = ff.findObjects("Spreadsheet::Sheet")
                Text = translate(
                    "TitleBlock Workbench",
                    "Please enter the name of the spreadsheet",
                )
                Input_SheetName = str(
                    Standard_Functions.Mbox(
                        text=Text,
                        title="TitleBlock Workbench",
                        style=21,
                        default="Settings",
                        stringList=Spreadsheet_List,
                    )
                )
                # if the user canceled, exit this function.
                if not Input_SheetName.strip():
                    return

                ExtSheet = ff.getObject(Input_SheetName)
            if SHEETNAME_SIMPLE_LIST != "":
                ExtSheet = ff.getObject(Input_SheetName)
        except Exception as e:
            if ENABLE_DEBUG is True:
                raise (e)
            Text = translate(
                "TitleBlock Workbench",
                "an problem occured while openening the FreeCAD file!\nDo you have it open in an another application",
            )
            Standard_Functions.Mbox(
                text=Text, title="TitleBlock Workbench", style=0
            )
            return

        try:
            # get the spreadsheet "TitleBlock"
            sheet = App.ActiveDocument.getObject("TitleBlock")

            # Get the startcolumn and the other three columns from there
            StartCellExt = EXTERNAL_FILE_SIMPLE_LIST
            if EXTERNAL_FILE_SIMPLE_LIST == "":
                # Set EXTERNAL_SOURCE_SHEET_NAME to the chosen sheetname
                preferences.SetString("SheetName", Input_SheetName)
                ExtSheet = ff.getObject(Input_SheetName)
                # Set the startcell with an inputbox
                Text = translate(
                    "TitleBlock Workbench",
                    "Please enter the name of the cell.\n"
                    + "Enter a single cell like 'A1', 'B2', etc. Other notations will be ignored!",
                )
                StartCellExt = str(
                    Standard_Functions.Mbox(
                        text=Text,
                        title="TitleBlock Workbench",
                        style=20,
                        default="A1",
                    )
                )
                if not StartCellExt.strip():
                    StartCellExt = "A1"

                # Set SHEETNAME_STARTCELL to the chosen sheetname
                preferences.SetString("StartCell", StartCellExt)

            if (Standard_Functions.GetA1fromR1C1(StartCellExt)).strip():
                StartCellExt = Standard_Functions.GetA1fromR1C1(StartCellExt)
            StartColumnExt = Standard_Functions.RemoveNumbersFromString(StartCellExt)
            # If debug mode is on, show the start colum and its number
            if ENABLE_DEBUG is True:
                Standard_Functions.Print(
                    translate(
                        "TitleBlock Workbench",
                        "Start column for the drawing list is: " + str(StartColumnExt),
                    ),
                    "Log",
                )
                Standard_Functions.Print(
                    translate(
                        "TitleBlock Workbench",
                        "Column number for the drawing list is: "
                        + str(Standard_Functions.GetNumberFromLetter(StartColumnExt)),
                    ),
                    "Log",
                )
            Column2 = Standard_Functions.GetLetterFromNumber(
                int(Standard_Functions.GetNumberFromLetter(StartColumnExt) + 1), True
            )

            # Get the start row
            StartRow = Standard_Functions.RemoveLettersFromString(
                EXTERNAL_FILE_SIMPLE_LIST
            )
            # if debug mode is on, show your start row
            if ENABLE_DEBUG is True:
                Text = translate(
                    "TitleBlock Workbench", "the start row for the drawing list is: " + str(StartRow)
                )
                Standard_Functions.Print(Text, "Log")

            # Define place holders for the property name and value in the excel list.
            PropertyNameExt = ""
            PropertyValueExt = ""

            # Get the pages in the document
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")

            # Go through the excel list and collect the property value based on the property name searched for.
            for i in range(1000):
                # Define the start row. This is the Header row + 1 + i as counter
                RowNumber = int(StartRow) + i + 1

                # Get the property name in the excel list. If it starts with "'", remove it
                PropertyNameExt = str(ExtSheet.getContents(f"{StartColumnExt}{RowNumber}"))
                if PropertyNameExt[:1] == "'":
                    PropertyNameExt = PropertyNameExt[1:]

                # Get the property value in the excel list. If it starts with "'", remove it
                PropertyValueExt = str(ExtSheet.getContents(f"{Column2}{RowNumber}"))
                if PropertyNameExt[:1] == "'":
                    PropertyNameExt = PropertyNameExt[1:]

                # If page names are not to be mapped, go here
                if USE_PAGE_NAMES_SIMPLE_LIST is False:
                    # If the property name in the drawing list matches the property name that
                    # must be search for in the title block, continue.
                    if PropertyNameExt == PROPERTY_NAME_SIMPLE_LIST:
                        # Go through all the pages
                        for page in pages:
                            # Get the editable texts
                            texts = page.Template.EditableTexts
                            # Get the property name in the titleblock spreadsheet and fill it with the property value from the drawing list
                            texts[PROPERTY_NAME_TITLEBLOCK_SIMPLE_LIST] = PropertyValueExt
                        return

                # If page names are to be mapped, go here
                if USE_PAGE_NAMES_SIMPLE_LIST is True:
                    # Go through the pages
                    for page in pages:
                        # If the Property name in the drawing list matches the page label:
                        # Fill in the desired editable text with the property value from the drawing list
                        if PropertyNameExt == page.Label:
                            # Get the editable texts
                            texts = page.Template.EditableTexts
                            texts[PROPERTY_NAME_TITLEBLOCK_SIMPLE_LIST] = PropertyValueExt
                            return

                # If the property name in the drawing list is empty, it is the end of the list.
                # Exit the function
                if PropertyNameExt == "":
                    break

            # Recomute the document
            App.ActiveDocument.recompute(None, True, True)

        except Exception as e:
            Text = translate(
                "TitleBlock Workbench", "TitleBlock Workbench: an error occurred!!\n"
            )
            if ENABLE_DEBUG is True:
                Text = translate(
                    "TitleBlock Workbench",
                    "TitleBlock Workbench: an error occurred!!\n"
                    + "See the report view for details",
                )
                raise e
            Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
    else:
        Text = translate("TitleBlock Workbench", "Use of a simple drawing list is not enabled!")
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
    return


# def MapAdvancedDrawingList_Excel():


# def MapAdvancedDrawingList_FreeCAD():
