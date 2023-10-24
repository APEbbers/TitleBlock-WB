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

import os
import sys
import FreeCAD
import FreeCADGui as Gui
import Standard_Functions
import FreeCADGui as Gui

# region defenitions
translate = FreeCAD.Qt.translate
preferences = FreeCAD.ParamGet(
    "User parameter:BaseApp/Preferences/TechDrawTitleBlockUtility"
)
# endregion

# region -- All settings from the UI

# External source
USE_EXTERNAL_SOURCE = preferences.GetBool("UseExternalSource")
EXTERNAL_SOURCE_PATH = preferences.GetString("ExternalFile")
EXTERNAL_SOURCE_SHEET_NAME = preferences.GetString("SheetName")
EXTERNAL_SOURCE_STARTCELL = preferences.GetString("StartCell")
AUTOFILL_TITLEBLOCK = preferences.GetBool("AutoFillTitleBlock")
IMPORT_SETTINGS_XL = preferences.GetBool("ImportSettingsXL")
SHEETNAME_SETTINGS_XL = preferences.GetString("SheetName_Settings")
SHEETNAME_STARTCELL_XL = preferences.GetString("StartCell_Settings")

# Use filename as drawingnumber
USE_FILENAME_DRAW_NO = preferences.GetBool("UseFileName")
DRAW_NO_FiELD = preferences.GetString("DrwNrFieldName")

# The values that are mapped
MAP_LENGTH = preferences.GetString("MapLength")
MAP_ANGLE = preferences.GetString("MapAngle")
MAP_MASS = preferences.GetString("MapMass")
MAP_NOSHEETS = preferences.GetString("MapNoSheets")

# Included values
INCLUDE_LENGTH = preferences.GetBool("IncludeLength")
INCLUDE_ANGLE = preferences.GetBool("IncludeAngle")
INCLUDE_MASS = preferences.GetBool("IncludeMass")
INCLUDE_NO_SHEETS = preferences.GetBool("IncludeNoOfSheets")

# Enable debug mode. This will enable additional report messages
ENABLE_DEBUG = preferences.GetBool("EnableDebug")

# All the settings in a List
SettingsList = [
    USE_EXTERNAL_SOURCE,
    EXTERNAL_SOURCE_PATH,
    EXTERNAL_SOURCE_SHEET_NAME,
    EXTERNAL_SOURCE_STARTCELL,
    AUTOFILL_TITLEBLOCK,
    IMPORT_SETTINGS_XL,
    SHEETNAME_SETTINGS_XL,
    SHEETNAME_STARTCELL_XL,
    USE_FILENAME_DRAW_NO,
    DRAW_NO_FiELD,
    MAP_LENGTH,
    MAP_ANGLE,
    MAP_MASS,
    MAP_NOSHEETS,
    INCLUDE_LENGTH,
    INCLUDE_ANGLE,
    INCLUDE_MASS,
    INCLUDE_NO_SHEETS,
    ENABLE_DEBUG,
]

# endregion


def ExportSettingsXL():
    from openpyxl import load_workbook
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo

    # region -- Get the workbook, create a new sheet and set the startcell (top left cell of table).
    wb = load_workbook(str(EXTERNAL_SOURCE_PATH))
    # Set the sheetname with a inputbox
    Worksheets_List = [i for i in wb.sheetnames if i != "TitleBlockData"]
    Input_SheetName = str(
        Standard_Functions.Mbox(
            text="Please enter the name of the worksheet",
            title="",
            style=3,
            default="Settings",
            stringList=Worksheets_List,
        )
    )
    # if the user canceled, exit this function.
    if not Input_SheetName.strip():
        return

    # Set SHEETNAME_SETTINGS_XL to the chosen sheetname
    preferences.SetString("SheetName_Settings", Input_SheetName)

    # Delete the current sheet if it exists
    for sheetname in wb.sheetnames:
        if sheetname == Input_SheetName:
            del wb[str(Input_SheetName)]

            break

    # create a new sheet
    ws = wb.create_sheet(title=Input_SheetName)

    # Set the startcell with an inputbox
    StartCell = str(
        Standard_Functions.Mbox(
            "Please enter the name of the cell.\nEnter a single cell like 'A1', 'B2', etc. Other notations will be ignored!",
            style=2,
            default="A1",
        )
    )
    if not StartCell.strip():
        StartCell = "A1"

    # Set SHEETNAME_STARTCELL_XL to the chosen sheetname
    preferences.SetString("StartCell_Settings", StartCell)

    # endregion

    # region -- Create the headers
    ws[StartCell] = "Name"
    TopRow = int(StartCell[1:])
    ValueCell = str(
        Standard_Functions.GetLetterFromNumber(
            Standard_Functions.GetNumberFromLetter(StartCell[:1]) + 2
        )
    ) + str(TopRow)
    ws[ValueCell] = "Value"
    # endregion

    # region -- Export the external source settings
    #
    # USE_EXTERNAL_SOURCE
    RowNumber = 1
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "UseExternalSource"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(USE_EXTERNAL_SOURCE).upper()
    RowNumber = RowNumber + 1

    # EXTERNAL_SOURCE_PATH
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "ExternalFile"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = EXTERNAL_SOURCE_PATH
    RowNumber = RowNumber + 1

    # EXTERNAL_SOURCE_SHEET_NAME
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "SheetName"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = EXTERNAL_SOURCE_SHEET_NAME
    RowNumber = RowNumber + 1

    # EXTERNAL_SOURCE_STARTCELL
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "StartCell"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = EXTERNAL_SOURCE_STARTCELL
    RowNumber = RowNumber + 1

    # AUTOFILL_TITLEBLOCK
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "AutoFillTitleBlock"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(AUTOFILL_TITLEBLOCK).upper()
    RowNumber = RowNumber + 1

    # IMPORT_SETTINGS_XL
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "ImportSettingsXL"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(IMPORT_SETTINGS_XL).upper()
    RowNumber = RowNumber + 1

    # SHEETNAME_SETTINGS_XL
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "SheetName_Settings"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = SHEETNAME_SETTINGS_XL
    RowNumber = RowNumber + 1

    # SHEETNAME_STARTCELL_XL
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "StartCell_Settings"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = SHEETNAME_STARTCELL_XL
    RowNumber = RowNumber + 1

    # endregion

    # region -- Export the filename settings
    #
    # USE_FILENAME_DRAW_NO
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "UseFileName"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(USE_FILENAME_DRAW_NO).upper()
    RowNumber = RowNumber + 1

    # DRAW_NO_FiELD
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "DrwNrFieldName"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = DRAW_NO_FiELD
    RowNumber = RowNumber + 1

    # endregion

    # region -- Export the Mapping settings
    #
    # MAP_LENGTH
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "MapLength"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = MAP_LENGTH
    RowNumber = RowNumber + 1

    # MAP_ANGLE
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "MapAngle"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = MAP_ANGLE
    RowNumber = RowNumber + 1

    # MAP_MASS
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "MapMass"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = MAP_MASS
    RowNumber = RowNumber + 1

    # MAP_NOSHEETS
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "MapNoSheets"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    ws[SettingValue].value = MAP_NOSHEETS
    RowNumber = RowNumber + 1
    # endregion

    # region -- Export the included value settings
    #
    # INCLUDE_LENGTH
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "IncludeLength"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(INCLUDE_LENGTH).upper()
    RowNumber = RowNumber + 1

    # INCLUDE_ANGLE
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "IncludeAngle"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(INCLUDE_ANGLE).upper()
    RowNumber = RowNumber + 1

    # INCLUDE_MASS
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "IncludeMass"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(INCLUDE_MASS).upper()
    RowNumber = RowNumber + 1

    # INCLUDE_NO_SHEETS
    ws[str(StartCell[:1] + str(TopRow + RowNumber))] = "IncludeNoOfSheets"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + RowNumber))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(INCLUDE_NO_SHEETS).upper()
    RowNumber = RowNumber + 1
    # endregion

    # Note: ENABLE_DEBUG is excluded from export.
    # This is a setting only needed for debuggin and thus a per user setting.

    # region Format the settings with the values as a Table
    #
    # Define the the last cell
    EndCell = str(
        Standard_Functions.GetLetterFromNumber(
            Standard_Functions.GetNumberFromLetter(StartCell[:1]) + 2
        )
    ) + str(RowNumber + TopRow - 1)

    # Define the table
    NumberOfSheets = str(len(wb.sheetnames))
    tab = Table(
        displayName="SettingsTable_" + NumberOfSheets,
        ref=f"{StartCell}:{EndCell}",
    )

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(
        name="TableStyleMedium11",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=True,
    )

    # add the style to the table
    tab.tableStyleInfo = style

    # add the table to the worksheet
    ws.add_table(tab)
    # endregion

    # Make the columns to autofit the date
    for col in ws.columns:
        SetLen = 0
        column = col[0].column_letter  # Get the column name

        for cell in col:
            if len(str(cell.value)) > SetLen:
                SetLen = len(str(cell.value))

        set_col_width = SetLen + 5
        # Setting the column width
        ws.column_dimensions[column].width = set_col_width

    # Save the workbook
    wb.save(EXTERNAL_SOURCE_PATH)


def ImportSettingsXL():
    from openpyxl import load_workbook

    # Get the workbook
    wb = load_workbook(str(EXTERNAL_SOURCE_PATH))

    # Get the sheetname
    ws = ""
    for sheetname in wb.sheetnames:
        if sheetname == SHEETNAME_SETTINGS_XL:
            ws = wb[str(SHEETNAME_SETTINGS_XL)]
            break

    # Get the startcell
    StartCell = SHEETNAME_STARTCELL_XL

    # Get the columns
    FirstColumn = int(Standard_Functions.GetNumberFromLetter(StartCell[:1]))
    SecondColumn = FirstColumn + 1

    # go through the excel until all settings are imported.
    counter = 0
    for i in range(1, 1000):
        Cell_Name = ws.cell(i, FirstColumn)
        Cell_Value = ws.cell(i, SecondColumn)

        # region -- Import the external source settings
        #
        # Import USE_EXTERNAL_SOURCE
        if Cell_Name.value == "UseExternalSource":
            preferences.SetBool("UseExternalSource", bool(Cell_Value.value))
            counter = counter + 1

        # Import EXTERNAL_SOURCE_PATH
        if Cell_Name.value == "ExternalFile":
            preferences.SetString("ExternalFile", str(Cell_Value.value))
            counter = counter + 1

        # Import EXTERNAL_SOURCE_SHEET_NAME
        if Cell_Name.value == "SheetName":
            preferences.SetString("SheetName", str(Cell_Value.value))
            counter = counter + 1

        # Import EXTERNAL_SOURCE_STARTCELL
        if Cell_Name.value == "StartCell":
            preferences.SetString("StartCell", str(Cell_Value.value))
            counter = counter + 1

        # Import AUTOFILL_TITLEBLOCK
        if Cell_Name.value == "AutoFillTitleBlock":
            preferences.SetBool("AutoFillTitleBlock", bool(Cell_Value.value))
            counter = counter + 1

        # Import IMPORT_SETTINGS_XL
        if Cell_Name.value == "ImportSettingsXL":
            preferences.SetBool("ImportSettingsXL", bool(Cell_Value.value))
            counter = counter + 1

        # Import SHEETNAME_SETTINGS_XL
        if Cell_Name.value == "SheetName_Settings":
            preferences.SetString("SheetName_Settings", str(Cell_Value.value))
            counter = counter + 1

        # Import SHEETNAME_STARTCELL_XL
        if Cell_Name.value == "StartCell_Settings":
            preferences.SetString("StartCell_Settings", str(Cell_Value.value))
            counter = counter + 1

        # endregion

        # region -- Import the filename settings
        #
        # Import USE_FILENAME_DRAW_NO
        if Cell_Name.value == "UseFileName":
            preferences.SetBool("UseFileName", bool(Cell_Value.value))
            counter = counter + 1

        # Import DRAW_NO_FiELD
        if Cell_Name.value == "DrwNrFieldName":
            preferences.SetString("DrwNrFieldName", str(Cell_Value.value))
            counter = counter + 1

        # endregion

        # region -- Import the mapping settings
        #
        # Import MAP_LENGTH
        if Cell_Name.value == "MapLength":
            preferences.SetString("MapLength", str(Cell_Value.value))
            counter = counter + 1

        # Import MAP_ANGLE
        if Cell_Name.value == "MapAngle":
            preferences.SetString("MapAngle", str(Cell_Value.value))
            counter = counter + 1

        # Import MAP_MASS
        if Cell_Name.value == "MAP_MASS":
            preferences.SetString("MapMass", str(Cell_Value.value))
            counter = counter + 1

        # Import MAP_NOSHEETS
        if Cell_Name.value == "MapNoSheets":
            preferences.SetString("MapNoSheets", str(Cell_Value.value))
            counter = counter + 1

        # endregion

        # region -- Import the mapping settings
        #
        # Import INCLUDE_LENGTH
        if Cell_Name.value == "IncludeLength":
            preferences.SetBool("IncludeLength", bool(Cell_Value.value))
            counter = counter + 1

        # Import INCLUDE_ANGLE
        if Cell_Name.value == "IncludeAngle":
            preferences.SetBool("IncludeAngle", bool(Cell_Value.value))
            counter = counter + 1

        # Import INCLUDE_MASS
        if Cell_Name.value == "IncludeMass":
            preferences.SetBool("IncludeMass", bool(Cell_Value.value))
            counter = counter + 1

        # Import INCLUDE_NO_SHEETS
        if Cell_Name.value == "IncludeNoOfSheets":
            preferences.SetBool("IncludeNoOfSheets", bool(Cell_Value.value))
            counter = counter + 1

        # endregion

        # Note: ENABLE_DEBUG is excluded from import.
        # This is a setting only needed for debuggin and thus a per user setting.

        if counter == len(SettingsList) + 1:
            break

    if counter > 0:
        print(
            f"Titleblock workbench settings imported from {EXTERNAL_SOURCE_PATH} and sheet: {sheetname} at {StartCell}"
        )
    if counter == 0:
        print("Settings are not imported")
