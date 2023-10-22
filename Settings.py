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
import Standard_Functions

# region defenitions
translate = FreeCAD.Qt.translate
preferences = FreeCAD.ParamGet(
    "User parameter:BaseApp/Preferences/TechDrawTitleBlockUtility"
)
# endregion

# region -- All settings from the UI
# Included units
INCLUDE_LENGTH = preferences.GetBool("IncludeLength")
INCLUDE_ANGLE = preferences.GetBool("IncludeAngle")
INCLUDE_MASS = preferences.GetBool("IncludeMass")

# Include total number of pages
INCLUDE_NO_SHEETS = preferences.GetBool("IncludeNoOfSheets")

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

# Enable debug mode. This will enable additional report messages
ENABLE_DEBUG = preferences.GetBool("EnableDebug")

# endregion


def ExportSettingsXL():
    from openpyxl import load_workbook
    from openpyxl.worksheet.datavalidation import DataValidation

    # region -- Get the workbook and sheet
    wb = load_workbook(str(EXTERNAL_SOURCE_PATH))
    # Set the sheetname with a inputbox
    Input_SheetName = str(
        Standard_Functions.Mbox(
            "Please enter the name of the worksheet", style=2, default="Settings"
        )
    )
    # Get the sheetname
    ws = ""
    counter = 0
    for sheetname in wb.sheetnames:
        if sheetname == Input_SheetName:
            ws = wb[str(Input_SheetName)]
            counter = 1
            break

    if counter == 0:
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
    ws[str(StartCell[:1] + str(TopRow + 1))] = "Use external source"
    # Write the value
    SettingValue = str(ValueCell[:1] + str(TopRow + 1))
    # Create a dropdown for the boolan
    dv = DataValidation(type="list", formula1='"True,False"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[SettingValue])
    ws[SettingValue] = str(USE_EXTERNAL_SOURCE)
    # endregion

    # Save the workbook
    wb.save(EXTERNAL_SOURCE_PATH)


def ImportSettingsXL():
    from openpyxl import load_workbook

    # Get the workbook
    wb = load_workbook(str(EXTERNAL_SOURCE_PATH))

    # Get the sheetname
    ws = ""
    counter = 0
    for sheetname in wb.sheetnames:
        if sheetname == SHEETNAME_SETTINGS_XL:
            ws = wb[str(SHEETNAME_SETTINGS_XL)]
            counter = 1
            break

    if counter == 0:
        ws = wb.create_sheet(title="TitleBlock_Settings")

    StartCell = SHEETNAME_STARTCELL_XL
