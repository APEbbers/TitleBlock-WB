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


def ImportSettingsXL():
    from openpyxl import load_workbook
