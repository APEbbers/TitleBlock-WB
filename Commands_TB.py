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

# FreeCAD init script of the Work Features module
import FreeCAD as App
import FreeCADGui as Gui
import Settings_TB
from Settings_TB import AUTOFILL_TITLEBLOCK

# Define the translation
translate = App.Qt.translate


def QT_TRANSLATE_NOOP(context, text):
    return text


# region - Core functions
# Export data from the titleblock to the spreadsheet
class FillSpreadsheet_Class:
    def GetResources(self):
        return {
            "Pixmap": "FillSpreadsheet.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "FillSpreadsheet", "Import data from titleblock"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "FillSpreadsheet",
                "Import data from titleblock to the titleblock spreadsheet",
            ),
        }

    def Activated(self):
        import FillSpreadsheet_TB

        FillSpreadsheet_TB.Start("FillSpreadsheet")
        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        # Set the default state
        result = False
        # Get for the active document.
        ActiveDoc = App.activeDocument()
        if ActiveDoc is not None:
            # Check if the document has any pages. If so the result is True and the command is activated.
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
            if pages is not None:
                result = True

        return result


# Import data from the spreadsheet to all pages
class FillTitleBlock_Class:
    def GetResources(self):
        return {
            "Pixmap": "FillTitleBlock.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("FillTitleBlock", "Populate titleblock"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "FillTitleBlock",
                "Imports data from the spreadsheet to titleblock of all pages",
            ),
        }

    def Activated(self):
        import FillTitleBlock_TB

        FillTitleBlock_TB.FillTitleBlock()

        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        # Set the default state
        result = False
        # Get for the active document.
        ActiveDoc = App.activeDocument()
        if ActiveDoc is not None:
            # Check if the document has any pages. If so the result is True and the command is activated.
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
            if pages is not None:
                result = True

        return result


# endregion


# region - External source commands
# Import data from the excel source to all pages
class ImportExcel_Class:
    def GetResources(self):
        return {
            "Pixmap": "ImportExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ImportExcel", "Import data from the excel source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportExcel",
                "Import data from the excel source file to the titleblock spreadsheet",
            ),
        }

    def Activated(self):
        import FillSpreadsheet_TB
        import FillTitleBlock_TB

        FillSpreadsheet_TB.Start("ImportExcel")
        if AUTOFILL_TITLEBLOCK is True:
            FillTitleBlock_TB.FillTitleBlock()
        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        # Set the default state
        result = False
        # Get for the active document.
        ActiveDoc = App.activeDocument()
        if ActiveDoc is not None:
            # Check if the document has any pages. If so the result is True and the command is activated.
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
            if pages is not None:
                result = True

        return result


# Import data from the FreeCAD source to all pages
class ImportFreeCAD_Class:
    def GetResources(self):
        return {
            "Pixmap": "ImportFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ImportFreeCAD", "Import data from the FreeCAD source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportFreeCAD",
                "Import data from the FreeCAD source file to the titleblock spreadsheet",
            ),
        }

    def Activated(self):
        import FillSpreadsheet_TB
        import FillTitleBlock_TB

        FillSpreadsheet_TB.Start("ImportFreeCAD")
        if AUTOFILL_TITLEBLOCK is True:
            FillTitleBlock_TB.FillTitleBlock()
        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        # Set the default state
        result = False
        # Get for the active document.
        ActiveDoc = App.activeDocument()
        if ActiveDoc is not None:
            # Check if the document has any pages. If so the result is True and the command is activated.
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
            if pages is not None:
                result = True

        return result


# Export data to an excel source file
class ExportSpreadsheet_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_Excel", "Export data to the excel source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_Excel",
                "Export data from the titleblock spreadsheet to the excel source file",
            ),
        }

    def Activated(self):
        import ExportSpreadsheet_TB

        ExportSpreadsheet_TB.ExportSpreadSheet_Excel()
        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        # Set the default state
        result = False
        # Get for the active document.
        ActiveDoc = App.activeDocument()
        if ActiveDoc is not None:
            # Check if the document has any pages. If so the result is True and the command is activated.
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
            if pages is not None:
                result = True

        return result


# Export data to an FreeCAD source file
class ExportSpreadsheet_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_FreeCAD", "Export data to the FreeCAD source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_FreeCAD",
                "Export data from the titleblock spreadsheet to the FreeCAD source file",
            ),
        }

    def Activated(self):
        import ExportSpreadsheet_TB

        ExportSpreadsheet_TB.ExportSpreadSheet_FreeCAD()
        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        # Set the default state
        result = False
        # Get for the active document.
        ActiveDoc = App.activeDocument()
        if ActiveDoc is not None:
            # Check if the document has any pages. If so the result is True and the command is activated.
            pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")
            if pages is not None:
                result = True

        return result


# Export settings to an excel source file
class ExportSettings_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportSettings.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ExportSettings_Excel", "Export settings to the excel source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSettings_Excel",
                "Exports all settings to the external excel workbook in its own sheet",
            ),
        }

    def Activated(self):
        Settings_TB.ExportSettings_XL()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE
    #     from Settings import IMPORT_SETTINGS_XL

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True and IMPORT_SETTINGS_XL is True:
    #         result = True

    #     return result


# Export settings to an FreeCAD source file
class ExportSettings_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportSettings_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ExportSettings_FreeCAD", "Export settings to the FreeCAD source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSettings_FreeCAD",
                "Exports all settings to the external FreeCAD file in its own sheet",
            ),
        }

    def Activated(self):
        Settings_TB.ExportSettings_FreeCAD()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE
    #     from Settings import IMPORT_SETTINGS_XL

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True and IMPORT_SETTINGS_XL is True:
    #         result = True

    #     return result


# Import the settings from an excel source file
class ImportSettings_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "ImportSettings.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ImportSettings_Excel", "Import settings from the excel source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportSettings_Excel",
                "Imports all settings from the external excel workbook",
            ),
        }

    def Activated(self):
        Settings_TB.ImportSettings_XL()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE
    #     from Settings import IMPORT_SETTINGS_XL

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True and IMPORT_SETTINGS_XL is True:
    #         result = True

    #     return result


# Import the settings from an FreeCAD source file
class ImportSettings_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "ImportSettings_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ImportSettings_FreeCAD", "Import settings from the FreeCAD source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportSettings_FreeCAD",
                "Imports all settings from the external FreeCAD file",
            ),
        }

    def Activated(self):
        Settings_TB.ImportSettings_FreeCAD()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE
    #     from Settings import IMPORT_SETTINGS_XL

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True and IMPORT_SETTINGS_XL is True:
    #         result = True

    #     return result


# Open the excel source file
class OpenExcel_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("OpenExcel", "Open the excel source file"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenExcel", "Open the excel workbook in it's default application"
            ),
        }

    def Activated(self):
        import Standard_Functions_TB as Standard_Functions
        from Settings_TB import EXTERNAL_SOURCE_PATH

        if EXTERNAL_SOURCE_PATH.lower().endswith(
            "xlsx"
        ) or EXTERNAL_SOURCE_PATH.lower().endswith("xlsm"):
            Standard_Functions.OpenFile(EXTERNAL_SOURCE_PATH)
        else:
            Standard_Functions.Print("Not an excel file!!", "Error")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    # return result


# Open the FreeCAD source file
class OpenFreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "OpenFreeCAD", "Open the FreeCAD source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenExcel", "Open the FreeCAD file with the titleblock data"
            ),
        }

    def Activated(self):
        import Standard_Functions_TB as Standard_Functions
        from Settings_TB import EXTERNAL_SOURCE_PATH

        if EXTERNAL_SOURCE_PATH.lower().endswith("fcstd"):
            Standard_Functions.OpenFreeCADFile(EXTERNAL_SOURCE_PATH)
        else:
            Standard_Functions.Print("Not an FreeCAD file!!", "Error")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


# Create a new empty excel source file
class NewExcel_class:
    def GetResources(self):
        return {
            "Pixmap": "NewExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "NewExcel", "Create an new excel titleblock source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewExcel", "Create an new excel titleblock source file"
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateTitleBlockData_Excel()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


# Create a new empty FreeCAD source file
class NewFreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "NewFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "NewFreeCAD", "Create a new FreeCAD source file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewFreeCAD",
                "Create an new FreeCAD file with an empty spreadsheet for the titleblock data and settings",
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateTitleBlockData_FreeCAD()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


# endregion


# region - Drawing list commands
class NewDrawingList_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "CreateDrawingList_Excel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("NewExcel", "Create a drawing list in Excel"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewExcel", "Create an new drawing list in an excel workbook"
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateSimpleDrawingList_Excel()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


class NewDrawingList_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "CreateDrawingList_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "NewFreeCAD", "Create a new drawing list in FreeCAD"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewFreeCAD",
                "Create an new empty drawing list in a FreeCAD file",
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateSimpleDrawingList_FreeCAD()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


class NewDrawingList_Internal_class:
    def GetResources(self):
        return {
            "Pixmap": "CreateDrawingList_Internal.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "NewFreeCAD", "Create a new drawing list spreadsheet"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewFreeCAD",
                "Create an new empty drawing list in a FreeCAD file",
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateSimpleDrawingList_Internal()

        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


class OpenDrawingList_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenDrawingList_Excel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "OpenDrawingList_Excel", "Open the excel drawing list"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenDrawingList_Excel",
                "Open the excel drawing list in it's default application",
            ),
        }

    def Activated(self):
        import Standard_Functions_TB as Standard_Functions
        from Settings_TB import EXTERNAL_FILE_SIMPLE_LIST

        if EXTERNAL_FILE_SIMPLE_LIST.lower().endswith(
            "xlsx"
        ) or EXTERNAL_FILE_SIMPLE_LIST.lower().endswith("xlsm"):
            Standard_Functions.OpenFile(EXTERNAL_FILE_SIMPLE_LIST)
        else:
            Standard_Functions.Print("Not an excel file!!", "Error")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    # return result


class OpenDrawingList_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenDrawingList_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "OpenDrawingList_FreeCAD", "Open the FreeCAD drawing list"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenDrawingList_FreeCAD", "Open the FreeCAD drawing list."
            ),
        }

    def Activated(self):
        import Standard_Functions_TB as Standard_Functions
        from Settings_TB import EXTERNAL_FILE_SIMPLE_LIST

        if EXTERNAL_FILE_SIMPLE_LIST.lower().endswith("fcstd"):
            Standard_Functions.OpenFreeCADFile(EXTERNAL_FILE_SIMPLE_LIST)
        else:
            Standard_Functions.Print("Not an FreeCAD file!!", "Error")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


class NewDrawingList_Advanced_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "CreateDrawingList_Advanced_Excel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("NewExcel", "Create a drawing list in Excel"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewExcel", "Create an new drawing list in an excel workbook"
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateAdvancedDrawingList_Excel()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


class NewDrawingList_Advanced_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "CreateDrawingList_Advanced_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "NewFreeCAD", "Create a new drawing list in FreeCAD"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewFreeCAD",
                "Create an new empty drawing list in a FreeCAD file",
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateAdvancedDrawingList_FreeCAD()
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


class NewDrawingList_Advanced_Internal_class:
    def GetResources(self):
        return {
            "Pixmap": "CreateDrawingList_Advanced_Internal.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "NewFreeCAD", "Create a new drawing list spreadsheet"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "NewFreeCAD",
                "Create an new empty drawing list in a FreeCAD file",
            ),
        }

    def Activated(self):
        import CreateNewFile_TB

        CreateNewFile_TB.CreateAdvancedDrawingList_Internal()

        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


class OpenDrawingList_Advanced_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenDrawingList_Advanced_Excel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "OpenDrawingList_Excel", "Open the excel drawing list"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenDrawingList_Excel",
                "Open the excel drawing list in it's default application",
            ),
        }

    def Activated(self):
        import Standard_Functions_TB as Standard_Functions
        from Settings_TB import EXTERNAL_FILE_ADVANCED_LIST

        if EXTERNAL_FILE_ADVANCED_LIST.lower().endswith(
            "xlsx"
        ) or EXTERNAL_FILE_ADVANCED_LIST.lower().endswith("xlsm"):
            Standard_Functions.OpenFile(EXTERNAL_FILE_ADVANCED_LIST)
        else:
            Standard_Functions.Print("Not an excel file!!", "Error")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    # return result


class OpenDrawingList_Advanced_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenDrawingList_Advanced_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "OpenDrawingList_FreeCAD", "Open the FreeCAD drawing list"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenDrawingList_FreeCAD", "Open the FreeCAD drawing list."
            ),
        }

    def Activated(self):
        import Standard_Functions_TB as Standard_Functions
        from Settings_TB import EXTERNAL_FILE_ADVANCED_LIST

        if EXTERNAL_FILE_ADVANCED_LIST.lower().endswith("fcstd"):
            Standard_Functions.OpenFreeCADFile(EXTERNAL_FILE_ADVANCED_LIST)
        else:
            Standard_Functions.Print("Not an FreeCAD file!!", "Error")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""
    #     from Settings import USE_EXTERNAL_SOURCE

    #     # Set the default state
    #     result = False
    #     # Check if the use of an external source is enabeled and if it is used for importing the settings
    #     if USE_EXTERNAL_SOURCE is True:
    #         result = True

    #     return result


# endregion


# region - additional commands
class ExpandToolbar_class:
    def GetResources(self):
        return {
            "Pixmap": "Expand Dots - #1.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("ExpandToolbar", "Expand toolbar"),
            "ToolTip": QT_TRANSLATE_NOOP("ExpandToolbar", "Expand toolbar"),
        }

    def Activated(self):
        import CreateUI_TB

        CreateUI_TB.toggleToolbars(ToolbarName="TitleBlock extra")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""

    #     return result


class SortFolder_AZ_class:
    def GetResources(self):
        return {
            "Pixmap": "SortGroup_AZ.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "SortFoldersAZ", "Sort the foldders ascending"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "SortFoldersAZ", "Sort the foldders ascending"
            ),
        }

    def Activated(self):
        import DrawingList_Functions_TB

        DrawingList_Functions_TB.SortGroups(False)
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""

    #     return result


class SortFolder_ZA_class:
    def GetResources(self):
        return {
            "Pixmap": "SortGroup_ZA.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "SortFoldersZA", "Sort the foldders descending"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "SortFoldersZA", "Sort the foldders descending"
            ),
        }

    def Activated(self):
        import DrawingList_Functions_TB

        DrawingList_Functions_TB.SortGroups(True)
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""

    #     return result


# endregion


# region - Add the commands to the Gui
Gui.addCommand("FillSpreadsheet", FillSpreadsheet_Class())
Gui.addCommand("FillTitleBlock", FillTitleBlock_Class())
Gui.addCommand("ImportExcel", ImportExcel_Class())
Gui.addCommand("ExportSpreadSheet_Excel", ExportSpreadsheet_Excel_class())
Gui.addCommand("ExportSettings_Excel", ExportSettings_Excel_class())
Gui.addCommand("ImportSettings_Excel", ImportSettings_Excel_class())
Gui.addCommand("OpenExcel", OpenExcel_class())
Gui.addCommand("NewExcel", NewExcel_class())
Gui.addCommand("ImportFreeCAD", ImportFreeCAD_Class())
Gui.addCommand("ExportSpreadSheet_FreeCAD", ExportSpreadsheet_FreeCAD_class())
Gui.addCommand("ExportSettings_FreeCAD", ExportSettings_FreeCAD_class())
Gui.addCommand("ImportSettings_FreeCAD", ImportSettings_FreeCAD_class())
Gui.addCommand("OpenFreeCAD", OpenFreeCAD_class())
Gui.addCommand("NewFreeCAD", NewFreeCAD_class())
Gui.addCommand("ExpandToolbar", ExpandToolbar_class())
Gui.addCommand("NewSimpleList_Excel", NewDrawingList_Excel_class())
Gui.addCommand("NewSimpleList_FreeCAD", NewDrawingList_FreeCAD_class())
Gui.addCommand("NewSimpleList_Internal", NewDrawingList_Internal_class())
Gui.addCommand("OpenSimpleList_Excel", OpenDrawingList_Excel_class())
Gui.addCommand("OpenSimpleList_FreeCAD", OpenDrawingList_FreeCAD_class())
Gui.addCommand("NewAdvancedList_Excel", NewDrawingList_Advanced_Excel_class())
Gui.addCommand("NewAdvancedList_FreeCAD", NewDrawingList_Advanced_FreeCAD_class())
Gui.addCommand("NewAdvancedList_Internal", NewDrawingList_Advanced_Internal_class())
Gui.addCommand("OpenAdvancedList_Excel", OpenDrawingList_Advanced_Excel_class())
Gui.addCommand("OpenAdvancedList_FreeCAD", OpenDrawingList_Advanced_FreeCAD_class())
Gui.addCommand("SortGroup_AZ", SortFolder_AZ_class())
Gui.addCommand("SortGroup_ZA", SortFolder_ZA_class())
# endregion
