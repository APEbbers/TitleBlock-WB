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
import Settings
from Settings import AUTOFILL_TITLEBLOCK

# Define the translation
translate = App.Qt.translate


def QT_TRANSLATE_NOOP(context, text):
    return text


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
        import FillSpreadsheet

        FillSpreadsheet.Start("FillSpreadsheet")
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
        import FillTitleBlock

        FillTitleBlock.FillTitleBlock()
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


class ImportExcel_Class:
    def GetResources(self):
        return {
            "Pixmap": "ImportExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ImportExcel", "Import data from an excel workbook"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportExcel",
                "Import data from an excel workbook to the titleblock spreadsheet",
            ),
        }

    def Activated(self):
        import FillSpreadsheet
        import FillTitleBlock

        FillSpreadsheet.Start("ImportExcel")
        if AUTOFILL_TITLEBLOCK is True:
            FillTitleBlock.FillTitleBlock()
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


class ImportFreeCAD_Class:
    def GetResources(self):
        return {
            "Pixmap": "ImportFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ImportFreeCAD", "Import data from an FreeCAD file with a spreadsheet"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportFreeCAD",
                "Import data from an FreeCAD file with a spreadsheet to the titleblock spreadsheet",
            ),
        }

    def Activated(self):
        import FillSpreadsheet
        import FillTitleBlock

        FillSpreadsheet.Start("ImportFreeCAD")
        if AUTOFILL_TITLEBLOCK is True:
            FillTitleBlock.FillTitleBlock()
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


class ExportSpreadsheet_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_Excel", "Export data to an excel workbook"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_Excel",
                "Export data from the titleblock spreadsheet to an excel workbook",
            ),
        }

    def Activated(self):
        import ExportSpreadsheet

        ExportSpreadsheet.ExportSpreadSheet_Excel()
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


class ExportSpreadsheet_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_FreeCAD", "Export data to an FreeCAD file"
            ),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSpreadSheet_FreeCAD",
                "Export data from the titleblock spreadsheet to an FreeCAD file",
            ),
        }

    def Activated(self):
        import ExportSpreadsheet

        ExportSpreadsheet.ExportSpreadSheet_FreeCAD()
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


class ExportSettings_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportSettings.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("ExportSettings_Excel", "Export settings excel"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSettings_Excel",
                "Exports all settings to the external excel workbook in its own sheet",
            ),
        }

    def Activated(self):

        Settings.ExportSettings_XL()
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


class ExportSettings_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "ExportSettings_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("ExportSettings_FreeCAD", "Export settings FreeCAD"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ExportSettings_FreeCAD",
                "Exports all settings to the external FreeCAD file in its own sheet",
            ),
        }

    def Activated(self):
        Settings.ExportSettings_FreeCAD()
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


class ImportSettings_Excel_class:
    def GetResources(self):
        return {
            "Pixmap": "ImportSettings.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("ImportSettings_Excel", "Import settings excel"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportSettings_Excel",
                "Imports all settings from the external excel workbook",
            ),
        }

    def Activated(self):
        Settings.ImportSettings_XL()
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


class ImportSettings_FreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "ImportSettings_FreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("ImportSettings_FreeCAD", "Import settings_FreeCAD"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "ImportSettings_FreeCAD",
                "Imports all settings from the external FreeCAD file",
            ),
        }

    def Activated(self):
        Settings.ImportSettings_FreeCAD()
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


class OpenExcel_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("OpenExcel", "Open the excel workbook"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenExcel", "Open the excel workbook in it's default application"
            ),
        }

    def Activated(self):
        import Standard_Functions_TitleBlock as Standard_Functions
        from Settings import EXTERNAL_SOURCE_PATH

        if EXTERNAL_SOURCE_PATH.lower().endswith("xlsx") or EXTERNAL_SOURCE_PATH.lower().endswith("xlsm"):
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


class OpenFreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "OpenFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("OpenFreeCAD", "Open the FreeCAD file"),
            "ToolTip": QT_TRANSLATE_NOOP(
                "OpenExcel", "Open the FreeCAD file with the titleblock data"
            ),
        }

    def Activated(self):
        import Standard_Functions_TitleBlock as Standard_Functions
        from Settings import EXTERNAL_SOURCE_PATH

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


class NewExcel_class:
    def GetResources(self):
        return {
            "Pixmap": "NewExcel.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("NewExcel", "Create an new excel workbook"),
            "ToolTip": QT_TRANSLATE_NOOP("NewExcel", "Create an new excel workbook"),
        }

    def Activated(self):
        import CreateNewFile

        CreateNewFile.createExcel()
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


class NewFreeCAD_class:
    def GetResources(self):
        return {
            "Pixmap": "NewFreeCAD.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("NewFreeCAD", "Create a new FreeCAD file"),
            "ToolTip": QT_TRANSLATE_NOOP("NewFreeCAD", "Create an new FreeCAD file with an empty spreadsheet for the titleblock data"),
        }

    def Activated(self):
        import CreateNewFile

        CreateNewFile.CreateFreeCAD()
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


class ExpandToolbar_class:
    def GetResources(self):
        return {
            "Pixmap": "Expand Dots - #1.svg",  # the name of a svg file available in the resources
            "MenuText": QT_TRANSLATE_NOOP("ExpandToolbar", "Expand toolbar"),
            "ToolTip": QT_TRANSLATE_NOOP("ExpandToolbar", "Expand toolbar"),
        }

    def Activated(self):
        import CreateUI

        CreateUI.toggleToolbars(ToolbarName="TitleBlock extra")
        return

    # def IsActive(self):
    #     """Here you can define if the command must be active or not (greyed) if certain conditions
    #     are met or not. This function is optional."""

    #     return result


# Add the commands to the Gui
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
