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
import os
import FreeCAD as App
import FreeCADGui as Gui
from inspect import getsourcefile
from Standard_Functions_TitleBlock import (
    StandardFunctions_FreeCAD as Standard_Functions,
)

__title__ = "TitleBlock Workbench"
__author__ = "A.P. Ebbers"
__url__ = "https://github.com/APEbbers/TechDrawTitleBlockUtility.git"

# get the path of the current python script
PATH_TB = file_path = os.path.dirname(getsourcefile(lambda: 0))

global PATH_TB_ICONS
global PATH_TB_RESOURCES
global PATH_TB_UI

PATH_TB_ICONS = os.path.join(PATH_TB, "Resources", "Icons").replace("\\", "/")
PATH_TB_RESOURCES = os.path.join(PATH_TB, "Resources").replace("\\", "/")
PATH_TB_UI = os.path.join(PATH_TB, PATH_TB_RESOURCES, "UI").replace("\\", "/")


class TitleBlockWB(Gui.Workbench):
    MenuText = "TitleBlock Workbench"
    ToolTip = "An extension for the TechDraw workbench to fill a TitleBlock"
    Icon = os.path.join(PATH_TB_ICONS, "TitleBlockWB.svg").replace("\\", "/")

    Gui.addIconPath(PATH_TB_ICONS)
    Gui.addPreferencePage(
        os.path.join(PATH_TB_UI, "PreferenceUI.ui"),
        "TitleBlock Workbench",
    )

    def GetClassName(self):
        # This function is mandatory if this is a full Python workbench
        # This is not a template, the returned string should be exactly "Gui::PythonWorkbench"
        return "Gui::PythonWorkbench"

    def Initialize(self):
        """This function is executed when the workbench is first activated.
        It is executed once in a FreeCAD session followed by the Activated function.
        """
        import TitleBlock_Commands  # import here all the needed files that create your FreeCAD commands
        import Settings
        from Settings import USE_EXTERNAL_SOURCE
        from Settings import IMPORT_SETTINGS_XL
        from Settings import ADD_TOOLBAR_TECHDRAW
        import CreateUI
        import TechDrawFunctions

        if IMPORT_SETTINGS_XL is True:
            Settings.ImportSettingsXL()

        if USE_EXTERNAL_SOURCE is True:
            ToolbarList = self.list = [
                "ImportExcel",
                "FillTitleBlock",
            ]
        if USE_EXTERNAL_SOURCE is False:
            ToolbarList = self.list = [
                "FillSpreadsheet",
                "FillTitleBlock",
            ]  # a list of command names created in the line above
        self.appendToolbar(
            "TitleBlock", ToolbarList
        )  # creates a new toolbar with your commands

        StandardList = self.list = [
            "FillTitleBlock",
            "FillSpreadsheet",
        ]
        ExcelList = self.list = [
            "ExportSpreadSheet",
            "ImportExcel",
            "OpenExcel",
        ]
        SettingsList = self.list = [
            "ExportSettings",
            "ImportSettings",
        ]
        self.appendMenu(
            "TitleBlock",
            StandardList,
        )  # creates a new menu
        self.appendMenu(["TitleBlock", "Excel functions"], ExcelList)
        self.appendMenu(["TitleBlock", "Settings"], SettingsList)

        if ADD_TOOLBAR_TECHDRAW is True:
            CreateUI.CreateTechDrawToolbar()
        if ADD_TOOLBAR_TECHDRAW is False:
            CreateUI.RemoveTechDrawToolbar()

        TechDrawFunctions.ImportTemplates()
        TechDrawFunctions.SetDefaultTemplate()

    def Activated(self):
        """This function is executed whenever the workbench is activated"""

        import TechDrawFunctions

        TechDrawFunctions.ImportTemplates()
        TechDrawFunctions.SetDefaultTemplate()
        return

    def Deactivated(self):
        """This function is executed whenever the workbench is deactivated"""
        return

    # def ContextMenu(self, recipient):
    #     """This function is executed whenever the user right-clicks on screen"""
    #     # "recipient" will be either "view" or "tree"
    #     self.appendContextMenu("My commands", self.list) # add commands to the context menu


Gui.addWorkbench(TitleBlockWB())
