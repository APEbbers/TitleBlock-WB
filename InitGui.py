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
import Settings_TB

# import Standard_Functions_TB as Standard_Functions

# Define the translation
translate = App.Qt.translate

__title__ = "TitleBlock Workbench"
__author__ = "A.P. Ebbers"
__url__ = "https://github.com/APEbbers/TechDrawTitleBlockUtility.git"

# get the path of the current python script
# PATH_TB = file_path = os.path.dirname(getsourcefile(lambda: 0))
PATH_TB = os.path.dirname(os.path.abspath(Settings_TB.__file__))

global PATH_TB_ICONS
global PATH_TB_RESOURCES
global PATH_TB_UI
global PATH_TRANSLATION

PATH_TB_ICONS = os.path.join(PATH_TB, "Resources", "Icons")
PATH_TB_RESOURCES = os.path.join(PATH_TB, "Resources")
PATH_TB_UI = os.path.join(PATH_TB, "Resources", "UI")
PATH_TRANSLATION = os.path.join(PATH_TB, "Translations")


class TitleBlockWB(Gui.Workbench):
    MenuText = "TitleBlock Workbench"
    ToolTip = "An extension for the TechDraw workbench to fill a TitleBlock"
    Icon = os.path.join(PATH_TB_ICONS, "TitleBlockWB.svg")

    Gui.addIconPath(PATH_TB_ICONS)
    Gui.addPreferencePage(
        os.path.join(PATH_TB_UI, "PreferenceUI.ui"),
        "TitleBlock Workbench",
    )

    # Import the observers
    from Events_TB import myObserver

    def GetClassName(self):
        # This function is mandatory if this is a full Python workbench
        # This is not a template, the returned string should be exactly "Gui::PythonWorkbench"
        return "Gui::PythonWorkbench"

    def Initialize(self):
        """This function is executed when the workbench is first activated.
        It is executed once in a FreeCAD session followed by the Activated function.
        """
        import Commands_TB  # import here all the needed files that create your FreeCAD commands
        import Settings_TB
        from Settings_TB import EXTERNAL_SOURCE_PATH
        from Settings_TB import IMPORT_SETTINGS_XL
        from Settings_TB import ADD_TOOLBAR_TECHDRAW
        import CreateUI_TB
        import TechDraw_Functions_TB
        import Events_TB

        # Add the observer(s)
        Events_TB.AddObservers()

        # Define the translation
        def QT_TRANSLATE_NOOP(context, text):
            return text

        # Add the translation path
        Gui.addLanguagePath(PATH_TRANSLATION)

        # import the settings with the correct function based on the extension
        if IMPORT_SETTINGS_XL is True and EXTERNAL_SOURCE_PATH.lower() == ".xlsx":
            Settings_TB.ImportSettings_XL()
        if IMPORT_SETTINGS_XL is True and EXTERNAL_SOURCE_PATH.lower() == ".fcstd":
            Settings_TB.ImportSettings_FreeCAD()

        # region - Create toolbars
        ToolbarListMain = CreateUI_TB.DefineToolbars()["ToolbarListMain"]
        ToolbarListExtra = CreateUI_TB.DefineToolbars()["ToolbarListExtra"]
        self.appendToolbar(
            QT_TRANSLATE_NOOP("Workbench", "TitleBlock"), ToolbarListMain
        )  # creates a new toolbar with your commands
        self.appendToolbar(
            QT_TRANSLATE_NOOP("Workbench", "TitleBlock extra"), ToolbarListExtra
        )  # creates a new toolbar with your commands

        # endregion

        # region - Create menus
        StandardList = CreateUI_TB.DefineMenus()["StandardList"]
        ExcelList = CreateUI_TB.DefineMenus()["ExcelList"]
        FreeCADList = CreateUI_TB.DefineMenus()["FreeCADList"]
        SettingsList = CreateUI_TB.DefineMenus()["SettingsList"]
        DrawingList_Simple = CreateUI_TB.DefineMenus()["DrawingList_Simple"]
        DrawingList_Advanced = CreateUI_TB.DefineMenus()["DrawingList_Advanced"]
        self.appendMenu(
            QT_TRANSLATE_NOOP("Workbench", "TitleBlock"),
            StandardList,
        )  # creates a new menu
        self.appendMenu(
            QT_TRANSLATE_NOOP("Workbench", ["TitleBlock", "External source"]), ExcelList
        )
        self.appendMenu(
            QT_TRANSLATE_NOOP("Workbench", ["TitleBlock", "External source"]),
            FreeCADList,
        )
        self.appendMenu(
            QT_TRANSLATE_NOOP("Workbench", ["TitleBlock", "Drawing list"]),
            DrawingList_Simple,
        )
        self.appendMenu(
            QT_TRANSLATE_NOOP("Workbench", ["TitleBlock", "Drawing list - Advanced"]),
            DrawingList_Advanced,
        )
        self.appendMenu(
            QT_TRANSLATE_NOOP("Workbench", ["TitleBlock", "Settings"]), SettingsList
        )
        # endregion

        # region - Create toolbar for TechDraw workbench
        if ADD_TOOLBAR_TECHDRAW is True:
            CreateUI_TB.CreateTechDrawToolbar()
        if ADD_TOOLBAR_TECHDRAW is False:
            CreateUI_TB.RemoveTechDrawToolbar()
        # endregion

        # region set the templates for the TechDraw workbench
        TechDraw_Functions_TB.ImportTemplates()
        TechDraw_Functions_TB.SetDefaultTemplate()
        # endregion

        return

    def Activated(self):
        """This function is executed whenever the workbench is activated"""
        import TechDraw_Functions_TB

        # set the techdraw templates
        TechDraw_Functions_TB.ImportTemplates()
        TechDraw_Functions_TB.SetDefaultTemplate()
        return

    def Deactivated(self):
        """This function is executed whenever the workbench is deactivated"""
        return

    # def ContextMenu(self, recipient):
    #     """This function is executed whenever the user right-clicks on screen"""
    #     # "recipient" will be either "view" or "tree"
    #     self.appendContextMenu("My commands", self.list) # add commands to the context menu


Gui.addWorkbench(TitleBlockWB())
