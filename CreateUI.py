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

# Module to create UI elements like toolbars for other workbenches then your own.
# of course you can probally create here your UI elements for your own workbench as well if you want.
# But this is not tested yet.
# I based the module on the code of the add-on manager:
# https://github.com/FreeCAD/FreeCAD/blob/main/src/Mod/AddonManager/install_to_toolbar.py

import FreeCAD as App
from Settings import USE_EXTERNAL_SOURCE
from Settings import EXTERNAL_SOURCE_PATH
import Standard_Functions_TitleBlock as Standard_Functions


# Define the translation
translate = App.Qt.translate


def DefineToolbars():
    ToolbarListMain = []
    ToolbarListExtra = []
    if USE_EXTERNAL_SOURCE is True:
        if EXTERNAL_SOURCE_PATH.lower().endswith(".fcstd") is False:
            ToolbarListMain = [
                "Separator",
                "ImportExcel",
                "FillTitleBlock",
                "ExpandToolbar",
            ]
            ToolbarListExtra = [
                "OpenExcel",
                "NewExcel",
                "Separator",
                "FillSpreadsheet",
                "Separator",
                "ExportSpreadSheet_Excel",
                "Separator",
                "ExportSettings_Excel",
                "ImportSettings_Excel",
            ]
        if EXTERNAL_SOURCE_PATH.lower().endswith(".fcstd") is True:
            ToolbarListMain = [
                "Separator",
                "ImportFreeCAD",
                "FillTitleBlock",
                "ExpandToolbar",
            ]
            ToolbarListExtra = [
                "OpenFreeCAD",
                "NewFreeCAD",
                "Separator",
                "FillSpreadsheet",
                "Separator",
                "ExportSpreadSheet_FreeCAD",
                "Separator",
                "ExportSettings_FreeCAD",
                "ImportSettings_FreeCAD",
            ]
    if USE_EXTERNAL_SOURCE is False:
        ToolbarListMain = [
            "Separator",
            "FillSpreadsheet",
            "FillTitleBlock",
            "ExpandToolbar",
        ]  # a list of command names created in the line above
        ToolbarListExtra = [
            "ExportSpreadSheet_Excel",
            "ExportSettings_Excel",
            "ImportSettings_Excel",
            "OpenExcel",
            "Separator",
            "ExportSpreadSheet_FreeCAD",
            "ExportSettings_FreeCAD",
            "ImportSettings_FreeCAD",
            "OpenFreeCAD",
        ]

    result = {"ToolbarListMain": ToolbarListMain, "ToolbarListExtra": ToolbarListExtra}

    return result


def DefineMenus():
    StandardList = [
        "FillTitleBlock",
        "FillSpreadsheet",
    ]
    ExcelList = [
        "NewExcel",
        "ExportSpreadSheet_Excel",
        "ImportExcel",
        "OpenExcel",
    ]
    FreeCADList = [
        "Separator",
        "NewFreeCAD",
        "ExportSpreadSheet_FreeCAD",
        "ImportFreeCAD",
        "OpenFreeCAD",
    ]
    SettingsList = [
        "ExportSettings_Excel",
        "ImportSettings_Excel",
        "Separator",
        "ExportSettings_FreeCAD",
        "ImportSettings_FreeCAD",
    ]
    result = {
        "StandardList": StandardList,
        "ExcelList": ExcelList,
        "FreeCADList": FreeCADList,
        "SettingsList": SettingsList
    }

    return result


def QT_TRANSLATE_NOOP(context, text):
    return text


def CreateTechDrawToolbar() -> object:
    """Creates a toolbar in the standard TechDraw WorkBench with the most importand commands"""
    import FreeCADGui as Gui
    from PySide2.QtWidgets import QToolBar

    # region -- define the names and folders

    # Define the name for the ToolbarGroup in the FreeCAD Parameters
    ToolbarGroupName = "TiTleBlock_Toolbar_TechDraw"
    # Define the name for the toolbar
    ToolBarName = "TitleBlock Toolbar"
    # define the parameter path for the toolbar
    TechDrawToolBarsParamPath = (
        "User parameter:BaseApp/Workbench/TechDrawWorkbench/Toolbar/"
    )

    # endregion

    # region -- check if the toolbar already exits.
    name_taken = get_toolbar_with_name(ToolBarName, TechDrawToolBarsParamPath)
    if name_taken:
        i = 2  # Don't use (1), start at (2)
        while True:
            if get_toolbar_with_name(ToolBarName, TechDrawToolBarsParamPath):
                ReplaceButtons()
                return
            i = i + 1
    # endregion

    # region -- Set the Toolbar up

    # add the ToolbarGroup in the FreeCAD Parameters
    TechDrawToolbar = App.ParamGet(
        "User parameter:BaseApp/Workbench/TechDrawWorkbench/Toolbar/" + ToolbarGroupName
    )

    # Set the name.
    TechDrawToolbar.SetString("Name", ToolBarName)

    # Set the toolbar active
    TechDrawToolbar.SetBool("Active", True)

    # add the commands
    if USE_EXTERNAL_SOURCE is True:
        if EXTERNAL_SOURCE_PATH.lower().endswith(".fcstd") is False:
            TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
            TechDrawToolbar.SetString("ImportExcel", "FreeCAD")
        if EXTERNAL_SOURCE_PATH.lower().endswith(".fcstd") is True:
            TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
            TechDrawToolbar.SetString("ImportFreeCAD", "FreeCAD")
    if USE_EXTERNAL_SOURCE is False:
        TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
        TechDrawToolbar.SetString("FillSpreadsheet", "FreeCAD")

    # endregion

    # Force the toolbars to be recreated
    wb = Gui.activeWorkbench()
    if int(App.Version()[0]) == 0 and int(App.Version()[1]) > 19:
        wb.reloadActive()
    return


def RemoveTechDrawToolbar() -> None:
    # Define the name for the ToolbarGroup in the FreeCAD Parameters
    ToolbarGroupName = "TiTleBlock_Toolbar_TechDraw"
    # define the parameter path for the toolbar
    TechDrawToolBarsParamPath = (
        "User parameter:BaseApp/Workbench/TechDrawWorkbench/Toolbar/"
    )

    custom_toolbars = App.ParamGet(TechDrawToolBarsParamPath)
    custom_toolbars.RemGroup(ToolbarGroupName)


def ReplaceButtons() -> None:
    # Define the name for the ToolbarGroup in the FreeCAD Parameters
    ToolbarGroupName = "TiTleBlock_Toolbar_TechDraw"
    # define the parameter path for the toolbar
    TechDrawToolBarsParamPath = (
        "User parameter:BaseApp/Workbench/TechDrawWorkbench/Toolbar/"
    )
    TechDrawToolbar = App.ParamGet(TechDrawToolBarsParamPath + ToolbarGroupName)
    TechDrawToolbar.RemString("ImportExcel")
    TechDrawToolbar.RemString("FillSpreadsheet")

    # add the commands
    if USE_EXTERNAL_SOURCE is True:
        if EXTERNAL_SOURCE_PATH.lower().endswith(".fcstd") is False:
            TechDrawToolbar.SetString("ImportExcel", "FreeCAD")
            TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
        if EXTERNAL_SOURCE_PATH.lower().endswith(".fcstd") is True:
            TechDrawToolbar.SetString("ImportFreeCAD", "FreeCAD")
            TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
    if USE_EXTERNAL_SOURCE is False:
        TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
        TechDrawToolbar.SetString("FillSpreadsheet", "FreeCAD")


# this is an modified verion of the one in:
# https://github.com/FreeCAD/FreeCAD/blob/main/src/Mod/AddonManager/install_to_toolbar.py
def get_toolbar_with_name(name: str, UserParam: str) -> bool:
    """Try to find a toolbar with a given name. Returns True if the preference group for the toolbar
    if found, or False if it does not exist."""
    top_group = App.ParamGet(UserParam)
    custom_toolbars = top_group.GetGroups()
    for toolbar in custom_toolbars:
        group = App.ParamGet(UserParam + toolbar)
        group_name = group.GetString("Name", "")
        if group_name == name:
            return True
    return False


def toggleToolbars(ToolbarName: str, WorkBench: str = ""):
    import FreeCADGui as Gui
    from PySide2.QtWidgets import QToolBar

    # Get the active workbench
    if WorkBench == "":
        WB = Gui.activeWorkbench()
    if WorkBench != "":
        WB = Gui.getWorkbench(WorkBench)

    # Get the list of toolbars present.
    ListToolbars = WB.listToolbars()
    # Go through the list. If the toolbar exists set ToolbarExists to True
    ToolbarExists = False
    for i in range(len(ListToolbars)):
        if ListToolbars[i] == ToolbarName:
            ToolbarExists = True

    # If ToolbarExists is True continue. Otherwise return.
    if ToolbarExists is True:
        # Get the main window
        mainWindow = Gui.getMainWindow()
        # Get the toolbar
        ToolBar = mainWindow.findChild(QToolBar, ToolbarName)
        # If the toolbar is not hidden, hide it and return.
        if ToolBar.isHidden() is False:
            ToolBar.setHidden(True)
            return
        # If the toolbar is hidden, set visible and return.
        if ToolBar.isHidden() is True:
            ToolBar.setVisible(True)
            return
    return
