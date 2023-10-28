import FreeCAD as App
import FreeCADGui as Gui
from Settings import USE_EXTERNAL_SOURCE


def CreateTechDrawToolbar() -> object:
    """Creates a toolbar in the standard TechDraw WorkBench with the most importand commands"""

    # region -- define the names and folders

    # Define the name for the ToolbarGroup in the FreeCAD Parameters
    ToolbarGroupName = "TiTleBlock_Toolbar_TechDraw"
    # Define the name for the toolbar
    ToolBarName = "TitleBlock_Toolbar"
    # define the parameter path for the toolbar
    TechDrawToolBarsParamPath = (
        "User parameter:BaseApp/Workbench/TechDrawWorkbench/Toolbar/"
    )

    # endregion

    # region -- check if the toolbar already exits. if so, exit.
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
    TechDrawToolbar.SetBool("Ative", True)

    # add the commands
    if USE_EXTERNAL_SOURCE is True:
        TechDrawToolbar.SetString("ImportExcel", "FreeCAD")
        TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
    if USE_EXTERNAL_SOURCE is False:
        TechDrawToolbar.SetString("FillSpreadsheet", "FreeCAD")
        TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")

    # endregion

    # Force the toolbars to be recreated
    wb = Gui.activeWorkbench()
    wb.reloadActive()


def RemoveTechDrawToolbar() -> None:
    # Define the name for the ToolbarGroup in the FreeCAD Parameters
    ToolbarGroupName = "TiTleBlock_Toolbar_TechDraw"
    # Define the name for the toolbar
    ToolBarName = "TitleBlock_Toolbar"
    # define the parameter path for the toolbar
    TechDrawToolBarsParamPath = (
        "User parameter:BaseApp/Workbench/TechDrawWorkbench/Toolbar/"
    )

    # region -- check if the toolbar already exits. if so, exit.
    name_taken = get_toolbar_with_name(ToolBarName, TechDrawToolBarsParamPath)
    if name_taken is True:
        TechDrawToolbar = App.ParamGet(
            "User parameter:BaseApp/Workbench/TechDrawWorkbench/Toolbar/"
            + ToolbarGroupName
        )
        TechDrawToolbar
    # endregion

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
        TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
        TechDrawToolbar.SetString("ImportExcel", "FreeCAD")
    if USE_EXTERNAL_SOURCE is False:
        TechDrawToolbar.SetString("FillTitleBlock", "FreeCAD")
        TechDrawToolbar.SetString("FillSpreadsheet", "FreeCAD")


# this is an modified verion of the one in:
# https://github.com/FreeCAD/FreeCAD/blob/main/src/Mod/AddonManager/install_to_toolbar.py
def get_toolbar_with_name(name: str, UserParam: str) -> object:
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
