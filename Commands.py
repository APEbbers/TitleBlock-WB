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

# FreeCAD init script of the Work Features module
import os
from inspect import getsourcefile
import FreeCADGui as Gui

# from _version import __version__

__title__ = "TitleBlock Workbench"
__author__ = "A.P. Ebbers"
__url__ = "https://github.com/APEbbers/TechDrawTitleBlockUtility.git"


# get the path of the current python script
PATH_TB = file_path = os.path.dirname(getsourcefile(lambda: 0))

global PATH_TB_ICONS
global PATH_TB_RESOURCES

PATH_TB_ICONS = os.path.join(PATH_TB, "Resources", "Icons")
PATH_TB_RESOURCES = os.path.join(PATH_TB, "Resources")


# Export data from the titleblock to the spreadsheet
class ListTitleBlock_Class:
    def GetResources(self):
        return {
            "Pixmap": "ListTitleBlock.svg",  # the name of a svg file available in the resources
            "MenuText": "Export data",
            "ToolTip": "Export data from titleblock to an spreadsheet",
        }

    def Activated(self):
        import ListTitleBlock

        ListTitleBlock()
        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        return True


# Import data from the spreadsheet to all pages
class FillTitleBlock_Class:
    def GetResources(self):
        return {
            "Pixmap": "FillTitleBlock.svg",  # the name of a svg file available in the resources
            "MenuText": "Import data",
            "ToolTip": "Imports data from the spreadsheet to titleblock of all pages",
        }

    def Activated(self):
        import FillTitleBlock

        FillTitleBlock()
        return

    def IsActive(self):
        """Here you can define if the command must be active or not (greyed) if certain conditions
        are met or not. This function is optional."""
        return True


# Add the commands to the Gui
Gui.addCommand("ListTitleBlock", ListTitleBlock_Class())
Gui.addCommand("FillTitleBlock", FillTitleBlock_Class())
