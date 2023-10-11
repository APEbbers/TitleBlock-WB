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
import FreeCAD as App
import FreeCADGui as Gui
from inspect import getsourcefile

# from _version import __version__

__title__ = "TitleBlock Workbench"
__author__ = "A.P. Ebbers"
__url__ = "https://github.com/APEbbers/TechDrawTitleBlockUtility.git"

""" ###############
M_DEBUG = False
###############
global TB_Release
TB_Release = __version__ """

# get the path of the current python script
PATH_TB = file_path = os.path.dirname(getsourcefile(lambda: 0))

global PATH_TB_ICONS
global PATH_TB_RESOURCES

PATH_TB_ICONS = os.path.join(PATH_TB, "Resources", "Icons")
PATH_TB_RESOURCES = os.path.join(PATH_TB, "Resources")

""" if M_DEBUG:
    print("DEBUG : PATH_TB           is " + str(PATH_TB))
    print("DEBUG : PATH_TB_ICONS     is " + str(PATH_TB_ICONS))
    print("DEBUG : PATH_TB_RESOURCES is " + str(PATH_TB_RESOURCES)) """


class TitleBlockWB(Workbench):
    MenuText = "TitleBlock Workbench"
    ToolTip = "An extension for the TechDraw workbench to fill a TitleBlock"
    Icon = os.path.join(PATH_TB_ICONS, "TitleBlockWB.svg")

    def Initialize(self):
        """This function is executed when the workbench is first activated.
        It is executed once in a FreeCAD session followed by the Activated function.
        """
        import Commands  # import here all the needed files that create your FreeCAD commands

        self.list = [
            "ListTitleBlock",
            "FillTitleBlock",
        ]  # a list of command names created in the line above
        self.appendToolbar(
            "TitleBlock", self.list
        )  # creates a new toolbar with your commands
        self.appendMenu("TitleBlock", self.list)  # creates a new menu

    def Activated(self):
        """This function is executed whenever the workbench is activated"""
        return

    def Deactivated(self):
        """This function is executed whenever the workbench is deactivated"""
        return

    # def ContextMenu(self, recipient):
    #     """This function is executed whenever the user right-clicks on screen"""
    #     # "recipient" will be either "view" or "tree"
    #     self.appendContextMenu("My commands", self.list) # add commands to the context menu

    def GetClassName(self):
        # This function is mandatory if this is a full Python workbench
        # This is not a template, the returned string should be exactly "Gui::PythonWorkbench"
        return "Gui::PythonWorkbench"


Gui.addWorkbench(TitleBlockWB())
