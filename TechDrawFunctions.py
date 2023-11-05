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
from inspect import getsourcefile
import FreeCAD as App

# get the path of the current python script
PATH_TB = os.path.dirname(getsourcefile(lambda: 0))
TEMPLATE_DIR = os.path.join(PATH_TB, "Example templates").replace("\\", "/")


def GetDefaultTemplate(selectedItem: int) -> str:
    if selectedItem == 0:
        return "A0_Landscape.svg"
    elif selectedItem == 1:
        return "A1_Landscape.svg"
    elif selectedItem == 2:
        return "A2_Landscape.svg"
    elif selectedItem == 3:
        return "A3_Landscape.svg"
    elif selectedItem == 4:
        return "A4_Landscape.svg"
    elif selectedItem == 5:
        return "A4_Portrait.svg"
    else:
        return "A4_Landscape.svg"


def ImportTemplates():
    from Settings import IMPORT_EXAMPLE_TEMPLATES

    if IMPORT_EXAMPLE_TEMPLATES is True:
        # Define the parameter path
        TechDrawTemplateDirParamPath = (
            "User parameter:BaseApp/Preferences/Mod/TechDraw/Files/"
        )

        # Get the parameter group
        TemplateDir = App.ParamGet(TechDrawTemplateDirParamPath)

        # Set the template directory
        TemplateDir.SetString("TemplateDir", TEMPLATE_DIR)


def SetDefaultTemplate():
    from Settings import DEFAULT_TEMPLATE
    from Settings import IMPORT_EXAMPLE_TEMPLATES
    from Settings import ENABLE_DEBUG

    ChosenTemplate = os.path.join(
        TEMPLATE_DIR, GetDefaultTemplate(DEFAULT_TEMPLATE).replace("\\", "/")
    )

    # Print the standard template when debug mode is active
    if ENABLE_DEBUG is True:
        print(ChosenTemplate)

    if IMPORT_EXAMPLE_TEMPLATES is True:
        # Define the parameter path
        TechDrawTemplateParamPath = (
            "User parameter:BaseApp/Preferences/Mod/TechDraw/Files/"
        )

        # Get the parameter group
        Template = App.ParamGet(TechDrawTemplateParamPath)

        # Set the template directory
        Template.SetString("TemplateFile", ChosenTemplate)
