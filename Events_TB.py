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

# https://github.com/FreeCAD/FreeCAD/blob/main/src/App/DocumentObserverPython.h

import FreeCAD as App
import FreeCADGui as Gui
import Standard_Functions_TB as Standard_Functions
import FillSpreadsheet_TB
import FillTitleBlock_TB
from Settings_TB import EXTERNAL_SOURCE_PATH
from Settings_TB import USE_EXTERNAL_SOURCE
from Settings_TB import ENABLE_DEBUG
from Settings_TB import ENABLE_RECOMPUTE_FILL_SPREADSHEET
from Settings_TB import ENABLE_RECOMPUTE_FILL_TITLEBLOCK


# Class for a observer that observers recompute events
class myObserver(object):
    # Observer for document recompute event ("CTRL+R")
    def slotRecomputedDocument(self, doc):
        ActiveWorkbench = Gui.activeWorkbench()
        AllowedWorkbenches = [
            "TitleBlockWB",
            "TechDrawWorkbench",
            "SpreadsheetWorkbench",
        ]
        for AllowedWorkBench in AllowedWorkbenches:
            if ActiveWorkbench.name() == AllowedWorkBench:
                if ENABLE_RECOMPUTE_FILL_SPREADSHEET is True:
                    if USE_EXTERNAL_SOURCE is True:
                        if EXTERNAL_SOURCE_PATH.lower().endswith(".xlsx") is True:
                            FillSpreadsheet_TB.Start("ImportExcel", doc, False)
                        if EXTERNAL_SOURCE_PATH.lower().endswith(".xlsx") is False:
                            FillSpreadsheet_TB.Start("ImportFreeCAD", doc, False)
                    if USE_EXTERNAL_SOURCE is False:
                        FillSpreadsheet_TB.Start("FillSpreadsheet", doc, False)
                    if ENABLE_DEBUG is True:
                        Standard_Functions.Print(
                            "The titleblock spreadsheet has been updated!"
                        )

                if ENABLE_RECOMPUTE_FILL_TITLEBLOCK is True:
                    FillTitleBlock_TB.FillTitleBlock(doc=doc, recompute=False)
                    if ENABLE_DEBUG is True:
                        Standard_Functions.Print(
                            "The titleblock in all the pages has been updated!"
                        )

                if ENABLE_DEBUG is True:
                    Standard_Functions.Print(
                        "%s has been recomputed\n" % doc.Label, "Log"
                    )
        return


# Add the observers


def AddObservers():
    obs = myObserver()
    App.addDocumentObserver(obs)
