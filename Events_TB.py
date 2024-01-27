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
import FreeCADGui
import Standard_Functions_TB as Standard_Functions


class myObserver(object):
    def __init__(self):
        self.target_doc = None

    def slotRecomputedObject(self, doc):
        if doc == self.target_doc:
            Standard_Functions.Print("%s has been recomputed\n" % doc.Label, "Log")


obs = myObserver()
App.addDocumentObserver(obs)

doc = Standard_Functions.Print("i_am_observed", "Log")
obs.target_doc = doc
