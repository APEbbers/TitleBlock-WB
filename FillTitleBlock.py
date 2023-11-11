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


# This macro fills the titleblock with all the data from the spreadsheet.
# The spreadsheet is generated by the macro "PopulateSpreadsheet".
# The data in spreadsheet can be changed by the user as long the data in column A remains unchanged.
# there a multiple senarios for using these macro's:
#  - When changing the template for an different size, you can quickly refill the titleblock.
#    Of course, the editable text in the new template must be equal.
#  - Link the date to model properties like parameters. Basicly all what is achievable with the Spreadsheet workbench
#  - Import data from excel or libre office
#  - Etc.

import FreeCAD as App
from Standard_Functions_TitleBlock import (
    StandardFunctions_FreeCAD as Standard_Functions,
)
import FillSpreadsheet
import Spreadsheet


def FillTitleBlock():
    from Settings import ENABLE_DEBUG
    from Settings import USE_EXTERNAL_SOURCE
    from Settings import EXTERNAL_SOURCE_SHEET_NAME
    from Settings import EXTERNAL_SOURCE_STARTCELL

    # Preset the value for the multiplier. This is used if an value has to be increased for every page.
    NumCounter = -1
    Multiplier = 1

    # Get the pages and go throug them one by one.
    try:
        pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")

        # Get the spreadsheet.
        sheet = Spreadsheet.Sheet(App.ActiveDocument.getObject("TitleBlock"))
        FillSpreadsheet.MapData(sheet=sheet, MapSpecific=4)

        for page in pages:
            # Get the editable texts
            texts = page.Template.EditableTexts
            # Fill the titleblock with the data from the spreadsheet named "Title block".
            # If the spreadsheet doesn't exist raise an error in the report view.
            try:
                # Increase the NumCounter
                NumCounter = NumCounter + 1

                # Go through the spreadsheet.
                for RowNum in range(1000):
                    # Start with x+1 first, to make sure that x is at least 1.
                    RowNum = RowNum + 2

                    # fill in the editable text based on the text name in column A and the value in column B.
                    try:
                        # check if there is a value. If there is an value, fill in.
                        str(sheet.getContents("B" + str(RowNum)))

                    except Exception:
                        pass
                    else:
                        try:
                            # check if there is a value. If so, this property value must be increased with every page
                            str(sheet.getContents("C" + str(RowNum))).strip()
                            try:
                                # check if there is a value in column D
                                str(sheet.getContents("D" + str(RowNum))).strip()
                                # convert it to a number and use it as multiplier
                                Multiplier = int(sheet.getContents("D" + str(RowNum)))

                                # if in debug mode. Show the value of the multiplier
                                if ENABLE_DEBUG is True:
                                    print("The values will be multiplied with: " + str(Multiplier))
                            except Exception as e:
                                # if debug mode is enabeled, print the exception
                                if ENABLE_DEBUG is True:
                                    # there is no int, so the multiplier is set to 1.
                                    print("No Int found!")
                                    print(e)
                                Multiplier = 1
                            try:
                                # check if the value in colom B is an number
                                int(str(sheet.getContents("B" + str(RowNum))))

                                # If Debug mode is enabled, show NumCounter and Multplier
                                if ENABLE_DEBUG is True:
                                    print("NumCounter is: " + str((NumCounter)) + ", Multiplier is: " + str(Multiplier))

                                # The page numbers will be calculated with the formula:
                                # -> the value in column B + (Multiplier*NumCounter).
                                # With Column B is the page number for the first page.
                                #
                                # Example: 1st pagenumber is 2 and the multiplier is 10. Page 1 has number 2.
                                # this results in:
                                # Page 1 has number 2. (as mentioned)
                                # Page 2 has number 12 [2+(10*1)] where 2 is the number of first page,
                                # 10 is the value of the multiplier and 1 is the number of the NumCounter.
                                # Page 3 has 2+(10*2)=22.
                                #
                                # When the 1st page has number 1, page 2 has number 11, page 3 has number 21,
                                # page 4 has 41, etc.
                                texts[str(sheet.getContents("A" + str(RowNum)))] = str(
                                    (int(sheet.getContents("B" + str(RowNum)))) + (Multiplier * NumCounter)
                                )

                            except Exception as e:
                                # if it is not an number, the value of column B will be added without calculation
                                texts[str(sheet.getContents("A" + str(RowNum)))] = str(
                                    sheet.getContents("B" + str(RowNum))
                                )
                                print("this is not a number!")
                                # if degbug mode is enabeled, print the exception
                                if ENABLE_DEBUG is True:
                                    print(e)
                        except Exception as e:
                            # if it is empty, the value of column B will be added without calculation
                            texts[str(sheet.getContents("A" + str(RowNum)))] = str(sheet.getContents("B" + str(RowNum)))
                            # if debug mode is enabeled, print the exception
                            if ENABLE_DEBUG is True:
                                print(e)

                    # Check if the next row exits. If not this is the end of all the available values.
                    try:
                        sheet.getContents("A" + str(RowNum + 1))
                    except Exception:
                        # print("end of range")
                        break

                # Write all the updated text to the page.
                page.Template.EditableTexts = texts

            except Exception as e:
                # raise an exeception if there is no spreadsheet.
                Standard_Functions.Mbox("No spreadsheet named 'TitleBlock'!!!", "TitleBlock Workbench", 0)
                # if degbug mode is enabeled, print the exception
                if ENABLE_DEBUG is True:
                    raise e

    except RuntimeError as e:
        # raise an exeception if there is no page.
        Standard_Functions.Mbox("No page present!!!", "", 0)
        # if degbug mode is enabeled, print the exception
        if ENABLE_DEBUG is True:
            print(e)
    except Exception as e:
        Text = "TitleBlock Workbench: an error occurred!!\n"
        if ENABLE_DEBUG is True:
            Text = "TitleBlock Workbench: an error occurred!!\n" + "See the report view for details"
            raise e
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
