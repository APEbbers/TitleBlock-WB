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
import Standard_Functions_TitleBlock as Standard_Functions

# Define the translation
translate = App.Qt.translate


def FillTitleBlock():
    from Settings import ENABLE_DEBUG
    from Settings import MAP_NOSHEETS

    # Preset the value for the multiplier. This is used if an value has to be increased for every page.
    NumCounter = -1
    Multiplier = 1

    # Get the pages and go throug them one by one.
    try:
        pages = App.ActiveDocument.findObjects("TechDraw::DrawPage")

        # Get the spreadsheet.
        sheet = App.ActiveDocument.getObject("TitleBlock")
        if sheet is None:
            Standard_Functions.Mbox("No titleblock spreadsheet present!", "TitleBlock Workbench", 0)
            return

        for page in pages:
            # Get the editable texts
            texts = page.Template.EditableTexts
            # Fill the titleblock with the data from the spreadsheet named "Title block".
            # If the spreadsheet doesn't exist raise an error in the report view.
            try:
                # Increase the NumCounter
                NumCounter = NumCounter + 1
                RowNum = 1
                # Go through the spreadsheet.
                for i in range(len(texts)):
                    # for RowNum in range(1000):
                    # Start with x+1 first, to make sure that x is at least 1.
                    RowNum = RowNum + 1

                    # fill in the editable text based on the text name in column A and the value in column B.
                    # check if there is a value. If there is an value, continue.
                    if str(sheet.getContents("B" + str(RowNum))).strip():
                        # Get the name of the editable field. if it starts with ', remove it.
                        textField = str(sheet.getContents("A" + str(RowNum)))
                        if textField[:1] == "'":
                            textField = textField[1:]

                        # define the string for the text value
                        textValue = ""

                        # if the value in B is not a number, just fill in
                        if (
                            str(sheet.getContents("B" + str(RowNum))).isnumeric()
                            is False
                        ):
                            # write the editable text
                            textValue = str(sheet.getContents("B" + str(RowNum)))
                            if textValue[:1] == "'":
                                textValue = textValue[1:]

                            if textField[:1] == "'":
                                textField = textField[1:]

                            texts[textField] = textValue

                        # If the value in B is a number continue:
                        if (
                            str(sheet.getContents("B" + str(RowNum))).isnumeric()
                            is True
                        ):
                            # Check if the total number of sheets must be filled in.
                            if MAP_NOSHEETS != "":
                                # define TextCompare as Value of MAP_NOSHEETs for comparison. if it starts with ', remove it.
                                TextCompare = MAP_NOSHEETS
                                if TextCompare[:1] == "'":
                                    TextCompare = TextCompare[1:]

                                #  set the value to the total number of pages.
                                if textField == TextCompare:
                                    textValue = str(len(pages))

                                # Write the editable text and set NoSheetsMapped to be true.
                                texts[textField] = textValue

                            # check if there is a value in C. if so, the number in B must be increased with a factor
                            if str(sheet.getContents("C" + str(RowNum))).strip():
                                # check if there is a value in column D, if not the muliplier will be 1.
                                Multiplier = 1
                                if str(sheet.getContents("D" + str(RowNum))).strip():
                                    # Check if the value in D is a number.
                                    if str(
                                        sheet.getContents("D" + str(RowNum))
                                    ).isnumeric():
                                        # convert it to a number and use it as multiplier
                                        Multiplier = int(
                                            sheet.getContents("D" + str(RowNum))
                                        )

                                # if in debug mode. Show the value of the multiplier
                                if ENABLE_DEBUG is True:
                                    Text = translate(
                                        "TitleBlock Workbench",
                                        "The values will be multiplied with: "
                                        + str(Multiplier),
                                    )
                                    Standard_Functions.Print(Text, "Log")

                                # check if the value in colom B is an number

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
                                textValue = str(
                                    (int(sheet.getContents("B" + str(RowNum))))
                                    + (Multiplier * NumCounter)
                                )
                                if textValue[:1] == "'":
                                    textValue = textValue[1:]

                                texts[textField] = textValue

                                # If Debug mode is enabled, show NumCounter and Multplier
                                if ENABLE_DEBUG is True:
                                    Text = translate(
                                        "TitleBlock Workbench",
                                        "NumCounter is: "
                                        + str((NumCounter))
                                        + ", Multiplier is: "
                                        + str(Multiplier)
                                        + "Text is:"
                                        + str(textValue),
                                    )
                                    Standard_Functions.Print(Text, "Log")

                    # Check if the next row exits. If not this is the end of all the available values.
                    try:
                        sheet.getContents("A" + str(RowNum + 1))
                    except Exception:
                        break

                # Write all the updated text to the page.
                page.Template.EditableTexts = texts

                # Recompute the page
                page.recompute()

            except Exception as e:
                # raise an exeception if there is no spreadsheet.
                Text = translate(
                    "TitleBlock Workbench", "An error occured when writing the values!!"
                )
                Standard_Functions.Mbox(
                    text=Text, title="TitleBlock Workbench", style=0
                )
                # if degbug mode is enabeled, print the exception
                if ENABLE_DEBUG is True:
                    raise e

    except Exception as e:
        Text = "TitleBlock Workbench: an error occurred!!\n"
        if ENABLE_DEBUG is True:
            Text = translate(
                "TitleBlock Workbench",
                "TitleBlock Workbench: an error occurred!!\n"
                + "See the report view for details",
            )
            raise e
        Standard_Functions.Mbox(text=Text, title="TitleBlock Workbench", style=0)
