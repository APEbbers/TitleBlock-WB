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


import FreeCAD as App
import Standard_Functions
from openpyxl import Workbook


def ExportSpreadSheet():
    try:
        # get the spreadsheet "TitleBlock"
        sheet = App.ActiveDocument.getObject("TitleBlock")
        sheet.recompute()

        # Create a workbook and activate the first sheet
        wb = Workbook()
        ws = wb.active

        # Set the headers
        ws["A1"].value = str(sheet.get("A1"))
        ws["B1"].value = str(sheet.get("B1"))
        ws["C1"].value = str(sheet.get("C1"))
        ws["D1"].value = str(sheet.get("D1"))

        # Go through the spreadsheet.
        for RowNum in range(1000):
            # Start with x+1 first, to make sure that x is at least 1.
            RowNum = RowNum + 2

            try:
                ws["A" + str(RowNum)] = str(sheet.get("A" + str(RowNum)))
            except Exception:
                ws["A" + str(RowNum)] = ""
            try:
                ws["B" + str(RowNum)] = str(sheet.get("B" + str(RowNum)))
            except Exception:
                ws["B" + str(RowNum)] = ""
            try:
                ws["C" + str(RowNum)] = str(sheet.get("C" + str(RowNum)))
            except Exception:
                ws["C" + str(RowNum)] = ""
            try:
                ws["D" + str(RowNum)] = str(sheet.get("D" + str(RowNum)))
            except Exception:
                ws["D" + str(RowNum)] = ""
            try:
                ws["E" + str(RowNum)] = str(sheet.get("E" + str(RowNum)))
            except Exception:
                ws["E" + str(RowNum)] = ""

            # Check if the next row exits. If not this is the end of all the available values.
            try:
                sheet.get("A" + str(RowNum + 1))
            except Exception:
                break

        # Make the columns to autofit the date
        for col in ws.columns:
            SetLen = 0
            column = col[0].column_letter  # Get the column name

            for cell in col:
                if len(str(cell.value)) > SetLen:
                    SetLen = len(str(cell.value))

            set_col_width = SetLen + 5
            # Setting the column width
            ws.column_dimensions[column].width = set_col_width

        # Save the excel file in a folder of your choosing
        Filter = [
            ("Excel", "*.xlsx"),
        ]
        FileName = Standard_Functions.SaveAsDialog(Filter)
        if FileName is not None:
            wb.save(str(FileName))

    except Exception as e:
        # raise Exception("No spreadsheet named 'TitleBlock'!!!")
        Standard_Functions.Mbox("No spreadsheet named 'TitleBlock'!!!", "", 0)
        raise (e)
