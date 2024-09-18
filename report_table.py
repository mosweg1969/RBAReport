#!/usr/bin/env python
# -*- coding: utf-8 -*-

from dict_functions import sortDictionary
from datetime import datetime
from rich.console import Console
from rich.table import Table
from rich import box
from rich.text import Text
from typing import Union, Any
import debug_routines as out
import subprocess


class ReportError(Exception):
    def __init__(self, sMessage):
        super().__init__(sMessage)


class Report:
    def __init__(self, sHeading=None, bZebra=False, bShowHeader=True, bLines=False):
        self.sReportHeading = sHeading
        self.iColumnCount = 0
        self.lColumns = []
        self.iRowCount = 0
        self.lRows = []
        self.lNewRow = []
        self.lBreaks = []
        self.iNewCellCount = 0
        self.bHasTotals = False
        self.lBreakColumns = []
        self.table = None
        self.bZebraStripes = bZebra
        self.bShowHeader = bShowHeader
        self.bLines = bLines
        self.lColumnHeadings = []

    def dumpReport(self) -> None:
        dDict = {'Heading': self.sReportHeading, 'Column Count': self.iColumnCount, 'Columns': self.lColumns, 'Rows': self.iRowCount}
        out.debug(dDict, 'Table Dump)')
        for iRow, Row in enumerate(self.lRows):
            for iCell, Cell in enumerate(Row):
                out.debug(Cell, f"Row {iRow+1}, Cell {iCell+1}")

    def setHeading(self, sHeading) -> None:
        self.sReportHeading = sHeading

    def setBreakColumn(self, *columns) -> None:
        for sCol in columns:
            self.lBreakColumns.append(self.__validateColumn__(sCol))

    def addColumn(self, sHeading, oDataType=str, sJust='left', sColour='cyan', bTotalled=False, bShow=True, bBreak=False, oDisplay=None) -> None:
        dColumn = {"heading": sHeading, "data_type": oDataType, "numeric": False, "width": 0}
        if sJust:
            dColumn["justify"] = sJust
        else:
            if oDataType is int:
                dColumn["justify"] = "right"
                dColumn["numeric"] = True
            else:
                dColumn["justify"] = "centre"
        dColumn["colour"] = sColour
        dColumn["totalled"] = bTotalled
        dColumn["total"] = 0
        dColumn['display'] = oDisplay
        if bTotalled:
            if oDataType is not int:
                raise ReportError("Can't total a column that isn't type integer.")
            self.bHasTotals = True

        dColumn["show"] = bShow
        dColumn["same"] = True
        dColumn["__first"] = ""

        if bBreak:
            self.lBreakColumns.append(len(self.lColumns))

        self.lColumns.append(dColumn)
        self.iColumnCount += 1
        self.lColumnHeadings.append(sHeading)

    def columnHeadingList(self) -> list:
        return [column['heading'] for column in self.lColumns]

    def columnSame(self, column) -> bool:
        return self.lColumns[self.__validateColumn__(column)]["same"]

    def showColumn(self, column, bShow):
        self.lColumns[self.__validateColumn__(column)]["show"] = bShow

    def deleteNewRow(self):
        self.addRow()
        self.deleteRow(self.iRowCount - 1)

    def initNewRow(self):
        if self.iNewCellCount:
            raise ReportError(f"Row already has {self.iNewCellCount} cells added.")

        for iColumn in range(self.iColumnCount):
            self.addCell("")

    def cellValue(self, column) -> str:
        iColumn = self.__validateColumn__(column)
        return self.lNewRow[iColumn]["value"]

    def addCells(self, *sValues):
        for sValue in sValues:
            self.addCell(sValue)

    def getCellValue(self, iRow: int, iCol: int):
        return self.lRows[iRow][iCol]['value']

    def updateCellValue(self, iRow: int, iCol: int, sValue: str):
        self.lRows[iRow][iCol]['value'] = sValue
        self.lRows[iRow][iCol]['display'] = sValue
        self.lRows[iRow][iCol]['width'] = len(sValue)
        if len(sValue) > self.lColumns[iCol]["width"]:
            self.lColumns[iCol]["width"] = len(sValue)

    def editCell(self, column: Union[int, str], sValue: Any, sColour: str = ""):
        iColumn = self.__validateColumn__(column)
        dColumn = self.lColumns[iColumn]

        if dColumn['data_type'] is bool:
            sValue = 'Yes' if sValue else 'No'
        if dColumn['data_type'] is datetime:
            if type(sValue) is datetime.date:
                pass
            else:
                try:
                    date_obj = datetime.strptime(sValue, '%Y-%m-%d %H:%M:%S')
                    sValue = date_obj.strftime("%e %b '%y")
                except ValueError:
                    sValue = ''
                except TypeError:
                    sValue = sValue.strftime("%e %b '%y")

        dCell = {'value': str(sValue), 'display': str(sValue), 'colour': sColour, "column": iColumn}
        if dColumn['display'] is None:
            dCell['display'] = str(sValue)
        else:
            if sValue:
                fFormat = dColumn['display']
                sValue = dCell['display'] = fFormat(sValue)

        self.lNewRow[iColumn] = dCell

        # check if column has all same values
        if dColumn["same"]:
            if len(self.lRows) == 0:
                dColumn["__first"] = sValue
            elif dColumn["__first"] != sValue:
                dColumn["same"] = False

        # does the total need to be updated
        if dColumn["totalled"] and sValue != "":
            dColumn["total"] += int(dCell['value'])

    def addCell(self, sValue, sColour=""):
        if self.iNewCellCount == len(self.lColumns):
            raise ReportError("Adding too many cells into row.")
        self.lNewRow.append({"value": "", 'display': '', "span": 1, "colour": "", "column": self.iNewCellCount})

        self.editCell(self.iNewCellCount, sValue, sColour)
        self.iNewCellCount += 1

    def addRowValues(self, *values) -> None:
        for val in values:
            self.addCell(val)
        self.addRow()

    def addPartialRowValues(self, *values) -> None:
        for val in values:
            self.addCell(val)
            if self.iNewCellCount == self.iColumnCount:
                self.addRow()

    def completePartialRow(self) -> None:
        if self.iNewCellCount:
            while self.iNewCellCount < self.iColumnCount:
                self.addCell('')
            self.addRow()

    def addRow(self, bBreak=False) -> None:
        if self.iNewCellCount != self.iColumnCount:
            raise ReportError(f"Wrong number of arguments to addRow - Column count = {self.iColumnCount}, Cells added = {self.iNewCellCount}")

        if bBreak:
            self.lBreaks.append(len(self.lRows))
        self.lRows.append(self.lNewRow)
        self.iRowCount += 1
        self.lNewRow = []
        self.iNewCellCount = 0

    def deleteRow(self, iRow) -> None:
        if not (0 <= iRow < self.iRowCount):
            raise ReportError(f"Row index {iRow} doesn't exist.")

        del self.lRows[iRow]
        self.iRowCount -= 1

    def columnIndex(self, sHeading) -> int:
        return self.__validateColumn__(sHeading)

    def __validateColumn__(self, column) -> int:
        if type(column) is int and 0 <= column < len(self.lColumns):
            return column
        if column.isnumeric() and 1 <= int(column) <= len(self.lColumns):
            return int(column) - 1
        for iColumn in range(self.iColumnCount):
            if self.lColumns[iColumn]["heading"].lower() == column.lower():
                return iColumn
        raise ReportError(f"Invalid column - '{column}' - Type - {type(column)}")

    def sortRows(self, column) -> None:
        # create a dictionary of the keys, with row number list, sort it
        iColumn = self.__validateColumn__(column)
        dSorts = {}

        iRow = 0
        for lRowValues in self.lRows:
            oSortKey = lRowValues[iColumn]["display"]
            if self.lColumns[iColumn]['numeric']:
                oSortKey = int(oSortKey)
                print('yes')
            if oSortKey in dSorts:
                dSorts[oSortKey].append(iRow)
            else:
                dSorts[oSortKey] = [iRow]
            iRow += 1

        dSorts = sortDictionary(dSorts)

        # now reconstruct the rows
        lNewRows = []
        for sSort in dSorts:
            for iRow in dSorts[sSort]:
                lNewRows.append(self.lRows[iRow])

        self.lRows = lNewRows

    def printReport(self, sSort=None, sFilter=None) -> None:
        if sSort:
            self.sortRows(sSort)

        if self.bHasTotals:
            for dColumn in self.lColumns:
                if dColumn["show"]:
                    if dColumn["totalled"]:
                        self.addCell(dColumn["total"])
                    else:
                        self.addCell("")
            self.editCell(0, "Totals", sColour='[purple]')
            self.addRow()

        dFilters = {}
        if sFilter:
            for sRule in sFilter.split(';'):
                iColumn = self.__validateColumn__(sRule.split(':')[0])
                lValues = sRule.split(':')[1].split(',')
                dFilters[int(iColumn)] = lValues

        self.table = Table(title=self.sReportHeading, box=box.DOUBLE, style='#808080', row_styles=(['cyan', 'white'] if self.bZebraStripes else None),
                           show_header=self.bShowHeader, show_lines=self.bLines)
        for dColumn in self.lColumns:
            if dColumn["show"]:
                sColour = dColumn['colour']
                sJust = dColumn['justify']
                if sJust == 'centre':
                    sJust = 'center'
                self.table.add_column(dColumn['heading'], style=sColour, justify=sJust)

        iBreakColumn = None
        for iRow, lRowValues in enumerate(self.lRows):
            lValues = [self.richColourFormat(dCell) for dCell in lRowValues if self.lColumns[dCell["column"]]["show"]]
            bEndSection = (self.bHasTotals and iRow == len(self.lRows)-1)

            # work out what column we are breaking on
            for iColumn in self.lBreakColumns:
                if lRowValues[iColumn]['value']:
                    iBreakColumn = iColumn
                    break

            # does it meet filter conditions?
            bFilterMet = True
            if self.bHasTotals and iRow == len(self.lRows) - 1:
                pass
            else:
                for iColumn in dFilters.keys():
                    if dFilters[iColumn][0].startswith('!'):
                        if lRowValues[iColumn]['value'] == dFilters[iColumn][0][1:]:
                            bFilterMet = False
                            break
                    elif dFilters[iColumn][0].startswith('<'):
                        if int(lRowValues[iColumn]['value']) >= int(dFilters[iColumn][0][1:]):
                            bFilterMet = False
                            break
                    elif dFilters[iColumn][0].startswith('>'):
                        if int(lRowValues[iColumn]['value']) <= int(dFilters[iColumn][0][1:]):
                            bFilterMet = False
                            break
                    elif lRowValues[iColumn]['value'] not in dFilters[iColumn]:
                        bFilterMet = False
                        break

            if bFilterMet:
                if iBreakColumn is not None and iRow > 0 and lRowValues[iBreakColumn]['value'] != self.lRows[iRow-1][iBreakColumn]['value']:
                    bEndSection = True

                if iRow in self.lBreaks:
                    bEndSection = True

                if bEndSection:
                    self.table.add_section()

                self.table.add_row(*lValues)
            elif bEndSection:
                self.table.add_section()

        Console().print(self.table, justify="left")

    def sendToClipboard(self):
        lLines = []
        lValues = [dColumn['heading'] for dColumn in self.lColumns if dColumn['show']]
        lLines.append('\t'.join(lValues))

        for iRow, lRowValues in enumerate(self.lRows):
            lValues = [dCell['value'] for dCell in lRowValues if self.lColumns[dCell["column"]]["show"]]
            lLines.append('\t'.join(lValues))

        sClipBoard = '\n'.join(lLines)
        subprocess.run('pbcopy', universal_newlines=True, input=sClipBoard)

    @staticmethod
    def richColourFormat(dCell: dict) -> Text:
        if dCell['colour']:
            return dCell['colour']+dCell['display']
        else:
            return dCell['display']
