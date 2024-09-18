#!/usr/bin/env python
# -*- coding: utf-8 -*-

# -------------------------------------------------------------------------------------------------------------------------------------------------------- #
#   All routines for providing terminal messaging, progress bars and messages
# -------------------------------------------------------------------------------------------------------------------------------------------------------- #

from datetime import date
from string_functions import quote
from rich import print
from rich.markup import escape
from rich.progress import Progress, TimeElapsedColumn, TextColumn, TaskID
from time import sleep
import os
import sys
import regex as re
from typing import Union

from globals import GlobalVars, GlobalConsts


lDebugLevels = ['Debug', 'Info', 'Output', 'Warning', 'Error', 'Critical']
sOutputColours = ['[bright_green]', '[cyan]', '[purple]', '[bright_yellow]', '[red]', '[bright_red on white]']
sMessagePrefix = ['Debug: ', 'Info: ', '', 'Warning: ', 'Error: ', 'CRITICAL: ']
iCurrLevel = None
iInitialLevel = None


def protectBrackets(sString: str) -> str:
    return sString.replace('[', r'\[')


def createProgressBars(sTextColumnField='none', **kwargs) -> Progress:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Creates a context manager for all subsequent progress bars
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    return Progress(*Progress.get_default_columns(), TimeElapsedColumn(), TextColumn('[cyan]{task.fields[' + sTextColumnField + ']}'), **kwargs)


def createProgressTask(sTask, **kwargs) -> Union[None, TaskID]:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Creates a task within the context manager
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    if 'total' in kwargs and kwargs.get('total') == 0:
        return None

    return GlobalVars.oProgress.add_task(f"[cyan]{sTask} ...", **kwargs)


def updateProgressTask(iTaskID, **kwargs):
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Updates the task message and/or progress indicator
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    if iTaskID is not None:
        GlobalVars.oProgress.update(iTaskID, **kwargs)


def destroyProgressTask(iTaskID) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Deletes the progress task
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    if iTaskID is not None:
        GlobalVars.oProgress.remove_task(iTaskID)
        sleep(0.1)


def restoreLevel() -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Restore the debug level to that which it was initially set as (i.e. from the command line)
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    global iCurrLevel
    iCurrLevel = iInitialLevel


def __checkLevel__() -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Internal function to set the debug level to default of 'Output' if it has not yet been set.
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    global iCurrLevel

    if iCurrLevel is None:
        setDebugLevel('Output')


def setDebugLevel(sLevel) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Set the debug level, if it is being set for the first time then record this initial level so level can be modified and reset to this.
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    global iCurrLevel
    global iInitialLevel

    if sLevel in lDebugLevels:
        if iCurrLevel is None:
            iInitialLevel = lDebugLevels.index(sLevel)
        iCurrLevel = lDebugLevels.index(sLevel)
    else:
        error(f"Cannot set debugging to level '{sLevel}'.")


def debug(*args) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   print arguments at debug level 0 = debug
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    __printOutput__(0, *args)


def info(*args) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   print arguments at debug level 1 = info
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    __printOutput__(1, *args)


def output(*args, **kwargs) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   print arguments at debug level 2 = output
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    __printOutput__(2, *args, **kwargs)


def warning(*args) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   print arguments at debug level 3 = warning
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    __printOutput__(3, *args)


def error(*args) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   print arguments at debug level 4 = error
    #   execution halts with return code = 1
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    __printOutput__(4, *args)
    exit(1)


def critical(*args) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   print arguments at debug level 5 = critical
    #   execution halts with return code = 4
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    __printOutput__(5, *args)
    exit(4)


def __printOutput__(iLevel, *messages, bRichPrint=False) -> None:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   print arguments at debug level 5 = critical
    #   execution halts with return code = 4
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    global iCurrLevel
    __checkLevel__()
    if iLevel >= iCurrLevel:
        if len(messages) and type(messages[0]) is dict:
            sTitle = '?'
            if len(messages) == 2 and type(messages[1]) is str:
                sTitle = messages[1]
            __printDict__(messages[0], sTitle, 0, sOutputColours[iLevel], '[yellow]')
        else:
            sPrint = sMessagePrefix[iLevel]
            if not bRichPrint:
                sPrint = sOutputColours[iLevel] + sPrint
            for oMsg in messages:
                sPrint += (str(oMsg) + ' ')
            print(sPrint)


def stripLength(sString: str) -> int:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   returns the length of a string without any of the markup
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    return len(re.sub('\[\w*]', '', sString))


def __printDict__(dDict, sName, iSpaces, cDictColour, cHighlight, sPrevDictKey=''):
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   pretty prints a dictionary
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    iMaxLen = 0
    for sKey in dDict:
        if type(sKey) == int:
            if len(str(sKey)) > iMaxLen:
                iMaxLen = len(str(sKey))
        else:
            if len(str(sKey)) + 2 > iMaxLen:
                iMaxLen = len(sKey) + 2

    lKeys = list(dDict)
    if len(lKeys) == 0:
        if iSpaces == 0:
            print(f'Dictionary : {cHighlight}{quote(sName)}[purple] {{}}')
            return
    else:
        if iSpaces == 0:
            print(f'Dictionary : {cHighlight}{quote(sName)}[purple] {{')

        sFirstKey = lKeys[0]
        sLastKey = lKeys[-1]
        for sKey in dDict:
            iTempSpaces = iSpaces
            sLeadIn = ''
            if sKey == sFirstKey:
                iSpaces = 0
                if sPrevDictKey:
                    sLeadIn = sPrevDictKey
            if type(sKey) == int:
                sLeadIn += ' ' * iSpaces + cDictColour + str(sKey) + ' ' * (iMaxLen - len(str(sKey))) + ' : '
            else:
                sLeadIn += ' ' * iSpaces + cDictColour + quote(sKey) + ' ' * (iMaxLen - (len(sKey) + 2)) + ' : '
            iSpaces = iTempSpaces
            sLeadOut = '[purple] }' + cDictColour if sKey == sLastKey and iSpaces > 0 else ''

            if type(dDict[sKey]) == dict:
                if dDict[sKey] is None:
                    print(f"{sLeadIn}[purple]" + "{[purple]}" + cDictColour)
                else:
                    __printDict__(dDict[sKey], '', iSpaces + iMaxLen + 5, cDictColour, cHighlight, sLeadIn + '[purple]{ ' + cDictColour)
            elif type(dDict[sKey]) is bool:
                print(f'{sLeadIn}[yellow]{dDict[sKey]}{cDictColour}{sLeadOut}')
            elif type(dDict[sKey]) is date:
                print(f'{sLeadIn}[dark_orange]{dDict[sKey]}{cDictColour}{sLeadOut}')
            elif type(dDict[sKey]) is str:
                print(f'{sLeadIn}[green]{escape(quote(dDict[sKey]))}{cDictColour}{sLeadOut}')
            elif type(dDict[sKey]) is list:
                print(f'{sLeadIn}{__listContents__(dDict[sKey], cDictColour, stripLength(sLeadIn) + 1, 0)}[bold] : ={len(dDict[sKey])}[/bold]{sLeadOut}')
            else:
                print(f'{sLeadIn}[white]{escape(str(dDict[sKey]))}{cDictColour}{sLeadOut}')

    if iSpaces == 0:
        print('[purple]}')


def __listContents__(lList: list, cDictColour, iLeadInLength: int = 0, iKeyLength: int = 0) -> str:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   internal routine to pretty print lists on a single line
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    lOut = []
    dRemap = {'year': 'yr', 'issue': 'is', 'ordinal': 'or', 'title': 'ti', 'index': 'ix', 'volume': 'vo', 'vol_index': 'vi',
              'number': '№', 'filename': 'fn', 'format': 'fo', 'date': 'dt', 'custom_name': 'cn', 'month': 'mn', 'extension': '…',
              'next_date': 'nd', 'special': 'sp', 'season': 'ss'}
    for oItem in lList:
        if type(oItem) is date:
            lOut.append(f"[white]{str(oItem.day)}/{oItem.strftime('%m')}{cDictColour}")
        elif type(oItem) is dict:
            lDictContents = []
            for sKey, sValue in oItem.items():
                if sKey in dRemap:
                    sKey = dRemap[sKey]

                if type(sValue) is int:
                    lDictContents.append(f"'{sKey}': [white]{sValue}{cDictColour}")
                elif type(sValue) is str:
                    lDictContents.append(f"'{sKey}': [white]'{sValue}'{cDictColour}")
                elif type(sValue) is date:
                    lDictContents.append(
                        f"'{sKey}': [white]{str(sValue.day)}/{sValue.strftime('%m')}{cDictColour}")
            sDictContents = '[purple]{' + cDictColour + ', '.join(lDictContents) + '[purple]}'
            if oItem == lList[0]:
                lOut.append(sDictContents)
            else:
                lOut.append('\n' + ' ' * (iLeadInLength + iKeyLength) + sDictContents)
        elif type(oItem) is str:
            lOut.append(f"[white]'{oItem}'{cDictColour}")
        else:
            lOut.append(f"[white]{oItem}{cDictColour}")

    return f"[blue]\[{cDictColour}{', '.join(lOut)}[blue]]{cDictColour}"


def oneLineDict(dDict: dict, bExtraOK=False) -> str:
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   returns a one line string of a dictionary contents, for use in report tables
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    dRemap = {'year': 'yr', 'issue': 'is', 'ordinal': 'or', 'title': 'ti', 'index': 'ix', 'volume': 'vo', 'vol_index': 'vi', 'number': '№',
              'filename': 'fn', 'format': 'fo', 'date': 'dt', 'custom_name': 'cn', 'month': 'mn', 'extension': '…', 'next_date': 'nd',
              'special': 'sp', 'month-2': 'm2', 'year-2': 'y2', 'season': 'sn', 'extra': 'ex', 'extra-2': 'x2', 'extra-3': 'x3', 'extra-4': 'x4',
              'day_of_month': 'dm', 'day_of_month-2': 'dm2', 'day': 'dy', 'day_or_month': 'd/m', 'numeric': '#', 'edition': 'ed', 'day-2': 'd2',
              'day_name': 'dn', 'price': '£$', 'title+': 't+', 'price-2': '$2', 'price-3': '$3', 'numeric-2': '#2', 'country': 'cn', 'price+': '$+'}
    lDictContents = []
    for sKey, sValue in dDict.items():
        if sKey not in ['extension']:
            if sKey.startswith('extra') and not bExtraOK:
                sKey = f"[red]{dRemap[sKey]}[cyan]"
            elif sKey in dRemap:
                sKey = dRemap[sKey]
            else:
                sKey = f"[red]???{sKey}???[cyan]"

            if type(sValue) is int:
                lDictContents.append(f"'{sKey}': [white]{sValue}[cyan]")
            elif type(sValue) is str:
                lDictContents.append(f"'{sKey}': [white]'{sValue}'[cyan]")
            elif type(sValue) is date:
                lDictContents.append(
                    f"'{sKey}': [white]{str(sValue.day)}/{sValue.strftime('%m')}[cyan]")
    return '[purple]{[cyan]' + ', '.join(lDictContents) + '[purple]}'
