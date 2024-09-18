#!/usr/bin/env python
# -*- coding: utf-8 -*-

def flatten(sStr) -> str:
    # remove spaces, punctuation and convert to lowercase
    sOut = ''
    for sChar in sStr:
        if sChar.isalpha():
            sOut += sChar.lower()
    return sOut


def isStringBoundedBy(sStr, charL, charR='') -> bool:
    if not charR:
        charR = charL
    if len(sStr) > 0:
        return sStr[0] == charL and sStr[len(sStr)-1] == charR
    return False


def quote(sString) -> str:
    return f'"{sString}"' if str(sString).count("'") else f"'{sString}'"


def integer(sString: str) -> int:
    try:
        return int(sString)
    except ValueError:
        return 0
