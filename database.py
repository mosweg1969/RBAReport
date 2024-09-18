#!/usr/bin/env python
# -*- coding: utf-8 -*-

# -------------------------------------------------------------------------------------------------------------------------------------------------------- #
#   Provides a wrapper layer for managing a SQLite3 database
#   Since the Virtual File System uses the database, the actual database file must be real
# -------------------------------------------------------------------------------------------------------------------------------------------------------- #

import sqlite3
import os

from dict_functions import replaceSingleQuotesInDict


class DatabaseError(Exception):
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    #   Custom error class for all database operations
    # ---------------------------------------------------------------------------------------------------------------------------------------------------- #
    def __init__(self, sMessage):
        super().__init__(sMessage)


class Database:
    oDBConnection = None
    oDBCursor = None

    def __init__(self, sDatabase: str) -> None:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Instantiate the database class object, create the connection and cursor.
        #   Errors if database is missing or if file doesn't exist.
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        if sDatabase == '':
            raise DatabaseError('Database name is missing.')

        if not os.path.exists(sDatabase):
            raise DatabaseError(f"Database '{sDatabase}' doesn't exist.")
        try:
            self.oDBConnection = sqlite3.connect(sDatabase)
            self.oDBCursor = self.oDBConnection.cursor()
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def close(self) -> None:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Closes the database connectio, errors if the database has already been closed.
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        if self.oDBConnection is None:
            raise DatabaseError('Database connection has already been closed.')

        try:
            self.oDBConnection.close()
            self.oDBConnection = None
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def errorIfClosed(self, sOperation: str) -> None:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Raise an error if the database has already been closed.
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        if self.oDBConnection is None:
            raise DatabaseError(f"Database connection has been closed. Cannot perform {sOperation}.")

    def execute(self, sSQL: str, tData: tuple = ()) -> None:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Execute a parameterised SQL query, using the substituted values in the tData tuple
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        self.errorIfClosed('execute')
        try:
            self.oDBCursor.execute(sSQL, tData)
            self.oDBConnection.commit()
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def fetchList(self, sSQL: str, tData: tuple = ()) -> list:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Execute a parameterised SQL query, using the substituted values in the tData tuple, and return a list of the database rows
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        self.errorIfClosed('fetch')
        try:
            result = self.oDBCursor.execute(sSQL, tData)
            return result.fetchall()
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def fetchValue(self, sSQL: str, tData: tuple = ()) -> any:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Execute a parameterised SQL query, using the values in the tData tuple, and return the single value expected from the SQL statement
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        self.errorIfClosed('fetch')
        try:
            result = self.oDBCursor.execute(sSQL, tData)
            lResults = result.fetchall()
            if len(lResults) == 0:
                return None
            if len(lResults) != 1 and len(lResults[0]) != 1:
                raise DatabaseError('More than one value or row was returned in fetchValue() function.')
            return lResults[0][0]
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def fetchValues(self, sSQL: str, tData: tuple = ()) -> any:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Execute a parameterised SQL query, using the values in the tData tuple, and return the tuple from the one row of data returned
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        self.errorIfClosed('fetch')
        try:
            result = self.oDBCursor.execute(sSQL, tData)
            lResults = result.fetchall()
            if len(lResults) == 0:
                return None
            if len(lResults) != 1:
                raise DatabaseError('More than one row was returned in fetchValues() function.')
            return tuple(lResults[0])
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def insertIntoTableUsingDict(self, sTableName: str, dFields: dict) -> int:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Insert into table sTableName using the {column name: value} pairs in dictionary dFields
        #   Single quotes are replaced by two single quotes
        #   Returns the rowID of the entry that was inserted, this is useful for autoupdate key IDs
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        self.errorIfClosed('insert')
        replaceSingleQuotesInDict(dFields)
        sSQL = f"INSERT INTO '{sTableName}' (" + ','.join([f"'{sField}'" for sField in dFields.keys()]) + ')\nVALUES (' +\
               ','.join([f"{sField}" for sField in dFields.values()]) + ')\n'
        try:
            self.execute(sSQL)
            return self.oDBCursor.lastrowid
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def updateTableUsingDict(self, sTableName: str, dUpdateFields: dict, dWhereFields: dict) -> None:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Update fields in table sTableName using the {column name: value} pairs in dictionary dUpdateFields selecting row(s) from dWhereFields.
        #   Single quotes are replaced by two single quotes in both dictionaries.
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        self.errorIfClosed('update')
        replaceSingleQuotesInDict(dUpdateFields)
        replaceSingleQuotesInDict(dWhereFields)
        sSQL = f"UPDATE '{sTableName}' \nSET " + ',\n    '.join([f"{sField} = {sValue}" for sField, sValue in dUpdateFields.items()]) + '\nWHERE ' +\
               'AND\n    '.join([f"{sField} = {sValue}" for sField, sValue in dWhereFields.items()])
        try:
            self.execute(sSQL)
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def lastRowInserted(self) -> int:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Returns the rowID of the entry that was last inserted, this is useful for autoupdate key IDs
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        return self.oDBCursor.lastrowid

    def attachDatabase(self, sPath: str, sAlias: str) -> None:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   Attach another SQLite3 database into the current database
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        self.errorIfClosed('attach')
        if sPath == '':
            raise DatabaseError('Database name to attach is missing.')

        if not os.path.exists(sPath):
            raise DatabaseError(f"Database to attach '{sPath}' doesn't exist.")

        sPath = sPath.replace("'", "''")
        try:
            self.execute(f"ATTACH DATABASE '{sPath}' AS {sAlias}")
        except sqlite3.Error as exp:
            raise DatabaseError(exp)

    def detachDatabase(self, sAlias: str) -> None:
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        #   detach a SQLite3 database from the current database based on the alias name
        # ------------------------------------------------------------------------------------------------------------------------------------------------ #
        try:
            self.execute(f"DETACH DATABASE {sAlias}")
        except sqlite3.Error as exp:
            raise DatabaseError(exp)
