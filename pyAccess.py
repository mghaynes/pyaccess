#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Author: Malcolm Haynes
Date: Decemner 2017

This module serves as a wrapper to an Access database. It uses the Windows database engine to Access
properties of the database not otherwise available such as various table properties.

Last Edit: Malcolm Haynes
Date: 03 NOV 2018
Separated the pure DB functions from the grading functions in DAOdbUtils
"""

__author__ = "Malcolm Haynes"
__copyright__ = "12/31/2017"
__version__ = "1.0"

# you'll need to import the win32com module. This can be done using pip (e.g., pip install pypiwin32)
import win32com.client
import collections
import logging

debug = 0  # Set from 0 or 2 to get varying levels of output; 0=no output, 2=very verbose (NOT IMPLEMENTED YET)
too_many_penalty = .05  # penalty for selecting too many items
max_misspelled = 2

Lookup = collections.namedtuple('Lookup', ['DisplayControl', 'RowSourceType', 'RowSource', 'BoundColumn',
                                           'ColumnCount', 'ColumnWidths', 'LimitToList'])
ColumnMeta = collections.namedtuple('ColumnMeta', ['Name', 'Type', 'Size'])

Relationship = collections.namedtuple('Relationship', ['Table', 'Field', 'RelatedTable', 'RelatedField',
                                                       'EnforceIntegrity', 'JoinType', 'Attributes'])


'''-----------------------------------------------------------------------------------------------------------------'''
'''                                               CLASS: DATABASE                                                   '''
'''    DataBase class loads key properties of database to include relationships, table, and query properties        '''


class DataBase:
    def __init__(self, dbPath, debug=0):
        self._dbEngine = win32com.client.Dispatch("DAO.DBEngine.120")
        self._ws = self._dbEngine.Workspaces(0)
        self._dbPath = dbPath
        self._db = self._ws.OpenDatabase(self._dbPath)
        self._debug = debug
        self.TableNames = self.TableList(debug=self._debug)
        self.QueryNames = self.TableList(isTable=False, debug=self._debug)
        self.Relationships = self.GetRelationships(debug=self._debug)
        self.Tables = self.LoadTables(self.TableNames, debug=self._debug)
        self.Queries = self.LoadTables(self.QueryNames, isTable=False, debug=self._debug)
        # self._db.Close()

    # For query list, isTable must be False
    def TableList(self, isTable=True, debug=0):
        table_list = []
        if isTable:
            tables = self._db.TableDefs
        else:
            tables = self._db.QueryDefs
        if debug and isTable:
            print('TABLES:')
        elif debug and not isTable:
            print('QUERIES')
        for table in tables:
            if not table.Name.startswith('MSys') and not table.Name.startswith('~'):
                table_list.append(table.Name)
                if debug:
                    print(table.Name)
        return table_list


    def LoadTables(self, table_list, isTable=True, debug=0):
        tables = {}
        for table in table_list:
            if isTable:
                tables[table] = Table(self._db.TableDefs(table), dbPath=self._dbPath)
                if table in self.Relationships:
                    tables[table].ForeignKeys = self.Relationships[table]
            else:
                tables[table] = Table(self._db.QueryDefs(table), isTable=isTable, dbPath=self._dbPath)
        return tables


    '''' Attributes translations (I THINK!)
        0 = Enforce referential integrity (RI), Inner join
        2 = Referential integrity (RI) not enforced, Inner join
        16777216 = RI, outer join on related table
        16777218 = No RI, outer join on related table
        33554434 = No RI, outer join on table
        33554432 = RI, outer join on table'''
    def GetRelationships(self, debug=1):
        relationships = dict()
        for rltn in self._db.Relations:
            if rltn.ForeignTable not in relationships:
                relationships[rltn.ForeignTable] = dict()
            if rltn.Table not in relationships[rltn.ForeignTable]:
                relationships[rltn.ForeignTable][rltn.Table] = dict()
            for field in rltn.Fields:
                if rltn.Attributes == 0:
                    JoinType = 'INNER'
                    ReferentialIntegrity = True
                elif rltn.Attributes == 2:
                    JoinType = 'INNER'
                    ReferentialIntegrity = False
                elif rltn.Attributes == 16777216:
                    JoinType = 'OUTER RELATED'
                    ReferentialIntegrity = True
                elif rltn.Attributes == 16777218:
                    JoinType = 'OUTER RELATED'
                    ReferentialIntegrity = False
                elif rltn.Attributes == 33554432:
                    JoinType = 'OUTER TABLE'
                    ReferentialIntegrity = True
                elif rltn.Attributes == 33554434:
                    JoinType = 'OUTER TABLE'
                    ReferentialIntegrity = False
                else:
                    JoinType = 'UNKNOWN'
                    ReferentialIntegrity = None
                new_rltn = Relationship(Table=rltn.ForeignTable, Field=field.ForeignName, RelatedTable=rltn.Table,
                                        RelatedField=field.Name, EnforceIntegrity=ReferentialIntegrity,
                                        JoinType=JoinType, Attributes=rltn.Attributes)
                relationships[rltn.ForeignTable][rltn.Table][field.ForeignName] = new_rltn
                # if debug:
                #     print(relationships)
        if debug:
            for table_name in relationships.keys():
                for foreign_name in relationships[table_name].keys():
                    for field_name in relationships[table_name][foreign_name].keys():
                        print(relationships[table_name][foreign_name][field_name])
        return relationships


'''-----------------------------------------------------------------------------------------------------------------'''
'''                                               CLASS: TABLE                                                      '''
''' DataBase class permits various operations on tables/queries to include getting records, SQL, lookups, keys,     '''
''' and more.                                                                                                       '''

class Table:
    def __init__(self, table_meta=None, isTable=True, dbPath=None, debug=0):
        if table_meta == None:
            return
        self._dbEngine = win32com.client.Dispatch("DAO.DBEngine.120")
        self._ws = self._dbEngine.Workspaces(0)
        self._dbPath = dbPath
        self._TableMetaData = table_meta
        self.Name = table_meta.Name
        self.debug = debug
        if isTable:
            self.TableType = 'TABLE'
            self.RecordCount = table_meta.RecordCount
            self.PrimaryKeys = self.GetPrimaryKeys()
            self.ForeignKeys = ''
        else:
            self.TableType = 'QUERY'
            self.SQL = self.GetSQL(table_meta)
            self.RecordCount = None
            # if dbPath != None:
            #     self.RecordCount = self.QueryRecordCount()
        self.ColumnMetaData = self.GetColumnMetaData(table_meta)
        self.ColumnCount = len(self.ColumnMetaData)

    def __str__(self):
        column_tuples = [(field.Name, field.Type, field.Size) for field in self.ColumnMetaData]
        if self.TableType == 'TABLE':
            if self.ForeignKeys:
                fk_list = [str(r2) for k, r in self.ForeignKeys.items() for k2, r2 in r.items()]
            else:
                fk_list = ['']
            return 'Table Name: {:25}Type: {:10}Row Count: {:<10}Column Count: {}\nColumns: {}\nPrimary Keys: ' \
                   '{}\nForeign Keys: {}'.format(self.Name, self.TableType ,self.RecordCount, self.ColumnCount,
                                                 column_tuples, ', '.join(self.PrimaryKeys),
                                                 '\n'.join(fk_list))
        elif self.TableType == 'QUERY':
            return 'Query Name: {:25}Type: {:10}Row Count: {:<10}Column Count: {}\nColumns: {}\nSQL: ' \
                   '{}'.format(self.Name, self.TableType ,self.RecordCount, self.ColumnCount, column_tuples, self.SQL)
        else:
            return ''
            # self._rows = self.RowCount(self.debug)

    def hasColumn(self, name):
        column_meta = self.ColumnMetaData
        found = False
        for col in column_meta:
            if name in col.Name:
                return True
        return False

    def QueryRecordCount(self):
        self._db = self._ws.OpenDatabase(self._dbPath)
        num_rows = self._db.OpenRecordset(self.Name).RecordCount
        self._db.Close()
        return num_rows



    # returns the names of the columns in a table
    def GetColumnMetaData(self, table_meta, debug=0):
        columns = []
        if debug:
            print('TABLE:', table_meta.Name)
        for Field in table_meta.Fields:
            if Field.Type == 1:
                type = 'Yes/No'
            elif Field.Type == 4:
                if Field.Attributes in [17 ,18]:
                    type = 'Autonumber'
                else:
                    type = 'LongInteger'
            elif Field.Type == 7:
                type = 'Double'
            elif Field.Type == 8:
                type = 'Date/Time'
            elif Field.Type == 10:
                type = 'ShortText'
            else:
                type = 'UNKNOWN'
            column_meta = ColumnMeta(Field.Name, type, Field.Size)
            columns.append(column_meta)
            if debug:
                print('Field Name:', column_meta.Name, 'Type:', column_meta.Type, 'Size', column_meta.Size)
        return columns

    def GetLookupProperties(self, fieldName, debug=0):
        # Note that the ColumnWidths uses twips a unit of measure where 1 in = 1440 twips, 1 cm = 567 twips
        LookupFields = ['RowSourceType', 'RowSource', 'BoundColumn', 'ColumnCount', 'ColumnWidths',
                        'LimitToList']
        column_widths = ''
        row_source = ''
        field_meta = self.GetFieldObject(fieldName)
        for property in field_meta.Properties:
            if property.Name == 'DisplayControl':
                if property.Value == 111:
                    display_control = 'Combo box'
                if property.Value == 110:
                    display_control = 'List box'
                if property.Value == 109:
                    display_control = 'Text box'
            if property.Name == 'RowSourceType':
                row_source_type = property.Value
            if property.Name == 'RowSource':
                row_source = property.Value
            if property.Name == 'BoundColumn':
                bound_column = property.Value
            if property.Name == 'ColumnCount':
                column_count = property.Value
            if property.Name == 'ColumnWidths':
                column_widths = property.Value
            if property.Name == 'LimitToList':
                limit_to_list = property.Value
            if debug > 1 and property.Name in LookupFields:
                print(property.Name ,': ', property.Value)
            if debug > 1 and property.Name == 'DisplayControl':
                print(property.Name, ': ', display_control)
        lookup = Lookup(display_control, row_source_type, row_source, bound_column, column_count, column_widths,
                        limit_to_list)
        return lookup


    def GetPrimaryKeys(self, debug=0):
        PKs = []
        for idx in self._TableMetaData.Indexes:
            if idx.Primary:
                for field in idx.Fields:
                    PKs.append(field.Name)
        if debug:
            print(self.Name.upper() ,'primary keys:', ','.join(PKs))
        return PKs

    def GetSQL(self, query, debug=0):
        if '~' not in query.Name:
            if debug:
                print('QUERY SQL for' ,query.Name)
                print(query.SQL)
            return query.SQL
        else:
            return 0

    def GetRecords(self, debug=0):
        self._db = self._ws.OpenDatabase(self._dbPath)
        table = self._db.OpenRecordset(self.Name)
        records = []
        while not table.EOF:
            temp_rec = []
            record = table.GetRows()
            for item in record:
                temp_rec.append(list(item)[0])
            records.append(temp_rec)
            if debug > 1:
                print(temp_rec)
        self._db.Close()
        return records

    def GetFieldObject(self, name):
        return self._TableMetaData.Fields(name)

    def GetFields(self):
        fields = []
        for column in self.ColumnMetaData:
            fields.append(column.Name)
        return fields

    def GetTypes(self):
        types = []
        for column in self.ColumnMetaData:
            types.append(column.Type)
        return types

    def GetSizes(self):
        sizes = []
        for column in self.ColumnMetaData:
            sizes.append(column.Size)
        return sizes


'''---------------------------------------------- END TABLE CLASS ------------------------------------------------'''


def main():
    SolnDBPath = r"./DBProject181_soln.accdb"
    StudentDBPath = r"./DBProject181.accdb"
    SolnDB = DataBase(SolnDBPath)
    StudentDB = DataBase(StudentDBPath)
    # Print meta data on all the tables in the database
    # for table in SolnDB.TableNames:
    #     print(SolnDB.Tables[table], '\n')
    # Print meta data on all the queries in the database
    # for query in SolnDB.QueryNames:
    #     print(SolnDB.Queries[query], '\n')
    # print all the relationships in the table
    # for relationship in SolnDB.Relationships:
    #    print(json.dumps(relationship))
    # print(json.dumps(SolnDB.Relationships))
    # print all the records in a table (Note: If debug < 2, it doesn't print anything. Just returns the records)
    # print('Platoon Table Records')
    # SolnDB.Tables['Platoon'].GetRecords(debug=2)
    # print()
    # print the lookups for a field (Note: If debug < 2, it doesn't print anything. Just returns the Lookup tuple)
    # print('Lookups for soldierTrained field in SoldierCompletesTraining')
    table = SolnDB.Tables['SoldierCompletesTraining']
    table2 = StudentDB.Tables['SoldierCompletesTraining']
    # table.GetLookupProperties('soldierTrained', debug=2)
    table.
    # print()
    # print the properties for some metadata (e.g. Table, Query, or Field)
    # print('Table Properties')
    # ListProperties(table._TableMetaData)
    # print('\nField Properties')
    # ListProperties(field)
    # print(field.Properties['ColumnHidden'].Value)
    # print('\nQuery Properties')
    # ListProperties(SolnDB.Queries['APFTStars']._TableMetaData)
    # print(SolnDB.Queries['APFTStars']._TableMetaData)


if __name__ == "__main__":
    main()