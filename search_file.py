import pyodbc
import os
import os.path
from os import path
from bidi import algorithm as bidialg
import csv
import logging

tables = ["Afik.mdb", "CommonTables.mdb", "DataSK.mdb",
          "db1.mdb", "db2.mdb", "db3.mdb",
          "db4.mdb", "db5.mdb", "db6.mdb",
          "db7.mdb", "db8.mdb", "db9.mdb",
          "db10.mdb", "db11.mdb", "ExportFile.mdb",
          "ExportFileSM.mdb", "ImportData.mdb", "ImportDataSM.mdb",
          "MolsaFU_DB.mdbb", "MolsaQA_DB.mdb", "MolsaReports.mdb",
          "MolsaReportsPLUS-ALEX.mdb", "MyLab.mdb", "PriceQuote.mdb",
          "Products.mdb", "reporteddata.mdb", "SekerMakdim.mdb",
          "SMCommonTables.mdb", "UpdateFile.mdb", "קליטת_נתונים5.mdb"]


def find_files(directory):
    """
    This function find all Access files in selected directory.
    :param directory: directory to seach Access file
    :type directory: string
    :return: Access file names.
    :rtype: string
    """
    files = []
    for file in os.listdir(directory):
        if file.endswith(".mdb"):
            files.append(file)
    if not files:
        files.append("No Access files in selected directory")
    return files


def print_table_in_file(file):
    """
    This function find all table names in Access file.
    :param file: directory to seach Access file
    :type file: string
    :return: table names in Access file.
    :rtype: string
    """
    if file.endswith(".mdb"):
        connection = pyodbc.connect(
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % (file))
        cursor = connection.cursor()
        table_names = [x[2] for x in cursor.tables(tableType='TABLE')]
        cursor.close()
        return table_names


def print_table_in_directory(directory):
    """
    This function find all table names in Access file.
    :param directory: directory to seach Access file
    :type directory: string
    :return: table names in Access file.
    :rtype: string
    """
    for file in os.listdir(directory):
        if file.endswith(".mdb"):
            md_file = os.path.join(directory, file)
            connection = pyodbc.connect(
                r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % (md_file))
            cursor = connection.cursor()
            table_names = [x[2] for x in cursor.tables(tableType='TABLE')]
            cursor.close()
            return table_names


def print_column(directory, table):
    """
    This function find all column names in Access table.
    :param directory: directory to seach Access file
    :type directory: string
    :param table: table to seach 
    :type table: string
    :return: column names in Access file.
    :rtype: string
    """
    for file in os.listdir(directory):
        if file.endswith(".mdb"):
            md_file = os.path.join(directory, file)
            connection = pyodbc.connect(
                r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % (md_file))
            cursor = connection.cursor()
            res = cursor.execute("SELECT * FROM " + table)
            columnList = [tuple[0] for tuple in res.description]
            cursor.close()
            return columnList


def find_columns_in_table(directory):
    """
    This function find all column names in Access table.
    :param directory: directory to seach Access file
    :type directory: string
    :return: column names in Access file.
    :rtype: string
    """
    connection = pyodbc.connect(
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % (directory))
    cursor = connection.cursor()
    table_names = [x[2] for x in cursor.tables(tableType='TABLE')]
    for tablen in range(len(table_names)):
        restable = cursor.execute(str("SELECT * FROM " +
                                      "[" + table_names[tablen] + " ]"))
        columnList = [tuple[0] for tuple in restable.description]
        cursor.close()

    return columnList


def find_value_in_directory(value, directory, excel):
    """
    This function find value in Access tables in given directry and write result to csv file
    :param value: value to seach Access file
    :type value: string
    :param directory: directory to seach Access file
    :type directory: string
    :param excel: csv result file 
    :type excel: string
    :return: csv result file.
    :rtype: string
    """
    try:
        # Repeat for each file in the directory
        for file in os.listdir(directory):
            # Apply file type filter
            if file.endswith(".mdb"):
                md_file = os.path.join(directory, file)
                connection = pyodbc.connect(
                    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % (md_file))
                cursor = connection.cursor()
                table_names = [x[2]
                               for x in cursor.tables(tableType='TABLE')]
                for tablen in range(len(table_names)):
                    restable = cursor.execute(str("SELECT * FROM " +
                                                  "[" + table_names[tablen] + " ]"))
                    columnList = [tuple[0] for tuple in restable.description]
                    for columns in range(len(columnList)):
                        cursor.execute(str("SELECT * FROM " + "[" + table_names[tablen] + " ]" +
                                           " where " + "[" + columnList[columns] + " ]" + " LIKE '%" + value + "%'"))
                        result = cursor.fetchone()

                        if result:
                            table_name = table_names[tablen]
                            column_name = columnList[columns]
                            update_values_to_csv_(excel,
                                                  file, table_names[tablen], columnList[columns], row_id, value)
                            return excel

                    cursor.close()

        return excel

    except Exception as error:
        print(error)
    finally:
        if connection:
            connection.close()
