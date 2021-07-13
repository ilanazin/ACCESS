import os.path
from os import path
import csv
from datetime import datetime
import time
import logging
# Writing to an excel sheet using Python
import xlwt
from xlwt import Workbook
from tabulate import tabulate
# Use colorama module to print in color
import colorama
from colorama import Fore, Back, Style

ASCII_ART_OPEN = """ _____                     _       _          ___                         ______ _ _       ______                                    
/  ___|                   | |     (_)        / _ \                        |  ___(_) |      | ___ \                                   
\ `--.  ___  __ _ _ __ ___| |__    _ _ __   / /_\ \ ___ ___ ___  ___ ___  | |_   _| | ___  | |_/ / __ ___   __ _ _ __ __ _ _ __ ___  
 `--. \/ _ \/ _` | '__/ __| '_ \  | | '_ \  |  _  |/ __/ __/ _ \/ __/ __| |  _| | | |/ _ \ |  __/ '__/ _ \ / _` | '__/ _` | '_ ` _ \ 
/\__/ /  __/ (_| | | | (__| | | | | | | | | | | | | (_| (_|  __/\__ \__ \ | |   | | |  __/ | |  | | | (_) | (_| | | | (_| | | | | | |
\____/ \___|\__,_|_|  \___|_| |_| |_|_| |_| \_| |_/\___\___\___||___/___/ \_|   |_|_|\___| \_|  |_|  \___/ \__, |_|  \__,_|_| |_| |_|
                                                                                                            __/ |                    
                                                                                                           |___/                     
   """

ASCII_ART_CLOSE = """______                                      _____ _                    _ 
| ___ \                                    /  __ \ |                  | |
| |_/ / __ ___   __ _ _ __ __ _ _ __ ___   | /  \/ | ___  ___  ___  __| |
|  __/ '__/ _ \ / _` | '__/ _` | '_ ` _ \  | |   | |/ _ \/ __|/ _ \/ _` |
| |  | | | (_) | (_| | | | (_| | | | | | | | \__/\ | (_) \__ \  __/ (_| |
\_|  |_|  \___/ \__, |_|  \__,_|_| |_| |_|  \____/_|\___/|___/\___|\__,_|
                 __/ |                                                   
                |___/                                                   
   """


def print_wellcome_screen():
    """
    This function print the welcome screen for Access Program
    :return: print the welcome screen for Access Program
    :rtype: None
    """
    print_message_with_color(ASCII_ART_OPEN, "blue")
    print_message_with_color(
        "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ", "blue")


def print_close_screen():
    """
    This function print the close screen for Access Program
    :return: print theclose screen for Access Program
    :rtype: None
    """
    print_message_with_color(ASCII_ART_CLOSE, "blue")
    print_message_with_color(
        "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ", "blue")


def print_message_with_color(msg, color="WHITE", back="RESET"):
    """
    This function prints string message to console with different font colors: BLACK, RED, GREEN, YELLOW, BLUE, MAGENTA, CYAN, WHITE, RESET. The default is white.
    :param screen: string to print
    :type screen: string
    :param screen: color of the string to print
    :type screen: string
    :param screen: back color of the string to print
    :type screen: string
    :return: print to console with different font colors
    :rtype: None
    """

    color = getattr(Fore, color.upper())
    back = getattr(Back, back.upper())
    print(color + msg + back + Style.RESET_ALL)


def create_files():
    """
    This function create files under predifined path than needed
    :return: Create files under predifined path han needed
    :rtype: Tuple
    """

    values_file_path = r'C:\Ilana\Python\PYTHON_PROJECT\Values.csv'
    log_file = r'C:\Ilana\Python\PYTHON_PROJECT\logging.txt'
    create_csv(values_file_path)
    my_logging(log_file)
    logging.info("Log file " + log_file + " created")
    return values_file_path, log_file


def my_logging(log_file):
    """
    This function define the structure of logger file that include the date, employee_id and name.
    :return: The logger file with prefefined structure
    :rtype: None
    """
    LOG_FORMAT = "%(levelname)s %(asctime)s - %(message)s"
    logging.basicConfig(filename=log_file,
                        level=logging.INFO,
                        format=LOG_FORMAT,
                        filemode='w')
    logger = logging.getLogger()
    # text messages
    logger.debug("Start logger")


def create_csv(values_file_path):
    """
    This function create csv file with header File name, Table, Column, Row, Value.
    :param values_file_path: file path to the csv file
    :type values_file_path: string
    :return: Crete csv file with header File_name, Table, Column, Row, Value.
    :rtype: None
    """

    try:
        if path.isfile(values_file_path) == False:
            with open(values_file_path, 'w', newline='') as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(
                    ["File_name", "Table", "Column", "Row", "Value"])
                writer.close
    except Exception as error:
        print(error)


def check_csv_header(values_file_path):
    """
    This function check that csv file have predefined header
    :param values_file_path: file path to the values csv file
    :type values_file_path: string
    :return: Check that csv file have predefined header. If yes return True, else False
    :rtype: Boolean
    """

    with open(values_file_path, 'r', newline='') as csv_file:
        reader = csv.reader(csv_file)
        # the first line is the header
        header = next(reader)

    if header[0] == 'File_name' and header[1] == 'Table' and header[2] == 'Column' and header[3] == 'Row' and header[4] == 'Value':
        return True
    else:
        print("Wrong csv header file. Please check that " + values_file_path +
              " file include the following columns: 'File_name', 'Table', 'Column', 'Row', 'Value'")
        return False


def update_values_to_csv_(values_file_path, file_name, table_name, column_name, row_id, value):
    """
    This function update values to the values file.
    :param values_file_path: file path to the csv file
    :type values_file_path: string
    :return: Add a new lines to the values file.
    :rtype: None
    """

    with open(values_file_path, 'a',  encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow([file_name, table_name, column_name, row_id, value])


def check_csv_header(csv_file_path):
    """
    This function check that csv file have predefined header
    :param employee_file_path: file path to the csv file
    :type employee_file_path: string
    :return: Check that csv file have predefined header. If yes return True, else False
    :rtype: Boolean
    """

    with open(employee_file_path, 'r', newline='') as csv_file:
        reader = csv.reader(csv_file)
        # the first line is the header
        header = next(reader)

    if header[0] == 'File' and header[1] == 'Table' and header[2] == 'Column' and header[3] == 'Value':
        return True
    else:
        print("Wrong csv header file. Please check that " + employee_file_path +
              " file include the following columns: 'File', 'Table', 'Column', 'Value'")
        return False
