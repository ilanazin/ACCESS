import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import Label
from tkinter import ttk, messagebox
from tkinter import Menu
from tkinter import filedialog as fd
from tkinter import Tk, Text
from tkinter.messagebox import showinfo
import sys
import search_file
import os.path
from os import path


def create_search_value(container):
    """
    This function create search frame for search value.
    :param container: the root window of tkinter.
    :type container: root window tkinter
    :return: container: the root window of tkinter.
    :rtype: root window tkinter
    """
    # Search value
    ttk.Label(container, text='Enter value to search:').grid(
        column=0, row=0, sticky=tk.W)
    keyword = ttk.Entry(container, width=40)
    keyword.grid(column=1, row=0, sticky=tk.W)
    keyword.focus()
    # Access Directory path
    ttk.Label(container, text='Enter Access directory path:').grid(
        column=0, row=1, sticky=tk.W)
    path = ttk.Entry(container, width=40)
    path.grid(column=1, row=1, sticky=tk.W)
    path.focus()
    # CSV directory path
    ttk.Label(container, text='Enter CSV file path:').grid(
        column=0, row=2, sticky=tk.W)
    ttk.Label(container, text='').grid(column=0, row=4, sticky=tk.W)
    CSV = ttk.Entry(container, width=40)
    CSV.grid(column=1, row=2, sticky=tk.W)
    CSV.focus()
    ttk.Button(container, text='Search', width=20, command=lambda: search_clicked(keyword, path, CSV)).grid(
        column=0, row=6, sticky=tk.W)
    ttk.Button(container, text='Find Access dir', width=20, command=lambda: select_Access(path)).grid(
        column=3, row=1, sticky=tk.W)
    ttk.Button(container, text='Find CSV file', width=20, command=lambda: select_CSV(CSV)).grid(
        column=3, row=2, sticky=tk.W)
    ttk.Button(container, text='Reset', width=20, command=lambda: reset_all_clicked(keyword, path, CSV)).grid(
        column=1, row=6, sticky=tk.W)
    ttk.Button(container, text='Close', width=20, command=lambda: close_clicked(container)).grid(
        column=3, row=6, sticky=tk.W)

    return container


def launchFindDialog(*args):
    """
    This function open message box with help
    :param: None
    :return: None
    """
    messagebox.showinfo(message="A program help to work with Access database")


def launchCreatedDialog(*args):
    """
    This function open message box with explanation who created it
    :param: None
    :return: None
    """
    messagebox.showinfo(
        message="A program creted by Ilana Zinger")


def select_file():
    """
    This function open window to select only Access file
    :param file_path: None
    :return: check that file exist according given path, not empty and with csv extention. If yes return True, else False
    :rtype: Boolean
    """
    filetypes = (
        ('Access files', '*.mdb'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    showinfo(
        title='Selected File is Access file',
        message=filename
    )


def select_Access(path):
    """
    This function open window to select Access file directory 
    :param file_path: None
    :return: None
    """
    dir = fd.askdirectory()
    path.insert(0, dir)


def select_CSV(CSV):
    """
    This function open window to select CSV file 
    :param file_path: None
    :return: None
    """
    filetypes = (
        ('CSV files', '*.csv'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    CSV.insert(0, filename)


def create_search_TEST(container):
    """
    This function create search frame for search value.
    :param container: the root window of tkinter.
    :type container: root window tkinter
    :return: container: the root window of tkinter.
    :rtype: root window tkinter
    """
    # Search value
    ttk.Label(container, text='Enter value to search:').grid(
        column=0, row=0, sticky=tk.W)
    keyword = ttk.Entry(container, width=40)
    keyword.grid(column=1, row=0, sticky=tk.W)
    keyword.focus()
    # Access Directory path
    ttk.Label(container, text='Enter Access directory path:').grid(
        column=0, row=1, sticky=tk.W)
    path = ttk.Entry(container, width=40)
    path.grid(column=1, row=1, sticky=tk.W)
    path.focus()
    # CSV directory path
    ttk.Label(container, text='Enter CSV file path:').grid(
        column=0, row=2, sticky=tk.W)
    ttk.Label(container, text='').grid(column=0, row=4, sticky=tk.W)
    CSV = ttk.Entry(container, width=40)
    CSV.grid(column=1, row=2, sticky=tk.W)
    CSV.focus()
    ttk.Button(container, text='Search', width=20, command=lambda: search_clicked(keyword, path, CSV)).grid(
        column=0, row=6, sticky=tk.W)
    ttk.Button(container, text='Find Access dir', width=20, command=lambda: select_Access(path)).grid(
        column=3, row=1, sticky=tk.W)
    ttk.Button(container, text='Find CSV file', width=20, command=lambda: select_Access(CSV)).grid(
        column=3, row=2, sticky=tk.W)
    ttk.Button(container, text='Reset', width=20, command=lambda: reset_all_clicked(keyword, path, CSV)).grid(
        column=1, row=6, sticky=tk.W)
    ttk.Button(container, text='Close', width=20, command=lambda: close_clicked(container)).grid(
        column=3, row=6, sticky=tk.W)

    return container


def create_search_files(container):
    """
    This function create search frame for search files.
    :param container: the root window of tkinter.
    :type container: root window tkinter
    :return: container: the root window of tkinter.
    :rtype: root window tkinter
    """
    ttk.Label(container, text='Enter Access directory path:').grid(
        column=0, row=1, sticky=tk.W)
    path = ttk.Entry(container, width=40)
    path.grid(column=1, row=1, sticky=tk.W)
    path.focus()
    ttk.Label(container, text='').grid(column=0, row=4, sticky=tk.W)
    ttk.Button(container, text='Search', width=20, command=lambda: search_files_clicked(path)).grid(
        column=0, row=6, sticky=tk.W)
    ttk.Button(container, text='Find Access dir', width=20, command=lambda: select_Access(path)).grid(
        column=3, row=1, sticky=tk.W)
    ttk.Button(container, text='Reset', width=20, command=lambda: reset_clicked(path)).grid(
        column=1, row=6, sticky=tk.W)
    ttk.Button(container, text='Close', width=20, command=lambda: close_clicked(container)).grid(
        column=3, row=6, sticky=tk.W)

    return container


def create_search_tables(container):
    """
    This function create search frame for search tables in given Access file.
    :param container: the root window of tkinter.
    :type container: root window tkinter
    :return: container: the root window of tkinter.
    :rtype: root window tkinter
    """
    # Access Directory path
    ttk.Label(container, text='Enter Access file path:').grid(
        column=0, row=1, sticky=tk.W)
    path = ttk.Entry(container, width=40)
    path.grid(column=1, row=1, sticky=tk.W)
    path.focus()
    ttk.Button(container, text='Search', width=20, command=lambda: search_tables_clicked(path)).grid(
        column=0, row=6, sticky=tk.W)
    ttk.Button(container, text='Find Access File', width=20, command=lambda: select_Access_file(path)).grid(
        column=3, row=1, sticky=tk.W)
    ttk.Button(container, text='Reset', width=20, command=lambda: reset_clicked(path)).grid(
        column=1, row=6, sticky=tk.W)
    ttk.Button(container, text='Close', width=20, command=lambda: close_clicked(container)).grid(
        column=3, row=6, sticky=tk.W)

    return container


def search_tables_clicked(path):
    """
    This function open window to select Access file.
    :param path: the directory to search.
    :type path: string
    :return: None
    """
    if os.path.isfile(path.get()):
        result = search_file.print_table_in_file(path.get())
        msg = result
        showinfo(
            title='Access Tables in selected fle',
            message=msg
        )
    else:
        msg = f'Wrong input, you entered: {path.get()} and it is not Access file'
        showinfo(
            title='Information',
            message=msg
        )


def search_columns_clicked(path):
    """
    This function open window to select Access file.
    :param path: the directory to search.
    :type path: string
    :return: None
    """
    if os.path.isfile(path.get()):
        result = search_file.find_columns_in_table(path.get())
        msg = result
        showinfo(
            title='Columns in selected file',
            message=msg
        )
    else:
        msg = f'Wrong input, you entered: {path.get()} and it is not Access file'
        showinfo(
            title='Information',
            message=msg
        )


def select_Access_file(path):
    """
    This function open window to select Access file 
    :param file_path: None
    :return: None
    """
    filetypes = (
        ('Access files', '*.mdb'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    path.insert(0, filename)


def create_search_column(container):
    """
    This function create search frame for columns value.
    :param container: the root window of tkinter.
    :type container: root window tkinter
    :return: container: the root window of tkinter.
    :rtype: root window tkinter
    """
    # Access Directory path
    ttk.Label(container, text='Enter Access file path:').grid(
        column=0, row=1, sticky=tk.W)
    path = ttk.Entry(container, width=40)
    path.grid(column=1, row=1, sticky=tk.W)
    path.focus()
    ttk.Button(container, text='Search', width=20, command=lambda: search_columns_clicked(path)).grid(
        column=0, row=6, sticky=tk.W)
    ttk.Button(container, text='Find Access File', width=20, command=lambda: select_Access_file(path)).grid(
        column=3, row=1, sticky=tk.W)
    ttk.Button(container, text='Reset', width=20, command=lambda: reset_clicked(path)).grid(
        column=1, row=6, sticky=tk.W)
    ttk.Button(container, text='Close', width=20, command=lambda: close_clicked(container)).grid(
        column=3, row=6, sticky=tk.W)

    return container


def search_files_clicked(path):
    """
    This function called when the search button is clicked.
    :param path: the directory to search.
    :type path: string
    :return: None
    """
    if os.path.isdir(path.get()):
        result = search_file.find_files(path.get())
        msg = result
        showinfo(
            title='Access files in selected directory',
            message=msg
        )
    else:
        msg = f'Wrong input, you entered: {path.get()} and it is not directory'
        showinfo(
            title='Information',
            message=msg
        )


def check_file(file_path):
    """
    This function check that file exist according given path, not empty and with csv extention
    :param file_path: file path to the csv file
    :type file_path: string
    :return: check that file exist according given path, not empty and with csv extention. If yes return True, else False
    :rtype: Boolean
    """

    if (path.isfile(file_path) and path.getsize(file_path) > 0):
        file_link = path.splitext(file_path)
        if file_link[1] == ".csv":
            return True
    else:
        return False


def search_clicked(keyword, path, CSV):
    """
    This function called when the search button is clicked.
    :param keyword: value for search.
    :type keyword: string
    :param path: the directory to search.
    :type path: string
    :param CSV: the csv file to save output.
    :type CSV: string
    :return: None
    """
    if not keyword.get():
        msg = f'Wrong input, you entered empty value'
        showinfo(
            title='Error',
            message=msg
        )
    elif check_file(CSV.get()) == FALSE:
        msg = f'Wrong input, you entered: {CSV.get()} and it is not CSV file'
        showinfo(
            title='Error',
            message=msg
        )

    elif os.path.isdir(path.get()):
        # value = r'אסותא-לב המפרץ חיפה'
        value = r'gvinot@bezeqint.net.il'
        search_file.find_value_in_directory(value, path.get(), CSV.get())
        msg = f'The result created in: {CSV.get()} file'
        showinfo(
            title='Information',
            message=msg
        )
    else:
        msg = f'Wrong input, you entered: {path.get()} and it is not directory'
        showinfo(
            title='Information',
            message=msg
        )


def reset_all_clicked(keyword, path, CSV):
    """
    This function called when the reset button is clicked and it empty all entries.
    :param keyword: value for search.
    :type keyword: string
    :param path: the directory to search.
    :type path: string
    :param CSV: the csv file to save output.
    :type CSV: string
    :return: None
    """
    keyword.delete(0, 'end')
    path.delete(0, 'end')
    CSV.delete(0, 'end')


def reset_clicked(path):
    """
    This function called when the reset button is clicked and it empty all entries.
    :param path: the directory to search.
    :type path: string
    :return: None
    """
    path.delete(0, 'end')


def close_clicked(container):
    """
    This function called when the close button is clicked and close all in search frame.
    :param container: the root window of tkinter.
    :type container: root window tkinter
    :return: None
    """
    list = container.grid_slaves()
    for l in list:
        l.destroy()


def create_main_window():
    """
    This function create main window, menus and executed in the loop.
    Param: None
    :return: None
    """
    # root window
    root = tk.Tk()
    root.title('Python GUI App')
    root.geometry('700x400')
    root.resizable(0, 0)
    # windows only (remove the minimize/maximize button)
    root.attributes('-toolwindow', True)
    # # layout on the root window
    root.columnconfigure(0, weight=4)
    root.columnconfigure(1, weight=1)
    menuBar = Menu(root)
    menuBar.columnconfigure(0, weight=1)
    root.config(menu=menuBar)
    # search Menu
    searchMenu = Menu(menuBar, tearoff=0)
    searchMenu.add_command(
        label="Search tables", command=lambda: create_search_tables(root))
    searchMenu.add_command(
        label="Search columns", command=lambda: create_search_column(root))
    searchMenu.add_command(
        label="Search Value in files", command=lambda: create_search_value(root))
    searchMenu.add_command(label="Exit", command=root.destroy)
    menuBar.add_cascade(label="Search", menu=searchMenu)
    # Create Report Menu todo
    createMenu = Menu(menuBar, tearoff=0)
    createMenu.add_command(label="Create")
    menuBar.add_cascade(label="Create report", menu=createMenu)
    # Export Report Menu todo
    exportMenu = Menu(menuBar, tearoff=0)
    exportMenu.add_command(label="Export")
    menuBar.add_cascade(label="Export to CSV", menu=exportMenu)
    # Open Access file Menu todo
    openMenu = Menu(menuBar, tearoff=0)
    openMenu.add_command(label="Open", command=select_file)
    menuBar.add_cascade(label="Open Access file", menu=openMenu)
    # Help Menu
    helpMenu = Menu(menuBar, tearoff=0)
    helpMenu.add_command(
        label="About", command=lambda: root.event_generate("<<OpenFindDialog>>"))
    helpMenu.add_command(
        label="Created by", command=lambda: root.event_generate("<<OpenCreatedDialog>>"))
    menuBar.add_cascade(label="Help", menu=helpMenu)
    root.bind("<<OpenFindDialog>>", launchFindDialog)
    root.bind("<<OpenCreatedDialog>>", launchCreatedDialog)
    root.mainloop()


try:
    if __name__ == "__main__":
        create_main_window()
except Exception as error:
    print(error)
