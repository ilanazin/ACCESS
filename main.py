import gui_module
import utils


def main_function():
    """
    This function is main function of the program.
    :return: Depend on the input in case of error the program return an error.
    :rtype: None
    """
    try:
        utils.print_wellcome_screen()
        utils.create_files()
        gui_module.create_main_window()
    except Exception as error:
        print(error)
    finally:
        utils.print_close_screen()


main_function()
