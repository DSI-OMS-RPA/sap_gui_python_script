import random
import sys
import time
import locale
import logging
import win32gui
import win32con
import subprocess
import pygetwindow as gw
import win32com.client as win32
from datetime import datetime, timedelta

# Define the SapGui class


class SapGui():
    """
    A class to manage SAP GUI interactions using the SAP GUI Scripting API and the win32 library.

    Attributes:
        system (str): SAP system platform.
        client (str): SAP client number.
        user (str): SAP username.
        password (str): SAP password.
        language (str): SAP language.
        path (str): Path to the SAPLogon executable.
        SapGuiAuto (object): SAP GUI Scripting engine object.
        connection (object): SAP connection object.
        session (object): SAP session object.

    Args:
        sap_args (dict): Dictionary containing SAP login credentials.

    Raises:
        Exception: If there is an error during initialization.

    Examples:
        sap_session = SapGui(sap_args)
    """

    def __init__(self, sap_args):
        """
        Initializes a SAP GUI session.

        Args:
            sap_args (dict): Dictionary containing SAP login credentials.

        Raises:
            Exception: If there is an error during initialization.
        """
        try:

            # Initialize instance variables for SAP configurations
            self.system = sap_args["platform"]
            self.client = sap_args['sap_client']
            self.user = sap_args["username"]
            self.password = sap_args["password"]
            self.language = sap_args['sap_language']

            # Path to SAPLogon executable
            self.path = sap_args['sap_path']

            # Open SAPLogon
            subprocess.Popen(self.path)
            time.sleep(2)  # Give it some time to open

            # Connect to the SAP GUI Scripting engine
            self.SapGuiAuto = win32.GetObject("SAPGUI")
            if not isinstance(self.SapGuiAuto, win32.CDispatch):
                return None

            # Get the SAP Scripting engine
            application = self.SapGuiAuto.GetScriptingEngine
            if not isinstance(application, win32.CDispatch):
                self.SapGuiAuto = None
                return None

            # Open a connection to the SAP system
            self.connection = application.OpenConnection(self.system, True)
            if not isinstance(self.connection, win32.CDispatch):
                application = None
                self.SapGuiAuto = None
                return None

            # Wait for the connection to be established
            time.sleep(3)

            # Get the first available session
            self.session = self.connection.Children(0)
            if not isinstance(self.session, win32.CDispatch):
                self.connection = None
                application = None
                self.SapGuiAuto = None
                return None

            # Resize the SAP GUI window
            self.session.findById("wnd[0]").resizeWorkingPane(169, 30, False)

            # Maximize the main window (window 0) in SAP GUI.
            # self.session.findById("wnd[0]").maximize()

        except Exception as e:
            logging.error(f"An exception occurred in __init__: {str(e)}")
            return None

    def handle_password_change(self):
        """
        Handles the password change prompt during SAP login.

        This function detects a password change popup, generates a new password in the format
        'Month#Year' (e.g., 'Maio#2024'), inputs the password in both fields, and attempts to
        log the user in. If successful, it returns True; otherwise, it returns False.

        Returns:
            bool: True if the password change was successful and the user is logged in, False otherwise.
        """
        try:
            # Ensure the locale is set to Portuguese for correct month formatting
            try:
                locale.setlocale(locale.LC_TIME, 'pt_PT.UTF-8')
            except locale.Error as e:
                logging.error(f"Locale setting failed: {str(e)}")
                return False

            # Check if a password change window is active (usually wnd[1])
            active_window = self.session.ActiveWindow
            if active_window.Name != "wnd[1]":
                logging.info("No password change prompt detected.")
                return True

            # Retrieve the popup window and its title
            popup_window = self.session.findById("wnd[1]")
            popup_title = popup_window.Text.lower()

            # Ensure it's a password change prompt by checking the title
            if "sap" not in popup_title:
                logging.error("Unexpected popup encountered, not a password change prompt.")
                return False

            # Verify the label to confirm it's asking for a new password
            try:
                input_label = popup_window.findById("usr/lblRSYST-NCODE_TEXT").Text.lower()
                if "nova senha" not in input_label:
                    logging.error("No valid password change prompt detected.")
                    return False
            except Exception as e:
                logging.error(f"Failed to find password change label: {str(e)}")
                return False

            logging.info("Password change prompt detected.")

            # Generate new password in 'NumberMonth#Year' format (e.g., '123Maio#2024')
            number = random.randrange(1, 999)
            new_password = f"{number}{datetime.now().strftime('%B').capitalize()}#{datetime.now().strftime('%Y')}"
            logging.info(f"Generated new password: {new_password}")

            # Input the new password into both fields
            popup_window.findById("usr/pwdRSYST-NCODE").text = new_password
            popup_window.findById("usr/pwdRSYST-NCOD2").text = new_password

            # Confirm the password change
            popup_window.findById("tbar[0]/btn[0]").press()

            # Wait for a short while to allow the change to take effect
            time.sleep(3)

            # Check if the user is logged in successfully after password change
            if self.session.findById("wnd[0]/tbar[0]/btn[15]", False):
                logging.info("Password changed successfully and user logged in.")
                return True
            else:
                logging.error("Password change failed or login unsuccessful.")
                return False

        except Exception as e:
            # Log any errors that occur during the process
            logging.error(f"Error during password change handling: {str(e)}")
            return False

    def sapLogin(self):
        """
        Logs in to SAP.

        Returns:
            bool: True if login is successful, False otherwise.
        """
        try:
            # Set the SAP login credentials and language in the GUI
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = self.client  # Mandante
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.user  # Utilizador
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password  # Password
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = self.language  # Idioma

            # Perform the login
            self.session.findById("wnd[0]").sendVKey(0)

            # Wait for a short time to see if any popup appears
            time.sleep(2)

            # Handle password change if required
            if not self.handle_password_change():
                # If password change handling fails, return False
                return False

            # Check for the specific popup window
            if self.session.ActiveWindow.Name == "wnd[1]":
                # Check if the popup is the multiple login warning by checking some unique text or title
                if "logon múltiplo" in self.session.findById("wnd[1]").Text:
                    logging.info("Multiple logins detected. Closing other sessions.")
                    # Select the option to close other sessions and click OK
                    self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # Check if login is successful by finding a UI element that appears only when logged in
            if self.session.findById("wnd[0]/tbar[0]/btn[15]"):
                # login sucessfull
                logging.info("Successfully connected to SAP.")
                return True
            else:
                # Login failed and Close the SAP GUI connection
                self.close_connection()
                return False

        except Exception as e:
            logging.error(f"Error during SAP login: {str(e)}")
            logging.error(sys.exc_info())

    def close_connection(self):
        """
        Closes the SAP connection.

        Returns:
            None
        """
        try:
            # Check if a connection object exists
            if self.connection is not None:
                self.connection.CloseSession('ses[0]')
                # Set the connection to None, indicating it's closed
                self.connection = None
                # Log a message indicating that the SAP connection is closed
                logging.info("SAP connection closed.")
            if self.SapGuiAuto is not None:
                self.SapGuiAuto = None
            logging.info("SAP connection closed safely.")
        except Exception as e:
            # Handle any exceptions that may occur during the closing process
            # Log an error message with details about the exception
            logging.error(f"Error closing SAP connection: {str(e)}")

    def sapLogout(self):
        """
        Logs out of SAP.

        Args:
            self: The instance of the SAP session.

        Raises:
            Exception: If there is an error during SAP logout.

        Examples:
            sap_session.sapLogout()
        """
        try:
            # Enter the logout command '/nex' in the command field
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            self.session.findById("wnd[0]").sendVKey(0)

            logging.info("Successfully logged out of SAP.")
        except Exception as e:
            logging.error(f"Error during SAP logout: {str(e)}")

    @staticmethod
    def get_dates():
        """
        Returns the start and end dates for the previous month.

        Returns:
            tuple: Start date and end date in the format "%d.%m.%Y".
        """
        current_date = datetime.now()

        # Get the first day of the current month
        first_day_current_month = current_date.replace(day=1)

        # Subtract one day from the first day of the current month to get the last day of the previous month
        previous_month_last_day = first_day_current_month - timedelta(days=1)

        # Set the start_date as the first day of the previous month
        start_date = previous_month_last_day.replace(
            day=1).strftime("%d.%m.%Y")
        end_date = current_date.strftime("%d.%m.%Y")

        return start_date, end_date

    # Waits for the SAP GUI element with the specified ID to appear within the given timeout.
    def wait_for_element(self, element_id, timeout=60):
        """
        Waits for the SAP GUI element with the specified ID to appear within the given timeout.
        Args:
            element_id (str): SAP GUI Scripting ID of the element to wait for.
            timeout (int, optional): The number of seconds to wait for the element. Defaults to 60.

        Returns:
            bool: True if the element appears within the timeout, otherwise False.
        """
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if self.session.findById(element_id):
                    return True
            except Exception:
                time.sleep(1)  # wait for 1 second before trying again
        return False

    def check_element_exists(self, element_path):
        """
        Checks if a SAP element exists.

        Args:
            element_path (str): The path of the SAP element.

        Returns:
            bool: True if the element exists, otherwise False.
        """

        try:
            self.session.findById(element_path)
            return True
        except Exception:
            return False

    def wait_for_save_as_dialog(self, title, max_attempts=10):
        """
        Waits for a dialog window with the specified title to appear.

        This function attempts to detect the appearance of a dialog window by checking if a window
        with the given title is present. It waits between each attempt for the window to appear,
        up to a specified maximum number of attempts.

        Parameters:
        title (str): The title of the dialog window to wait for.
        max_attempts (int): The maximum number of attempts to check for the window. Default is 10.

        Returns:
        bool: True if the dialog window is detected within the max_attempts, otherwise False.
        """
        for attempt in range(max_attempts):
            if gw.getWindowsWithTitle(title):
                return True
            time.sleep(1)  # Wait for 1 second before the next attempt

        return False

    def get_sap_element_text(self, element_path):
        """
        Retrieves the text of a SAP element identified by the given element_path.

        Args:
            element_path (str): The path of the SAP element.

        Returns:
            str: The text of the SAP element, or None if the element is not found or an error occurs.
        """
        try:
            element = self.session.FindById(element_path)
            return element.Text
        except Exception as e:
            print(f"Error: {e}")
            return None

    def bring_dialog_to_top(self, title):
        """
        Brings a dialog window with the specified title to the top of the screen.

        This function checks if a dialog window is open and, if found, restores it if minimized,
        shows it, and brings it to the top of the screen.

        Parameters:
        title (str): The title of the dialog window to bring to the top.

        Returns:
        bool: True if the window was found and brought to the top, False otherwise.
        """
        save_as_window = gw.getWindowsWithTitle(title)
        if save_as_window:
            window_handle = save_as_window[0]._hWnd
            try:
                # Restore the window if minimized
                win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
                # Show the window
                win32gui.ShowWindow(window_handle, win32con.SW_SHOWNORMAL)
                # Bring the window to the top
                win32gui.BringWindowToTop(window_handle)
                return True
            except Exception as e:
                print(f"Error bringing window to top: {e}")
                return False
        return False

    def scroll_to_field(self, field_path):
        try:
            self.session.findById(field_path).setFocus()
        except Exception:
            self.session.findById(field_path.split("/")[:-1].join("/")).verticalScrollbar.position += 1

    def set_cell_value(self, column_path, text):
        """
        Find the first empty cell in a table column and update it with a given text.

        Args:
            column_path (str): The path of the column in the table.
            text (str): The text to update the empty cell with.

        Returns:
            None
        """
        # Find the first empty row
        row_number = 0
        cell_path = column_path.format(row_number)
        cell = self.session.findById(cell_path)

        while cell and cell.Text != "":
            row_number += 1
            cell_path = column_path.format(row_number)
            cell = self.session.findById(cell_path)

        # Check if a blank cell was found
        if cell:
            # If the cell is empty, write "ZREC"
            cell.Text = text
            cell.setFocus()
            cell.caretPosition = 4
            self.session.findById("wnd[0]").sendVKey(0)

        return row_number

    # Enter a command in the SAP command field, submit, and perform additional operations.
    def perform_operation(self, command):
        """
        Performs the specified command in the SAP GUI.

        Args:
            command (str): The command to be executed.

        Returns:
            None
        """
        try:
            # Set the value of the specified field and Submit the command
            self.session.findById("wnd[0]/tbar[0]/okcd").text = command
            self.session.findById("wnd[0]").sendVKey(0)

            # Wait for the specific element from the new page to be present
            element_to_wait_for = "element_here"
            if self.wait_for_element(element_to_wait_for):
                logging.info("Element found.")
            else:
                logging.error(f"Element {element_to_wait_for} not found.")
        except Exception as e:
            logging.error(f"Error during command execution: {str(e)}")
