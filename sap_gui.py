import random
import sys
import time
import locale
import logging
import win32gui
import win32con
import win32api
import subprocess
import pygetwindow as gw
import win32com.client as win32
import os
import psutil
from datetime import datetime, timedelta
from typing import Optional, Tuple, Any, Dict
from dataclasses import dataclass

class SapConfigError(Exception):
    """Custom exception for SAP configuration errors."""
    pass

@dataclass
class SapConfig:
    """
    Configuration for SAP connection.

    Attributes:
        platform: SAP system identifier (e.g., 'PRD', 'QAS')
        client: SAP client number (e.g., '100')
        username: SAP login username
        password: SAP login password
        language: SAP interface language code
        path: Path to SAP executable
    """
    platform: str
    client: str
    username: str
    password: str
    language: str
    path: str = "saplogon.exe"

    def __post_init__(self):
        """Validate configuration parameters."""
        if not all([self.platform, self.client, self.username, self.password, self.language, self.path]):
            raise SapConfigError("All SAP configuration fields are required")
        if not self.client.isdigit():
            raise SapConfigError("SAP client must be numeric")
        if len(self.language) != 2:
            raise SapConfigError("Language code must be 2 characters")

class SapGuiError(Exception):
    """Custom exception for SAP GUI related errors"""
    pass

def run_application(software: str) -> bool:
    """
    Search for and run an application on the system.

    Args:
        software (str): Application executable name to find and run

    Returns:
        bool: True if application found and launched successfully
    """
    def search_drive(drive: str, software: str) -> Optional[str]:
        for root, _, files in os.walk(drive):
            if software in files:
                return os.path.join(root, software)
        return None

    def is_process_running(process_name: str) -> bool:
        for proc in psutil.process_iter(['pid', 'name']):
            if process_name.lower() in proc.info['name'].lower():
                return True
        return False

    if is_process_running(software):
        logging.info(f"{software} is already running")
        return True

    start_time = time.time()
    drives = win32api.GetLogicalDriveStrings().split('\000')[:-1]

    for drive in drives:
        software_path = search_drive(drive, software)
        if software_path:
            elapsed_time = time.time() - start_time
            logging.info(f"{software} found at {software_path}. Starting...")
            subprocess.Popen(software_path)
            logging.info(f"Time taken: {elapsed_time:.2f} seconds")
            return True

    elapsed_time = time.time() - start_time
    logging.warning(f"{software} not found. Search time: {elapsed_time:.2f} seconds")
    return False

class SapGui:
    """
    Manages SAP GUI automation using SAP GUI Scripting API and win32 library.

    This class provides a comprehensive interface for automating SAP GUI operations,
    including session management, authentication, window handling, and element interaction.
    It implements robust error handling and retry mechanisms for reliable automation.

    Key Features:
        - Automated login/logout with multi-session handling
        - Password change management with secure password generation
        - Window and dialog management with focus control
        - Element interaction with timeout and retry mechanisms
        - Comprehensive error handling and logging
        - Resource cleanup and connection management

    Example Usage:
        sap_config = {
            "platform": "PRD",
            "sap_client": "100",
            "username": "user",
            "password": "pass",
            "sap_language": "PT",
            "sap_path": "C:/Path/To/SAPLogon.exe"
        }

        sap = SapGui(sap_config)
        try:
            if sap.sapLogin():
                # Perform SAP operations
                sap.perform_operation("/nVA01")
            sap.sapLogout()
        finally:
            sap.close_connection()

    Note:
        Requires SAP GUI Scripting to be enabled in the SAP system
        and appropriate user permissions for automation.
    """

    # Class constants
    DEFAULT_TIMEOUT = 60
    DEFAULT_RETRY_ATTEMPTS = 3
    RETRY_DELAY = 1

    def __init__(self, sap_args: dict):
        """
        Initialize a SAP GUI session with enhanced error handling and typing.

        Args:
            sap_args: Dictionary containing SAP configuration parameters

        Raises:
            SapGuiError: If initialization fails
        """
        try:
            config = SapConfig(**sap_args)
            self._initialize_connection(config)
            self._setup_logging()
        except Exception as e:
            raise SapGuiError(f"Failed to initialize SAP GUI: {str(e)}") from e

    def _initialize_connection(self, config: SapConfig) -> None:
        """
        Initialize SAP GUI connection with retry mechanism.

        Establishes connection to SAP system using provided configuration.
        Implements retry logic to handle intermittent connection issues.

        Args:
            config: SapConfig instance containing connection parameters

        Raises:
            SapGuiError: If connection fails after maximum retry attempts

        Note:
            Performs connection steps with delays to ensure stable initialization:
            1. Launches SAPLogon executable
            2. Connects to SAP GUI automation engine
            3. Opens connection to specified SAP system
            4. Initializes session and configures window
        """
        self.config = config
        retry_count = 0

        while retry_count < self.DEFAULT_RETRY_ATTEMPTS:
            try:

                # Find and execute SAPLogon
                runner = run_application(config.path)
                if not runner:
                    raise Exception("Failed to run SAPLogon.")

                # subprocess.Popen(self.path)
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
                self.connection = application.OpenConnection(config.platform, True)
                time.sleep(3)

                self.session = self.connection.Children(0)
                self.session.findById("wnd[0]").resizeWorkingPane(169, 30, False)
                return

            except Exception as e:
                retry_count += 1
                if retry_count == self.DEFAULT_RETRY_ATTEMPTS:
                    raise SapGuiError(f"Failed to initialize after {retry_count} attempts: {str(e)}")
                time.sleep(self.RETRY_DELAY)

    def _setup_logging(self) -> None:
        """Configure logging with appropriate format and level"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

    @staticmethod
    def generate_password() -> str:
        """Generate a new password following the required format"""
        try:
            locale.setlocale(locale.LC_TIME, 'pt_PT.UTF-8')
        except locale.Error:
            logging.warning("Failed to set Portuguese locale, using system default")

        number = random.randrange(1, 999)
        return f"{number}{datetime.now().strftime('%B').capitalize()}#{datetime.now().strftime('%Y')}"

    def scroll_to_field(self, field_path: str) -> None:
        """
        Scrolls to make the specified field visible in the SAP GUI window.

        Args:
            field_path: Path to the SAP GUI field element

        Note:
            If direct focus fails, incrementally scrolls the vertical scrollbar
        """
        try:
            self.session.findById(field_path).setFocus()
        except Exception:
            parent_path = "/".join(field_path.split("/")[:-1])
            self.session.findById(parent_path).verticalScrollbar.position += 1

    def handle_password_change(self) -> bool:
        """
        Handles SAP password change prompt with secure password generation.

        Detects password change window and automates the process by:
        - Generating secure password following SAP requirements
        - Filling both password fields
        - Submitting the change
        - Verifying success

        Returns:
            bool: True if password change successful, False otherwise

        Note:
            Generated password format: {number}{Month}#{Year}
            Requires Portuguese locale for month formatting
        """
        try:
            # Set Portuguese locale for month formatting
            try:
                locale.setlocale(locale.LC_TIME, 'pt_PT.UTF-8')
            except locale.Error:
                logging.warning("Portuguese locale unavailable, using system default")

            # Check for password change window
            active_window = self.session.ActiveWindow
            if active_window.Name != "wnd[1]":
                return True

            popup_window = self.session.findById("wnd[1]")
            if not "nova senha" in popup_window.findById("usr/lblRSYST-NCODE_TEXT").Text.lower():
                return False

            # Generate and set new password
            new_password = self.generate_password()
            popup_window.findById("usr/pwdRSYST-NCODE").text = new_password
            popup_window.findById("usr/pwdRSYST-NCOD2").text = new_password
            popup_window.findById("tbar[0]/btn[0]").press()

            time.sleep(3)
            return bool(self.session.findById("wnd[0]/tbar[0]/btn[15]", False))

        except Exception as e:
            logging.error(f"Password change failed: {str(e)}")
            return False


    def sapLogin(self) -> bool:
        """
        Perform SAP system login with comprehensive error handling.

        Executes the complete login sequence including:
        - Setting login credentials (client, username, password, language)
        - Handling password change prompts if required
        - Managing multiple login scenarios
        - Verifying successful login state

        Returns:
            bool: True if login successful and session is active,
                 False if any step of the login process fails

        Raises:
            SapGuiError: If critical error occurs during login process

        Note:
            - Automatically handles password change requirements
            - Implements resolution for multiple active sessions
            - Performs validation of login state through UI element checks
            - Ensures proper cleanup on login failure
        """
        try:
            # Set login credentials
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = self.config.client
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.config.username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.config.password
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = self.config.language

            # Submit login credentials
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)

            if not self.handle_password_change():
                return False

            # Handle multiple login scenario
            if self._handle_multiple_login():
                return self._verify_login()

            return False

        except Exception as e:
            logging.error(f"Login failed: {str(e)}")
            self.close_connection()
            return False

    def _handle_multiple_login(self) -> bool:
        """Handle multiple login scenario"""
        try:
            if self.session.ActiveWindow.Name == "wnd[1]":
                if "logon múltiplo" in self.session.findById("wnd[1]").Text:
                    self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            return True
        except Exception as e:
            logging.error(f"Multiple login handling failed: {str(e)}")
            return False

    def _verify_login(self) -> bool:
        """Verify successful login"""
        try:
            return bool(self.session.findById("wnd[0]/tbar[0]/btn[15]"))
        except Exception:
            return False

    def close_connection(self) -> None:
        """Safely close SAP connection with resource cleanup"""
        try:
            if self.connection:
                self.connection.CloseSession('ses[0]')
                self.connection = None
            if self.SapGuiAuto:
                self.SapGuiAuto = None
            logging.info("SAP connection closed safely")
        except Exception as e:
            logging.error(f"Error closing connection: {str(e)}")

    def sapLogout(self) -> None:
        """Perform SAP logout"""
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            self.session.findById("wnd[0]").sendVKey(0)
            logging.info("Successfully logged out")
        except Exception as e:
            logging.error(f"Logout failed: {str(e)}")

    @staticmethod
    def get_dates() -> Tuple[str, str]:
        """
        Get start and end dates for the previous month.

        Returns:
            Tuple containing start_date and end_date strings
        """
        current_date = datetime.now()
        first_day_current = current_date.replace(day=1)
        prev_month_last = first_day_current - timedelta(days=1)

        return (
            prev_month_last.replace(day=1).strftime("%d.%m.%Y"),
            current_date.strftime("%d.%m.%Y")
        )

    def wait_for_element(self, element_id: str, timeout: int = DEFAULT_TIMEOUT) -> bool:
        """
        Wait for SAP GUI element to become available with timeout.

        Repeatedly attempts to find specified element until timeout occurs.
        Uses polling mechanism with 1-second intervals between attempts.

        Args:
            element_id: SAP GUI scripting ID of target element
                       Format: "wnd[0]/usr/txtField" or similar SAP path
            timeout: Maximum wait time in seconds before giving up
                    Defaults to DEFAULT_TIMEOUT class constant

        Returns:
            bool: True if element found within timeout period,
                 False if element not found after timeout

        Example:
            >>> # Wait for input field to appear
            >>> if sap.wait_for_element("wnd[0]/usr/txtVBELN"):
            >>>     # Element found, proceed with interaction
            >>> else:
            >>>     # Handle timeout case

        Note:
            - Implements non-blocking wait with regular polling
            - Silently handles exceptions during element search
            - Use for synchronizing automation with SAP UI state
        """
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if self.session.findById(element_id):
                    return True
            except Exception:
                time.sleep(1)
        return False

    def check_element_exists(self, element_path: str) -> bool:
        """Check if SAP element exists"""
        try:
            self.session.findById(element_path)
            return True
        except Exception:
            return False

    def wait_for_save_as_dialog(self, title: str, max_attempts: int = 10) -> bool:
        """Wait for save dialog with specified title"""
        for _ in range(max_attempts):
            if gw.getWindowsWithTitle(title):
                return True
            time.sleep(1)
        return False

    def get_sap_element_text(self, element_path: str) -> Optional[str]:
        """Get text from SAP element with error handling"""
        try:
            element = self.session.FindById(element_path)
            return element.Text
        except Exception as e:
            logging.error(f"Failed to get element text: {str(e)}")
            return None

    def bring_dialog_to_top(self, title: str) -> bool:
        """Bring dialog window to top of screen"""
        try:
            save_as_window = gw.getWindowsWithTitle(title)
            if save_as_window:
                window_handle = save_as_window[0]._hWnd
                win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
                win32gui.ShowWindow(window_handle, win32con.SW_SHOWNORMAL)
                win32gui.BringWindowToTop(window_handle)
                return True
            return False
        except Exception as e:
            logging.error(f"Failed to bring dialog to top: {str(e)}")
            return False

    def perform_operation(self, command: str, element_to_wait_for: Optional[str] = None, timeout: int = DEFAULT_TIMEOUT) -> bool:
        """
        Executes SAP command and verifies operation completion.

        Args:
            command: SAP transaction code or command (e.g., '/nVA01')
            element_to_wait_for: Optional element ID to verify operation success
            timeout: Maximum wait time for element appearance

        Returns:
            bool: True if operation successful and element found (if specified)

        Raises:
            SapGuiError: If command execution fails or element not found
        """
        try:
            # Execute command
            self.session.findById("wnd[0]/tbar[0]/okcd").text = command
            self.session.findById("wnd[0]").sendVKey(0)

            # Verify operation if element specified
            if element_to_wait_for:
                if not self.wait_for_element(element_to_wait_for, timeout):
                    raise SapGuiError(f"Element {element_to_wait_for} not found after command {command}")
                logging.info(f"Command {command} executed, element {element_to_wait_for} found")
            else:
                logging.info(f"Command {command} executed")
            return True

        except Exception as e:
            logging.error(f"Operation failed - Command: {command}, Error: {str(e)}")
            raise SapGuiError(f"Operation failed: {str(e)}")

    def set_cell_value(self, column_path: str, text: str) -> int:
        """
        Set value in first empty cell of specified column.

        Args:
            column_path: Path to table column
            text: Text to set in cell

        Returns:
            int: Row number where value was set
        """
        row_number = 0
        try:
            while True:
                cell_path = column_path.format(row_number)
                cell = self.session.findById(cell_path)

                if not cell or cell.Text == "":
                    cell.Text = text
                    cell.setFocus()
                    cell.caretPosition = len(text)
                    self.session.findById("wnd[0]").sendVKey(0)
                    return row_number

                row_number += 1

        except Exception as e:
            logging.error(f"Failed to set cell value: {str(e)}")
            raise SapGuiError(f"Failed to set cell value: {str(e)}")
