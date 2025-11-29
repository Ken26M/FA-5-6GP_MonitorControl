"""
AFCOM - Serial Communication GUI Program
Cannot be used directly, it is a part of main.py
"""
"""
Customized for FA-5 application by Ken26M 2025.11
github.com/Ken26M/FA-5-6GP_MonitorControl
"""

__author__ = 'Mehmet Cagri Aksoy - github.com/mcagriaksoy'
__annotations__ = 'AFCOM - Serial Communication GUI Program'
__version__ = '2024.05'
__license__ = 'MIT'
__status__ = 'Research'

# IMPORTS
from sys import argv
import time
import os
import fa5usbdata as fa5
from fa5usbdata import Commands
import traceback
import logging
import sys
from subprocess import call
from decimal import Decimal
import json
from enum import Enum
from PyQt6.QtCore import QObject, QThread, pyqtSignal, Qt, QMargins, QEvent


class Device_Type(Enum):
    FA_2 = "FA-2"
    FA_3 = "FA-3"
    FA_5 = "FA-5"
    TinyGTC = "TinyGTC"


# Use %APPDATA% on Windows (fall back to user's home if APPDATA not set)
_APPDATA = os.getenv('LOCALAPPDATA') or os.path.join(os.path.expanduser("~"), 'AppData', 'Local')
_SETTINGS_DIR = os.path.join(_APPDATA, 'FA5_MonitorControl')
_SETTINGS_FILENAME = 'settings.json'
SETTINGS_FILE = os.path.join(_SETTINGS_DIR, _SETTINGS_FILENAME)

# Single source of truth for default settings to avoid duplication
DEFAULT_SETTINGS = {
    "saved_command_1": Commands.GET_GATE_TIME_SETTING.value,
    "saved_command_2": "",
    "saved_command_3": "",
    "saved_command_4": Commands.SET_GATE_TIME_100MS.value,
    "device_type": Device_Type.FA_5.value,
}


def load_user_settings():
    """Load persisted settings from disk, return dict with defaults."""
    defaults = DEFAULT_SETTINGS.copy()
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    defaults.update(data)
        else:
            # Create directory if needed
            os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
            # Save defaults to file
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(defaults, f, indent=4)
    except Exception:
        logging.exception("Failed to read or create settings file, using defaults")
    return defaults


def save_user_settings(settings_dict):
    """Persist the settings dict to disk. Creates settings dir if needed. Errors are logged but not raised."""
    try:
        # Ensure settings directory exists
        try:
            os.makedirs(_SETTINGS_DIR, exist_ok=True)
        except Exception:
            # If we can't create the folder, still attempt to write (may fail)
            logging.debug(f"Could not create settings dir {_SETTINGS_DIR}")
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings_dict, f, indent=2)
    except Exception:
        logging.exception("Failed to save user settings")


# Runtime Type Checking
PROGRAM_TYPE_DEBUG = False
PROGRAM_TYPE_RELEASE = True

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Attempt to import PySerial-related modules
try:
    import serial.tools.list_ports
    from serial import SerialException, Serial
except ImportError:
    logging.error("Import Error: PySerial library is not installed.")
    logging.info("Attempting to install PySerial...")
    try:
        call([sys.executable, "-m", "pip", "install", "pyserial"])
        logging.info("PySerial installed successfully. Please restart the program.")
        sys.exit(0)
    except Exception as install_error:
        logging.error(f"Failed to install PySerial: {install_error}")
        sys.exit(1)

# Attempt to import PyQt6 core and widgets (required)
try:
    from PyQt6.QtCore import QObject, QThread, pyqtSignal, Qt, QMargins, QEvent
    from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QInputDialog, QSizePolicy
    from PyQt6 import uic
    from PyQt6.QtGui import QPainter

    if PROGRAM_TYPE_DEBUG:
        from PyQt6.uic import loadUi
    else:
        from ui_config import Ui_main_window

except ImportError as e:
    logging.critical("Failed to import essential PyQt6 modules. Please install PyQt6 before running this application.")
    logging.critical("You can usually install it with: python -m pip install PyQt6")
    logging.critical(f"ImportError details: {e}")
    sys.exit(1)

CHARTS_AVAILABLE = True
try:
    from PyQt6.QtCharts import QChart, QChartView, QValueAxis, QLineSeries
except ImportError as e:
    logging.critical(
        "Failed to import essential PyQt6-Charts modules. Please install PyQt6-Charts before running this application.")
    logging.critical("You can usually install it with: python -m pip install PyQt6-Charts")
    logging.critical(f"ImportError details: {e}")
    CHARTS_AVAILABLE = False
    pass

# GLOBAL VARIABLES
SERIAL_CON = Serial()
PORTS = []
is_serial_port_established = False
ml = fa5.MeasureLog()
receive_time_of_freq = time.time()
waiting_for_command = True
command_queue = []


def get_serial_port():
    """ Lists serial port names

        :raises EnvironmentError:
            On unsupported or unknown platforms
        :returns:
            A list of the serial ports available on the system
    """

    ports = serial.tools.list_ports.comports()

    # COM ID GB7TBL (
    # ID = p.vid : p.pid
    # IDhex=0403:6001
    # IDdec=1027:24577

    result = []
    for port in ports:
        try:
            s = Serial(port.device)
            s.close()
            if (port.vid == 1027 and port.pid == 24577) or (port.vid == 1155 and port.pid == 22337):  # Filter out other com devices
                result.insert(0,port.device)
        except SerialException:
            pass
    return result


# MULTI-THREADING
class Worker(QObject):
    """ Worker Thread """
    finished = pyqtSignal()
    serial_data = pyqtSignal(str, object)

    def __init__(self):
        super(Worker, self).__init__()
        self.working = True

    def read_line(self):
        # Read data from the serial port
        # 1 line of measure data is not send in 1 go, if the gate time is 10 sec 2 text parts will be 10 sec apart
        # This makes sure we only start processing the received data when the text line is complete
        buffer = ""
        while True:
            if SERIAL_CON.in_waiting > 0:
                # Read one byte at a time
                byte = SERIAL_CON.readline()
                datastring = byte.decode('utf-8', errors='ignore')
                # print('byte:', byte)
                if datastring:
                    # Append the byte to the buffer
                    buffer += datastring

                    # Check if the buffer ends with CR LF
                    if datastring.endswith('\n'):
                        # Extract the line and reset the buffer
                        line = buffer  # .rstrip('\r\n')
                        # print('line:', line)
                        return line
            # Small delay to prevent busy-waiting
            else:
                time.sleep(0.1)
            # print('sleep 0.01')

    def work(self):
        """ Read data from serial port """
        while self.working:
            try:
                line = self.read_line()  # get full text line from FA-5
                is_data_stream = ml.add_string(line)  # add measurement to log and update settings if applicable
                metadata = {"datastream": is_data_stream}  # example extra parameter
                self.serial_data.emit(line, metadata)
            except SerialException as e:
                print(e)
                # Emit last error message before die!
                self.serial_data.emit("ERROR_SERIAL_EXCEPTION", {"error": str(e)})
                self.working = False
        # Only emit finished when the worker actually stops (after loop exits)
        self.finished.emit()
        SERIAL_CON.close()


class MainWindow(QMainWindow):
    """ Main Window """

    def setup_frequency_plot(self):
        """Initialize the frequency plot"""
        # If QtCharts isn't available, skip chart setup to avoid NameError at import/runtime
        if not CHARTS_AVAILABLE:
            logging.warning("QtCharts not available — skipping frequency chart setup.")
            try:
                # Hide existing placeholder graphics view if present
                if hasattr(self, 'graphicsViewBottom') and self.graphicsViewBottom is not None:
                    try:
                        self.graphicsViewBottom.hide()
                    except Exception:
                        pass
            except Exception:
                pass
            return

        self.freq_series = QLineSeries()
        pen = self.freq_series.pen()
        pen.setWidthF(0.5)  # thinner line for better visibility
        self.freq_series.setPen(pen)
        self.freq_series.setName("Frequency")

        self.chart = QChart()
        self.chart.addSeries(self.freq_series)
        # Remove title (user doesn't want it) and hide legend to save space
        self.chart.setTitle("")
        if self.chart.legend() is not None:
            self.chart.legend().setVisible(False)

        self.axis_x = QValueAxis()
        # remove axis title to save vertical space
        self.axis_x.setTitleText("")
        self.axis_y = QValueAxis()
        self.axis_y.setTitleText("")

        self.chart.addAxis(self.axis_x, Qt.AlignmentFlag.AlignBottom)
        self.chart.addAxis(self.axis_y, Qt.AlignmentFlag.AlignLeft)
        self.freq_series.attachAxis(self.axis_x)
        self.freq_series.attachAxis(self.axis_y)

        self.chart_view = QChartView(self.chart)
        self.chart_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        # Remove spacing around the chart view widget
        self.chart_view.setContentsMargins(0, 0, 0, 0)
        try:
            # make the chart view expand to fill its layout cell
            self.chart_view.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        except Exception:
            pass
        try:
            self.chart_view.setViewportMargins(0, 0, 0, 0)
        except Exception:
            pass

        # Try to remove internal chart margins / padding (may not be supported on all Qt bindings)
        try:
            self.chart.setContentsMargins(0, 0, 0, 0)
            self.chart.setMargins(QMargins(0, 0, 0, 0))
        except Exception:
            pass

        # Clear layout margins and spacing around the chart to minimize empty space
        try:
            self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
            self.gridLayout_2.setSpacing(0)
        except Exception:
            pass

        # Replace the graphicsViewBottom widget with the chart view
        try:
            self.gridLayout_2.replaceWidget(self.graphicsViewBottom, self.chart_view)
            self.graphicsViewBottom.hide()
            self.chart_view.show()
        except Exception:
            # If layout replacement fails, fallback to hiding the placeholder
            try:
                if hasattr(self, 'graphicsViewBottom') and self.graphicsViewBottom is not None:
                    self.graphicsViewBottom.hide()
            except Exception:
                pass

    def clear_frequency_plot(self):
        """Clear the frequency plot and related displays so the graph appears fresh.

        This removes all points from the series, clears the data text fields
        and resets the axes to a minimal default range.
        """
        # If charts are not available, just clear text fields and exit
        if not CHARTS_AVAILABLE:
            try:
                if hasattr(self, 'data_textEdit') and self.data_textEdit is not None:
                    try:
                        self.data_textEdit.clear()
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                self.start_time = time.time()
            except Exception:
                pass
            return

        try:
            # Remove the existing series from the chart (if present) and create a fresh one.
            if hasattr(self, 'chart') and self.chart is not None:
                try:
                    if hasattr(self, 'freq_series') and self.freq_series is not None:
                        try:
                            self.chart.removeSeries(self.freq_series)
                        except Exception:
                            pass
                    # create a fresh series and attach axes
                    self.freq_series = QLineSeries()
                    pen = self.freq_series.pen()
                    pen.setWidthF(0.5)  # thinner line for better visibility
                    self.freq_series.setPen(pen)
                    self.freq_series.setName("Frequency")
                    self.chart.addSeries(self.freq_series)
                    try:
                        # Re-attach axes if available
                        # Attach existing axes to the fresh series (don't add axes again)
                        if hasattr(self, 'axis_x') and self.axis_x is not None:
                            try:
                                self.freq_series.attachAxis(self.axis_x)
                            except Exception:
                                pass
                        if hasattr(self, 'axis_y') and self.axis_y is not None:
                            try:
                                self.freq_series.attachAxis(self.axis_y)
                            except Exception:
                                pass
                    except Exception:
                        pass
                except Exception:
                    pass

            # Clear measurement text areas if present
            if hasattr(self, 'data_textEdit') and self.data_textEdit is not None:
                try:
                    self.data_textEdit.clear()
                except Exception:
                    pass

            # Reset axes ranges to a small default
            if hasattr(self, 'axis_x') and self.axis_x is not None:
                try:
                    self.axis_x.setRange(0, 1)
                except Exception:
                    pass
            if hasattr(self, 'axis_y') and self.axis_y is not None:
                try:
                    self.axis_y.setRange(0, 1)
                except Exception:
                    pass

            # Force a repaint of the chart view so UI updates immediately
            try:
                if hasattr(self, 'chart_view') and self.chart_view is not None:
                    self.chart_view.update()
                    try:
                        self.chart_view.repaint()
                    except Exception:
                        pass
            except Exception:
                pass

            # reset start time so plotted timestamps appear relative if needed
            try:
                self.start_time = time.time()
            except Exception:
                pass
        except Exception:
            traceback.print_exc()

    def __init__(self):
        """ Initialize Main Window """
        super(MainWindow, self).__init__()
        # Ensure settings dict exists early so other methods can safely update it
        try:
            self._user_settings = load_user_settings()
            # Get the current device type from settings
            device_str = self._user_settings.get("device_type", Device_Type.FA_5.value)
            self.device_type = Device_Type(device_str)
        except Exception:
            # fallback to DEFAULT_SETTINGS (single source of truth) if loading fails
            self._user_settings = DEFAULT_SETTINGS.copy()

        # In release builds, instantiate and apply the generated UI here so
        # all widget attributes (start_button, end_button, etc.) exist before
        # we try to connect signals or manipulate them.
        if PROGRAM_TYPE_RELEASE:
            try:
                ui = Ui_main_window()
                ui.setupUi(self)
                # Copy attributes from generated ui instance to this MainWindow instance so
                # existing code can access widgets as self.widget_name (e.g. self.start_button).
                try:
                    for name, val in ui.__dict__.items():
                        # Don't overwrite existing attributes on self
                        if not hasattr(self, name):
                            setattr(self, name, val)
                except Exception:
                    logging.exception("Failed to copy UI attributes onto MainWindow instance")
            except Exception:
                logging.exception("Failed to setup UI from Ui_main_window")
        # Ensure the window is shown in both debug and release modes
        try:
            self.show()
        except Exception:
            pass

        if PROGRAM_TYPE_DEBUG:
            ui_path = self.resource_path('main_window.ui')
            uic.loadUi(ui_path, self)
            self.show()  # Show the GUI

        # Make the settings label clickable to change device type
        try:
            if hasattr(self, 'label_com_settings'):
                # show current device type
                self.label_com_settings.setText("Settings " + str(self.device_type.value) + ":")
                self.label_com_settings.installEventFilter(self)
                try:
                    # make it look clickable
                    self.label_com_settings.setCursor(Qt.CursorShape.PointingHandCursor)
                    self.label_com_settings.setToolTip("Click to change device type")
                except Exception:
                    pass
        except Exception:
            logging.exception("Failed to prepare label_com_settings for device selection")

        # Plot configuration: sliding window in seconds (0 = unlimited)
        self.plot_window_seconds = 1000  # show last 1000 seconds; set to 0 to disable

        ports = get_serial_port()
        self.thread = None
        self.worker = None
        self.start_time = 0
        self.get_FA_settings = True

        self.start_button.clicked.connect(self.start_loop)
        self.end_button.clicked.connect(self.stop_loop)
        self.end_button.setEnabled(False)
        self.refresh_button.clicked.connect(self.refresh_port)
        self.resetstats_button.clicked.connect(self.reset_stats)
        self.resethold_button.clicked.connect(self.resethold_stats)
        self.button_command_ch1.clicked.connect(self.channel_change_ch1)
        self.button_command_ch2.clicked.connect(self.channel_change_ch2)
        self.button_command_intosc.clicked.connect(self.channel_change_intosc)
        self.button_command_get_settings.clicked.connect(self.get_settings_from_device)
        self.button_command_precision.clicked.connect(self.toggle_precision_mode)
        self.button_command_lpf.clicked.connect(self.toggle_lpf)
        self.button_command_impedance.clicked.connect(self.toggle_imp50)
        self.button_command_version.clicked.connect(self.get_prod_version_info)
        self.save_txt_button.clicked.connect(self.save_to_csv)

        self.command_edit_1.clicked.connect(self.command1)
        self.command_edit_2.clicked.connect(self.command2)
        self.command_edit_3.clicked.connect(self.command3)
        self.command_edit_4.clicked.connect(self.command4)

        self.saved_command_1.clicked.connect(self.move_command1_to_text)
        self.saved_command_2.clicked.connect(self.move_command2_to_text)
        self.saved_command_3.clicked.connect(self.move_command3_to_text)
        self.saved_command_4.clicked.connect(self.move_command4_to_text)

        self.port_comboBox.addItems(ports)
        self.comboBox_gatetime.clear()
        self.comboBox_gatetime.addItems(
            ["0.1 sec.", "0.2 sec.", "0.5 sec.", "1 sec.", "2 sec.", "5 sec.", "10 sec.", "20 sec."])
        self.button_command_setgatetime.clicked.connect(self.set_gate_time)
        self.clipboardcopy_buton.clicked.connect(self.copy_stats_to_clipboard)

        self.send_data_button.clicked.connect(self.write_data_button)
        # Apply persisted saved commands (use settings loaded earlier into self._user_settings)
        try:
            # apply loaded values (fall back to sensible defaults)
            if hasattr(self, 'saved_command_1'):
                self.saved_command_1.setText(
                    self._user_settings.get("saved_command_1", Commands.GET_GATE_TIME_SETTING.value))
            if hasattr(self, 'saved_command_2'):
                self.saved_command_2.setText(self._user_settings.get("saved_command_2", "Command 2"))
            if hasattr(self, 'saved_command_3'):
                self.saved_command_3.setText(self._user_settings.get("saved_command_3", "Command 3"))
            if hasattr(self, 'saved_command_4'):
                self.saved_command_4.setText(
                    self._user_settings.get("saved_command_4", Commands.SET_GATE_TIME_100MS.value))
            if hasattr(self, 'label_com_settings'): # make mode/device type visible in GUI
                device_type = self._user_settings.get("device_type", Device_Type.FA_5.value)
                self.label_com_settings.setText("Settings " + device_type + ":")


        except Exception:
            logging.exception("Failed to apply user saved command settings")
        self.setup_frequency_plot()

    def update_frequency_plot(self, timestamp, frequency):
        """Update the frequency plot with new data"""
        # If charts aren't available, skip plotting to avoid NameError
        if not CHARTS_AVAILABLE:
            return

        self.freq_series.append(timestamp, frequency)

        # Auto-adjust axes ranges
        points = self.freq_series.points()
        if len(points) > 0:
            x_min = points[0].x()
            x_max = points[-1].x()
            y_values = [p.y() for p in points]
            y_min = min(y_values)
            y_max = max(y_values)

            # Add padding to ranges
            x_range = x_max - x_min if x_max != x_min else 1.0
            y_range = y_max - y_min if y_max != y_min else 1.0

            # smaller padding to reduce empty space
            pad_frac = 0.02

            # Optionally use sliding window to focus on recent data
            if getattr(self, 'plot_window_seconds', 0) and self.plot_window_seconds > 0:
                window = float(self.plot_window_seconds)
                x_start = max(x_min, x_max - window)
                self.axis_x.setRange(x_start, x_max + x_range * pad_frac)
            else:
                self.axis_x.setRange(x_min, x_max + x_range * pad_frac)

            self.axis_y.setRange(y_min - y_range * pad_frac, y_max + y_range * pad_frac)

            # Limit number of points
            if len(points) > 2000:
                self.freq_series.remove(0)

    def resource_path(self, relative_path):
        """ Get the absolute path to the resource, works for dev and PyInstaller """
        # If the application is run as a bundle, the PyInstaller bootloader
        # extracts all files and sets this attribute to the path of the extracted files.
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        else:
            return os.path.join(os.path.abspath("."), relative_path)

    def command1(self):
        """ Open the text input popup to save command for button 1 """
        self.command_edit(1)

    def command2(self):
        """ Open the text input popup to save command for button 2 """
        self.command_edit(2)

    def command3(self):
        """ Open the text input popup to save command for button 3 """
        self.command_edit(3)

    def command4(self):
        """ Open the text input popup to save command for button 4 """
        self.command_edit(4)

    def command_edit(self, button_number):
        """ Open the text input popup to save command """
        # Create a text input popup
        text, ok = QInputDialog.getText(
            self, 'Set your command', 'Please enter the command that you want to save:')
        if ok:
            new_text = str(text)
            try:
                if button_number == 1 and hasattr(self, 'saved_command_1'):
                    self.saved_command_1.setText(new_text)
                    self._user_settings['saved_command_1'] = new_text
                elif button_number == 2 and hasattr(self, 'saved_command_2'):
                    self.saved_command_2.setText(new_text)
                    self._user_settings['saved_command_2'] = new_text
                elif button_number == 3 and hasattr(self, 'saved_command_3'):
                    self.saved_command_3.setText(new_text)
                    self._user_settings['saved_command_3'] = new_text
                elif button_number == 4 and hasattr(self, 'saved_command_4'):
                    self.saved_command_4.setText(new_text)
                    self._user_settings['saved_command_4'] = new_text
                # Persist updated settings to disk
                try:
                    save_user_settings(self._user_settings)
                except Exception:
                    logging.exception("Failed to persist user settings after edit")
            except Exception:
                logging.exception("Error while applying new saved command text")

    def move_command1_to_text(self):
        """ Move the saved command to the text box """
        self.send_to_command_buffer(self.saved_command_1.text())

    def move_command2_to_text(self):
        """ Move the saved command to the text box """
        self.send_to_command_buffer(self.saved_command_2.text())

    def move_command3_to_text(self):
        """ Move the saved command to the text box """
        self.send_to_command_buffer(self.saved_command_3.text())

    def move_command4_to_text(self):
        """ Move the saved command to the text box """
        self.send_to_command_buffer(self.saved_command_4.text())

    def refresh_port(self):
        """ Refresh the serial port list """
        ports = get_serial_port()
        self.port_comboBox.clear()
        self.port_comboBox.addItems(ports)

    def print_message_on_screen(self, text):
        """ Print the message on the screen """
        msg = QMessageBox()
        msg.setWindowTitle("Warning!")
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setText(text)
        msg.exec()

    def channel_change_ch1(self):
        # print('changing channel ch1')
        self.send_to_command_buffer(Commands.SET_CH1_FREQ_AND_POWER)
        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)

    def channel_change_ch2(self):
        # print('changing channel ch2')
        self.send_to_command_buffer(Commands.SET_CH2_FREQ_AND_POWER)
        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)

    def channel_change_intosc(self):
        # print('changing channel int osc')
        self.send_to_command_buffer(Commands.SET_INTERNAL_10M_REF)
        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)

    def get_settings_from_device(self):
        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)
        self.send_to_command_buffer(Commands.GET_GATE_TIME_SETTING)

    def toggle_precision_mode(self):
        if ml.latest_settings.precision:
            self.send_to_command_buffer(Commands.HIGH_PRECISION_OFF)
        else:
            self.send_to_command_buffer(Commands.HIGH_PRECISION_ON)
        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)

    def toggle_lpf(self):
        if ml.latest_settings.lpf:
            self.send_to_command_buffer(Commands.CH1_LPF_OFF)
        else:
            self.send_to_command_buffer(Commands.CH1_LPF_ON)
        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)

    def toggle_imp50(self):
        if ml.latest_settings.imp50:
            self.send_to_command_buffer(Commands.CH1_1M_IMPEDANCE)
        else:
            self.send_to_command_buffer(Commands.CH1_50_IMPEDANCE)
        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)

    def get_prod_version_info(self):
        self.send_to_command_buffer(Commands.GET_PROD_VERSION)
        print('changing channel impedance')

    def set_gate_time(self):
        time_text = self.comboBox_gatetime.currentText()
        command_gate_time = fa5.make_gate_time_command(time_text[0:-5])
        self.send_to_command_buffer(command_gate_time)

    def reset_stats(self):
        if self.resethold_button.isChecked():
            ml.reset(False)
        else:
            ml.reset_on_next_read()
        # Clear the plotted data and measurement displays so the user sees a fresh graph
        try:
            self.clear_frequency_plot()
        except Exception:
            pass

    def resethold_stats(self):
        if self.resethold_button.isChecked():
            ml.reset()
            self.label_time.setStyleSheet("color: blue;")
            try:
                self.clear_frequency_plot()
            except Exception:
                pass
        else:
            self.label_time.setStyleSheet("")

    def establish_serial_communication(self):
        """ Establish serial communication """
        port = self.port_comboBox.currentText()
        baudrate = self.baudrate_comboBox.currentText()
        timeout = 2  # self.timeout_comboBox.currentText()
        length = '8'  # self.len_comboBox.currentText()
        # parity = self.parity_comboBox.currentText()
        stopbits = '1'  # self.bit_comboBox.currentText()
        global SERIAL_CON

        # If the port looks like a socket:// URL or host:port, use serial_for_url
        try:
            if isinstance(port, str) and (port.startswith('socket://') or ':' in port):
                # Normalize host:port to socket://host:port if needed
                url = port
                if not port.startswith('socket://'):
                    url = 'socket://' + port
                # Use serial_for_url to support pyserial socket transport
                try:
                    SERIAL_CON = serial.serial_for_url(url, baudrate=int(baudrate, 10), timeout=float(timeout))
                except Exception:
                    # Fall back to creating a standard Serial (may fail)
                    SERIAL_CON = serial.Serial(port=str(port),
                                               baudrate=int(baudrate, base=10),
                                               timeout=float(timeout),
                                               bytesize=int(length, base=10),
                                               parity=serial.PARITY_NONE,
                                               stopbits=float(stopbits),
                                               xonxoff=False,
                                               rtscts=False,
                                               dsrdtr=False,
                                               )
            else:
                SERIAL_CON = serial.Serial(port=str(port),
                                           baudrate=int(baudrate, base=10),
                                           timeout=float(timeout),
                                           bytesize=int(length, base=10),
                                           parity=serial.PARITY_NONE,
                                           stopbits=float(stopbits),
                                           xonxoff=False,
                                           rtscts=False,
                                           dsrdtr=False,
                                           )
        except Exception:
            # Re-raise to caller so they can show an error
            raise
        # Use the standard pyserial attribute is_open (fallback with getattr for compatibility)
        if not getattr(SERIAL_CON, 'is_open', False):
            try:
                SERIAL_CON.open()
            except Exception:
                pass

    def start_loop(self):
        """ Start the loop """
        self.start_time = time.time()
        # Clear any previous custom stylesheet so the widget uses the platform default look
        self.port_comboBox.setStyleSheet("")
        self.end_button.setEnabled(True)

        # If the serial port is not selected, print a message
        if self.port_comboBox.currentText() == "":
            self.print_message_on_screen("Please select a serial port first!")
            # Set port_comboBox background color to red
            self.port_comboBox.setStyleSheet('background-color: red')
            return

        try:
            # Clear any custom stylesheet to restore the widget's default color
            self.port_comboBox.setStyleSheet("")
            self.establish_serial_communication()
        except SerialException:
            self.print_message_on_screen(
                "Exception occured while trying establish serial communication!")
            return

        global is_serial_port_established
        is_serial_port_established = True
        self.get_FA_settings = True

        try:
            self.worker = Worker()  # a new worker to perform those tasks
            self.thread = QThread()  # a new thread to run our background tasks in
            # move the worker into the thread, do this first before connecting the signals
            self.worker.moveToThread(self.thread)
            # begin our worker object's loop when the thread starts running
            self.thread.started.connect(self.worker.work)
            self.worker.serial_data.connect(self.read_data_from_thread)
            # stop the loop on the stop button click
            self.end_button.clicked.connect(self.stop_loop)
            # tell the thread it's time to stop running
            self.worker.finished.connect(self.thread.quit)
            # have worker mark itself for deletion
            self.worker.finished.connect(self.worker.deleteLater)
            # have thread mark itself for deletion
            self.thread.finished.connect(self.thread.deleteLater)
            self.thread.start()
        except RuntimeError:
            self.print_message_on_screen("Exception in Worker Thread!")

    def stop_loop(self):
        """ Stop the loop """
        self.worker.working = False
        self.on_end_button_clicked()

    def update_gui_settings(self, settings):
        # self.label_freq_channel.setText(settings.channel)
        if settings.channel == "Channel: 1":
            self.button_command_ch1.setStyleSheet("background-color: green; color: white;")
            self.button_command_ch2.setStyleSheet("")
            self.button_command_intosc.setStyleSheet("")
        if settings.channel == "Channel: 2":
            self.button_command_ch1.setStyleSheet("")
            self.button_command_ch2.setStyleSheet("background-color: green; color: white;")
            self.button_command_intosc.setStyleSheet("")
        if settings.channel == "Channel: Internal Clock":
            self.button_command_ch1.setStyleSheet("")
            self.button_command_ch2.setStyleSheet("")
            self.button_command_intosc.setStyleSheet("background-color: green; color: white;")
        if settings.ext_reference_osc:
            self.label_freq_clockref.setText("Ref Clock: External")
        else:
            self.label_freq_clockref.setText("Ref Clock: Internal")
        self.label_gatetime.setText("Gate Time: " + str(settings.gate_time) + " ms")

        if settings.precision:
            self.button_command_precision.setText("Precision is ON")
            self.button_command_precision.setStyleSheet("background-color: green; color: white;")
        else:
            self.button_command_precision.setText("Precision is OFF")
            self.button_command_precision.setStyleSheet("")  # default color

        if settings.lpf:
            self.button_command_lpf.setText("LPF is ON")
            self.button_command_lpf.setStyleSheet("background-color: green; color: white;")
        else:
            self.button_command_lpf.setText("LPF is OFF")
            self.button_command_lpf.setStyleSheet("")  # default color

        if settings.imp50:
            self.button_command_impedance.setText("Zin = 50Ω")
            self.button_command_impedance.setStyleSheet("background-color: green; color: white;")
        else:
            self.button_command_impedance.setText("Zin = 1MΩ")
            self.button_command_impedance.setStyleSheet("")  # default color

    def read_data_from_thread(self, serial_data, meta=None):
        """ Write the result to the text edit box"""
        # self.data_textEdit.append("{}".format(i))
        if "ERROR_SERIAL_EXCEPTION" in serial_data:
            self.print_message_on_screen(
                "Serial Port Exception! Please check the serial port"
                " Possibly it is not connected or the port is not available!")
            self.status_label.setText("NOT CONNECTED!")
            self.status_label.setStyleSheet('color: red')
        else:
            if is_serial_port_established:
                self.baudrate_comboBox.setEnabled(False)
                self.port_comboBox.setEnabled(False)
                self.start_button.setEnabled(False)
                self.status_label.setText("CONNECTED!")
                self.status_label.setStyleSheet('color: green')
                if self.get_FA_settings:
                    if self.device_type == Device_Type.FA_5:
                        self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)
                        self.send_to_command_buffer(Commands.GET_GATE_TIME_SETTING)
                    self.get_FA_settings = False
                if (meta and meta.get("datastream")) or serial_data[-4:-2] == 'OK':
                    self.send_command()  # execute next after receiving OK or 1 line measurement data
                if meta and meta.get("datastream"):
                    # print('info received cont')
                    # continuous  measurements
                    global receive_time_of_freq
                    receive_time_of_freq = time.time()
                    try:
                        frequency, timestamp = ml.latest_value("frequency")
                        if frequency > 0:
                            self.update_frequency_plot(timestamp, frequency)
                            target_frequency, offset, ppm = ml.get_freq_difference()
                        else:
                            target_frequency = Decimal('0.0')
                            offset = 0
                            ppm = 0
                        hours, remainder = divmod(timestamp, 3600)
                        minutes, seconds = divmod(remainder, 60)
                        seconds = int(round(seconds))
                        formatted_time = f"{int(hours)}h{int(minutes)}m{seconds}s"
                        self.label_time.setText("Time: " + formatted_time)
                        timeinterval = ml.get_time_interval("frequency")
                        self.data_textEdit.append(
                            "{}".format(str(round(timestamp, 2)).zfill(8) + ", " + serial_data.strip()))
                        self.label_freq.setText("Freq: " + fa5.group_spaces(frequency).zfill(19) + ' Hz')
                        self.label_freq_avg.setText(
                            "Freq Avg: " + fa5.group_spaces(round(ml.average_value("frequency"), 7)) + ' Hz')
                        self.label_freq_stdev.setText(
                            "Freq STDev: " + str(round(1000 * ml.std_dev_value("frequency"), 4)) + ' mHz')
                        self.label_target_freq.setText("Target Freq: " + target_frequency.to_eng_string() + ' Hz')
                        if offset < 0:  # avoid value to jump around in GUI with and without '-'
                            sign = '- '
                        else:
                            sign = '+'
                        self.label_freq_offset.setText(
                            "Offset: " + sign + fa5.group_spaces(abs(round(offset, 9))).zfill(11) + ' Hz')
                        self.label_freq_count.setText("Count: " + str(ml.count("frequency"))
                                                      + "     Interval: " + str(timeinterval) + " ms")
                        self.label_freq_max.setText(
                            "Max: " + fa5.group_spaces(ml.max_value("frequency")).zfill(19) + ' Hz')
                        self.label_freq_min.setText(
                            "Min: " + fa5.group_spaces(ml.min_value("frequency")).zfill(19) + ' Hz')
                        self.label_freq_pp.setText(
                            "Pk-Pk: " + fa5.group_spaces(1000 * ml.peak_to_peak("frequency")) + ' mHz')
                        self.label_freq_power.setText("Power: " + str(ml.latest_value("power")[0]) + ' dBm')
                        self.label_freq_ppm.setText(
                            "Rel. Offset: " + fa5.group_spaces(round(ppm, 9)) + ' ppm')
                        self.update_gui_settings(ml.latest_settings)
                    except Exception as e:
                        print("error on data:", serial_data)
                        traceback.print_exc()
                else:  # serial_data[-4:-2] == 'OK' and
                    # one time message
                    # print('info received unique')
                    self.data_textEdit_all.append("{}".format(serial_data.strip()))

    def send_to_command_buffer(self, command):
        global waiting_for_command
        command_queue.append(command)
        # if waiting_for_command:  # if no previous command is waiting for answer then send command immediately
        #     self.send_command()

    def send_command(self, command_str=""):
        global waiting_for_command
        """ Send data to serial port """
        if command_str != "":
            command = command_str
            waiting_for_command = False
        elif len(command_queue) > 0:
            command = command_queue.pop(0)  # FIFO
            waiting_for_command = False
        else:
            waiting_for_command = True
            # print('no command to process')
            return

        if not isinstance(command, str):
            command = str(command.value)

        global is_serial_port_established
        if is_serial_port_established:
            if len(command_queue) == 0:
                time.sleep(0.2)
            # print(time.strftime("%Hh%Mm%Ss", time.localtime())," ",time.time(), 'commands', command_queue)
            SERIAL_CON.write(command.encode())
            self.data_textEdit_all.append("{}".format("** Command Send: " + command))
        else:
            self.print_message_on_screen(
                "Serial Port is not established yet! Please establish the serial port first!")

    def save_to_csv(self):
        """ Save the values to the CSV file"""
        if ml.save_to_csv(self.note_textEdit.toPlainText()):
            self.save_txt_button.setStyleSheet("")
        else:
            self.save_txt_button.setStyleSheet("background-color: red; color: white;")

    def on_end_button_clicked(self):
        """ Stop the process """
        global is_serial_port_established
        is_serial_port_established = False
        self.baudrate_comboBox.setEnabled(True)
        self.port_comboBox.setEnabled(True)
        self.end_button.setEnabled(False)
        self.start_button.setEnabled(True)
        self.status_label.setText("Disconnected")
        self.status_label.setStyleSheet('color: red')

    def write_data_button(self):
        """ Send data to serial port """
        global is_serial_port_established
        if is_serial_port_established:
            mytext = self.send_data_text.text()
            self.send_to_command_buffer(mytext)
        else:
            self.print_message_on_screen(
                "Serial Port is not established yet! Please establish the serial port first!")

    def copy_stats_to_clipboard(self):
        # format statistics lineµ
        if ml.copy_stats_to_clipboard(self.note_textEdit.toPlainText()):
            self.clipboardcopy_buton.setStyleSheet("")
        else:
            self.clipboardcopy_buton.setStyleSheet("background-color: red; color: white;")

    def eventFilter(self, source, event):
        """ Install event filter on label_com_settings so user can click it to change Device_Type """
        # Only interested in mouse click events
        # Show the prompt on double-click of the label
        if event.type() == QEvent.Type.MouseButtonPress:
            try:
                # Check if the clicked widget is the label_com_settings
                if source == self.label_com_settings:
                    # Open device type selection dialog
                    self.prompt_device_type_selection()
            except Exception:
                logging.exception("Error in eventFilter")
        return super(MainWindow, self).eventFilter(source, event)

    def prompt_device_type_selection(self):
        """ Prompt the user to select or confirm the device type """
        # Present a simple dropdown to choose the device type and persist selection
        try:
            options = [t.value for t in Device_Type]
            try:
                current_index = options.index(self.device_type.value)
            except Exception:
                current_index = 0

            item, ok = QInputDialog.getItem(self, "Select Device Type",
                                            "Device type:", options, current_index, False)
            if not ok or not item:
                return

            # Map chosen item back to Device_Type enum (Device_Type(item) uses the value)
            try:
                self.device_type = Device_Type(item)
            except Exception:
                # If mapping fails, leave unchanged
                logging.exception("Failed to map selected device type to enum")
                return

            # Update label
            if hasattr(self, 'label_com_settings'):
                self.label_com_settings.setText("Settings " + str(self.device_type.value) + ":")

            # Persist selection
            self._user_settings['device_type'] = self.device_type.value
            try:
                save_user_settings(self._user_settings)
            except Exception:
                logging.exception("Failed to save user settings after device type selection")

            # Mark that we need to re-fetch device settings when next connected
            self.get_FA_settings = True

            # Inform the user
            # try:
            #     self.print_message_on_screen("Device type set to " + self.device_type.value)
            # except Exception:
            #     pass
        except Exception:
            logging.exception("Error while prompting for device type")

def start_ui_design():
    """ Start the UI Design """
    app = QApplication(argv)  # Create an instance
    window_object = MainWindow()  # Create an instance of our class

    app.exec()  # Start the application
