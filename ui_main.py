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

"""
AFCOM - Serial Communication GUI Program
Cannot be used directly, it is a part of main.py
"""
"""
Customized for FA-5 application by Ken26M 2024.08
github.com/Ken26M/FA-5-6GP_MonitorControl
"""

__author__ = 'Mehmet Cagri Aksoy - github.com/mcagriaksoy'
__annotations__ = 'AFCOM - Serial Communication GUI Program'
__version__ = '2024.05'
__license__ = 'MIT'
__status__ = 'Research'


# Runtime Type Checking
PROGRAM_TYPE_DEBUG = True
PROGRAM_TYPE_RELEASE = False

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

# Attempt to import PyQt and PyQtGraph modules
try:
    from PyQt6.QtCore import QObject, QThread, pyqtSignal
    from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QInputDialog
    from PyQt6 import uic
    # import pyqtgraph as pg

    if PROGRAM_TYPE_DEBUG:
        from PyQt6.uic import loadUi
    else:
        from ui_config import Ui_main_window
except ImportError as e:
    missing_module = str(e).split(" ")[-1]
    logging.error(f"Import Error: {missing_module} library is not installed.")
    logging.info(f"Attempting to install the missing module: {missing_module}...")
    try:
        call([sys.executable, "-m", "pip", "install", missing_module])
        logging.info(f"{missing_module} installed successfully. Please restart the program.")
        sys.exit(0)
    except Exception as install_error:
        logging.error(f"Failed to install {missing_module}: {install_error}")
        sys.exit(1)

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
            if port.vid == 1027 and port.pid == 24577:  # Filter out other com devices
                result.append(port.device)
        except SerialException:
            pass
    return result


# MULTI-THREADING
class Worker(QObject):
    """ Worker Thread """
    finished = pyqtSignal()
    serial_data = pyqtSignal(str)

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
                    if buffer.endswith('\r\n'):
                        # Extract the line and reset the buffer
                        line = buffer  # .rstrip('\r\n')
                        # print('line:', line)
                        return line
            # Small delay to prevent busy-waiting
            time.sleep(0.1)
            # print('sleep 0.01')

    def work(self):
        """ Read data from serial port """
        while self.working:
            try:
                line = self.read_line()  # get full text line from FA-5
                ml.add_string(line)  # add measurement to log and update settings if applicable
                self.serial_data.emit(line)
            except SerialException as e:
                print(e)
                # Emit last error message before die!
                self.serial_data.emit("ERROR_SERIAL_EXCEPTION")
                self.working = False
            self.finished.emit()
        SERIAL_CON.close()


def reset_stats():
    ml.reset_on_next_read()


class MainWindow(QMainWindow):
    """ Main Window """

    def __init__(self):
        """ Initialize Main Window """
        super(MainWindow, self).__init__()
        if PROGRAM_TYPE_DEBUG:
            ui_path = self.resource_path('main_window.ui')
            uic.loadUi(ui_path, self)
            self.show()  # Show the GUI

        ports = get_serial_port()
        self.thread = None
        self.worker = None
        self.start_time = 0
        self.get_FA_settings = True

        self.start_button.clicked.connect(self.start_loop)
        self.end_button.clicked.connect(self.stop_loop)
        self.end_button.setEnabled(False)
        self.refresh_button.clicked.connect(self.refresh_port)
        self.resetstats_button.clicked.connect(reset_stats)
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
        self.saved_command_4.setText(Commands.EE_AUTO_SEND_FREQ_POWER.value)

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
            if button_number == 1:
                self.saved_command_1.setText(str(text))
            elif button_number == 2:
                self.saved_command_2.setText(str(text))
            elif button_number == 3:
                self.saved_command_3.setText(str(text))
            elif button_number == 4:
                self.saved_command_4.setText(str(text))

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

    def establish_serial_communication(self):
        """ Establish serial communication """
        port = self.port_comboBox.currentText()
        baudrate = self.baudrate_comboBox.currentText()
        timeout = self.timeout_comboBox.currentText()
        length = '8'  # self.len_comboBox.currentText()
        # parity = self.parity_comboBox.currentText()
        stopbits = '1'  # self.bit_comboBox.currentText()
        global SERIAL_CON
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
        if not SERIAL_CON.isOpen():
            SERIAL_CON.open()

    def start_loop(self):
        """ Start the loop """
        self.start_time = time.time()
        self.port_comboBox.setStyleSheet('background-color: white')
        self.end_button.setEnabled(True)

        # If the serial port is not selected, print a message
        if self.port_comboBox.currentText() == "":
            self.print_message_on_screen("Please select a serial port first!")
            # Set port_comboBox background color to red
            self.port_comboBox.setStyleSheet('background-color: red')
            return

        try:
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
            self.button_command_precision.setStyleSheet("")  # default color

    def read_data_from_thread(self, serial_data):
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
                self.timeout_comboBox.setEnabled(False)
                self.baudrate_comboBox.setEnabled(False)
                self.port_comboBox.setEnabled(False)
                self.start_button.setEnabled(False)
                self.status_label.setText("CONNECTED!")
                self.status_label.setStyleSheet('color: green')
                if self.get_FA_settings:
                    self.send_to_command_buffer(Commands.GET_FREQ_POWER_SETTINGS)
                    self.send_to_command_buffer(Commands.GET_GATE_TIME_SETTING)
                    self.get_FA_settings = False
                if serial_data[0] == "$" or serial_data[-4:-2] == 'OK':
                    self.send_command()  # execute next after receiving OK or 1 line measurement data
                if serial_data[0] == "$":
                    # print('info received cont')
                    # continuous  measurements
                    global receive_time_of_freq
                    receive_time_of_freq = time.time()
                    try:
                        frequency, timestamp = ml.latest_value("frequency")
                        if frequency > 0:
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
                            "{}".format(str(round(timestamp, 1)).zfill(7) + ", " + serial_data.strip()))
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
                            "Rel. Offset: " + fa5.group_spaces(round(ppm, 7)) + ' ppm')
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
        if ml.save_to_csv():
            self.save_txt_button.setStyleSheet("")
        else:
            self.save_txt_button.setStyleSheet("background-color: red; color: white;")

    def on_end_button_clicked(self):
        """ Stop the process """
        global is_serial_port_established
        is_serial_port_established = False
        self.timeout_comboBox.setEnabled(True)
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
        if ml.copy_stats_to_clipboard():
            self.clipboardcopy_buton.setStyleSheet("")
        else:
            self.clipboardcopy_buton.setStyleSheet("background-color: red; color: white;")


def start_ui_design():
    """ Start the UI Design """
    app = QApplication(argv)  # Create an instance
    window_object = MainWindow()  # Create an instance of our class

    if PROGRAM_TYPE_RELEASE:
        ui = Ui_main_window()
        ui.setupUi(window_object)

    app.exec()  # Start the application
