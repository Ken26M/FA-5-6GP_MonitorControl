__author__ = "Ken26M, github.com/Ken26M"
__copyright__ = "Copyright 2024, The BG7TBL FA-5-6GP Application Project"
__license__ = "MIT"
__version__ = "2024.08"
__status__ = "Work in Progress"

import time
from decimal import Decimal, InvalidOperation
import statistics
import re
from enum import Enum
import math
import pyperclip as pc
import os


# keep commands easy to get and change if needed
def make_gate_time_command(gatetime):  # gatetime is string in ms
    gatetimestr = str(int(float(gatetime) * 1000)).zfill(5)
    commandstr = "$A" + gatetimestr + "*"
    return commandstr


class Commands(Enum):
    GET_GATE_TIME_SETTING = "$G*"
    GET_FREQ_POWER_SETTINGS = "$D*"
    GET_PROD_VERSION = "$V*"
    RESET = "$R*"
    SET_INTERNAL_10M_REF = "$C0000*"
    SET_CH1_FREQ_AND_POWER = "$C0101*"
    SET_CH2_FREQ_AND_POWER = "$C0202*"
    CH1_50_IMPEDANCE = "$C0303*"
    CH1_1M_IMPEDANCE = "$C0404*"
    CH1_LPF_ON = "$C0505*"
    CH1_LPF_OFF = "$C0606*"
    HIGH_PRECISION_ON = "$C0707*"
    HIGH_PRECISION_OFF = "$C0808*"
    SEND_FREQ_AFTER_MEASURE = "$C0909*"
    SEND_FREQ_AND_POWER = "$C1010*"
    USE_D_FOR_FREQ_POWER = "$C1111*"
    SET_BAUD_RATE_4800_BPS = "$C2020*"
    SET_BAUD_RATE_9600_BPS = "$C2121*"
    SET_BAUD_RATE_19200_BPS = "$C2222*"
    SET_BAUD_RATE_38400_BPS = "$C2323*"
    SET_BAUD_RATE_57600_BPS = "$C2424*"
    SET_BAUD_RATE_115200_BPS = "$C2525*"
    SET_GATE_TIME_100MS = "$A00100*"
    SET_GATE_TIME_1000MS = "$A01000*"
    SET_GATE_TIME_2000MS = "$A02000*"
    SET_GATE_TIME_20000MS = "$A20000*"
    # SET_BAUD_RATE_BPS = "$BXXXXX*"
    # CH2_FREQ_MASK_ABOVE = "$HXXXXX*"
    # CH2_FREQ_MASK_BELOW = "$LXXXXX*"
    EE_SELF_TEST_ON_POWER_UP = "$E2020*"
    EE_CH1_FREQ_ON_POWER_UP = "$E2121*"
    EE_CH2_FREQ_ON_POWER_UP = "$E2222*"
    EE_CH1_50_OHMS_SAVE_EEPROM = "$E3030*"
    EE_CH1_1MOHMS_SAVE_EEPROM = "$E3131*"
    EE_BEEP_ON_SAVE_EEPROM = "$E3232*"
    EE_BEEP_OFF_SAVE_EEPROM = "$E3333*"
    EE_PRECISION_MODE_ON = "$E3434*"
    EE_PRECISION_MODE_OFF = "$E3535*"
    EE_TURN_ON_LPF_SAVE_EEPROM = "$E3636*"
    EE_TURN_OFF_LPF_SAVE_EEPROM = "$E3737*"
    EE_CALIBRATE_CH1_50_OHMS_0_DBM = "$E4040*"
    EE_CALIBRATE_CH1_50_OHMS_MINUS_20_DBM = "$E4141*"
    EE_CALIBRATE_CH1_50_OHMS_5_DBM = "$E4242*"
    EE_CALIBRATE_CH2_0_DBM = "$E4343*"
    EE_CALIBRATE_CH2_MINUS_20_DBM = "$E4444*"
    EE_SAVE_CH1_NOISE_FLOOR = "$E4545*"
    EE_SAVE_CH2_NOISE_FLOOR = "$E4646*"
    EE_AUTO_SEND_FREQ = "$E6060*"
    EE_AUTO_SEND_FREQ_POWER = "$E6161*"
    EE_QUERY_FREQ_POWER_MANUALLY = "$E6262*"

    # note "Set baud rate 115200 BPS" is in reality 12800 BPS current firmware 2024.7 (bug reported)


class Category(Enum):
    POWER = "power"
    FREQUENCY = "frequency"
    UNKNOWN = "unknown"


class FASettings:
    def __init__(self, channel=None, imp50=None, precision=None, ext_reference_osc=None, lpf=None, gate_time=0):
        self.channel = channel
        self.imp50 = imp50
        self.precision = precision
        self.ext_reference_osc = ext_reference_osc
        self.lpf = lpf
        self.gate_time = gate_time

    def update_settings(self, settings):
        if settings.channel is not None:
            self.channel = settings.channel
        if settings.imp50 is not None:
            self.imp50 = settings.imp50
        if settings.precision is not None:
            self.precision = settings.precision
        if settings.ext_reference_osc is not None:
            self.ext_reference_osc = settings.ext_reference_osc
        if settings.lpf is not None:
            self.lpf = settings.lpf
        if settings.gate_time > 0:
            self.gate_time = settings.gate_time

    def __repr__(self):
        return (f"Settings(channel={self.channel}, imp={self.imp50}, precision={self.precision}, "
                f"reference_osc={self.ext_reference_osc}, lpf={self.lpf}, gate_time={self.gate_time})")


def group_spaces(decimal_number: Decimal) -> str:
    # Convert the Decimal object to a string preserving format
    if decimal_number != 0:

        decimal_str = str(decimal_number)

        # Split the number into integer and fractional parts
        integer_part, fractional_part = decimal_str.split('.')

        # Function to group the digits into chunks of three from the left
        def chunk_string(s, size):
            return re.findall(f'.{{1,{size}}}', s)

        # Group the integer part from the right (start at the decimal separator)
        integer_groups = ' '.join(chunk_string(integer_part[::-1], 3))[::-1]

        # Group the fractional part by 3 from the left
        fractional_groups = ' '.join(chunk_string(fractional_part, 3))

        # Combine the formatted parts
    else:
        integer_groups = '0'
        fractional_groups = '0'
    return f"{integer_groups}.{fractional_groups}"


def channel_to_text(channel_char):
    if channel_char == 'A':
        return 'Channel: 1'
    elif channel_char == 'B':
        return 'Channel: 2'
    elif channel_char == 'T':
        return 'Channel: Internal Clock'


def preprocess_string(string):
    # Example preprocessing logic to determine category
    # Remove leading and trailing whitespace
    new_settings = FASettings()
    preprocessed_string = string.strip()

    # Determine if it's a frequency or power
    # print("ID this:", preprocessed_string)
    # $E6161* (get freq and power)
    # $A0010000000.001000001,+00129,
    measurements = []

    if 'OK' in preprocessed_string:  # check if it was a return of command
        if 'POK' in preprocessed_string:
            category = Category.POWER
            preprocessed_cut_number = preprocessed_string[2:-4]
            measurements.append((category, preprocessed_cut_number))
        elif 'AOK' == preprocessed_string[-3:] or 'GOK' == preprocessed_string[-3:]:
            new_settings.gate_time = int(preprocessed_string[-14:-7])
        elif 'DOK' == preprocessed_string[-3:]:  # $APIR ,0009999999.999870000,+00127,DOK
            channel_char = preprocessed_string[1:2]
            precisionmode_char = preprocessed_string[2:3]
            refosc_char = preprocessed_string[3:4]
            impedance_char = preprocessed_string[4:5]
            lpf_char = preprocessed_string[5:6]
            new_settings.channel = channel_to_text(channel_char)
            new_settings.precision = precisionmode_char == 'P'
            new_settings.ext_reference_osc = refosc_char == 'E'
            new_settings.imp50 = impedance_char == 'R'
            new_settings.lpf = lpf_char == 'L'
        elif 'EOK' == preprocessed_string[-3:]:
            if 'LPF' == preprocessed_string[0:3]:  # "LPF ON EOK and LPF OFF EOK"
                new_settings.lpf = "LPF ON EOK" == preprocessed_string
        elif 'COK' == preprocessed_string[-3:]:
            if 'LPF' == preprocessed_string[4:7]:  # "CH1 LPF ON COK and CH1 LPF OFF COK"
                new_settings.lpf = "CH1 LPF ON COK" == preprocessed_string
        else:
            category = Category.UNKNOWN  # This will not be added to frequencies or power
    else:
        if '$' in preprocessed_string:  # continuous data stream
            str_list = preprocessed_string.split(',')
            category = Category.FREQUENCY
            preprocessed_cut_number = str_list[0][2:-3]  # get number but cut of Double error
            measurements.append((category, preprocessed_cut_number))
            if len(str_list) == 3:  # check if power is also present
                category = Category.POWER
                preprocessed_cut_number = str_list[1]
                measurements.append((category, preprocessed_cut_number))
        else:
            category = Category.UNKNOWN  # This will not be added to frequencies or power
    return measurements, new_settings


class MeasureLog:  # 'add_string' is the main function to be used, add the provided string from the FA-5
    def __init__(self):
        self.strings = []
        self.frequencies = []
        self.power = []
        self.start_time = time.time()
        self.reset_on_the_next_read = True
        self.latest_settings = FASettings()

    def add_string(self, single_line_string):
        if self.reset_on_the_next_read:
            self.reset()

        timestamp = time.time() - self.start_time

        # Analyse string content
        measurements, new_settings = preprocess_string(single_line_string)

        # Add to string list
        self.strings.append((single_line_string, timestamp))

        # update settings with acquired data, not always all settings, sometimes only one value.
        self.latest_settings.update_settings(new_settings)

        for measurement in measurements:
            category = measurement[0]
            preprocessed_string = measurement[1]
            # Add to the appropriate list based on the category
            try:
                number = Decimal(preprocessed_string)  # decimal to avoid adding rounding errors
                # make sure to use a string to define the decimal number (a float will add an error into the Decimal)
                # .getcontext: Context(prec=28, rounding=ROUND_HALF_EVEN, Emin=-999999, Emax=999999, capitals=1,
                # clamp=0, flags=[], traps=[InvalidOperation, DivisionByZero, Overflow])
                if category == Category.FREQUENCY:
                    self.frequencies.append((number, timestamp))
                elif category == Category.POWER:
                    self.power.append((number / 10, timestamp))
            except InvalidOperation:
                print('not a number, need to look into why')
                pass  # Do nothing if the string is not a valid Decimal

    def get_freq_difference(self, freq=0, target_freq=0):
        freq = Decimal(freq)
        target_freq = Decimal(target_freq)
        if freq == 0 and len(self.frequencies) > 0:
            frequency = self.average_value("frequency")  # error of average freq
            # frequency = self.frequencies[-1][0]   # error of latest measurement
        else:
            frequency = freq

        if target_freq == 0:
            decimals = -int(math.floor(math.log10(abs(Decimal(frequency) + Decimal('0.1'))))) + 3
            target_frequency = round(frequency, decimals)
        else:
            target_frequency = target_freq
        difference = frequency - target_frequency
        if target_frequency > 0:
            ppm = 1000000 * difference / target_frequency
        else:
            ppm = 0.0

        return target_frequency, difference, ppm

    def get_strings(self):
        return self.strings

    def get_frequencies(self):
        return self.frequencies

    def get_power(self):
        return self.power

    def get_start_time(self):
        return self.start_time

    def reset(self, reset_time=True):
        self.strings = []
        self.frequencies = []
        self.power = []
        if reset_time:
            self.start_time = time.time()
        self.reset_on_the_next_read = False

    def reset_on_next_read(self):
        self.reset_on_the_next_read = True

    def min_value(self, category):
        data = self._get_data(category)
        if data:
            return min(value for value, _ in data)
        return 0.0

    def max_value(self, category):
        data = self._get_data(category)
        if data:
            return max(value for value, _ in data)
        return 0.0

    def average_value(self, category):
        data = self._get_data(category)
        if data:
            values = [value for value, _ in data]
            return sum(values) / len(values)
        return 0.0

    def std_dev_value(self, category):
        data = self._get_data(category)
        if data and len(data) > 1:
            values = [value for value, _ in data]
            return statistics.stdev(values)
        return 0.0

    def peak_to_peak(self, category):
        min_val = self.min_value(category)
        max_val = self.max_value(category)
        if min_val is not None and max_val is not None:
            return max_val - min_val
        return 0.0

    def latest_value(self, category):
        data = self._get_data(category)
        if data:
            return data[-1][0], data[-1][1]
        return 0.0, 0.0

    def get_freq_list(self, category):
        data = self._get_data(category)
        if data:
            values, timestamp = zip(*data)
            return values, timestamp
        return None

    def get_time_interval(self, category):
        data = self._get_data(category)
        interval = 0
        if len(data) > 1:
            interval = data[-1][1] - data[-2][1]
        return int(round(interval, 3) * 1000)  # ms

    def count(self, category):
        data = self._get_data(category)
        if data:
            return len(data)
        return 0

    def _get_data(self, category):
        if category == 'frequency':
            return self.frequencies
        elif category == 'power':
            return self.power
        return []

    def copy_stats_to_clipboard(self, notes=""):
        # Copy to clipboard and write row to log file.
        header = 'Notes,Start Time,Stop Time,Target Freq,Average Freq (Hz),Difference (mHz),STDev (mHz),Gate Time,' \
                 'Pk-Pk (mHz), Power (dBm),Minimum,Maximum,Channel,Count,50 Ohm,ppm'
        avg_frequency = round(self.average_value("frequency"), 7)
        target_frequency, difference, ppm = self.get_freq_difference(avg_frequency)
        stats_text = notes.replace("\n", "; ")[:-2]
        stats_text = stats_text + "," + time.strftime("%d/%m/%y %H:%M:%S", time.localtime(self.start_time))  # start
        stats_text = stats_text + "," + time.strftime("%d/%m/%y %H:%M:%S", time.localtime())  # end
        stats_text = stats_text + "," + target_frequency.to_eng_string()
        stats_text = stats_text + "," + str(avg_frequency)
        stats_text = stats_text + "," + str(1000 * difference)
        stats_text = stats_text + "," + str(round(1000 * self.std_dev_value("frequency"), 4))
        stats_text = stats_text + "," + str(self.latest_settings.gate_time / 1000)
        stats_text = stats_text + "," + str(1000 * self.peak_to_peak("frequency"))
        stats_text = stats_text + "," + str(self.latest_value("power")[0])
        stats_text = stats_text + "," + str(self.min_value("frequency"))
        stats_text = stats_text + "," + str(self.max_value("frequency"))
        stats_text = stats_text + "," + str(self.latest_settings.channel)
        stats_text = stats_text + "," + str(self.count("frequency"))
        stats_text = stats_text + "," + str(self.latest_settings.imp50)
        stats_text = stats_text + "," + str(ppm)

        try:
            pc.copy(stats_text)  # copy to clipboard
            file_path = 'FA5_Statistics_log.csv'
            file_existed_before = os.path.exists(file_path)
            with open(file_path, 'a') as file:
                if not file_existed_before:
                    file.write(header + '\n')
                file.write(stats_text + '\n')
            return True
        except PermissionError:
            print('Write file failed, already open?')
            return False

    def save_to_csv(self, notes=""):
        """ Save the values to the CSV file"""
        try:
            with open('Measured_Frequencies.csv', 'a', encoding='utf-8') as f:
                f.writelines([str(round(i[1], 2)) + ',' + str(i[0]) for i in self.frequencies[0:1]])
                f.writelines(',' + time.strftime("%d/%m/%y %H:%M:%S", time.localtime(self.start_time))
                             + ', N:' + str(self.count('frequency')) + "," + notes.replace("\n", ",")[:-1] + '\n')
                f.writelines([str(round(i[1], 2)) + ',' + str(i[0]) + '\n' for i in self.frequencies[1:]])
                f.close()
            return True
        except PermissionError:
            print('Write file failed, already open?')
            return False

    def __str__(self):
        strings_repr = "\n".join([f"{s} (added at {t:.2f} seconds)" for s, t in self.strings])
        frequencies_repr = "\n".join([f"{n} (added at {t:.2f} seconds)" for n, t in self.frequencies])
        power_repr = "\n".join([f"{n} (added at {t:.2f} seconds)" for n, t in self.power])
        return (f"Strings:\n{strings_repr}\n\n"
                f"Frequencies:\n{frequencies_repr}\n\n"
                f"Power:\n{power_repr}")
