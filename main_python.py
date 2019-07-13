import eel
import time
from analyser import *

print(test_import())
eel.init('web')

c = AnaConnection(excel=True, plot=False)
file_name = c.excel_filename

@eel.expose
def measure_voltages():
    eel.changeProgress(1)
    # c.connect()
    time.sleep(1.5)
    eel.changeProgress(2)
    """
    buffer_size = 1024
    c.sample_voltage(frequency=68000.0, buffer_size=buffer_size, extension_scaling_factor=18)
    c.disconnect()
    """
    time.sleep(1.5)
    eel.changeProgress(3)
    eel.alertSwitch()

@eel.expose
def measure_impedance():
    eel.changeProgress(4)
    """
    c.connect()
    c.imp_mag_phase_cap(steps=100, start=1e2, stop=1e6, reference=1e2)
    c.disconnect()
    """
    time.sleep(1)
    eel.changeProgress(5)
    # make_plots()
    time.sleep(1)
    eel.changeProgress(6)

@eel.expose
def open_excel():
    import os
    os.system('start excel.exe {}'.format(file_name))

eel.start('index.html')