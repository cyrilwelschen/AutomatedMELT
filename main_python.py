import eel
import time
from analyser import *

eel.init('web')

c = AnaConnection(excel=True, plot=False)

prod = True

@eel.expose
def measure_voltages():
    c.create_wb()
    eel.changeProgress(1)
    if prod:
        successfully_connected = c.connect()
        if not successfully_connected:
            eel.alertConnectionFailed()
            return False
    else:
        time.sleep(1.5)
    eel.changeProgress(2)
    if prod:
        buffer_size = 1024
        c.sample_voltage(frequency=68000.0, buffer_size=buffer_size, extension_scaling_factor=18)
        c.disconnect()
    else:
        time.sleep(1.5)
    eel.changeProgress(3)
    time.sleep(0.5)
    eel.alertSwitch()

@eel.expose
def measure_impedance():
    eel.changeProgress(4)
    if prod:
        c.connect()
        c.imp_mag_phase_cap(steps=100, start=1e2, stop=1e6, reference=1e2)
        c.disconnect()
    else:
        time.sleep(1.5)
    eel.changeProgress(5)
    if prod:
        c.make_plots()
    else:
        time.sleep(1)
    eel.changeProgress(6)

@eel.expose
def open_excel():
    import os
    os.system('start excel.exe {}'.format(c.file_name))

eel.start('main.html')