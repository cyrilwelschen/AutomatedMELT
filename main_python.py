import eel
import time
from analyser import *

print(test_import())
eel.init('web')

c = AnaConnection(excel=True, plot=False)
file_name = c.excel_filename

@eel.expose
def meas_test():
    print("measuring...")
    for i in range(7):
        eel.changeProgress(i)
        time.sleep(1.5)
    print("measuring done")
    """
    c.connect()
    c.get_voltage()
    c.disconnect()
    c.connect()
    buffer_size = 1024
    c.sample_voltage(frequency=68000.0, buffer_size=buffer_size, extension_scaling_factor=18)
    c.disconnect()
    # ask_continue()
    c.connect()
    c.imp_mag_phase_cap(steps=100, start=1e2, stop=1e6, reference=1e2)
    c.disconnect()
    make_plots()
    """

eel.start('index.html')