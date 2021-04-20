from windrose import WindroseAxes
from matplotlib import pyplot as plt
import matplotlib.cm as cm
import numpy as np
from windrose import WindroseAxes
from matplotlib import pyplot as plt
import numpy as np


kmh_to_ms = 0.277778
# Decembre 2020
ws = [5.6*kmh_to_ms, 6.8*kmh_to_ms, 6.4*kmh_to_ms, 4.5*kmh_to_ms, 3.5*kmh_to_ms, 5*kmh_to_ms, 5.1*kmh_to_ms]
wd = [45, 45, 67.5, 67.5, 45, 45, 45]

# Février 2021
# ws = [5.60*kmh_to_ms, 2.90*kmh_to_ms, 2.90*kmh_to_ms, 3.40*kmh_to_ms, 5.00*kmh_to_ms, 6.30*kmh_to_ms, 3.70*kmh_to_ms]
# wd = [135, 45, 90, 90, 90, 112.5, 112.5]



fig = plt.figure()
rect = [0.125, 0.125, 0.75, 0.75]
# rect = [0.5, 0.5, 0.5, 0.5]
wa = WindroseAxes(fig, rect)
fig.add_axes(wa)
wa.bar(wd, ws, normed=True, opening=0.8, edgecolor='white')
wa.set_legend(bbox_to_anchor=(1, -0.08), title="WIND SPEED\n(m/s)")

plt.show()
