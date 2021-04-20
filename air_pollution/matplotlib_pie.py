import matplotlib.pyplot as plt

# Pie chart, where the slices will be ordered and plotted counter-clockwise:
labels = 'Frogs', 'Hogs', 'Dogs', 'Logs'
sizes = [200, 30, 45, 30]
explode = (0, 0.01, 0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')

fig1, ax1 = plt.subplots()
color_set = ('#00E400', '#FFFF00', '#FF7E00', '#FF0000')
ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.2f%%', shadow=False, startangle=90, colors = color_set)
ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

plt.show()
