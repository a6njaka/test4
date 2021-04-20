import numpy as np
import matplotlib.pyplot as plt

fig3, (ax_3a, ax_3b) = plt.subplots(figsize=(12, 4), nrows=1, ncols=2)

comp_pm25_oms = [41, 7]
comp_pm10_oms = [18, 30]
ingredients = ["Inferieur aux Normes", "Supérieur aux Normes"]

AQI_COLOR = ['#00F600', 'red']

plt.subplot(131)
wedges, texts, autotexts = plt.pie(comp_pm25_oms,
                                   textprops=dict(color="w"),
                                   startangle=90,
                                   autopct=lambda p: '{:.1f}%'.format(round(p)) if p > 0 else '',
                                   radius=1,
                                   colors=AQI_COLOR,
                                   shadow=True)
plt.setp(autotexts, size=10)
for autotext in autotexts:
    autotext.set_color('#000000')
plt.axis('equal')
plt.text(-0.1, -1.25, "PM2.5")
plt.text(0, 1.25, "Comparaison du PM2.5 et PM10 aux Normes de l'OMS", fontsize=14)

plt.subplot(132)
wedges, texts, autotexts = plt.pie(comp_pm10_oms,
                                   textprops=dict(color="w"),
                                   startangle=90,
                                   autopct=lambda p: '{:.1f}%'.format(round(p)) if p > 0 else '',
                                   radius=1,
                                   colors=AQI_COLOR,
                                   shadow=True)

plt.legend(wedges, ingredients,
           title="",
           loc="center left",
           bbox_to_anchor=(1, 0.5, 0, 0))

plt.setp(autotexts, size=10)
for autotext in autotexts:
    autotext.set_color('#000000')
plt.axis('equal')
plt.text(-0.1, -1.25, "PM10")
folder = r"C:\Users\NJAKA\Desktop\Air pollution RM1/"
fig3.savefig(f'{folder}images/comp_oms.png')
fig3.savefig(f'{folder}images/comp_oms.pdf')

plt.show()
