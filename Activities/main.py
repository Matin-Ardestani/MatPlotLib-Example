import matplotlib.pyplot as plt
import openpyxl
from pathlib import Path

directory = str(Path.cwd())

temp = directory.split('\\')
if temp[-1] != 'Activities':
    directory += '\\Activities'

f = ('%s/activity.xlsx' % directory)
print(directory)
ob = openpyxl.load_workbook(f)

sheet = ob.active

lst = []
dates = []
all_time = []
python = []

for row in sheet.iter_rows():
    for cell in row:
        lst.append(cell.value)


counter_date = 1
counter_all = 2
counter_python = 3
result = []
counter = 0

for this in range(70):

    counter_date += 4
    counter_all += 4
    counter_python += 4

    if lst[counter_python] != 0:

        d = (lst[counter_date]).split('/')
        dat = d[0] + '\n' + d[1]

        dic = {
            'date' : dat,
            'all' : lst[counter_all],
            'python' : lst[counter_python]
        }
        result.append(dic)



python_times = []
for this in result:
    python_times.append(this['python'])


all_times = []
for this in result:
    all_times.append(this['all'])


dates = []
for this in result:
    dates.append(this['date'])



plt.fill_between(dates , all_times, label='All Activities' , color='#2a99b0')
plt.plot(dates ,all_times , color="#022b33")

plt.fill_between(dates , python_times, label='Python Activities' , color='#54991f' , alpha=0.7)
plt.plot(dates ,python_times , color="#254a09" , alpha=0.5)

plt.tick_params(labelsize=8)
plt.title('Matin Ardestani Activities')
plt.xticks(dates , dates)
plt.xlabel('Days')
plt.ylabel('Times(minute)')

plt.legend()

plt.show()
