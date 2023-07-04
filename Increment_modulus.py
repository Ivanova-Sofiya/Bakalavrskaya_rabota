import numpy as np
import pandas as pd
import xlwt
import matplotlib.pyplot as plt

def get_differences(data):
    size = len(data)
    list_Step_1 = []
    for i in range(size):
        if i == 0:
            list_Step_1.append(data[0])
        else:
            list_Step_1.append(data[i] - data[i-1])
    return list_Step_1

def get_mean_diff(list_Step_1):
    list_Step_2 = []
    mean = np.mean(list_Step_1)
    list_Step_2.append(mean)
    return mean

def get_AM_values(list_Step_1, mean_diff):
    AM_values = []
    N = len(list_Step_1)

    for m in range(2, int(N/2)):
        k = N // m
        Xk = np.zeros(k)
        for i in range(k):
            start_idx = i * m
            end_idx = start_idx + m 
            Xk[i] = (np.sum(list_Step_1[start_idx:end_idx])) / m
        Nm = N / m
        AM = (1/(Nm)) * np.sum(np.abs(Xk -  mean_diff))

        AM_values.append(AM)
    return AM_values

def plot_AM_values(AM_values):
    m_values = np.arange(1, len(AM_values)+1)
    plt.plot(m_values, AM_values)
    plt.xscale("log")
    plt.yscale("log")
    plt.xlabel("m")
    plt.ylabel("AM")
    #plt.show()

def get_hurst_exponent(AM_values):
    log_AM = np.log(AM_values)
    log_m = np.log(np.arange(1, len(AM_values)+1))
    
    mx = (log_m).sum()/len(log_m)
    my = (log_AM).sum()/len(log_AM)
    a2 = np.dot(log_m.T, log_m)/len(log_m)
    a11 = np.dot(log_m.T, log_AM)/len(log_m)

    H = (a11 - mx*my)/(a2 - mx**2)
    a = my - H*mx
    
    tg = np.tan(H/2)
    return tg + 1

DATA = pd.read_excel('data/fractal_series.xlsx', sheet_name = 'Data_3')
############### Считывание данных из excel ###############
times = DATA.iloc[:, 0].tolist()
resistance = DATA.iloc[:, 1].tolist()
tau = 1
count = 0
step = 100

length_1 = 32326
length_2 = 22400
length_3 = 46500


list_diffs = []
list_mean_diff = []
list_hurs = []
list_df = []


while tau <= length_3:
    srez = resistance[tau:tau+step]

    diffs = get_differences(srez)

    mean_diff = get_mean_diff(diffs)
    list_mean_diff.append(mean_diff)

    AM_values = get_AM_values(diffs, mean_diff)

    plot_AM_values(AM_values)

    hurst_exponent = get_hurst_exponent(AM_values)

    print("Показатель Херста равен:", hurst_exponent)
    
    list_hurs.append(hurst_exponent)
    list_df.append(2-hurst_exponent)

    tau += step
    count += 1

from bokeh.plotting import figure, show
from bokeh.models import ColumnDataSource
from bokeh.layouts import gridplot

x_values = list(range(len(list_hurs)))
x1_values = list(range(len(list_df)))

source = ColumnDataSource(data=dict(x=x_values, y=list_hurs))
source_1 = ColumnDataSource(data=dict(x=x1_values, y=list_df))

p = figure(title="Показатель Херста [H]", x_axis_label="Номер среза", y_axis_label="H", width=1250, height= 600, margin = (50, 50, 50, 50))
p1 = figure(title="Показатель фрактальной размерности [D]", x_axis_label="Номер среза", y_axis_label="D", width=1250, height= 600, margin = (50, 50, 50, 50))

p.line('x', 'y', source=source, line_width=1)
p1.line('x', 'y', source=source_1, line_width=1)

show(gridplot([[p], [p1]]))

wb = xlwt.Workbook()
ws = wb.add_sheet('Данные')

i = 0
j = 1
for n, x in enumerate(list_hurs):
    if n > 0:
        j += 1
        i = 0
    ws.write(j, i, x)
ws.write(0, 0, 'Херст - Метод модулей приращений')

i = 1
j = 1
for n, x in enumerate(list_df):
    if n > 0:
        j += 1
        i = 1
    ws.write(j, i, x)
ws.write(0, 1, 'Фрактальная размерность - Метод модулей приращений')

i = 3 # индекс столбца, куда будут записываться значения временного ряда
j = 0 # индекс строки, с которой начинается запись

ws.write(j, i, 'Значения временного ряда')

# записываем значения временного ряда для каждого среза
srez = []
for n, x in enumerate(resistance):
    if n % step == 0:
        j += 1
        ws.write(j, i, str(srez)) # записываем значения временного ряда для предыдущего среза
        srez = []
    srez.append(x) # добавляем текущее значение временного ряда в список значений для текущего среза

# записываем последний срез после завершения цикла
j += 1
ws.write(j, i, str(srez))

wb.save("data/Increment_modulus.xls")
