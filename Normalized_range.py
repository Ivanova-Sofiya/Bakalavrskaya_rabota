import numpy as np
import pandas as pd
import xlwt
import matplotlib.pyplot as plt
from scipy.stats import linregress
from sklearn.impute import KNNImputer

def hurst_exponent(time_series):
    M = len(time_series)
    N = M - 1
    n = 2
    hurst_values = []
    while n <= N // 2:
        A = N // n
        I = np.zeros((A, n))
        for a in range(A):
            I[a] = time_series[a * n : (a + 1) * n]
        e = np.mean(I, axis=1)
        X = np.cumsum(I - e[:, np.newaxis], axis=1)
        R = np.max(X, axis=1) - np.min(X, axis=1)
        S = np.std(I, axis=1)
        R_S = np.mean(R / S)
        hurst_values.append(np.log(R_S) / np.log(n))
        n += 15
    x = np.log(range(2, len(hurst_values) + 2))
    y = hurst_values
    coeffs = np.polyfit(x, y, 1)
    slope = coeffs[0]
    return (slope)

DATA = pd.read_excel('data/fractal_series.xlsx', sheet_name='Data_1')
times = DATA.iloc[:, 0].tolist()
resistance = DATA.iloc[:, 1].tolist()

tau = 1
count = 0
step = 100
length = len(resistance)
list_diffs = []
list_mean_diff = []
list_hurs = []
list_df = []

while tau <= length-1:
    srez = resistance[tau:tau+step]
    hurst_ex = hurst_exponent(srez)

    # проверяем наличие nan значений в результатах
    if np.isnan(hurst_ex):
        # заполняем пропущенные значения на основе ближайших соседей
        imputer = KNNImputer(n_neighbors=2)
        hurst_values = np.array(list_hurs).reshape(-1, 1)
        imputed_values = imputer.fit_transform(hurst_values)
        hurst_ex = imputed_values[-1][0]

    print("Показатель Херста равен:", hurst_ex)
    list_hurs.append(hurst_ex)
    list_df.append(2 - hurst_ex)
    tau += step
    count += 1

from bokeh.plotting import figure, show
from bokeh.models import ColumnDataSource
from bokeh.layouts import gridplot

source1 = ColumnDataSource(data=dict(x=times, y=resistance))

p1 = figure(title="Исходные данные", x_axis_label="t, с", y_axis_label="dR/R, %", width=1000, height= 600, margin = (50, 50, 50, 50))
p1.line('x', 'y', source=source1, line_width=2, color = 'red')

x = np.arange(len(list_hurs))
y = np.array(list_hurs[:-1])
source2 = ColumnDataSource(data=dict(x=x, y=y))

p2 = figure(title="Значение критерия херста [H]", x_axis_label="Номер среза", y_axis_label="H", width=1250, height= 600, margin = (50, 50, 50, 50))
p2.line('x', 'y', source=source2, line_width=2)

x_1 = np.arange(len(list_df))
y_1 = np.array(list_df[:-1])
source3 = ColumnDataSource(data=dict(x=x_1, y=y_1))

p3 = figure(title="Значение фрактальной размерности [D]", x_axis_label="Номер среза", y_axis_label="D", width=1250, height= 600, margin = (50, 50, 50, 50))
p3.line('x', 'y', source=source3, line_width=2)

grid = gridplot([[p1], [p2], [p3]])

show(grid)

wb = xlwt.Workbook()
ws = wb.add_sheet('Данные')

i = 0
j = 1
for n, x in enumerate(list_hurs):
    if n > 0:
        j += 1
        i = 0
    ws.write(j, i, x)
ws.write(0, 0, 'Метод нормированного размаха - H')

i = 1
j = 1
for n, x in enumerate(list_df):
    if n > 0:
        j += 1
        i = 1
    ws.write(j, i, x)
ws.write(0, 1, 'Метод нормированного размаха - D')

i = 3 # индекс столбца, куда будут записываться значения временного ряда
j = 1 # индекс строки, с которой начинается запись

# записываем заголовок для столбца с значениями временного ряда
ws.write(j, i, 'Значения временного ряда')

# записываем значения временного ряда для каждого среза
srez = [] # очищаем список значений временного ряда
for n, x in enumerate(resistance):
    if n % step == 0:
        j += 1
        ws.write(j, i, str(srez)) # записываем значения временного ряда для предыдущего среза
        srez = [] # очищаем список значений временного ряда для нового среза
    srez.append(x) # добавляем текущее значение временного ряда в список значений для текущего среза

# записываем последний срез после завершения цикла
j += 1
ws.write(j, i, str(srez))

wb.save("./data/Normalized_range.xls")