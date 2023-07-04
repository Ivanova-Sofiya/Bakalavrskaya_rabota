import numpy as np
import pandas as pd
import xlwt
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit

DATA = pd.read_excel('./data/fractal_series.xlsx', sheet_name = 'Data_3')

############### Считывание данных из excel ###############
times = DATA.iloc[:, 0].tolist()
resistance = DATA.iloc[:, 1].tolist()


def amplitude_variation(a, b, m, data):
    del_ = (b - a) / m 
    V = 0 
    for i in range(1, m + 1): 
        t1 = (i - 1) * del_ + a 
        t2 = i * del_ + a 
        data_range = data[int(t1):int(t2)] 
        A = np.max(data_range) - np.min(data_range)
        V += A 
    return V 

def linear_func(x, M, C):
    return -M * x + C

window_size = 100
start = 1
fractal_indices = []
fractal_dimensions = []

while start <= len(resistance) - window_size:
    a, b = start, start + window_size - 1 
    m_values = np.arange(1, 100, 1)
    V_values = [amplitude_variation(a, b, m, resistance) for m in m_values]

    log_m = np.log(m_values)
    log_V = np.log(np.array(V_values) + 1e-10)

    M, C = curve_fit(linear_func, log_V, log_m)[0]
    fractal_index = -M
    fractal_indices.append(fractal_index)

    fractal_dimension = 1 + fractal_index
    fractal_dimensions.append(fractal_dimension)

    start += 100

    print(f'Fractal indices: {fractal_index}')
    # print(f'Fractal dimensions: {fractal_dimension}')



################ ФРАКТАЛЬНАЯ РАЗМЕРНОСТЬ - D ################
from bokeh.plotting import figure, show
from bokeh.layouts import column
import numpy as np

fig1 = figure(title="Индекс фракатльной размерности [M]", x_axis_label='', y_axis_label='M', width=1250, height= 600, margin=(50, 50, 50, 50), 
              x_range=(0, len(fractal_indices)), y_range=(min(fractal_indices), max(fractal_indices)),
              tools=['pan','box_zoom','wheel_zoom','reset'])
fig1.line(list(range(len(fractal_indices))), fractal_indices[::1], legend_label='M - метод минимального покрытия', color='red', line_width=2)
fig1.legend.location = "top_left"


fig2 = figure(title="Фрактальная размерность [D]", x_axis_label='', y_axis_label='D', width=1250, height= 600, margin=(50, 50, 50, 50),
              x_range=(0, len(fractal_dimensions)), y_range=(min(fractal_dimensions), max(fractal_dimensions)),
              tools=['pan','box_zoom','wheel_zoom','reset'])
fig2.line(list(range(len(fractal_dimensions))), fractal_dimensions, legend_label='D - метод минимального покрытия', color='green', line_width=2)
fig2.legend.location = "top_left"


col1 = column(fig1, fig2)
show(column(col1))


wb = xlwt.Workbook()
ws = wb.add_sheet('Данные')

# i = 0
# j = 1
# for n, x in enumerate(fractal_indices):
#     if n > 0:
#         j += 1
#         i = 0
#     ws.write(j, i, x)
# ws.write(0, 0, 'Индекс - Метод минимального покрытия')

# i = 1
# j = 1
# for n, x in enumerate(fractal_dimensions):
#     if n > 0:
#         j += 1
#         i = 1
#     ws.write(j, i, x)
# ws.write(0, 1, 'Фрактальная размерность- Метод минимального покрытия')

# wb = xlwt.Workbook()
# ws = wb.add_sheet('Данные')

ws.write(0, 0, 'Временной ряд')
# Записываем заголовок для индекса фрактальности
ws.write(0, 1, 'Индекс - Метод минимального покрытия')
# Записываем заголовок для фрактальной размерности
ws.write(0, 2, 'Фрактальная размерность- Метод минимального покрытия')

# Записываем данные по каждому окну
for n, (fractal_index, fractal_dimension) in enumerate(zip(fractal_indices, fractal_dimensions)):
    a, b = n*window_size, (n+1)*window_size
    # Записываем значение временного ряда для данного окна
    ws.write(n+1, 3, str(resistance[a:b]))
    # Записываем значение индекса фрактальности для данного окна
    ws.write(n+1, 0, fractal_index)
    # Записываем значение фрактальной размерности для данного окна
    ws.write(n+1, 1, fractal_dimension)

wb.save("data/DATA_3-INDEX.xls")
