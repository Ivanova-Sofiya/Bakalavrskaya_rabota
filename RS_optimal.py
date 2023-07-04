import xlwt
import time 
import math
import pandas as pd
import statistics
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import nnls
import plotly.graph_objects as go
import seaborn as sns
from bokeh.plotting import figure, output_file, show
from bokeh.embed import file_html
from bokeh.resources import CDN
from bokeh.layouts import column

#######################################################################
#Количество данных в каждом наборе данных
#32326  22420   46502

#Данные находятся на разных страницах: Data_1; Data_2; Data_3
DATA = pd.read_excel('data/fractal_series.xlsx', sheet_name = 'Data_1')

############### Считывание данных из excel ###############
times = DATA.iloc[:, 0].tolist()
resistance = DATA.iloc[:, 1].tolist()

len_times = len(times) #длина столбца времени
len_resistance = len(resistance) #длина столбца сопротивления

################ ЗАСЕЧКА ВРЕМЕНИ ################
start = time.time() 
################################################
# Все списки
list_Tau = [] 
list_S = []
list_S_t_t = []
list_R = []
list_S_t_0 = []
list_R_t_t = []
list_RS_full = []
list_AVG_RS = []

list_H_MNK = []
list_H_const = []
list_H_const_MNK = []
list_H_power_dep = [] # Степенная зависимость

list_D_t_0 = []
list_D_power_dep = []
list_D_H_MNK= []

list_T = []
list_a = []

flag = False
################################################
length_1 = 32326
length_2 = 22420
length_3 = 46500
################################################
Tau, step = 1, 100
count = 0
t_t = 101
t_0 = 0
################################################
# Определяем размах для всего временного ряда
sub_srez = resistance[Tau:len(resistance)]
current_avg_1 = np.mean(sub_srez)
list_Deviations_1 = []
for j in range (len(sub_srez)):
    list_Deviations_1. append(np.sum(sub_srez[:j+1] - current_avg_1))
S = np.std(resistance)
list_S_t_t.append(S)
################################################
#Основной цикл, определения H для каждого среза временного ряда
while Tau <= length_1:
    srez = resistance[Tau:Tau+step] #делим временной ряд на отрезки
#########################################################################    
    current_avg = np.mean(srez) #Текущее среднее
    #Список накомпленных отклонений для всего отрезка
    list_Deviations = []
    for t in range(len(srez)):
        list_Deviations.append(np.sum(srez[:t+1] - current_avg))
   #########################################################################
    R = (max(list_Deviations) - min(list_Deviations)) #Размах
    list_R.append(R)

    R_t_t = (max(list_Deviations_1) - min(list_Deviations_1))
    list_R_t_t.append(R_t_t)

    print("")
    S = np.std(srez) #Стандартное отклонение
    list_S.append(S)

    list_RS_full.append(R/S)
    
    list_RS_now = []
    list_RS_now.append(R/S)

    list_AVG_RS.append(np.mean(list_RS_full))

    list_Tau.append(Tau)

    list_step = []
    list_step.append(step)
################################ MNK #####################################

    mx = (np.log(np.array(list_step))).sum()/step
    my = (np.log(np.array(list_RS_now))).sum()/step
    a2 = np.dot(np.log(np.array(list_step)).T, np.log(np.array(list_step)))/step
    a11 = np.dot(np.log(np.array(list_step)).T, np.log(np.array(list_RS_now)))/step

    H = (a11 - mx*my)/(a2 - mx**2)
    a = my - H*mx
    
    a_const = ((list_RS_now[0])**(1/H)) / step

    list_H_MNK.append(H)
    list_a.append(a_const)
 ################################ Hurst #####################################
    alfa = 1 #const для вычисления показатель Hurst через формулу

    #Const
    H_const = ((math.log(list_RS_full[count]))/((math.log(alfa*step)))) 
    #MHK
    H_const_MNK = (math.log(list_RS_full[count]))/((math.log(((a_const)*step)))) 
    # D = 2-H
    if (t_0 == 0):
        H_power_dep = 1
    else:
        H_power_dep = (np.log(((list_R_t_t[0])/(list_S_t_t[0]))/((list_R[count - 1])/(list_S[count - 1])))) / (np.log(len(resistance)-1/t_0)) 

    list_H_const.append(H_const)
    list_H_const_MNK.append(H_const_MNK)
    list_H_power_dep.append(-H_power_dep)

########################## ФРАКТАЛЬНАЯ РАЗМЕРНОСТЬ ##########################
    list_D_power_dep.append(2 - (list_H_power_dep[count]))
    list_D_H_MNK.append(2 - list_H_MNK[count])
    if not flag:
        flag = True
    else:
        if (t_0 == 0):
            list_D_t_0 = 1
        else:
            list_D_t_0.append(2 - H_power_dep)

    print(list_R_t_t[count], "R_t_t")
    print(list_R[count], "list R")
    print(t_t, "t_t")
    print(t_0, "t_0")
        
    list_T.append(times[Tau - 1])
    Tau = Tau + step
    t_t += step
    t_0 += step
    count = count + 1
#############################################################################
print(count, " - число итераций ")

end = time.time() - start
print("Время выполнения программы: ")
print(end)

#################################### ГРАФИКИ ################################
from bokeh.plotting import figure, show
from bokeh.layouts import column
import numpy as np

# Исходные данные
fig1 = figure(title="Исходные данные", x_axis_label='t, с', y_axis_label='dR/R, %', width=1250, height= 600, margin = (10, 50, 50, 50))
fig1.line(times, resistance)

################ ЗАВИСИМОСТЬ LOG(R/S) от LOG(N) ################
fig2 = figure(title="Зависимость log(R/S) от log(N)", x_axis_label='log(N)', y_axis_label='log(R/S(N))', width=1250, height= 600,  margin = (50, 50, 50, 50))
fig2.line(np.log(list_Tau), np.log(list_RS_full))

################ ГРАФИК №1 - ЗАВИСИМОСТЬ LOG(R/S) от времени ################
fig3 = figure(title="Зависимость log(R/S) от врем.ряда", x_axis_label='t, c', y_axis_label='log(R/S(N))', width=1250, height= 600,  margin = (50, 50, 50, 50))
fig3.line(list_T, np.log10(list_RS_full))

################ ХЕРСТ - H ################
fig4 = figure(title="Показатель Херста [H]", x_axis_label='', y_axis_label='H', width=1250, height= 600, margin = (50, 50, 50, 50))
fig4.line(list(range(len(list_H_power_dep))), list_H_power_dep[1:], legend_label='H - R/S (Через формулу D)', color = 'red')
# fig4.line(list(range(len(list_H_const_MNK))),list_H_const_MNK,  legend_label='Н - константа из МНК', color = 'blue' )
# fig4.line(list(range(len(list_H_const))),list_H_const,  legend_label='H - константа', color = 'orange')
fig4.line(list(range(len(list_H_MNK))), list_H_MNK[:-1],  legend_label='H - R/S', color = 'green')
fig4.legend.location = "top_left"

################ ФРАКТАЛЬНАЯ РАЗМЕРНОСТЬ - D ################
fig5 = figure(title="Фрактальная размерность [D]", x_axis_label='', y_axis_label='D', width=1250, height= 600,  margin = (50, 50, 50, 50))
fig5.line(list(range(len(list_D_power_dep))), list_D_power_dep[1:], legend_label='D - R/S (Через формулу D)', color = 'red')
fig5.line(list(range(len(list_D_H_MNK))), list_D_H_MNK[:-1], legend_label='D - R/S (2 - H)', color = 'green')
fig5.legend.location = "top_left"

# объединение графиков в колонки
col1 = column(fig1, fig2)
col2 = column(fig3, fig4, fig5)

# отображение графиков в браузере
show(column(col1, col2))

################ Запись в EXCEL ################
wb = xlwt.Workbook()
ws = wb.add_sheet('Данные')

i = 0
j = 1
for n, x in enumerate(list_H_MNK):
    if n > 0:
        j += 1
        i = 0
    ws.write(j, i, x)
ws.write(0, 0, 'Н через МНК')

i = 3
j = 1
for n, x in enumerate(list_D_power_dep):
    if n > 0:
        j += 1
        i = 3
    ws.write(j, i, x)
ws.write(0, 3, 'H через степен. завис.')

i = 5
j = 1
for n, x in enumerate(list_H_power_dep):
    if n > 0:
        j += 1
        i = 5
    ws.write(j, i, x)
ws.write(0, 5, 'D - степен. завис.')

i = 6
j = 1
for n, x in enumerate(list_D_H_MNK):
    if n > 0:
        j += 1
        i = 6
    ws.write(j, i, x)
ws.write(0, 6, 'D=2-H(MNK)')

i = 7  # Индекс столбца, куда будут записываться значения временного ряда
j = 0  # Индекс строки, с которой начинается запись

ws.write(j, i, 'Значения временного ряда')

# Записываем значения временного ряда для каждого среза
srez = [] 
for n, x in enumerate(resistance):
    if n % step == 0:
        j += 1
        # Записываем значения временного ряда для предыдущего среза
        for k, value in enumerate(srez):
            ws.write(j, i + k, value)
        srez = [] 
    srez.append(x)

# Записываем последний срез после завершения цикла
j += 1
for k, value in enumerate(srez):
    ws.write(j, i + k, value)

wb.save("./data/Segment_part.xls")