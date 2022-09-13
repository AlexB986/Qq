
import os, sys
from tkinter import *
from tkinter import ttk
import math, decimal
from docxtpl import DocxTemplate,RichText
import sympy as smp



root = Tk()
root.title('Испытательная лаборатория')
root.geometry('1700x900+100+150')
# root.resizable(True, True)
root.columnconfigure(0,weight=0)
# root.rowconfigure(1,weight=3)
os.chdir(sys.path[0])
doc = DocxTemplate('lq.docx')


# def select_level_conditions_var():
#     level = level_conditions_var.get()
#     if level == 1:
#         print('yes')
#     elif level == 2:
#         print('no')


# level_conditions_var = IntVar()

protocol_label = Label(root,text='Протокол №')
protocol_label.place(x=10,y=10)
protocol_text = Entry(root)
protocol_text.place(x=260,y=10)

position_surname_label = Label(root, text='Должность, фамилия')
position_surname_label.place(x=10,y=40)
position_surname_text = Entry(root)
position_surname_text.place(x=260,y=40)

laboratory_name_label = Label(root,text='Наименование лаборатории')
laboratory_name_label.place(x=10,y=70)
laboratory_name_text = Entry(root)
laboratory_name_text.place(x=260,y=70)

organization_name_label = Label(root,text='Наименование\n проверяемой организации')
organization_name_label.place(x=10,y=100)
organization_name_text = Entry(root)
organization_name_text.place(x=260,y=100)

city_label = Label(root, text='Город')
city_label.place(x=10,y=130)
city_text = Entry(root)
city_text.place(x=260,y=130)

type_label = Label(root, text='Тип (класс) объекта проверки')
type_label.place(x=10,y=160)
type_text = Entry(root)
type_text.place(x=260,y=160)

model_label = Label(root, text='Модель')
model_label.place(x=10,y=190)
model_text = Entry(root)
model_text.place(x=260,y=190)

serial_number_label = Label(root, text='Серийный номер, год выпуска')
serial_number_label.place(x=10,y=220)
serial_number_text = Entry(root)
serial_number_text.place(x=260,y=220)

check_type_label = Label(root, text='Тип проверки')
check_type_label.place(x=10,y=250)
check_type_text = Entry(root)
check_type_text.place(x=260,y=250)

time_data_label = Label(root, text='Дата и время проверки')
time_data_label.place(x=10,y=280)
time_data_text = Entry(root)
time_data_text.place(x=260,y=280)

place_installation_label = Label(root, text='Место установки')
place_installation_label.place(x=10,y=310)
place_installation_text = Entry(root)
place_installation_text.place(x=260,y=310)

equipment_parameters_label = Label(root, text='Параметры работы оборудования')
equipment_parameters_label.place(x=10,y=340)
equipment_parameters_text = Entry(root)
equipment_parameters_text.place(x=260,y=340)

Organization_inn_label = Label(root, text='Организация-заказчик, ИНН')
Organization_inn_label.place(x=10,y=370)
Organization_inn_text = Entry(root)
Organization_inn_text.place(x=260,y=370)

Legal_address_label = Label(root, text='Юридический адрес')
Legal_address_label.place(x=10,y=400)
Legal_address_text = Entry(root)
Legal_address_text.place(x=260,y=400)

contact_person_label = Label(root, text='Контактное лицо')
contact_person_label.place(x=10,y=430)
contact_person_text = Entry(root)
contact_person_text.place(x=260,y=430)

contact_details_label = Label(root, text='Контактные данные')
contact_details_label.place(x=10,y=460)
contact_details_text = Entry(root)
contact_details_text.place(x=260,y=460)

location_object_label = Label(root, text='Адрес местонахождения объекта проверки')
location_object_label.place(x=10,y=490)
location_object_text = Entry(root)
location_object_text.place(x=260,y=490)

air_temperature_label = Label(root, text='Температура воздуха')
air_temperature_label.place(x=410,y=10)
air_temperature_text = Entry(root)
air_temperature_text.place(x=650,y=10)

relative_humidity_label = Label(root, text='Относительная влажность воздуха, %')
relative_humidity_label.place(x=410,y=40)
relative_humidity_text = Entry(root)
relative_humidity_text.place(x=650,y=40)

absolute_pressure_label = Label(root, text='Абсолютное давление, гПа')
absolute_pressure_label.place(x=410,y=70)
absolute_pressure_text = Entry(root)
absolute_pressure_text.place(x=650, y=70)

speed_meter_label = Label(root,text='Измеритель скорости')
speed_meter_label.place(x=410,y=100)
speed_meter=('testo425','testo440/bluetooth','testo440/provod')
speed_meter_combo = ttk.Combobox(root,value=speed_meter)
speed_meter_combo.place(x=650,y=100)


vn1_label = Label(root, text='Vнисх1')
vn1_label.place(x=410,y=130)
vn1_text = Entry(root)
vn1_text.place(x=650, y=130)

vn2_label = Label(root, text='Vнисх2')
vn2_label.place(x=410,y=160)
vn2_text = Entry(root)
vn2_text.place(x=650, y=160)

vn3_label = Label(root, text='Vнисх3')
vn3_label.place(x=410,y=190)
vn3_text = Entry(root)
vn3_text.place(x=650,y=190)

vn4_label = Label(root, text='Vнисх4')
vn4_label.place(x=410,y=220)
vn4_text = Entry(root)
vn4_text.place(x=650,y=220)

vn5_label = Label(root, text='Vнисх5')
vn5_label.place(x=410,y=250)
vn5_text = Entry(root)
vn5_text.place(x=650, y=250)

vn6_label = Label(root, text='Vнисх6')
vn6_label.place(x=410,y=280)
vn6_text = Entry(root)
vn6_text.place(x=650, y=280)

vn7_label = Label(root, text='Vнисх7')
vn7_label.place(x=410,y=310)
vn7_text = Entry(root)
vn7_text.place(x=650,y=310)

vn8_label = Label(root, text='Vнисх8')
vn8_label.place(x=410,y=340)
vn8_text = Entry(root)
vn8_text.place(x=650,y=340)

v1_label = Label(root, text='V1')
v1_label.place(x=410,y=370)
v1_text = Entry(root)
v1_text.place(x=650,y=370)

v2_label = Label(root, text='V2')
v2_label.place(x=410,y=400)
v2_text = Entry(root)
v2_text.place(x=650,y=400)

v3_label = Label(root, text='V3')
v3_label.place(x=410,y=430)
v3_text = Entry(root)
v3_text.place(x=650,y=430)

v4_label = Label(root, text='V4')
v4_label.place(x=410,y=460)
v4_text = Entry(root)
v4_text.place(x=650,y=460)

v5_label = Label(root, text='V5')
v5_label.place(x=410,y=490)
v5_text = Entry(root)
v5_text.place(x=650,y=490)

v6_label = Label(root, text='V6')
v6_label.place(x=410,y=520)
v6_text = Entry(root)
v6_text.place(x=650,y=520)

v7_label = Label(root, text='V7')
v7_label.place(x=410,y=550)
v7_text = Entry(root)
v7_text.place(x=650,y=550)

v8_label = Label(root, text='V8')
v8_label.place(x=410,y=580)
v8_text = Entry(root)
v8_text.place(x=650,y=580)

v9_label = Label(root, text='V9')
v9_label.place(x=410,y=610)
v9_text = Entry(root)
v9_text.place(x=650,y=610)

v10_label = Label(root, text='V10')
v10_label.place(x=410,y=640)
v10_text = Entry(root)
v10_text.place(x=650,y=640)

v11_label = Label(root, text='V11')
v11_label.place(x=410,y=670)
v11_text = Entry(root)
v11_text.place(x=650,y=670)

v12_label = Label(root, text='V12')
v12_label.place(x=410,y=700)
v12_text = Entry(root)
v12_text.place(x=650,y=700)

v13_label = Label(root, text='V13')
v13_label.place(x=410,y=730)
v13_text = Entry(root)
v13_text.place(x=650,y=730)

v14_label = Label(root, text='V14')
v14_label.place(x=410,y=760)
v14_text = Entry(root)
v14_text.place(x=650,y=760)

v15_label = Label(root, text='V15')
v15_label.place(x=410,y=790)
v15_text = Entry(root)
v15_text.place(x=650,y=790)

v16_label = Label(root, text='V16')
v16_label.place(x=410,y=820)
v16_text = Entry(root)
v16_text.place(x=650,y=820)
rt = RichText()
def select_level_conditions_var():
    level = level_conditions_var.get()
    if level == 1:
        rt.add('СООТВЕТСТВУЕТ /', underline=True,size=20,font='Times New Roman')
        rt.add('\n НЕ СООТВЕТСТВУЕТ ', strike=True,size=20,font='Times New Roman')
        rt_embedded = RichText()
        rt_embedded.add(rt)
    elif level == 2:
        rt.add('СООТВЕТСТВУЕТ ', strike=True,size=20,font='Times New Roman' )
        rt.add('/\n НЕ СООТВЕТСТВУЕТ ',underline=True,size=20,font='Times New Roman')
        rt_embedded = RichText()
        rt_embedded.add(rt)

level_conditions_var = IntVar()

conditions_label = Label(root, text='Условия окружающей среды')
conditions_label.place(x=780,y=10)
conditions_y_radio = Radiobutton(root, text='Соответствует', variable=level_conditions_var, value=1,command=select_level_conditions_var)
conditions_y_radio.place(x=960,y=10)
conditions_n_radio = Radiobutton(root, text='Не соответствует', variable=level_conditions_var, value=2,command=select_level_conditions_var)
conditions_n_radio.place(x=1070,y=10)
rt1 = RichText()
def select_level_box_var():
    level = level_box_var.get()
    if level == 3:
        rt1.add('Да/', underline=True,size=24,font='Times New Roman')
        rt1.add('Нет ', strike=True,size=24,font='Times New Roman')
        rt_embedded1 = RichText()
        rt_embedded1.add(rt1)
    elif level == 4:
        rt1.add('Да', strike=True,size=24,font='Times New Roman' )
        rt1.add('/Нет',underline=True,size=24,font='Times New Roman')
        rt_embedded1 = RichText()
        rt_embedded1.add(rt1)

level_box_var = IntVar()
box_label = Label(root, text='Готовность бокса к проверке')
box_label.place(x=780,y=40)
box_y_radio = Radiobutton(root, text='Да', variable=level_box_var, value=3,command=select_level_box_var)
box_y_radio.place(x=960,y=40)
box_n_radio = Radiobutton(root, text='Нет', variable=level_box_var, value=4,command=select_level_box_var)
box_n_radio.place(x=1000,y=40)

supply_filter_class_label = Label(root, text='Класс установленного приточного фильтра ')
supply_filter_class_label.place(x=780,y=70)
supply_filter_class_text = Entry(root)
supply_filter_class_text.place(x=1040,y=70)

number_of_filter_supply_label = Label(root,text='Количество фильтров')
number_of_filter_supply_label.place(x=800,y=100)
number_of_filter_supply=('1','2','3')
number_of_filter_supply_combo = ttk.Combobox(root,value=number_of_filter_supply)
number_of_filter_supply_combo.place(x=1040,y=100)

rt2 = RichText()
def select_leak_box_var():
    level = level_leak_var.get()
    if level == 5:
        rt2.add('☑Да/❎', underline=True,color='# 000000',size=20,font='Times New Roman' )
        rt2.add('Нет', strike=True,color='# 000000',size=20,font='Times New Roman' ) #⮽
        rt_embedded2 = RichText()
        rt_embedded2.add(rt2)
    elif level == 6:
        rt2.add('❎', color='# 000000', size=20, font='Times New Roman')
        rt2.add('Да', strike=True,color='# 000000',size=20,font='Times New Roman' )
        rt2.add('/☑Нет  ',underline=True,color='# 000000',size=20,font='Times New Roman' )
        rt_embedded2 = RichText()
        rt_embedded2.add(rt2)

level_leak_var = IntVar()
leak_label = Label(root, text='Подозрения на возможную утечку')
leak_label.place(x=780,y=130)
leak_y_radio = Radiobutton(root, text='Да', variable=level_leak_var, value=5, command=select_leak_box_var)
leak_y_radio.place(x=1000,y=130)
leak_n_radio = Radiobutton(root, text='Нет ', variable=level_leak_var, value=6, command=select_leak_box_var)
leak_n_radio.place(x=1060,y=130)

exhaust_filter_label = Label(root, text='Класс установленного выпускного фильтра')
exhaust_filter_label.place(x=780,y=160)
exhaust_filter_text = Entry(root)
exhaust_filter_text.place(x=1040,y=160)

number_of_filter_exhaust_label = Label(root,text='Количество фильтров')
number_of_filter_exhaust_label.place(x=780,y=190)
number_of_filter_exhaust=('1','2','3')
number_of_filter_exhaust_combo = ttk.Combobox(root,value=number_of_filter_supply)
number_of_filter_exhaust_combo.place(x=1040,y=190)

place_measurement_label = Label(root, text='Зона или место измерения')
place_measurement_label.place(x=780,y=220)
place_measurement_text = Entry(root)
place_measurement_text.place(x=1040,y=220)

rt3 = RichText()
def select_suspicions_leak_var():
    level = level_suspicions_leak_var.get()
    if level == 7:
        rt3.add('☑ Да/❎', underline=True,color='# 000000',size=20,font='Times New Roman' )
        rt3.add('Нет', strike=True,color='# 000000',size=20,font='Times New Roman' ) #⮽
        rt_embedded3 = RichText()
        rt_embedded3.add(rt3)
    elif level == 8:
        rt3.add('❎', color='# 000000', size=20, font='Times New Roman')
        rt3.add('Да', strike=True,color='# 000000',size=20,font='Times New Roman' )
        rt3.add('/☑Нет  ',underline=True,color='# 000000',size=20,font='Times New Roman' )
        rt_embedded3 = RichText()
        rt_embedded3.add(rt3)

level_suspicions_leak_var = IntVar()

suspicions_leak_label = Label(root, text='Подозрения на возможную утечку')
suspicions_leak_label.place(x=780,y=250)
suspicions_leak_y_radio = Radiobutton(root, text='Да', variable=level_suspicions_leak_var, value=7, command=select_suspicions_leak_var)
suspicions_leak_y_radio.place(x=1000,y=250)
suspicions_leak_n_radio = Radiobutton(root, text='Нет ', variable=level_suspicions_leak_var, value=8, command=select_suspicions_leak_var)
suspicions_leak_n_radio.place(x=1040,y=250)

additions_deviations_label = Label(root, text='Дополнения, отклонения или исключения\nпо проведенным методам испытаний')
additions_deviations_label.place(x=780,y=280)
additions_deviations_text = Entry(root)
additions_deviations_text.place(x=1040,y=280)

executor_label = Label(root, text='Ответственный исполнитель')
executor_label.place(x=780,y=310)

executor_position_label = Label(root, text='Должность')
executor_position_label.place(x=960,y=310)
executor_position_text = Entry(root)
executor_position_text.place(x=1040,y=310)

executor_name_label = Label(root, text='ФИО')
executor_name_label.place(x=1190,y=310)
executor_name_text = Entry(root)
executor_name_text.place(x=1240, y=310)

conclusion_results_label = Label(root, text='Заключение по результатам проверок')
conclusion_results_label.place(x=780,y=340)
conclusion_results_text = Entry(root)
conclusion_results_text.place(x=1040, y=340)

rt4 = RichText()
def select_requirements_var():
    level = level_requirements_var.get()
    if level == 9:
        rt4.add('☑ Подтверждены ❎ ', underline=True,color='# 000000',size=20,font='Times New Roman' )
        rt4.add(' Не подтверждены', strike=True,color='# 000000',size=20,font='Times New Roman' ) #⮽
        rt_embedded4 = RichText()
        rt_embedded4.add(rt4)
    elif level == 10:
        rt4.add('❎ ', color='# 000000', size=20, font='Times New Roman')
        rt4.add('Подтверждены ', strike=True,color='# 000000',size=20,font='Times New Roman' )
        rt4.add('☑ Не подтверждены  ',underline=True,color='# 000000',size=20,font='Times New Roman' )
        rt_embedded4 = RichText()
        rt_embedded4.add(rt4)

level_requirements_var = IntVar()

requirements_label = Label(root, text='Согласно требований СанПиН 3.3686-21')
requirements_label.place(x=780,y=370)
requirements_y_radio = Radiobutton(root, text='Подтверждены', variable=level_requirements_var, value=9,command=select_requirements_var)
requirements_y_radio.place(x=1040,y=370)
requirements_n_radio = Radiobutton(root, text='Не подтверждены ', variable=level_requirements_var, value=10,command=select_requirements_var)
requirements_n_radio.place(x=1190,y=370)

recommended_frequency_label = Label(root,text='Рекомендуемая периодичность проверки\nэксплуатационных характеристик БМБ II класса ')
recommended_frequency_label.place(x=780,y=400)
recommended_frequency_text = Entry(root)
recommended_frequency_text.place(x=1040,y=400)

remarks_label = Label(root, text='Замечания и дальнейшие рекомендации')
remarks_label.place(x=780,y=430)
remarks_text = Entry(root)
remarks_text.place(x=1040,y=430)
##############################################################
###########константы######## δ прибора
instrument1 = 0.12
instrument2 = 0.13
Generator_performance= 1.78*(10**8)   #Производительность генератора, Nквд:
Expanded_Uncertainty_generator= 0.2  #Расш.неопредел. генератора, Nквд:
Sampler_dimensions= [7.8,1.2]        #Габариты пробоотборника, см:
# print(Generator_performance)
supply = ['H14',1,5*(10**-5)]   #Приточный ф-р, класс, шт.
outlet = ['H14',1,5*(10**-5)]   #Выпускной ф-р, класс, шт.
##Приточный фильтр
tscan_supply =135
Nint_supply =560
NAR_supply = None
##Выпускной фильтр
tscan_outlet =75
Nint_outlet = 192
NAR_outlet = None
zone = {
    'working_chamber':{'width':0.905,'lenght':0.580},
    'working_opening':{'width':0.880,'lenght':0.195},
    'outlet_opening' :{'width':0.700,'lenght':0.235}
}

q = zone['working_opening']['lenght']
s1_m2 = float(zone['working_chamber']['width']) * float(zone['working_chamber']['lenght'])
s2_m2= float(zone['working_opening']['width']) * float(zone['working_opening']['lenght'])
s3_m2= round(float(zone['outlet_opening']['width']) * float(zone['outlet_opening']['lenght']),4)


def data():

    context={}

    vn1 = float( vn1_text.get())
    vn1_f = '{0:.2f}'.format(vn1)
    context['vn1'] = vn1_f

    vn2 = float( vn2_text.get())
    vn2_f = '{0:.2f}'.format(vn2)
    context['vn2'] = vn2_f

    vn3 = float( vn3_text.get())
    vn3_f = '{0:.2f}'.format(vn3)
    context['vn3'] = vn3_f

    vn4 = float( vn4_text.get())
    vn4_f = '{0:.2f}'.format(vn4)
    context['vn4'] = vn4_f

    vn5 = float( vn5_text.get())
    vn5_f = '{0:.2f}'.format(vn5)
    context['vn5'] = vn5_f

    vn6 = float( vn6_text.get())
    vn6_f = '{0:.2f}'.format(vn6)
    context['vn6'] = vn6_f

    vn7 = float( vn7_text.get())
    vn7_f = '{0:.2f}'.format(vn7)
    context['vn7'] = vn7_f

    vn8 = float( vn8_text.get())
    vn8_f = '{0:.2f}'.format(vn8)
    context['vn8'] = vn8_f


    v1 = float(v1_text.get())
    v1_f = '{0:.2f}'.format(v1)
    context['v1'] = v1_f

    v2 = float(v2_text.get())
    v2_f = '{0:.2f}'.format(v2)
    context['v2'] = v2_f

    v3 = float(v3_text.get())
    v3_f = '{0:.2f}'.format(v3)
    context['v3'] = v3_f

    v4 = float(v4_text.get())
    v4_f = '{0:.2f}'.format(v4)
    context['v4'] = v4_f

    v5 = float(v5_text.get())
    v5_f = '{0:.2f}'.format(v5)
    context['v5'] = v5_f

    v6 = float(v6_text.get())
    v6_f = '{0:.2f}'.format(v6)
    context['v6'] = v6_f

    v7 = float(v7_text.get())
    v7_f = '{0:.2f}'.format(v7)
    context['v7'] = v7_f

    v8 = float(v8_text.get())
    v8_f = '{0:.2f}'.format(v8)
    context['v8'] = v8_f

    v9 = float(v9_text.get())
    v9_f = '{0:.2f}'.format(v9)
    context['v9'] = v9_f

    v10 = float(v10_text.get())
    v10_f = '{0:.2f}'.format(v10)
    context['v10'] = v10_f

    v11 = float(v11_text.get())
    v11_f = '{0:.2f}'.format(v11)
    context['v11'] = v11_f

    v12 = float(v12_text.get())
    v12_f = '{0:.2f}'.format(v12)
    context['v12'] = v12_f

    v13 = float(v13_text.get())
    v13_f = '{0:.2f}'.format(v13)
    context['v13'] = v13_f

    v14 = float(v14_text.get())
    v14_f = '{0:.2f}'.format(v14)
    context['v14'] = v14_f

    v15 = float(v15_text.get())
    v15_f = '{0:.2f}'.format(v15)
    context['v15'] = v15_f

    v16 = float(v16_text.get())
    v16_f = '{0:.2f}'.format(v16)
    context['v16'] = v16_f

    x_cp = (vn1 + vn2 + vn3 + vn4 + vn5 + vn6 + vn7 + vn8) / 8
    x_cp = round(x_cp, 2)
    x_cp_f = '{0:.2f}'.format(x_cp)
    context['x_cp'] = x_cp_f

    ##U a(Xср.)
    u_a = round((smp.sqrt(vn1 - x_cp) ** 2 + (vn2 - x_cp) ** 2 + (vn3 - x_cp) ** 2 + (vn4 - x_cp) ** 2 + (
                vn5 - x_cp) ** 2 + (vn6 - x_cp) ** 2 + (vn7 - x_cp) ** 2 + (vn8 - x_cp) ** 2) / (8 * 7), 2)

    ##U b(Xср.)
    u_b = round(instrument1 / 2 / 3 ** (0.5), 2)  # smp.sqrt(3),2)
    ##U b(Xср.в)
    u_b2 = round(instrument2 / 2 / smp.sqrt(3), 2)

    ##U сумм
    u = round(smp.sqrt(u_a ** 2 + u_b ** 2), 3)  # !!!!!!!!!!!!!!!2

    ##U расш.
    u_pa = round(u * 2, 2)
    u_pa_f = '{0:.2f}'.format(u_pa)
    context['u_pa'] = u_pa_f

    ##X +20%
    v_otk_up = round(x_cp * 1.2, 2)
    v_otk_up_f = '{0:.2f}'.format(v_otk_up)
    context['v_otb'] = v_otk_up_f

    ##X -20%
    v_otk_down = round(x_cp * 0.8, 2)
    v_otk_down_f = '{0:.2f}'.format(v_otk_down)
    context['v_oth'] = v_otk_down_f

    ## max min
    vn_list = [vn1, vn2, vn3, vn4, vn5, vn6, vn7, vn8]
    #????? vn_list_f = '{0:.2f}'.format(vn_list)????
    context['v_max'] = max(vn_list)
    context['v_min'] = min(vn_list)

    ###Q=
    q1 = round(x_cp * s1_m2, 4)
    q1_f = '{0:.4f}'.format(q1)
    context['q'] = q1_f

    ###U расш. Расход нисходящего воздушного потока, м3\с
    u_enter = round(u_pa * s1_m2, 4)
    u_enter_f = '{0:.4f}'.format(u_enter)
    context['u_pac'] = u_enter_f
    ##U сумм.Расход нисходящего воздушного потока, м3\с
    u_sum = round(u * s1_m2, 4)  # посмотреть позже!!!!!!!!!!!!!!!

    ##################################################Скорость выходящего воздушного потока, м\с

    ##Xср. выход
    x_cp_exit = round((v1 + v2 + v3 + v4 + v5 + v6 + v7 + v8 + v9 + v10 + v11 + v12 + v13 + v14 + v15 + v16) / 16, 2)
    x_cp_exit_f = '{0:.2f}'.format(x_cp_exit)
    context['v_exit'] = x_cp_exit_f
    ##Xср. вход
    x_cp_ent = round(x_cp_exit * float(zone['outlet_opening']['width']) * float(zone['outlet_opening']['lenght']) / (
                float(zone['working_opening']['width']) * float(zone['working_opening']['lenght'])), 2)
    x_cp_ent_f = '{0:.2f}'.format(x_cp_ent)
    context['x_cp_ent'] = x_cp_ent_f
    # U расш вход.
    u_pac_ent = round(u_b2 * 2, 2)
    u_pac_ent_f = '{0:.2f}'.format(u_pac_ent)
    context['u_pac_ent'] = u_pac_ent_f
    ##Q= Расход входящего воздушного потока, м3\с
    q2 = round(x_cp_exit * s3_m2, 4)
    q2_f = '{0:.4f}'.format(q2)
    context['q_ent'] = q2_f

    ##U расш.Расход входящего воздушного потока, м3\с
    u1_pac_ent = round(u_pac_ent * s2_m2, 4)
    u1_pac_ent_f = '{0:.4f}'.format(u1_pac_ent)
    context['u1_pac_ent'] = u1_pac_ent_f
    ##U сумм.Расход входящего воздушного потока, м3\с
    u2_sum = round(u_b2 * s2_m2, 5)  # !!!!!!!!!!!!!!!4

    ########################3колонка
    air_flow = round(q1 + q2, 4)  # Расход воздуха в QКВД, м3/с
    filter_concentration = int(
        round(Generator_performance / (air_flow * 10 ** 6), 0))  # Концентрация до фильтра,Сс, частиц/см3
    # filter_concentration_f = '{0:.2f}'.format(filter_concentration)
    context['filter'] = filter_concentration

    ##u(NКВД)
    u_nkbd = round((Generator_performance * Expanded_Uncertainty_generator / 3 ** (0.5)), 2)

    ##u(QКВД)
    u_qkbd = round(smp.sqrt(u_sum ** 2 + u2_sum ** 2), 4)  # посмотреть позже!!!!!!!!!!!!!!!

    ##u(CC)
    u_cc = int(round(smp.sqrt((1 / (air_flow * 1000000)) ** 2 * u_nkbd ** 2 + (
                Generator_performance / (air_flow ** 2 * 1000000)) ** 2 * u_qkbd ** 2), 0))
    u_cc_pash = u_cc * 2  # u(Сс) расш.
    # u_cc_pash_f = '{0:.2f}'.format(u_cc_pash)
    context['u_cc_pash'] = u_cc_pash

    context['nint_ent'] = Nint_supply  # Приточный фильтр
    context['nint_exi'] = Nint_outlet  # Выпускной фильтр
    ##NPR Приточный фильтр
    npr_ent = int(filter_concentration * supply[2] * 472 * 10)
    ##NPR Выпускной фильтр
    npr_exi = int(filter_concentration * outlet[2] * 472 * 10)
    ###NAR Приточный фильтр
    nar_ent = int(npr_ent - 2 * smp.sqrt(npr_ent))
    # nar_ent_f = '{0:.2f}'.format(nar_ent)
    context['nar_ent'] = nar_ent
    ###NAR Выпускной фильтр
    nar_exi = int(npr_exi - 2 * smp.sqrt(npr_exi))
    # nar_exi_f = '{0:.2f}'.format(nar_exi)
    context['nar_exi'] = nar_exi

    ###Cпосле ф.Приточный фильтр
    c_after = round(float((Nint_supply / 472 / tscan_supply)) / 10 ** -3, 1)
    c_after_f = '{0:.1f}'.format(c_after)
    context['c_after'] = c_after_f

    ##u(Nint).Приточный фильтр
    u_nint = round(float(Nint_supply * 0.2 / 2 / smp.sqrt(3)) / 10, 1)

    ##u(Cпосле ф.)расш.Приточный фильтр
    u_after = round(float((u_nint / 472 / tscan_supply * 2)) / 10 ** -4, 1)
    u_after_f = '{0:.1f}'.format(u_after)
    context['u_after'] = u_after_f

    ###NAR .Выпускной фильтр
    nar_ent1 = int(npr_exi - 2 * smp.sqrt(npr_exi))
    # nar_ent1_f = '{0:.1f}'.format(nar_ent1)
    context['nar_ent1'] = nar_ent1

    ##Cпосле ф.Выпускной фильтр
    c_after1 = round(float((Nint_outlet / 472 / tscan_outlet)) / 10 ** -3, 1)
    c_after1_f = '{0:.1f}'.format(c_after1)
    context['c_after1'] = c_after1_f

    ##u(Nint).Выпускной фильтр
    u_nint1 = round(float(Nint_outlet * 0.2 / 2 / smp.sqrt(3)) / 10, 1)

    ##u(Cпосле ф.)расш.Выпускной фильтр
    u_after1 = round(float((u_nint1 / 472 / tscan_outlet * 2)) / 10 ** -5, 1)
    u_after1_f = '{0:.1f}'.format(u_after1)
    context['u_after1'] = u_after1_f


    rt_embedded = RichText()
    rt_embedded.add(rt)

    rt_embedded1 = RichText()
    rt_embedded1.add(rt1)

    rt_embedded2 = RichText()
    rt_embedded2.add(rt2)

    rt_embedded3 = RichText()
    rt_embedded3.add(rt3)

    rt_embedded4 = RichText()
    rt_embedded4.add(rt4)
    context1 = { 'conditions':rt_embedded,
                 'box':rt_embedded1,
                 'leak':rt_embedded2,
                 'leak1':rt_embedded3,
                 'requirements':rt_embedded4}
    for name,get_data in context.items():
        # get_data = '{0:.2f}'.format(get_data)
        get_data = str(get_data).replace('.',',')
        context1[name] = get_data

    context['colont'] = protocol_text.get()
    context['position'] = position_surname_text.get()
    context['subdivisions'] = laboratory_name_text.get()
    context['organization'] = organization_name_text.get()
    context['city'] = city_text.get()
    context['class'] = type_text.get()
    context['model'] = model_text.get()
    context['serial'] = serial_number_text.get()
    context['examination'] = check_type_text.get()
    context['data'] = time_data_text.get()
    context['time'] = time_data_text.get()
    context['cabinet'] = place_installation_text.get()
    context['options'] = equipment_parameters_text.get()
    context['inn'] = Organization_inn_text.get()
    context['address'] = Legal_address_text.get()
    context['contact'] = contact_person_text.get()
    context['telephone'] = contact_details_text.get()
    context['location_address'] = location_object_text.get()
    context['temperature'] = air_temperature_text.get()
    context['humidity'] = relative_humidity_text.get()
    context['pressure'] = absolute_pressure_text.get()
    context['position_name'] = executor_name_text.get()
    context['position_after'] = executor_position_text.get()
    context['performance_characteristics'] = conclusion_results_text.get()
    context['recommended_checks'] = recommended_frequency_text.get()
    context['recommendations_notes'] = remarks_text.get()
    context['deviations']= additions_deviations_text.get()


    doc.render(context1)
    doc.save('lex_Tem.docx')

data


##############################################################

btn_ok = ttk.Button(root, text='Ok',command=data)
btn_ok.place(x=800, y=530)
btn_cancel = ttk.Button(root, text='Cancel')
btn_cancel.place(x=800, y=630)

root.mainloop()


