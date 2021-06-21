import xlrd
import numpy as np
import math
from math import fabs
import datetime
import xlsxwriter

wb = xlrd.open_workbook('input/Пример3_.xlsx')


def load():
    sheet = wb.sheet_by_index(0)
    _matrix = []
    for row in range(sheet.nrows):
        _row = []
        for col in range(sheet.ncols):
            _row.append(sheet.cell_value(row, col))
        _matrix.append(_row)
    """Считаем количество значений обводненности"""

    #def wc_count():
    #    global countof_wc
    #    countof_wc = 0
    #    i = 2
    #    while (_matrix[i][3] != ''):
    #        countof_wc += 1
    #        i += 1

    #wc_count()
    global countof_wc
    countof_wc = 4

    """Считываем даты"""
    #print(_matrix)
    ###
    ### вход
    ###def dateSearch():
    ###    qD = list()
    ###    for i in range(97):
    ###        qD.append(_matrix[i + 7][1])
    ###
    def dateSearch():
        qD = list()
        for i in range(len(_matrix) - 7):
            qD.append(_matrix[i + 7][0])
        ##i = 2
        ##while (i < len(_matrix)):
        ##    qD.append(_matrix[i][0])
        ##    i += 1
        #print(qD)

        for i in range(len(qD)):
            qD[i] = int(qD[i])
            qD[i] = datetime.datetime.fromordinal(qD[i] + 693594)
            qD[i] = datetime.datetime.date(qD[i])

        tupeYear = []
        tupeMonth = []
        tupeDay = []

        global givenDate
        givenDate = []

        for i in range(len(qD)):
            tupeYear.append(int(qD[i].year))
            tupeMonth.append(int(qD[i].month))
            tupeDay.append(int(qD[i].day))
            givenDate.append([tupeDay[i], tupeMonth[i], tupeYear[i]])
            givenDate[i] = ('.'.join(map(str, givenDate[i])))
            givenDate[i] = givenDate[i][:0] + '0' + givenDate[i][0:]
            if givenDate[i][4] == '.':
                givenDate[i] = givenDate[i][:3] + '0' + givenDate[i][3:]

    dateSearch()
    global q_oil
    print(_matrix)
    q_oil = []
    ##for i in range(len(_matrix) - 2):
    for i in range(len(_matrix) - 7):
        t = _matrix[i + 7][5]
        ##t = _matrix[i + 2][1]
        ##t = t.replace(',', '.')
        if t != '':
            t = float(t)
        else:
            t = 1111

        q_oil.append(t)
        q_oil[i] = float(q_oil[i])
    ###
    ###while q_oil[-1] == 1111:
    ###    q_oil.pop()
    ###
    q_oil.pop(0)
    ###for i in range(4):
    ###    q_oil.pop(0)
    q_oil = np.array(q_oil)

    global q_liq
    q_liq = []
    for i in range(len(_matrix) - 7):
    ##for i in range(len(_matrix) - 2):
        ##t = _matrix[i + 2][2]
        t = _matrix[i + 7][6]
        ##t = t.replace(',', '.')

        if t != '':
            t = float(t)
        else:
            t = 1111
        q_liq.append(t)
        ##q_liq[i] = float(q_liq[i])

    nulls_count = 0
    ###
    ###while q_liq[-1] == 1111:
    ###    q_liq.pop()
    ###    nulls_count += 1
    ###
    q_liq.pop(0)
    ###for i in range(4):
    ###    q_liq.pop(0)
    q_liq = np.array(q_liq)

    givenDate.pop(0)
    ####for i in range(4):
    ####    givenDate.pop(0)
    global wc
    #wc = []
    #for i in range(countof_wc):
    #    t = _matrix[i + 2][3]
    #    wc.append(t)
    wc = [0.98, 0.99, 0.995, 0.999]


def franc(x, y, wc):

    mat = np.vstack([x, np.ones(len(x))]).T
    a, b = np.linalg.lstsq(mat, y, rcond=None)[0]

    comp_q_liq = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq = np.append(comp_q_liq, a * (q_oil[i] ** 2) + b * q_oil[i])

    graph1 = np.array([])
    for i in range(len(q_oil)):
        graph1 = np.append(graph1, q_liq[i] / q_oil[i])

    graph2 = np.array([])
    for i in range(len(q_oil)):
        graph2 = np.append(graph2, a * q_oil[i] + b)

    def act_forec(wc, a, b):
        """Нахождение активных извлекаемых запасов нефти """
        f_oil = 1 - wc
        active = (1 / (2 * a * f_oil)) - (b / (2 * a))

        """Нахождение прогнозной накопленной добычи жидкости """
        forec_liq = a * (active ** 2) + b * active

        """Нахождение прогнозной накопленной добычи нефти """
        forec_oil = forec_liq - (a * (active ** 2) + (b - 1) * active)
        return (active, forec_liq, forec_oil)


    active, forec_liq, forec_oil = act_forec(wc, a, b)
    return (a, b, active, forec_liq, forec_oil, comp_q_liq, graph1, graph2)

def sippas(x, y, wc):

    mat = np.vstack([x, np.ones(len(x))]).T
    a, b = np.linalg.lstsq(mat, y, rcond=None)[0]

    comp_q_liq = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq = np.append(comp_q_liq, (b * q_oil[i]) / (1 - a * q_oil[i]))

    graph1 = np.array([])
    for i in range(len(q_oil)):
        graph1 = np.append(graph1, q_liq[i] / q_oil[i])

    graph2 = np.array([])
    for i in range(len(q_oil)):
        graph2 = np.append(graph2, a * q_liq[i] + b)

    def act_forec(wc, a, b):
        """Нахождение активных извлекаемых запасов нефти """
        f_oil = 1 - wc
        active = (1 / a) - math.sqrt((b / (a ** 2)) * f_oil)

        """Нахождение прогнозной накопленной добычи жидкости """
        forec_liq = ((b * active) / (1 - a * active))


        """Нахождение прогнозной накопленной добычи нефти """
        forec_oil = forec_liq - active
        return (active, forec_liq, forec_oil)

    active, forec_liq, forec_oil = act_forec(wc, a, b)
    return (a, b, active, forec_liq, forec_oil, comp_q_liq, graph1, graph2)


def abyzbaev_main(x, y, wc):

    mat = np.vstack([x, np.ones(len(x))]).T
    a, b = np.linalg.lstsq(mat, y, rcond=None)[0]

    comp_q_liq = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq = np.append(comp_q_liq, np.exp(b) * (q_oil[i] ** a))

    graph1 = np.array([])
    for i in range(len(q_oil)):
        graph1 = np.append(graph1, np.log(q_liq[i]))

    graph2 = np.array([])
    for i in range(len(q_oil)):
        graph2 = np.append(graph2, a * np.log(q_oil[i]) + b)

    def act_forec(wc, a, b):
        """Нахождение активных извлекаемых запасов нефти """
        f_oil = 1 - wc
        active = (1 / (a * np.exp(b) * f_oil)) ** (1 / (a - 1))

        """Нахождение прогнозной накопленной добычи жидкости """
        forec_liq = np.exp(b) * (active ** a)

        """Нахождение прогнозной накопленной добычи нефти """
        temp = forec_liq - (np.exp(b) * (active ** a))
        forec_oil = forec_liq - temp
        return (active, forec_liq, forec_oil)

    active, forec_liq, forec_oil = act_forec(wc, a, b)
    return (a, b, active, forec_liq, forec_oil, comp_q_liq, graph1, graph2)

def abyzbaev_mod1(x, y, wc):

    mat = np.vstack([x, np.ones(len(x))]).T
    a, b = np.linalg.lstsq(mat, y, rcond=None)[0]

    comp_q_liq = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq = np.append(comp_q_liq, np.exp(b) * (q_oil[i] ** (a + 1)))

    graph1 = np.array([])
    for i in range(len(q_oil)):
        graph1 = np.append(graph1, np.log(q_liq[i] / q_oil[i]))

    graph2 = np.array([])
    for i in range(len(q_oil)):
        graph2 = np.append(graph2, a * np.log(q_oil[i]) + b)

    def act_forec(wc, a, b):
        """Нахождение активных извлекаемых запасов нефти """
        f_oil = 1 - wc
        active = (1 / ((a + 1) * np.exp(b) * f_oil)) ** (1 / a)

        """Нахождение прогнозной накопленной добычи жидкости """
        forec_liq = np.exp(b) * (active ** (a + 1))

        """Нахождение прогнозной накопленной добычи нефти """
        forec_oil = np.array([])
        temp = np.exp(b) * (active ** (a + 1))
        temp = temp - active
        forec_oil = forec_liq - temp
        return (active, forec_liq, forec_oil)

    active, forec_liq, forec_oil = act_forec(wc, a, b)
    return (a, b, active, forec_liq, forec_oil, comp_q_liq, graph1, graph2)

def abyzbaev_mod2(x, y, wc):

    mat = np.vstack([x, np.ones(len(x))]).T
    a, b = np.linalg.lstsq(mat, y, rcond=None)[0]

    comp_q_liq = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq = np.append(comp_q_liq, np.exp(b / (1 - a)) * (q_oil[i] ** (1 / (1 - a))))

    graph1 = np.array([])
    for i in range(len(q_oil)):
        graph1 = np.append(graph1, np.log(q_liq[i] / q_oil[i]))

    graph2 = np.array([])
    for i in range(len(q_oil)):
        graph2 = np.append(graph2, a * np.log(q_liq[i]) + b)

    def act_forec(wc, a, b):
        """Нахождение активных извлекаемых запасов нефти """
        f_oil = 1 - wc
        active = ((1 - a) / ((np.exp(b / (1 - a))) * f_oil)) ** (1 - a)

        """Нахождение прогнозной накопленной добычи жидкости """
        forec_liq = np.exp(b / (1 - a)) * (active ** (1 / (1 - a)))

        """Нахождение прогнозной накопленной добычи нефти """
        temp = np.exp(b / (1 - a)) * (active ** (1 / (1 - a)))
        temp = temp - active
        forec_oil = forec_liq - temp
        return (active, forec_liq, forec_oil)

    active, forec_liq, forec_oil = act_forec(wc, a, b)
    return (a, b, active, forec_liq, forec_oil, comp_q_liq, graph1, graph2)

def govor(x, y, wc):

    mat = np.vstack([x, np.ones(len(x))]).T
    a, b = np.linalg.lstsq(mat, y, rcond=None)[0]

    x2 = np.array(np.exp(a * np.log(q_oil) + b) + q_oil)
    y2 = np.array(q_liq)

    mat2 = np.vstack([x2, np.ones(len(x))]).T
    a2, b2 = np.linalg.lstsq(mat, y2, rcond=None)[0]

    comp_q_liq2 = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq2 = np.append(comp_q_liq2, np.exp(a * np.log(q_oil[i]) + b) + q_oil[i])

    comp_q_liq = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq = np.append(comp_q_liq, a * np.log(q_oil[i]) + b)

    graph1 = np.array([])
    for i in range(len(q_oil)):
        graph1 = np.append(graph1, np.log(q_liq[i] - q_oil[i]))

    graph2 = np.array([])
    for i in range(len(q_oil)):
        graph2 = np.append(graph2, a * np.log(q_oil[i]) + b)

    def act_forec(wc, a, b):
        """Нахождение активных извлекаемых запасов нефти """
        f_oil = 1 - wc
        active = ((1 - f_oil) / (a * np.exp(b) * f_oil)) ** (1 / (a - 1))

        """Нахождение прогнозной накопленной добычи жидкости """
        forec_liq = np.exp(b) * (active ** a)

        """Нахождение прогнозной накопленной добычи нефти """
        forec_oil = forec_liq - (np.exp(b) * (active ** a))
        return (active, forec_liq, forec_oil)

    active, forec_liq, forec_oil = act_forec(wc, a, b)
    return (a, b, active, forec_liq, forec_oil, comp_q_liq, comp_q_liq2, graph1, graph2)


def nazarov_sipachev(x, y, wc):

    mat = np.vstack([x, np.ones(len(x))]).T
    a, b = np.linalg.lstsq(mat, y, rcond=None)[0]

    comp_q_liq = np.array([])
    for i in range(len(q_oil)):
        comp_q_liq = np.append(comp_q_liq, a * (q_liq[i] - q_oil[i]) * q_oil[i] + b * q_oil[i])

    graph1 = np.array([])
    for i in range(len(q_oil)):
        graph1 = np.append(graph1, q_liq[i] / q_oil[i])

    graph2 = np.array([])
    for i in range(len(q_oil)):
        graph2 = np.append(graph2, a * (q_liq[i] - q_oil[i]) + b)

    def act_forec(wc, a, b):
        """Нахождение активных извлекаемых запасов нефти """
        f_oil = 1 - wc
        active = (1 / a) * (1 - math.sqrt((b * f_oil - f_oil) / (1 - f_oil)))

        """Нахождение прогнозной накопленной добычи жидкости """
        forec_liq = ((b * active - a * (active ** 2)) / (1 - a * active))

        """Нахождение прогнозной накопленной добычи нефти """
        forec_oil = forec_liq - (((b - 1) * active) / (1 - a * active))
        return (active, forec_liq, forec_oil)

    active, forec_liq, forec_oil = act_forec(wc, a, b)
    return (a, b, active, forec_liq, forec_oil, comp_q_liq, graph1, graph2)

def check_qual(forec_liq):
    n = len(forec_liq)
    mape = 0

    for i in range(n):
        mape = mape + fabs((q_liq[i] - forec_liq[i]) / q_liq[i])
    mape = (1 / n) * mape

    return mape


def output():
    global mapeCnts
    mapeCnts = np.array([])
    global mapeNames
    mapeNames = np.array([])
    needed_columns = 44

    workbook = xlsxwriter.Workbook('Выходные данные.xlsx')
    worksheet = workbook.add_worksheet('Выходные данные')
    bold = workbook.add_format({'bold': 1})
    worksheet = workbook.add_worksheet('Активные запасы')
    worksheet = workbook.get_worksheet_by_name('Выходные данные')

    """ Вывод """
    for i in range(countof_wc):
        worksheet.write((needed_columns * i), 0, 'Обводненность', bold)
        worksheet.write(1 + (needed_columns * i), 0, '-=-=-=-=-=-=-=-=-=-=-', bold)
        worksheet.write(2 + (needed_columns * i), 0, wc[i], bold)
        for j in range(needed_columns - 4):
            worksheet.write(j + 2 + (needed_columns * i), 1, '             ||', bold)

        colum = 0

        '''Французский Институт'''
        name = 'Французский Институт'
        x = np.array(q_oil)
        y = np.array(q_liq / q_oil)
        a, b, active, fore_liq, fore_oil, comp_q_liq, graph1, graph2 = franc(x, y, wc[i])
        print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active,
                     fore_liq, fore_oil, comp_q_liq)


        worksheet = workbook.get_worksheet_by_name('Активные запасы')

        for j in range(len(graph1)):
            worksheet.write(35 + j, 3 + colum, graph1[j])

        for j in range(len(graph2)):
            worksheet.write(35 + j, 5 + colum, graph2[j])

        worksheet.write(0 + i * 5, 0, wc[i])

        worksheet.write(2 + i * 5, 3 + colum, 'Французский институт')
        worksheet.write(3 + i * 5, 3 + colum, 'Активные запасы')
        worksheet.write(3 + i * 5, 4 + colum, active, bold)

        worksheet = workbook.get_worksheet_by_name('Выходные данные')

        if (len(mapeCnts) < 7):
            mapeCnts = np.append(mapeCnts, check_qual(comp_q_liq))
            mapeNames = np.append(mapeNames, name)
        colum += 6
        """####################################################################################"""

        '''Cипачев - Посевич'''

        name = 'Cипачев - Посевич'
        x = np.array(q_liq)
        y = np.array(q_liq / q_oil)
        a, b, active, fore_liq, fore_oil, comp_q_liq, graph1, graph2 = sippas(x, y, wc[i])

        print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active,
                     fore_liq, fore_oil, comp_q_liq)
        if (len(mapeCnts) < 7):
            mapeCnts = np.append(mapeCnts, check_qual(comp_q_liq))
            mapeNames = np.append(mapeNames, name)

        worksheet = workbook.get_worksheet_by_name('Активные запасы')

        for j in range(len(graph1)):
            worksheet.write(35 + j, 3 + colum, graph1[j])

        for j in range(len(graph2)):
            worksheet.write(35 + j, 5 + colum, graph2[j])

        worksheet.write(2 + i * 5, 3 + colum, 'Cипачев - Посевич')
        worksheet.write(3 + i * 5, 3 + colum, 'Активные запасы')
        worksheet.write(3 + i * 5, 4 + colum, active, bold)

        worksheet = workbook.get_worksheet_by_name('Выходные данные')


        colum += 6

        """####################################################################################"""

        ''' Назаров - Сипачев '''

        name = 'Назаров - Сипачев'
        x = np.array(q_liq - q_oil)
        y = np.array(q_liq / q_oil)
        a, b, active, fore_liq, fore_oil, comp_q_liq, graph1, graph2 = nazarov_sipachev(x, y, wc[i])

        print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active,
                     fore_liq, fore_oil, comp_q_liq)
        if (len(mapeCnts) < 7):
            mapeCnts = np.append(mapeCnts, check_qual(comp_q_liq))
            mapeNames = np.append(mapeNames, name)

        worksheet = workbook.get_worksheet_by_name('Активные запасы')

        for j in range(len(graph1)):
            worksheet.write(35 + j, 3 + colum, graph1[j])

        for j in range(len(graph2)):
            worksheet.write(35 + j, 5 + colum, graph2[j])

        worksheet.write(2 + i * 5, 3 + colum, 'Назаров - Сипачев')
        worksheet.write(3 + i * 5, 3 + colum, 'Активные запасы')
        worksheet.write(3 + i * 5, 4 + colum, active, bold)

        worksheet = workbook.get_worksheet_by_name('Выходные данные')


        colum += 6

        """####################################################################################"""

        '''Стандартный метод Абызбаева'''

        name = 'Стандартный метод Абызбаева'
        x = np.array(np.log(q_oil))
        y = np.array(np.log(q_liq))
        a, b, active, fore_liq, fore_oil, comp_q_liq, graph1, graph2 = abyzbaev_main(x, y, wc[i])

        print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active,
                     fore_liq, fore_oil, comp_q_liq)
        if (len(mapeCnts) < 7):
            mapeCnts = np.append(mapeCnts, check_qual(comp_q_liq))
            mapeNames = np.append(mapeNames, name)

        worksheet = workbook.get_worksheet_by_name('Активные запасы')

        for j in range(len(graph1)):
            worksheet.write(35 + j, 3 + colum, graph1[j])

        for j in range(len(graph2)):
            worksheet.write(35 + j, 5 + colum, graph2[j])

        worksheet.write(2 + i * 5, 3 + colum, 'Стандартный метод Абызбаева')
        worksheet.write(3 + i * 5, 3 + colum, 'Активные запасы')
        worksheet.write(3 + i * 5, 4 + colum, active, bold)

        worksheet = workbook.get_worksheet_by_name('Выходные данные')


        colum += 6

        """####################################################################################"""

        '''Первый мод. метод Абызбаева'''

        name = 'Первый мод. метод Абызбаева'
        x = np.array(np.log(q_oil))
        y = np.array(np.log(q_liq / q_oil))
        a, b, active, fore_liq, fore_oil, comp_q_liq, graph1, graph2 = abyzbaev_mod1(x, y, wc[i])

        print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active,
                     fore_liq, fore_oil, comp_q_liq)
        if (len(mapeCnts) < 7):
            mapeCnts = np.append(mapeCnts, check_qual(comp_q_liq))
            mapeNames = np.append(mapeNames, name)

        worksheet = workbook.get_worksheet_by_name('Активные запасы')

        for j in range(len(graph1)):
            worksheet.write(35 + j, 3 + colum, graph1[j])

        for j in range(len(graph2)):
            worksheet.write(35 + j, 5 + colum, graph2[j])

        worksheet.write(2 + i * 5, 3 + colum, 'Первый мод. метод Абызбаева')
        worksheet.write(3 + i * 5, 3 + colum, 'Активные запасы')
        worksheet.write(3 + i * 5, 4 + colum, active, bold)

        worksheet = workbook.get_worksheet_by_name('Выходные данные')

        colum += 6

        """####################################################################################"""

        '''Второй мод. метод Абызбаева'''

        name = 'Второй мод. метод Абызбаева'
        x = np.array(np.log(q_liq))
        y = np.array(np.log(q_liq / q_oil))
        a, b, active, fore_liq, fore_oil, comp_q_liq, graph1, graph2 = abyzbaev_mod2(x, y, wc[i])

        print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active,
                     fore_liq, fore_oil, comp_q_liq)
        if (len(mapeCnts) < 7):
            mapeCnts = np.append(mapeCnts, check_qual(comp_q_liq))
            mapeNames = np.append(mapeNames, name)

        worksheet = workbook.get_worksheet_by_name('Активные запасы')

        for j in range(len(graph1)):
            worksheet.write(35 + j, 3 + colum, graph1[j])

        for j in range(len(graph2)):
            worksheet.write(35 + j, 5 + colum, graph2[j])

        worksheet.write(2 + i * 5, 3 + colum, 'Второй мод. метод Абызбаева')
        worksheet.write(3 + i * 5, 3 + colum, 'Активные запасы')
        worksheet.write(3 + i * 5, 4 + colum, active, bold)

        worksheet = workbook.get_worksheet_by_name('Выходные данные')

        colum += 6

        """####################################################################################"""

        '''Метод Говоровой - Рябининой'''

        name = 'Метод Говоровой - Рябининой'
        x = np.array(np.log(q_oil))
        y = np.array(np.log(q_liq - q_oil))
        a, b, active, fore_liq, fore_oil, comp_q_liq, comp_q_liq2, graph1, graph2 = govor(x, y, wc[i])

        print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active,
                     fore_liq, fore_oil,
                     comp_q_liq, 1, comp_q_liq2)
        if (len(mapeCnts) < 7):
            mapeCnts = np.append(mapeCnts, check_qual(comp_q_liq2))
            mapeNames = np.append(mapeNames, name)

        worksheet = workbook.get_worksheet_by_name('Активные запасы')

        for j in range(len(graph1)):
            worksheet.write(35 + j, 3 + colum, graph1[j])

        for j in range(len(graph2)):
            worksheet.write(35 + j, 5 + colum, graph2[j])

        worksheet.write(2 + i * 5, 3 + colum, 'Метод Говоровой - Рябининой')
        worksheet.write(3 + i * 5, 3 + colum, 'Активные запасы')
        worksheet.write(3 + i * 5, 4 + colum, active, bold)

        worksheet = workbook.get_worksheet_by_name('Выходные данные')


        colum += 6

    worksheet = workbook.add_worksheet('Входные данные')

    """Вывод Даты"""
    worksheet.write(0, 0, 'Дата', bold)
    for i in range(len(givenDate)):
        worksheet.write(i + 2, 0, givenDate[i])

    """Вывод Накопленной нефти и Накопленной жидкости"""
    worksheet.write(0, 1, 'Накопленная добыча нефти', bold)
    for i in range(len(givenDate)):
        worksheet.write(i + 2, 1, q_oil[i])

    worksheet.write(0, 2, 'Накопленная добыча жидкости', bold)
    for i in range(len(givenDate)):
        worksheet.write(i + 2, 2, q_liq[i])

    # Подсчет наилучших MAPE
    best_of_mape_counts = []
    best_of_mape_names = []
    temp_mape = []
    for ji in range(len(mapeCnts)):
        temp_mape.append((mapeCnts[ji]))

    for sss in range(3):
        val, idx = min((val, idx) for (idx, val) in enumerate(temp_mape))
        best_of_mape_counts.append(val)
        best_of_mape_names.append(mapeNames[idx])
        temp_mape[idx] = 100

    worksheet = workbook.add_worksheet('Графики')
    for j in range(4):
        for i in range(7):
            wcc = [0.98, 0.99, 0.995, 0.999]
            chart = workbook.add_chart({'type': 'scatter',
                                     'subtype': 'straight'})

            chart.add_series({
                'name': 'Модель',
                'categories': ['Входные данные', 52, 1, 97, 1],
                'values': ['Активные запасы', 85,  3 + (i * 6), 130,  3 + (i * 6)],
                'line': {'color': 'blue'},
            })

            chart.add_series({
                'name': 'Факт',
                'categories': ['Входные данные', 52, 1, 97, 1],
                'values': ['Активные запасы', 85,  5 + (i * 6), 130,  5 + (i * 6)],
                'line': {'color': 'red'},
            })

            chart.set_title({'name': ['Выходные данные', 0 + (i * 6), 2]})
            chart.set_x_axis({'name': 'Нефть'})
            chart.set_y_axis({'name': 'Характеристики вытеснения'})

            worksheet.write(7, 1 + (9 * j), 'wc = ', bold)  # Вывод обводненности
            worksheet.write(7, 2 + (9 * j), wcc[j], bold)

            worksheet.write(7 + (i * 17), 4, 'mape =', bold)    # Вывод MAPE
            worksheet.write(7 + (i * 17), 5, mapeCnts[i], bold)

            worksheet.write(1, 1, 'Лучший метод по показателям MAPE')
            worksheet.write('F2', best_of_mape_names[0], bold)
            worksheet.write('F4', 'MAPE = ')
            worksheet.write('G5', best_of_mape_counts[0], bold)

            worksheet.write('J2', 'На втором месте метод')
            worksheet.write('K3', best_of_mape_names[1], bold)
            worksheet.write('J4', 'MAPE = ')
            worksheet.write('K5', best_of_mape_counts[1], bold)

            worksheet.write('O2', 'На третьем месте метод')
            worksheet.write('P3', best_of_mape_names[2], bold)
            worksheet.write('O4', 'MAPE = ')
            worksheet.write('P5', best_of_mape_counts[2], bold)

            """Рамки"""
            for bu in range(3):
                worksheet.write(6 + (i * 17), 0 + (12 * bu), '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
                                                            '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'
                                                            '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
                                                            '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=', bold)
            for bu in range(6):
                worksheet.write(bu, 0, '             ||', bold)
                worksheet.write(bu, 19, '||', bold)


            chart.set_style(13)
            worksheet.insert_chart(8 + (i * 17), 1 + (9 * j), chart, {'x_offset': 25, 'y_offset': 10})

    workbook.close()

def print_method(worksheet, bold, i, q_liq, name, needed_columns, colum, a, b, active, fore_liq, fore_oil, comp_q_liq,
                 calc=0, comp_q_liq2 = 0):

    if (calc == 1):
        mapeCount = check_qual(comp_q_liq2)
    else:
        mapeCount = check_qual(comp_q_liq)


    worksheet.write(colum + (needed_columns * i), 2, name, bold)
    worksheet.write(3 + colum + (needed_columns * i), 2, 'a =')
    worksheet.write(3 + colum + (needed_columns * i), 5, a)
    worksheet.write(4 + colum + (needed_columns * i), 2, 'b =')
    worksheet.write(4 + colum + (needed_columns * i), 5, b)
    worksheet.write(1 + colum + (needed_columns * i), 6, 'Активные извлекаемые запасы')
    worksheet.write(2 + colum + (needed_columns * i), 6, active)
    worksheet.write(1 + colum + (needed_columns * i), 7, 'Прогн жидк')
    worksheet.write(2 + colum + (needed_columns * i), 7, fore_liq)
    worksheet.write(1 + colum + (needed_columns * i), 8, 'Прогн нефть')
    worksheet.write(2 + colum + (needed_columns * i), 8, fore_oil)
    for bu in range(9):
        worksheet.write(5 + colum + (needed_columns * i), 2 + (12 * bu), '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
                                                                        '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'
                                                                        '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
                                                                        '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=', bold)
    worksheet.write(1 + colum + (needed_columns * i), 10, 'Факт')
    worksheet.write(2 + colum + (needed_columns * i), 10, 'Выч')
    worksheet.write(4 + colum + (needed_columns * i), 8, 'mape =')
    worksheet.write(4 + colum + (needed_columns * i), 9, mapeCount)

    for j in range(len(q_liq)):
        if (calc == 2):
            worksheet.write(1 + colum + (needed_columns * i), 11 + j, np.log((q_liq[j] - q_oil[j]) / q_oil[j]))
        elif (calc == 3):
            worksheet.write(1 + colum + (needed_columns * i), 11 + j, (q_liq[j] - q_oil[j]) / q_oil[j])
        else:
            worksheet.write(1 + colum + (needed_columns * i), 11 + j, q_liq[j])
    for j in range(len(q_liq)):
        worksheet.write(2 + colum + (needed_columns * i), 11 + j, comp_q_liq[j])

    for j in range(len(q_liq)):
        if (calc == 1):
            worksheet.write(1 + colum + (needed_columns * i), 11 + j, q_liq[j])
            worksheet.write(2 + colum + (needed_columns * i), 11 + j, comp_q_liq2[j])


def main():
    load()
    output()

if __name__ == "__main__":
    main()
