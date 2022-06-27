# -*- coding: utf-8 -*-
import win32com.client
import tkinter
import re
import os
from datetime import datetime
from dadata import Dadata


# Функция для горячих клавиш на русской клавиатуре
def key_release(event):
    ctrl = (event.state & 0x4) != 0
    if event.keycode == 88 and ctrl and event.keysym.lower() != "x":
        event.widget.event_generate("<<Cut>>")

    if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
        event.widget.event_generate("<<Paste>>")

    if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
        event.widget.event_generate("<<Copy>>")


# Маска даты
def date_mask(text, valid, entry):
    ip = re.findall("^\d{0,2}\.\d{0,2}\.\d{0,4}$", text.get())
    if len(ip) != 1:
        text.set(valid[0])
    if ip:
        valid[0] = ip[0]
        cursor_position = entry.index("insert")
        index = ip[0][:cursor_position].rfind(u".")
        if cursor_position - index == 3:
            entry.icursor(cursor_position + 1)


# Маска времени
def time_mask(text, valid, entry):
    ip = re.findall("^\d{0,2}:\d{0,2}$", text.get())
    if len(ip) != 1:
        text.set(valid[0])
    if ip:
        valid[0] = ip[0]
        cursor_position = entry.index("insert")
        index = ip[0][:cursor_position].rfind(u":")
        if cursor_position - index == 3:
            entry.icursor(cursor_position + 1)


def string_lines(name, cells):
    """Функция разбивает name на слова помещающиеся в cell и возвращает в виде списка"""
    name = name.split()
    cache = None
    string = []

    for i_word in name:
        if cache is None:
            cache = i_word
        elif len(cache + ' ' + i_word) <= cells:
            cache += ' ' + i_word
        else:
            string.append(cache)
            cache = i_word
    string.append(cache)

    return string


def one_way(content, first_hor_cell, first_ver_cell, cells, sheet):
    """Функция записывает разбитые строки в файл"""
    hor_cell = first_hor_cell
    ver_cell = first_ver_cell
    if len(content) <= cells:
        for rec in content:
            sheet.Cells(hor_cell, ver_cell).Value = rec
            ver_cell += 3
    else:
        lines = string_lines(content, cells)
        for line in lines:
            for rec in line:
                sheet.Cells(hor_cell, ver_cell).Value = rec
                ver_cell += 3
            ver_cell = first_ver_cell
            hor_cell += 2


def scroll_func(event):
    canvas.configure(scrollregion=canvas.bbox("all"), width=450, height=600)


def _on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def show():
    lst = [rereg_on, check_1, check_2, check_3, check_4, check_5, check_6, check_7, check_8]
    if reg_var.get() == 2:
        i = 1
        for each in lst:
            each.grid(row=i, column=0, columnspan=4, sticky='w')
            i += 1
        kkt_rn_lbl.grid(row=20, column=0, sticky='w')
        kkt_rn_txt.grid(row=20, column=1, columnspan=3, padx=5, pady=5)
    else:
        for each in lst:
            each.grid_forget()
        kkt_rn_lbl.grid_forget()
        kkt_rn_txt.grid_forget()


def show_fd_info():
    if var4.get() == 1:
        fd_info_lbl.grid(row=51, column=0, columnspan=4)
        fn_crash_lbl.grid(row=52, column=0, sticky='w')
        rad1_crash.grid(row=52, column=1, sticky='w')
        rad2_crash.grid(row=52, column=1, sticky='e')
        fd_num_lbl.grid(row=53, column=0, sticky='w')
        fd_num_txt.grid(row=53, column=1, sticky='w')
        reg_date_lbl.grid(row=53, column=2, sticky='w')
        reg_date_txt.grid(row=53, column=3, sticky='w')
        reg_time_lbl.grid(row=54, column=2, sticky='w')
        reg_time_txt.grid(row=54, column=3, sticky='w')
        fp_lbl.grid(row=54, column=0, sticky='w')
        fp_txt.grid(row=54, column=1, sticky='w')
        close_fd_info_lbl.grid(row=55, column=0, columnspan=4)
        close_fd_num_lbl.grid(row=56, column=0, sticky='w')
        close_fd_num_txt.grid(row=56, column=1, sticky='w')
        close_reg_date_lbl.grid(row=56, column=2, sticky='w')
        close_reg_date_txt.grid(row=56, column=3, sticky='w')
        close_reg_time_lbl.grid(row=57, column=2, sticky='w')
        close_reg_time_txt.grid(row=57, column=3, sticky='w')
        close_fp_lbl.grid(row=57, column=0, sticky='w')
        close_fp_txt.grid(row=57, column=1, sticky='w')
    else:
        fd_info_lbl.grid_forget()
        fn_crash_lbl.grid_forget()
        rad1_crash.grid_forget()
        rad2_crash.grid_forget()
        fd_num_lbl.grid_forget()
        fd_num_txt.grid_forget()
        reg_date_lbl.grid_forget()
        reg_date_txt.grid_forget()
        reg_time_lbl.grid_forget()
        reg_time_txt.grid_forget()
        fp_lbl.grid_forget()
        fp_txt.grid_forget()
        close_fd_info_lbl.grid_forget()
        close_fd_num_lbl.grid_forget()
        close_fd_num_txt.grid_forget()
        close_reg_date_lbl.grid_forget()
        close_reg_date_txt.grid_forget()
        close_reg_time_lbl.grid_forget()
        close_reg_time_txt.grid_forget()
        close_fp_lbl.grid_forget()
        close_fp_txt.grid_forget()


def show_ofd():
    if opt1.get() == 1:
        ofd_lbl.grid_forget()
        ofd_list.grid_forget()
    else:
        ofd_lbl.grid(row=37, column=0, sticky='w')
        ofd_list.grid(row=37, column=1, columnspan=3, sticky='w')


def callback(*args):
    ofd = ofd_var.get()
    return ofd


def define_index():
    token = "dd8e2bcbf21a08ddd423a0704f5dd4a4c3001d82"
    secret = "d0d05ba0d05963a7bb34069950ca52d219471c51"
    dadata = Dadata(token, secret)

    address_parts = (region_txt.get(), rayon_txt.get(), city_txt.get(), punkt_txt.get(), street_txt.get(),
                     house_txt.get(), housing_txt.get())
    address = ' '.join(address_parts)
    result = dadata.clean(name="address", source=address)
    index_var.set(result['postal_code'])
    dadata.close()


def clicked():
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\#
    #                                         СТРАНИЦА 1                                              #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\#

    # Открытие файла
    file_dir = os.path.dirname(os.path.realpath('__file__'))
    filename = os.path.join(file_dir, 'dist', 'main', 'reg_copy.xls')

    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(Filename=filename)
    sheet_1 = wb.Sheets("s1")
    sheet_2 = wb.Sheets("s2")
    sheet_3 = wb.Sheets("s3")
    sheet_4 = wb.Sheets("s4")
    sheet_5 = wb.Sheets("s5")
    sheet_6 = wb.Sheets("s6")
    sheet_7 = wb.Sheets("s7")
    sheet_8 = wb.Sheets("s8")
    sheet_9 = wb.Sheets("s9")
    sheet_10 = wb.Sheets("s10")

    # Тип документа
    sheet_1.Cells(12, 27).Value = reg_var.get()

    # Коды причины перерегистрации
    if reg_var.get() == 2:
        sheet_1.Cells(15, 33).Value = var1.get()
        sheet_1.Cells(15, 39).Value = var2.get()
        sheet_1.Cells(15, 45).Value = var3.get()
        sheet_1.Cells(15, 51).Value = var4.get()
        sheet_1.Cells(15, 57).Value = var5.get()
        sheet_1.Cells(15, 63).Value = var6.get()
        sheet_1.Cells(15, 69).Value = var7.get()
        sheet_1.Cells(15, 75).Value = var8.get()
    else:
        sheet_1.Cells(15, 33).Value = " "
        sheet_1.Cells(15, 39).Value = " "
        sheet_1.Cells(15, 45).Value = " "
        sheet_1.Cells(15, 51).Value = " "
        sheet_1.Cells(15, 57).Value = " "
        sheet_1.Cells(15, 63).Value = " "
        sheet_1.Cells(15, 69).Value = " "
        sheet_1.Cells(15, 75).Value = " "

    # ИНН
    one_way(inn_txt.get(), 4, 40, 12, sheet_1)

    # ОГРН
    one_way(ogrn_txt.get(), 1, 40, 15, sheet_1)

    # КПП
    one_way(kpp_txt.get(), 6, 40, 9, sheet_1)

    # ФИО ИП / Наименование организации
    one_way(name_txt.get(), 17, 1, 40, sheet_1)

    # Заявитель
    sheet_1.Cells(33, 2).Value = user_var.get()

    # ФИО заявителя
    i = 1
    n = 36
    for rec in user_name_txt.get():
        if rec == " ":
            i = 1
            n = n + 2
            continue
        sheet_1.Cells(n, i).Value = rec
        i = i + 3

    # Сегодняшняя дата
    td_date = datetime.now().strftime('%d.%m.%Y')
    td = td_date.split('.')

    # День
    i = 28
    for rec in td[0]:
        sheet_1.Cells(45, i).Value = rec
        i = i + 3

    # Месяц
    i = 37
    for rec in td[1]:
        sheet_1.Cells(45, i).Value = rec
        i = i + 3

    # Год
    i = 46
    for rec in td[2]:
        sheet_1.Cells(45, i).Value = rec
        i = i + 3

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\#
    #                                         СТРАНИЦА 2                                              #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\#

    # Документ заявителя
    one_way(doc_txt.get(), 12, 1, 20, sheet_2)

    # Регистрационный номер
    one_way(kkt_rn_txt.get(), 21, 22, 20, sheet_2)

    # Дата внизу страницы
    sheet_2.Cells(48, 95).Value = td_date

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 3                                               #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    # Наименование ККТ
    one_way(kkt_name_txt.get(), 13, 55, 20, sheet_3)

    # ЗН ККТ
    one_way(kkt_zn_txt.get(), 17, 55, 20, sheet_3)

    # Наименование ФН
    one_way(fn_name_txt.get(), 21, 55, 20, sheet_3)

    # ЗН ФН
    one_way(fn_zn_txt.get(), 33, 55, 20, sheet_3)

    # Адрес установки ККТ
    one_way(index_txt.get(), 38, 31, 6, sheet_3)  # индекс
    one_way(region_txt.get(), 38, 115, 2, sheet_3)  # регион
    one_way(rayon_txt.get(), 40, 31, 30, sheet_3)  # район
    one_way(city_txt.get(), 42, 31, 30, sheet_3)  # город
    one_way(punkt_txt.get(), 44, 31, 30, sheet_3)  # населенный пункт

    # Дата внизу страницы
    sheet_3.Cells(47, 95).Value = td_date

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 4                                               #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    one_way(street_txt.get(), 9, 31, 30, sheet_4)  # улица
    one_way(house_txt.get(), 11, 31, 8, sheet_4)  # номер дома
    one_way(housing_txt.get(), 13, 31, 8, sheet_4)  # номер строения
    one_way(room_txt.get(), 15, 31, 8, sheet_4)  # номер квартиры

    # Место установки ККТ
    one_way(place_txt.get(), 18, 52, 20, sheet_4)

    # Параметры ККТ
    sheet_4.Cells(27, 52).Value = opt1.get()  # Автономный режим

    # Дата внизу страницы
    sheet_4.Cells(46, 95).Value = td_date

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 5                                               #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    sheet_5.Cells(11, 58).Value = opt2.get()  # Проведение лотерей
    sheet_5.Cells(18, 58).Value = opt3.get()  # Проведение азартных игр
    sheet_5.Cells(23, 58).Value = opt4.get()  # Обмен игорных знаков/денег
    sheet_5.Cells(29, 58).Value = opt5.get()  # Банковский платежный агент (субагент)
    sheet_5.Cells(33, 58).Value = opt6.get()  # Платежный агент (субагент)
    sheet_5.Cells(37, 58).Value = opt7.get()  # Автоматический режим
    sheet_5.Cells(42, 58).Value = opt8.get()  # Маркировка

    # Дата внизу страницы
    sheet_5.Cells(51, 95).Value = td_date

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 6                                               #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    sheet_6.Cells(10, 58).Value = opt9.get()  # Расчеты в интернете
    sheet_6.Cells(15, 58).Value = opt10.get()  # Развозная торговля
    sheet_6.Cells(19, 58).Value = opt11.get()  # БСО
    sheet_6.Cells(24, 58).Value = opt12.get()  # Подакцизные товары

    # Дата внизу страницы
    sheet_6.Cells(47, 95).Value = td_date

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 7                                               #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    if opt7.get() == 2:
        wb.Worksheets('s7').Delete()
        wb.Worksheets('s8').Delete()
    else:
        # Дата внизу страницы
        sheet_7.Cells(49, 95).Value = td_date
        sheet_8.Cells(49, 95).Value = td_date

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 8                                               #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 9                                               #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    if callback() == "OFD.ru":
        name_ofd = "ООО «ПЕТЕР-СЕРВИС СПЕЦТЕХНОЛОГИИ»"
        inn_ofd = "7841465198"
    elif callback() == "Платформа (Эвотор) ОФД":
        name_ofd = "ООО «ЭВОТОР ОФД»"
        inn_ofd = "9715260691"
    elif callback() == "Первый ОФД":
        name_ofd = "ООО «ЭНЕРГИТИЧЕСКИЕ СИСТЕМЫ И КОММУНИКАЦИИ»"
        inn_ofd = "7709364346"
    elif callback() == "Контур ННТ":
        name_ofd = "ООО «КОНТУР НТТ»"
        inn_ofd = "6658497833"
    elif callback() == "Яндекс ОФД":
        name_ofd = "ООО «Яндекс.ОФД»"
        inn_ofd = "7704358518"
    elif callback() == "Тензор":
        name_ofd = "ООО «Компания «Тензор»"
        inn_ofd = "7605016030"
    elif callback() == "Калуга Астрал":
        name_ofd = "ЗАО «КАЛУГА АСТРАЛ»"
        inn_ofd = "4029017981"
    elif callback() == "Ярус":
        name_ofd = "ООО «ЯРУС»"
        inn_ofd = "7728699517"
    elif callback() == "Дримкас":
        name_ofd = "ООО «ДРИМКАС»"
        inn_ofd = "7802870820"
    elif callback() == "Гарант":
        name_ofd = "ООО «ЭЛЕКТРОННЫЙ ЭКСПРЕСС»"
        inn_ofd = "7729633131"
    elif callback() == "Тандер":
        name_ofd = "АО «ТАНДЕР»"
        inn_ofd = "2310031475"
    elif callback() == "ИнитПро":
        name_ofd = "ООО УДОСТОВЕРЯЮЩИЙ ЦЕНТР «ИНИТПРО»"
        inn_ofd = "5902034504"
    elif callback() == "е-ОФД":
        name_ofd = "ООО «ГРУППА ЭЛЕМЕНТ»"
        inn_ofd = "7729642175"
    elif callback() == "ЭнвижнГруп (МТС)":
        name_ofd = "АО «ЭНВИЖНГРУП»"
        inn_ofd = "7703282175"
    elif callback() == "Билайн ОФД":
        name_ofd = "ПАО «ВЫМПЕЛ-КОММУНИКАЦИИ»"
        inn_ofd = "7713076301"
    elif callback() == "МультиКарта":
        name_ofd = "ООО «МУЛЬТИКАРТА»"
        inn_ofd = "7710007966"
    elif callback() == "МультиКарта":
        name_ofd = "ООО «МУЛЬТИКАРТА»"
        inn_ofd = "7710007966"
    elif callback() == "ОФД Онлайн":
        name_ofd = "ООО «ОПЕРАТОР ФИСКАЛЬНЫХ ДАННЫХ «ОНЛАЙН»"
        inn_ofd = "6686089392"
    elif callback() == "Информационный центр":
        name_ofd = "АО «ИНФОРМАЦИОННЫЙ ЦЕНТР»"
        inn_ofd = "7701553038"
    else:
        name_ofd = " "
        inn_ofd = " "
    one_way(name_ofd, 12, 55, 20, sheet_9)  # Наименование ОФД
    one_way(inn_ofd, 21, 55, 12, sheet_9)  # ИНН ОФД

    # Дата внизу страницы
    sheet_9.Cells(46, 95).Value = td_date

    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
    #                                         СТРАНИЦА 10                                              #
    # /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

    if var4.get() == 2:
        wb.Worksheets('s10').Delete()
    else:
        # ФН поврежден
        sheet_10.Cells(11, 56).Value = fn_crash.get()

        # № ФД
        one_way(fd_num_txt.get(), 18, 56, 8, sheet_10)

        # Дата
        reg_date_pr = reg_date_txt.get()
        if len(reg_date_pr) == 10:
            sheet_10.Cells(22, 56).Value = reg_date_pr[0]
            sheet_10.Cells(22, 59).Value = reg_date_pr[1]
            sheet_10.Cells(22, 65).Value = reg_date_pr[3]
            sheet_10.Cells(22, 68).Value = reg_date_pr[4]
            sheet_10.Cells(22, 74).Value = reg_date_pr[6]
            sheet_10.Cells(22, 77).Value = reg_date_pr[7]
            sheet_10.Cells(22, 80).Value = reg_date_pr[8]
            sheet_10.Cells(22, 83).Value = reg_date_pr[9]

        # Время
        reg_time_pr = reg_time_txt.get()
        if len(reg_time_pr) == 5:
            sheet_10.Cells(26, 56).Value = reg_time_pr[0]
            sheet_10.Cells(26, 59).Value = reg_time_pr[1]
            sheet_10.Cells(26, 65).Value = reg_time_pr[3]
            sheet_10.Cells(26, 68).Value = reg_time_pr[4]

        # ФП
        one_way(fp_txt.get(), 29, 56, 10, sheet_10)

        # № ФД закрытия ФН
        one_way(close_fd_num_txt.get(), 34, 56, 8, sheet_10)

        # Дата закрытия ФН
        close_reg_date_pr = close_reg_date_txt.get()
        if len(close_reg_date_pr) == 10:
            sheet_10.Cells(38, 56).Value = close_reg_date_pr[0]
            sheet_10.Cells(38, 59).Value = close_reg_date_pr[1]
            sheet_10.Cells(38, 65).Value = close_reg_date_pr[3]
            sheet_10.Cells(38, 68).Value = close_reg_date_pr[4]
            sheet_10.Cells(38, 74).Value = close_reg_date_pr[6]
            sheet_10.Cells(38, 77).Value = close_reg_date_pr[7]
            sheet_10.Cells(38, 80).Value = close_reg_date_pr[8]
            sheet_10.Cells(38, 83).Value = close_reg_date_pr[9]

        # Время закрытия ФН
        close_reg_time_pr = close_reg_time_txt.get()
        if len(close_reg_time_pr) == 5:
            sheet_10.Cells(42, 56).Value = close_reg_time_pr[0]
            sheet_10.Cells(42, 59).Value = close_reg_time_pr[1]
            sheet_10.Cells(42, 65).Value = close_reg_time_pr[3]
            sheet_10.Cells(42, 68).Value = close_reg_time_pr[4]

        # ФП закрытия ФН
        one_way(close_fp_txt.get(), 45, 56, 10, sheet_10)

        # Дата внизу страницы
        sheet_10.Cells(51, 95).Value = td_date

    # Страницы
    for i, sheet_obj in enumerate(wb.Sheets):
        sheet = wb.Sheets(sheet_obj.Name)
        if len(str(i + 1)) == 1:
            number = '00{}'.format(i + 1)
        else:
            number = '0{}'.format(i + 1)
        one_way(number, 6, 76, 3, sheet)
    one_way(number, 26, 36, 3, sheet_1)

    # Сохранение и закрытие файла
    file_way = wb.Application.GetSaveAsFilename("reg - {0}".format(name_txt.get().replace('"', '')),
                                                "Файл Excel 2007 (*.xls), *.xls")
    print(file_way)
    wb.SaveAs(file_way)
    wb.Close()
    excel.Quit()


# /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #
#                                    ИНТЕРФЕЙС ПРОГРАММЫ                                           #
# /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ #

root = tkinter.Tk()
root.title("Заявление на регистрацию/перерегистрацию ККТ")
root.geometry('460x600')
root.resizable(False, False)

canvas = tkinter.Canvas(root)
frame = tkinter.Frame(canvas)
scroll = tkinter.Scrollbar(root, orient='vertical', command=canvas.yview)
canvas.configure(yscrollcommand=scroll.set)

scroll.pack(side="right", fill="y")
canvas.pack(side="left")
canvas.create_window((0, 0), window=frame, anchor='nw')
frame.bind("<Configure>", scroll_func)
canvas.bind_all("<MouseWheel>", _on_mousewheel)

# Горячие клавиши на русской клавиатуре
root.bind_all("<Key>", key_release, "+")
# ------------Конец-------------

# -------------Тип заявления--------------
reg_lbl = tkinter.Label(frame, text="Тип заявления")
reg_lbl.grid(row=0, column=0, sticky='w')

reg_var = tkinter.IntVar()
reg_var.set(1)
rad1 = tkinter.Radiobutton(frame, text="Регистрация", variable=reg_var, value=1, command=show)
rad1.grid(row=0, column=1, sticky='w')
rad2 = tkinter.Radiobutton(frame, text="Перерегистрация", variable=reg_var, value=2, command=show)
rad2.grid(row=0, column=2, columnspan=3, sticky='w')
# -----------------Конец------------------

# -------------Причина перерегистрации--------------
rereg_on = tkinter.Label(frame, text="Причина перерегистрации")

var1 = tkinter.IntVar()
var1.set(2)
check_1 = tkinter.Checkbutton(frame, text="изменение адреса и (или) места установки ККТ", variable=var1, onvalue=1,
                              offvalue=2)

var2 = tkinter.IntVar()
var2.set(2)
check_2 = tkinter.Checkbutton(frame, text="смена ОФД", variable=var2, onvalue=1, offvalue=2)

var3 = tkinter.IntVar()
var3.set(2)
check_3 = tkinter.Checkbutton(frame, text="изменениями сведений об автоматическом устройстве", variable=var3, onvalue=1,
                              offvalue=2)

var4 = tkinter.IntVar()
var4.set(2)
check_4 = tkinter.Checkbutton(frame, text="замена ФН", variable=var4, onvalue=1, offvalue=2, command=show_fd_info)

var5 = tkinter.IntVar()
var5.set(2)
check_5 = tkinter.Checkbutton(frame, text="переход из автономного режима в режим передачи данных", variable=var5,
                              onvalue=1,
                              offvalue=2)

var6 = tkinter.IntVar()
var6.set(2)
check_6 = tkinter.Checkbutton(frame, text="переход из режима передачи данных в автономный режимм", variable=var6,
                              onvalue=1,
                              offvalue=2)

var7 = tkinter.IntVar()
var7.set(2)
check_7 = tkinter.Checkbutton(frame, text="изменение названия ЮЛ или ФИО пользователя", variable=var7, onvalue=1,
                              offvalue=2)

var8 = tkinter.IntVar()
var8.set(2)
check_8 = tkinter.Checkbutton(frame, text="иные причины", variable=var8, onvalue=1, offvalue=2)
# ------------ОГРН--------------
ogrn_lbl = tkinter.Label(frame, text="ОГРН/ОГРНИП")
ogrn_lbl.grid(row=10, column=0, sticky='w')

ogrn_txt = tkinter.Entry(frame, width=50)
ogrn_txt.grid(row=10, column=1, columnspan=3, padx=5, pady=5)
# ------------Конец-------------

# -------------ИНН--------------
inn_lbl = tkinter.Label(frame, text="ИНН")
inn_lbl.grid(row=11, column=0, sticky='w')

inn_txt = tkinter.Entry(frame, width=50)
inn_txt.grid(row=11, column=1, columnspan=3, padx=5, pady=5)
# ------------Конец-------------

# -------------КПП--------------
kpp_lbl = tkinter.Label(frame, text="КПП")
kpp_lbl.grid(row=12, column=0, sticky='w')

kpp_txt = tkinter.Entry(frame, width=50)
kpp_txt.grid(row=12, column=1, columnspan=3, padx=5, pady=5)
# ------------Конец-------------

# -------------ФИО ИП--------------
name_lbl = tkinter.Label(frame, text="ФИО ИП/\nНаименование\nорганизации", justify='left')
name_lbl.grid(row=13, column=0, sticky='w')

name_txt = tkinter.Entry(frame, width=50)
name_txt.grid(row=13, column=1, columnspan=3, padx=5, pady=5)
# ------------Конец-------------

# -------------ФИО заявителя--------------
user_lbl = tkinter.Label(frame, text="Заявитель")
user_lbl.grid(row=14, column=0, sticky='w')

user_var = tkinter.IntVar()
user_var.set(1)
rad1 = tkinter.Radiobutton(frame, text="Пользователь", variable=user_var, value=1)
rad1.grid(row=14, column=1, sticky='w')
rad2 = tkinter.Radiobutton(frame, text="Представитель пользователя", variable=user_var, value=2)
rad2.grid(row=14, column=2, columnspan=3, sticky='w')

user_name_lbl = tkinter.Label(frame, text="ФИО заявителя")
user_name_lbl.grid(row=16, column=0, sticky='w')
user_name_txt = tkinter.Entry(frame, width=50)
user_name_txt.grid(row=16, column=1, columnspan=3, padx=5, pady=5)
# -----------------Конец------------------

# ----------------Документ----------------
doc_lbl = tkinter.Label(frame, text="Документ")
doc_lbl.grid(row=17, column=0, sticky='w')

doc_txt = tkinter.Entry(frame, width=50)
doc_txt.grid(row=17, column=1, columnspan=3, padx=5, pady=5)
# -----------------Конец------------------

# ------------Наименование ККТ------------
kkt_name_lbl = tkinter.Label(frame, text="Наименование ККТ")
kkt_name_lbl.grid(row=18, column=0, sticky='w')

kkt_name_txt = tkinter.Entry(frame, width=50)
kkt_name_txt.grid(row=18, column=1, columnspan=3, padx=5, pady=5)
# -----------------Конец------------------

# ------------ЗН ККТ------------
kkt_zn_lbl = tkinter.Label(frame, text="ЗН ККТ")
kkt_zn_lbl.grid(row=19, column=0, sticky='w')

kkt_zn_txt = tkinter.Entry(frame, width=50)
kkt_zn_txt.grid(row=19, column=1, columnspan=3, padx=5, pady=5)
# -----------------Конец------------------

# ------------РН ККТ------------
kkt_rn_lbl = tkinter.Label(frame, text="РН ККТ")
kkt_rn_txt = tkinter.Entry(frame, width=50)
# -----------------Конец------------------

# ------------Наименование ФН------------
fn_name_lbl = tkinter.Label(frame, text="Наименование ФН")
fn_name_lbl.grid(row=21, column=0, sticky='w')

fn_name_txt = tkinter.Entry(frame, width=50)
fn_name_txt.grid(row=21, column=1, columnspan=3, padx=5, pady=5)
# -----------------Конец------------------

# ------------ЗН ФН------------
fn_zn_lbl = tkinter.Label(frame, text="ЗН ФН")
fn_zn_lbl.grid(row=22, column=0, sticky='w')

fn_zn_txt = tkinter.Entry(frame, width=50)
fn_zn_txt.grid(row=22, column=1, columnspan=3, padx=5, pady=5)
# -----------------Конец------------------

# ------------Адрес установки ККТ------------
address_lbl = tkinter.Label(frame, text="Адрес установки ККТ")
address_lbl.grid(row=23, column=0, columnspan=4, sticky='w' + 'e')

# Индекс
index_lbl = tkinter.Label(frame, text="Почтовый индекс")
index_lbl.grid(row=24, column=0, sticky='w')

index_var = tkinter.StringVar()
index_txt = tkinter.Entry(frame, textvariable=index_var, width=15)
index_txt.grid(row=24, column=1, sticky='w', padx=5, pady=5)

index_btn = tkinter.Button(frame, text="<-", command=define_index)
index_btn.grid(row=24, column=2, sticky='w', padx=0, pady=5)

# Регион
region_lbl = tkinter.Label(frame, text="Регион (код)")
region_lbl.grid(row=24, column=2, sticky='e')

region_txt = tkinter.Entry(frame, width=15)
region_txt.grid(row=24, column=3, padx=5, pady=5)

# Район
rayon_lbl = tkinter.Label(frame, text="Район")
rayon_lbl.grid(row=25, column=0, sticky='w')

rayon_txt = tkinter.Entry(frame, width=50)
rayon_txt.grid(row=25, column=1, columnspan=3, padx=5, pady=5)

# Город
city_lbl = tkinter.Label(frame, text="Город")
city_lbl.grid(row=26, column=0, sticky='w')

city_txt = tkinter.Entry(frame, width=50)
city_txt.grid(row=26, column=1, columnspan=3, padx=5, pady=5)

# Населенный пункт
punkt_lbl = tkinter.Label(frame, text="Населенный пункт")
punkt_lbl.grid(row=27, column=0, sticky='w')

punkt_txt = tkinter.Entry(frame, width=50)
punkt_txt.grid(row=27, column=1, columnspan=3, padx=5, pady=5)

# Улица
street_lbl = tkinter.Label(frame, text="Улица")
street_lbl.grid(row=28, column=0, sticky='w')

street_txt = tkinter.Entry(frame, width=50)
street_txt.grid(row=28, column=1, columnspan=3, padx=5, pady=5)

# Номер дома
house_lbl = tkinter.Label(frame, text="Номер дома")
house_lbl.grid(row=29, column=1, sticky='w')

house_txt = tkinter.Entry(frame, width=15)
house_txt.grid(row=30, column=1, sticky='w', padx=5, pady=5)

# Номер корпуса
housing_lbl = tkinter.Label(frame, text="Номер корпуса")
housing_lbl.grid(row=29, column=2, sticky='w')

housing_txt = tkinter.Entry(frame, width=15)
housing_txt.grid(row=30, column=2, sticky='w', padx=5, pady=5)

# Квартиры
room_lbl = tkinter.Label(frame, text="Квартира (офис)")
room_lbl.grid(row=29, column=3, sticky='w')

room_txt = tkinter.Entry(frame, width=15)
room_txt.grid(row=30, column=3, sticky='w', padx=5, pady=5)

# Место установки
place_lbl = tkinter.Label(frame, text="Место установки ККТ")
place_lbl.grid(row=31, column=0, sticky='w')

place_txt = tkinter.Entry(frame, width=50)
place_txt.grid(row=31, column=1, columnspan=3, padx=5, pady=5)

# -------------------Параметры ККТ---------------------
options_lbl = tkinter.Label(frame, text="Параметры ККТ")
options_lbl.grid(row=32, column=0, columnspan=4)

opt1 = tkinter.IntVar()
opt1.set(2)
check_opt_1 = tkinter.Checkbutton(frame, text="Автономный режим", variable=opt1, onvalue=1, offvalue=2,
                                  command=show_ofd)
check_opt_1.grid(row=33, column=0, columnspan=2, sticky='w')

opt2 = tkinter.IntVar()
opt2.set(2)
check_opt_2 = tkinter.Checkbutton(frame, text="Проведение лотерей", variable=opt2, onvalue=1, offvalue=2)
check_opt_2.grid(row=34, column=0, columnspan=2, sticky='w')

opt3 = tkinter.IntVar()
opt3.set(2)
check_opt_3 = tkinter.Checkbutton(frame, text="Проведение азартных игр", variable=opt3, onvalue=1, offvalue=2)
check_opt_3.grid(row=35, column=0, columnspan=2, sticky='w')

opt4 = tkinter.IntVar()
opt4.set(2)
check_opt_4 = tkinter.Checkbutton(frame, text="Обмен игорных знаков/денег", variable=opt4, onvalue=1,
                                  offvalue=2, justify='left')
check_opt_4.grid(row=36, column=0, columnspan=2, sticky='w')

opt5 = tkinter.IntVar()
opt5.set(2)
check_opt_5 = tkinter.Checkbutton(frame, text="Банковский платежный агент\n(субагент)", variable=opt5, onvalue=1, offvalue=2)
check_opt_5.grid(row=37, column=0, columnspan=2, sticky='w')

opt6 = tkinter.IntVar()
opt6.set(2)
check_opt_6 = tkinter.Checkbutton(frame, text="Платежный агент (субагент)", variable=opt6, onvalue=1, offvalue=2)
check_opt_6.grid(row=38, column=0, columnspan=2, sticky='w')

opt7 = tkinter.IntVar()
opt7.set(2)
check_opt_7 = tkinter.Checkbutton(frame, text="Автоматический режим", variable=opt7, onvalue=1, offvalue=2)
check_opt_7.grid(row=33, column=2, columnspan=4, sticky='w')

opt8 = tkinter.IntVar()
opt8.set(2)
check_opt_8 = tkinter.Checkbutton(frame, text="Маркировка", variable=opt8, onvalue=1, offvalue=2)
check_opt_8.grid(row=34, column=2, columnspan=4, sticky='w')

opt9 = tkinter.IntVar()
opt9.set(2)
check_opt_9 = tkinter.Checkbutton(frame, text="Расчеты только в Интернете", variable=opt9, onvalue=1, offvalue=2)
check_opt_9.grid(row=35, column=2, columnspan=4, sticky='w')

opt10 = tkinter.IntVar()
opt10.set(2)
check_opt_10 = tkinter.Checkbutton(frame, text="Развозная торговля", variable=opt10, onvalue=1, offvalue=2)
check_opt_10.grid(row=36, column=2, columnspan=4, sticky='w')

opt11 = tkinter.IntVar()
opt11.set(2)
check_opt_11 = tkinter.Checkbutton(frame, text="Только для услуг (БСО)", variable=opt11, onvalue=1, offvalue=2)
check_opt_11.grid(row=37, column=2, columnspan=4, sticky='w')

opt12 = tkinter.IntVar()
opt12.set(2)
check_opt_12 = tkinter.Checkbutton(frame, text="Продажа подакцизных товаров", variable=opt12, onvalue=1, offvalue=2)
check_opt_12.grid(row=38, column=2, columnspan=4, sticky='w')
# -----------------------Конец------------------------

# ------------------------ОФД-------------------------
ofd_lbl = tkinter.Label(frame, text="ОФД")

ofd_var = tkinter.StringVar()
ofd_var.trace('w', callback)
ofd_var.set("Выберите ОФД")
ofd_list = tkinter.OptionMenu(frame, ofd_var, "OFD.ru", "Платформа (Эвотор) ОФД", "Первый ОФД", "Контур ННТ", "Такском",
                              "Яндекс ОФД", "Тензор", "Калуга Астрал", "Ярус", "Дримкас", "Гарант", "Тандер", "ИнитПро",
                              "е-ОФД", "ЭнвижнГруп (МТС)", "Билайн ОФД", "МультиКарта", "ОФД Онлайн",
                              'Информационный центр')
ofd_lbl.grid(row=50, column=0, sticky='w')
ofd_list.grid(row=50, column=1, columnspan=3, sticky='w')
# -----------------------Конец------------------------

# ------------------------Сведения о ФД-------------------------
fd_info_lbl = tkinter.Label(frame, text="Сведения из отчета о регистрации")

# ФН поврежден
fn_crash_lbl = tkinter.Label(frame, text="ФН поврежден")

fn_crash = tkinter.IntVar()
fn_crash.set(2)
rad1_crash = tkinter.Radiobutton(frame, text="да", variable=fn_crash, value=1)
rad2_crash = tkinter.Radiobutton(frame, text="нет", variable=fn_crash, value=2)

# № ФД регистрации/перерегистрации
fd_num_lbl = tkinter.Label(frame, text="№ ФД отчета\nо регистрации", justify='left')
fd_num_txt = tkinter.Entry(frame, width=15)

# Дата отчета о регистрации/перерегистрации
reg_date_lbl = tkinter.Label(frame, text="Дата отчета\nо регистрации", justify='left')

reg_date_var = tkinter.StringVar()
reg_date_txt = tkinter.Entry(frame, textvariable=reg_date_var, width=10)

reg_date_last_valid = [u".."]
reg_date_var.trace("w", lambda *args: date_mask(reg_date_var, reg_date_last_valid, reg_date_txt))
reg_date_var.set("")

# Время
reg_time_lbl = tkinter.Label(frame, text="Время отчета\nо регистрации", justify='left')

reg_time_var = tkinter.StringVar()
reg_time_txt = tkinter.Entry(frame, textvariable=reg_time_var, width=10)

reg_time_last_valid = [u":"]
reg_time_var.trace("w", lambda *args: time_mask(reg_time_var, reg_time_last_valid, reg_time_txt))
reg_time_var.set("")

# Фискальный признак отчета регистрации/перерегистрации
fp_lbl = tkinter.Label(frame, text="ФП отчета\nо регистрации", justify='left')
fp_txt = tkinter.Entry(frame, width=15)

# Сведения из отчета о закрытии ФН
close_fd_info_lbl = tkinter.Label(frame, text="Сведения из отчета о закрытии ФН")

# № ФД закрытия ФН
close_fd_num_lbl = tkinter.Label(frame, text="№ ФД отчета\nо закрытии ФН", justify='left')
close_fd_num_txt = tkinter.Entry(frame, width=15)

# Дата отчета о закрытии ФН
close_reg_date_lbl = tkinter.Label(frame, text="Дата отчета\nо закрытии ФН", justify='left')

close_reg_date_var = tkinter.StringVar()
close_reg_date_txt = tkinter.Entry(frame, textvariable=close_reg_date_var, width=10)

close_reg_date_last_valid = [u".."]
close_reg_date_var.trace("w",
                         lambda *args: date_mask(close_reg_date_var, close_reg_date_last_valid, close_reg_date_txt))
close_reg_date_var.set("")

# Время
close_reg_time_lbl = tkinter.Label(frame, text="Время отчета\nо закрытии ФН", justify='left')

close_reg_time_var = tkinter.StringVar()
close_reg_time_txt = tkinter.Entry(frame, textvariable=close_reg_time_var, width=10)

close_reg_time_last_valid = [u":"]
close_reg_time_var.trace("w",
                         lambda *args: time_mask(close_reg_time_var, close_reg_time_last_valid, close_reg_time_txt))
close_reg_time_var.set("")

# Фискальный признак отчета закрытия ФН
close_fp_lbl = tkinter.Label(frame, text="ФП отчета\nо закрытии ФН", justify='left')
close_fp_txt = tkinter.Entry(frame, width=15)
# ------------------------------Конец-------------------------------

# ------------Кнопка Сформировать и Закрыть-------------
btn = tkinter.Button(frame, text="Сформировать", command=clicked)
btn.grid(row=100, column=0, sticky='w', padx=5, pady=5)

# btn_exit = tkinter.Button(frame, text="Закрыть", command=quit)
# btn_exit.grid(row=100, column=3, sticky='e', padx=5, pady=5)
# -----------------Конец------------------

tkinter.mainloop()
