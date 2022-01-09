import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from tkcalendar import *
import pandas as pd
from csv import reader
import py_win_keyboard_layout
from datetime import datetime
from time import sleep
from selenium.webdriver.common.keys import Keys
from sheet2dict import Worksheet
import chromedriver_autoinstaller
import babel.numbers
chromedriver_autoinstaller.install(cwd=True)

py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)
today = datetime.today().strftime("%d-%m-%Y")

def codes():
    global codes_list, names_eik
    initial_list = []
    with open("codes.csv") as file:
        csv_reader = reader(file)
        next(csv_reader)
        for row in csv_reader:
            initial_list.append(row)

    codes_list = []
    for sublist in initial_list:
        for item in sublist:
            codes_list.append(item)

    names_eik1 = Worksheet()
    names_eik1.xlsx_to_dict(path='names1.xlsx')
    names_eik1 = names_eik1.sheet_items
    names_eik = {}
    for x in names_eik1:
        names_eik[x['EIK']] = x['Name']
    return codes_list, names_eik


codes()


def diff_month(card_date):
    todays_date = datetime.today().strftime("%d-%m-%Y")
    td = int(todays_date.split("-")[2])
    cd = int(card_date.split("-")[2])
    td_m = int(todays_date.split("-")[1])
    cd_m = int(card_date.split("-")[1])
    return ((td - cd) * 12) + td_m - cd_m


def date_str(date):
    date = list(date)
    if date[0] == '0':
        if date[3] == '0':
            date[3] = ""
        date[0] = ''

    if date[3] == '0':
        date[3] = ""
    new_date = "".join(date)
    return new_date


today = date_str(today)


def twopointeight():
    df_ml = pd.read_excel("unishtozhenie_ML.xlsx",
                          dtype={"EIK": str, "Added_to_Sum": str, 'Submitted': str})
    not_submitted = df_ml.loc[df_ml["Submitted"].isnull()]

    list_dates = not_submitted["Data"].tolist()
    list_dates = list(set(list_dates))
    for date in list_dates:
        date_f = not_submitted.loc[not_submitted["Data"] == date]
        sum_ml = date_f.sum(axis=0)["Koli4estvo_obshto"]
        if sum_ml > 2.8:
            result = 'problem'
            date = date.strftime('%d/%m/%Y')
            label['text'] = f'Общото количество\n за дата: {date}\nнадвишава 2.8 тона'
            tk.messagebox.showerror(title='Грешка',
                                    message=f'Общото количество за дата:\n {date}надвишава 2.8 тона')
            return result


def unishtozhenie(firm, code, kol, *args):
    file = "unishtozhenie_drugi.xlsx"
    if firm == 'Екосейф':
        eik_uni = '204712082'
        osn_text = '03-ДО-666-00/23.08.2019г.'
        dein = ' '
    elif firm == 'ПУДООС':
        eik_uni = '131045382'
        osn_text = '12- ДО-1202-00/06.12.2012 г.'
        dein = 'наземно горене'

    if kol == '' or float(kol) == 0:
        label['text'] = 'Моля въведете количество'
        return None

    date = cal.get()
    month_move = diff_month(date)
    entry_kol.delete(0, len(entry_kol.get()))

    df_dest = pd.read_excel(file)
    fin_df = df_dest.loc[df_dest["Code"] == code]
    indx = df_dest.loc[df_dest["Code"] == code].index

    final_sum = fin_df.sum(axis=0)["Koli4estvo_obshto"] - float(kol)
    if float(final_sum) < 0:
        label['text'] = "Надвишавате наличното количество\nпо този код"
        tk.messagebox.showerror(title='Грешка',
                                message=f'Общото количество за дата:\n {date}надвишава 2.8 тона')
        return None
    df_dest.at[indx, "Koli4estvo_obshto"] = final_sum

    try:
        web = webdriver.Chrome()
        web.maximize_window()
        url = "https://nwms.eea.government.bg/app/base/home"
        web.get(url)
        sleep(2)
        vhod = web.find_element_by_xpath(
            "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
        vhod.click()
        sleep(2)
        el_akt = web.find_element_by_xpath("/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
        el_akt.click()

        sleep(10)
        firm1 = None
        i = 0
        while firm1 == None and i < 6:
            try:
                firm1 = web.find_element_by_xpath(
                    '/html/body/app-root/app-auth-main-page/app-login-page/app-organization-selector/div/div/div[2]/div/button')
            except:
                sleep(2)
                web.close()
                web.get(url)
                sleep(2)
                vhod = web.find_element_by_xpath(
                    "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
                vhod.click()
                sleep(2)
                el_akt = web.find_element_by_xpath(
                    "/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
                el_akt.click()
                sleep(15)
        firm1.click()
        sleep(5)
        otcheti = web.find_element_by_xpath(
            '/html/body/app-root/app-messages-main-page/div/div[2]/app-subheader/nav/div/ul/li[2]/a')
        otcheti.click()
        sleep(10)
        otchetni_knigi = web.find_element_by_xpath(
            '/html/body/app-root/app-reports-main-page/div/div[1]/app-sidebar-menu/nav/app-tree-view/ul/li[1]/button/div/div[2]')
        otchetni_knigi.click()
        sleep(5)
        web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
        sleep(2)
        tursene0 = web.find_element_by_xpath(
            '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/div[2]/div/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
        tursene0.send_keys('Събиране')
        sleep(2)
        transp0 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
        transp0.click()
        sleep(2)
        tursene = web.find_element_by_xpath(
            '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-search-and-navigation-bar/ul/li[3]/button')
        tursene.click()
        sleep(2)
        opolz = web.find_element_by_xpath(
            '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-data-viewer-as-table/div[2]/table/tbody/tr[1]/td/div/button/i')
        opolz.click()
        sleep(2)
        predaden_otpaduk = web.find_element_by_xpath('/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/ul/li[3]/a')
        predaden_otpaduk.click()
        sleep(7)
        x = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[3]/app-fi-select-dropdown/div/div/ng-select'
        inputc = web.find_element_by_xpath(x)
        inputc.click()
        sleep(2)
        uridi4esko_lice = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div[1]')
        sleep(2)
        uridi4esko_lice.click()
        code_inp = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[1]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
        code1 = web.find_element_by_xpath(code_inp)
        code1.send_keys(code)
        sleep(2)
        code2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
        sleep(2)
        code2.click()
        sleep(3)
        if today != date:
            cal1 = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[2]/app-fi-input-date/div/div/div[2]/button[1]')
            cal1.click()
            sleep(2)
            if month_move >= 1:
                sleep(3)
                cal_mv = web.find_element_by_xpath(
                    '/html/body/ngb-datepicker/div[1]/ngb-datepicker-navigation/div[1]/button')
                for x in range(month_move):
                    cal_mv.click()
                    sleep(1)
            # elif month_move >= 1:
            #     tk.messagebox.showerror(title='Грешка',
            #                             message="Невалидна дата\n(датата е от преди повече от 1 месец)")
            #     web.close()
            #     return 0
            data_pos = web.find_element_by_css_selector(f"div[aria-label='{date}']")
            data_pos.click()
            sleep(2)
        sleep(3)
        eik_u = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
        eik1 = web.find_element_by_xpath(eik_u)
        eik1.send_keys(eik_uni)
        sleep(4)
        eik2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
        sleep(4)
        eik2.click()
        lice = web.find_element_by_xpath(
            '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[5]/app-fi-input-text/div/div/input')
        try:
            lice.send_keys(names_eik[eik_uni])
        except KeyError:
            tk.messagebox.showerror(title='Грешка',
                                    message="В names.txt файла липсва \nимето срещу това ЕИК")
            return 0
        sleep(3)

        sleep(2)
        koli4estvo = web.find_element_by_xpath(
            '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[6]/app-fi-input-float/div/div/input')
        koli4estvo.send_keys(kol)
        sleep(2)
        web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
        sleep(2)
        osnovanie = web.find_element_by_xpath(
            '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[7]/app-fi-input-long-text/div/div/textarea')
        osnovanie.send_keys(osn_text)
        sleep(2)
        deinc = web.find_element_by_xpath(
            '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[8]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
        deinc.send_keys('D10')
        sleep(2)
        dein10 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
        dein10.click()
        sleep(2)
        dein_op = web.find_element_by_xpath(
            '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[9]/app-fi-input-long-text/div/div/textarea')
        dein_op.send_keys(dein)
        sleep(2)
        zapis = web.find_element_by_xpath(
            '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[10]/div[2]/button')
        sleep(2)

        if not zapis.is_enabled():
            # label['text'] = "Възникна проблем при \nподаването на картата"
            tk.messagebox.showerror(title='Грешка',
                                    message="Възникна проблем при \nподаването на картата")
            return 0
        sleep(5)
        zapis.click()
        df_dest.to_excel(file, index=False)
        sleep(5)

        web.close()
    except Exception as e:
        tk.messagebox.showerror(title='Грешка',
                                message="Възникна проблем при \nподаването на картата")
        print(e)
        return 0

    label['text'] = "Подадено"
    tk.messagebox.showinfo(title=None,
                           message='Данните бяха подадени успешно')


def import_stuff(k=1):
    new_arch()
    destruction_filing()
    label['text'] = ''
    df = pd.read_excel("individual_info.xlsx",
                       dtype={'EIK_Tovarodatel': str, 'EIK_Polu4atel': str, 'Submitted': str})
    nepod = df.loc[df["Submitted"].isnull()]
    nepod = nepod.loc[df['EIK_Polu4atel'].notnull()]
    if nepod.empty:
        if k!=1:
            return None
        label['text'] = 'Всички карти са подадени и сумирани\n(транспортиране)'
        tk.messagebox.showinfo(title=None,
                               message='Данните бяха подадени успешно')

        return None
    global web
    web = webdriver.Chrome()
    web.maximize_window()

    url = "https://nwms.eea.government.bg/app/base/home"
    web.get(url)
    sleep(2)
    vhod = web.find_element_by_xpath(
        "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
    vhod.click()
    sleep(2)
    el_akt = web.find_element_by_xpath("/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
    el_akt.click()

    sleep(15)
    firm = None
    i = 0
    while firm == None and i < 6:
        try:
            firm = web.find_element_by_xpath(
                '/html/body/app-root/app-auth-main-page/app-login-page/app-organization-selector/div/div/div[2]/div/button')
        except:
            sleep(2)
            web.close()
            web.get(url)
            sleep(2)
            vhod = web.find_element_by_xpath(
                "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
            vhod.click()
            sleep(2)
            el_akt = web.find_element_by_xpath(
                "/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
            el_akt.click()
            sleep(15)

    firm.click()
    sleep(5)
    otcheti = web.find_element_by_xpath(
        '/html/body/app-root/app-messages-main-page/div/div[2]/app-subheader/nav/div/ul/li[2]/a')
    otcheti.click()
    sleep(10)
    otchetni_knigi = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[1]/app-sidebar-menu/nav/app-tree-view/ul/li[1]/button/div/div[2]')
    otchetni_knigi.click()
    sleep(5)
    web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
    sleep(2)
    tursene0 = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/div[2]/div/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
    tursene0.send_keys('транспортиране')
    sleep(2)
    transp0 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
    transp0.click()
    sleep(2)
    tursene = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-search-and-navigation-bar/ul/li[3]/button')
    tursene.click()
    sleep(2)
    trans = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-data-viewer-as-table/div[2]/table/tbody/tr[1]/td/div/button/i')
    trans.click()
    sleep(2)
    polu4en_otpaduk = web.find_element_by_xpath('/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/ul/li[2]/a')
    polu4en_otpaduk.click()
    sleep(10)
    for index, row in nepod.iterrows():
        code_enter = nepod.at[index, "Code"]
        eik_enter = str(nepod.at[index, "EIK_Tovarodatel"])
        amnt = str(nepod.at[index, "Koli4estvo"])
        card_date_0 = nepod.at[index, "Data"].strftime('%d/%m/%Y')
        card_date = datetime.strptime(card_date_0, '%d/%m/%Y')
        card_date = card_date.strftime('%d-%m-%Y')
        try:
            month_move = diff_month(card_date)
            card_date = date_str(card_date)
        except IndexError:
            label['text'] = "Невалиден формат на датата\n (форматът трябва изглежда така: 09/02/2021)"
            label.config(font='Roboto 16 italic bold')
            return 0
        try:

            # sleep(3)
            # web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
            sleep(3)
            code_inp = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[1]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            code1 = web.find_element_by_xpath(code_inp)
            code1.send_keys(code_enter)
            sleep(2)
            code2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            code2.click()
            sleep(2)
            x = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[3]/app-fi-select-dropdown/div/div/ng-select'
            inputc = web.find_element_by_xpath(x)
            inputc.click()
            sleep(2)
            uridi4esko_lice = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div[1]')
            sleep(2)
            uridi4esko_lice.click()

            eik_u = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            eik1 = web.find_element_by_xpath(eik_u)
            eik1.send_keys(eik_enter)
            sleep(4)
            eik2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(4)
            eik2.click()
            sleep(3)

            if today != card_date:
                cal = web.find_element_by_xpath(
                    '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[2]/app-fi-input-date/div/div/div[2]/button[1]')
                cal.click()
                sleep(2)
                if month_move >= 1:
                    sleep(3)
                    cal_mv = web.find_element_by_xpath(
                        '/html/body/ngb-datepicker/div[1]/ngb-datepicker-navigation/div[1]/button')
                    for x in range(month_move):
                        cal_mv.click()
                        sleep(1)
                # elif month_move >= 1:
                #     # label['text'] = "Невалидна дата\n(датата е от преди повече от 1 месец)"
                #     tk.messagebox.showerror(title='Грешка',
                #                             message=f'Невалидна дата\n(датата е от преди повече от 1 месец')
                #     web.close()
                #     return 0
                data_pos = web.find_element_by_css_selector(f"div[aria-label='{card_date}']")
                data_pos.click()
                sleep(2)

            lice = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[5]/app-fi-input-text/div/div/input')
            try:
                lice.send_keys(names_eik[eik_enter])
            except KeyError:
                # label['text'] = "В names.txt файла липсва \nимето срещу това ЕИК"
                tk.messagebox.showerror(title='Грешка',
                                        message=f'В names1.xlsx файла липсва \nимето срещу това ЕИК:\n{eik_enter}')
                return 0

            koli4estvo = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[6]/app-fi-input-float/div/div/input')
            koli4estvo.send_keys(amnt)
            sleep(2)
            koli4estvo.send_keys(Keys.PAGE_DOWN)

            proizhod = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[7]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            pr = web.find_element_by_xpath(proizhod)
            pr.send_keys('извън')
            sleep(3)
            pr2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            pr2.click()
            sleep(2)
            if code_enter == '16 03 05*' or code_enter == '16 03 03*':
                osn_text = ' '
                op_text = ' '
            elif code_enter == '15 01 10*':
                osn_text = 'От дейността на фирмата'
                op_text = 'Празни опаковки'
            elif code_enter == '20 01 21*':
                osn_text = 'От периодична промяна на луминисцентните лампи'
                op_text = 'Луминисцентни лампи'
            else:
                osn_text = 'Здравеопазване'
                op_text = 'Клинични отпадъци'
            osnovanie = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[8]/app-fi-input-long-text/div/div/textarea')
            osnovanie.send_keys(osn_text)
            sleep(2)
            opisanie = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[9]/app-fi-input-long-text/div/div/textarea')
            opisanie.send_keys(op_text)
            zapis = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[10]/div[2]/button')
            sleep(5)
            if not zapis.is_enabled():
                tk.messagebox.showerror(title='Грешка',
                                        message=f'Възникна проблем при \nподаването на картата')
                return 0
            sleep(3)
            zapis.click()
            sleep(5)
            df.at[index, "Submitted"] = 'Podadeno'
            df.to_excel('individual_info.xlsx', index=False)
            sleep(3)
            web.find_element_by_tag_name('body').click()
            web.find_element_by_tag_name('body').send_keys(Keys.HOME)
            ### Mahni Posle

        except Exception as e:
            web.close()
            tk.messagebox.showerror(title='Грешка',
                                    message=f'Възникна проблем при \nподаването на картата')
            print(e)
            return 0
    sleep(2)
    df = pd.read_excel("suhranenie.xlsx",
                       dtype={'EIK': str, 'Submitted': str, 'Submitted_Predaden': str})
    nepod = df.loc[df["Submitted_Predaden"].isnull()]
    sleep(2)
    predaden = web.find_element_by_xpath('/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/ul/li[3]/a')
    predaden.click()
    sleep(5)

    for index, row in nepod.iterrows():
        code_enter = nepod.at[index, "Code"]
        eik_enter = str(nepod.at[index, "EIK"])
        amnt = str(nepod.at[index, "Koli4estvo_obshto"])
        card_date_0 = nepod.at[index, "Data"].strftime('%d/%m/%Y')
        card_date = datetime.strptime(card_date_0, '%d/%m/%Y')
        card_date = card_date.strftime('%d-%m-%Y')
        try:
            month_move = diff_month(card_date)
            card_date = date_str(card_date)
        except IndexError:
            label['text'] = "Невалиден формат на датата\n (форматът трябва изглежда така: 09/02/2021)"
            label.config(font='Roboto 16 italic bold')
            return 0
        try:
            code_inp = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[1]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            code1 = web.find_element_by_xpath(code_inp)
            code1.send_keys(code_enter)
            sleep(2)
            code2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            code2.click()
            sleep(2)
            x = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[3]/app-fi-select-dropdown/div/div/ng-select'
            inputc = web.find_element_by_xpath(x)
            inputc.click()
            sleep(2)
            uridi4esko_lice = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div[1]')
            sleep(2)
            uridi4esko_lice.click()

            eik_u = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            eik1 = web.find_element_by_xpath(eik_u)
            eik1.send_keys(eik_enter)
            sleep(4)
            eik2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(4)
            eik2.click()
            sleep(3)

            if today != card_date:
                cal = web.find_element_by_xpath(
                    '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[2]/app-fi-input-date/div/div/div[2]/button[1]')
                cal.click()
                sleep(2)
                if month_move >= 1:
                    sleep(3)
                    cal_mv = web.find_element_by_xpath(
                        '/html/body/ngb-datepicker/div[1]/ngb-datepicker-navigation/div[1]/button')
                    for x in range(month_move):
                        cal_mv.click()
                        sleep(1)
                # elif month_move >= 1:
                #     label['text'] = "Невалидна дата\n(датата е от преди повече от 1 месец)"
                #     web.close()
                #     return 0
                data_pos = web.find_element_by_css_selector(f"div[aria-label='{card_date}']")
                data_pos.click()
                sleep(2)

            lice = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[5]/app-fi-input-text/div/div/input')
            try:
                lice.send_keys(names_eik[eik_enter])
            except KeyError:
                # label['text'] = "В names.txt файла липсва \nимето срещу това ЕИК"
                tk.messagebox.showerror(title='Грешка',
                                        message=f'В names1.xlsx файла липсва \nимето срещу това ЕИК:\n{eik_enter}')
                return 0

            koli4estvo = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[6]/app-fi-input-float/div/div/input')
            koli4estvo.send_keys(amnt)
            sleep(2)
            koli4estvo.send_keys(Keys.PAGE_DOWN)
            sleep(2)
            osnovanie = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[7]/app-fi-input-long-text/div/div/textarea')
            osn = '07-ДО-321-03/24.01.2019г.'
            osnovanie.send_keys(osn)
            sleep(2)
            deinost = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[8]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
            deinost.send_keys('D15')
            sleep(2)
            deinost15 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            deinost15.click()
            sleep(2)
            deinost_op = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[9]/app-fi-input-long-text/div/div/textarea')
            deinost_op.send_keys('Съхраняване до извършване на коя да е дейност от D1 до D14')
            sleep(2)
            zapis = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[10]/div[2]/button')
            sleep(5)
            if not zapis.is_enabled():
                tk.messagebox.showerror(title='Грешка',
                                        message=f'Възникна проблем при \nподаването на картата')
                return 0
            zapis.click()
            df.at[index, "Submitted_Predaden"] = 'Podadeno'
            df.to_excel('suhranenie.xlsx', index=False)
            sleep(5)
            web.find_element_by_tag_name('body').send_keys(Keys.HOME)
            sleep(1)

        except Exception as e:
            tk.messagebox.showerror(title='Грешка',
                                    message=f'Възникна проблем при \nподаването на картата')
            print(e)
            return 0

    sleep(2)
    df = pd.read_excel("unishtozhenie_ML.xlsx",
                       dtype={'EIK': str, 'Submitted': str, 'Submitted_Predaden': str, 'Submitted_Tretiran': str})
    nepod = df.loc[df["Submitted_Predaden"].isnull()]
    sleep(2)
    predaden = web.find_element_by_xpath('/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/ul/li[3]/a')
    predaden.click()
    sleep(5)

    for index, row in nepod.iterrows():
        code_enter = nepod.at[index, "Code"]
        eik_enter = str(nepod.at[index, "EIK"])
        amnt = str(nepod.at[index, "Koli4estvo_obshto"])
        card_date_0 = nepod.at[index, "Data"].strftime('%d/%m/%Y')
        card_date = datetime.strptime(card_date_0, '%d/%m/%Y')
        card_date = card_date.strftime('%d-%m-%Y')
        try:
            month_move = diff_month(card_date)
            card_date = date_str(card_date)
        except IndexError:
            label['text'] = "Невалиден формат на датата\n (форматът трябва изглежда така: 09/02/2021)"
            label.config(font='Roboto 16 italic bold')
            return 0
        try:
            code_inp = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[1]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            code1 = web.find_element_by_xpath(code_inp)
            code1.send_keys(code_enter)
            sleep(2)
            code2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            code2.click()
            sleep(2)
            x = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[3]/app-fi-select-dropdown/div/div/ng-select'
            inputc = web.find_element_by_xpath(x)
            inputc.click()
            sleep(2)
            uridi4esko_lice = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div[1]')
            sleep(2)
            uridi4esko_lice.click()

            eik_u = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            eik1 = web.find_element_by_xpath(eik_u)
            eik1.send_keys(eik_enter)
            sleep(4)
            eik2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(4)
            eik2.click()
            sleep(3)

            if today != card_date:
                cal = web.find_element_by_xpath(
                    '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[2]/app-fi-input-date/div/div/div[2]/button[1]')
                cal.click()
                sleep(2)
                if month_move >= 1:
                    sleep(3)
                    cal_mv = web.find_element_by_xpath(
                        '/html/body/ngb-datepicker/div[1]/ngb-datepicker-navigation/div[1]/button')
                    for x in range(month_move):
                        cal_mv.click()
                        sleep(1)
                # elif month_move >= 1:
                #     label['text'] = "Невалидна дата\n(датата е от преди повече от 1 месец)"
                #     web.close()
                #     return 0
                data_pos = web.find_element_by_css_selector(f"div[aria-label='{card_date}']")
                data_pos.click()
                sleep(2)

            lice = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[5]/app-fi-input-text/div/div/input')
            try:
                lice.send_keys(names_eik[eik_enter])
            except KeyError:
                # label['text'] = "В names.txt файла липсва \nимето срещу това ЕИК"
                tk.messagebox.showerror(title='Грешка',
                                        message=f'В names.txt файла липсва \nимето срещу това ЕИК:\n{eik_enter}')
                return 0

            koli4estvo = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[6]/app-fi-input-float/div/div/input')
            koli4estvo.send_keys(amnt)
            sleep(2)
            koli4estvo.send_keys(Keys.PAGE_DOWN)
            sleep(2)
            osnovanie = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[7]/app-fi-input-long-text/div/div/textarea')
            osn = '07-ДО-321-03/24.01.2019г.'
            osnovanie.send_keys(osn)
            sleep(2)
            deinost = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[8]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
            deinost.send_keys('D09')
            sleep(2)
            deinost09 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            deinost09.click()
            sleep(2)
            deinost_op = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[9]/app-fi-input-long-text/div/div/textarea')
            deinost_op.send_keys('Физико-химично третиране, чрез автоклавиране	')
            sleep(2)
            zapis = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[10]/div[2]/button')
            sleep(5)
            if not zapis.is_enabled():
                tk.messagebox.showerror(title='Грешка',
                                        message=f'Възникна проблем при \nподаването на картата')
                return 0
            sleep(5)
            zapis.click()
            df.at[index, "Submitted_Predaden"] = 'Podadeno'
            df.to_excel('unishtozhenie_ML.xlsx', index=False)
            sleep(5)
            web.find_element_by_tag_name('body').click()
            web.find_element_by_tag_name('body').send_keys(Keys.HOME)
            sleep(1)
            ### Mahni Posle
        except Exception as e:
            tk.messagebox.showerror(title='Грешка',
                                    message=f'Възникна проблем при \nподаването на картата')
            print(e)
            return 0
    if k==1:
        web.close()
        # df.to_excel('individual_info.xlsx', index=False)

        label['text'] = 'Всички карти са подадени и сумирани'
        tk.messagebox.showinfo(title=None,
                               message='Данните бяха подадени успешно')
    elif k!=1:
        try:
            sleep(2)
            nazad = web.find_element_by_xpath('/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/div[1]/button')
            nazad.click()
        except:
            pass
    return None


def destruction_filing():
    df_ot = pd.read_excel("suhranenie.xlsx",
                          dtype={"EIK": str, "Added_to_Sum": str, 'Submitted': str})
    df_uni_ot = pd.read_excel("unishtozhenie_drugi.xlsx",
                              dtype={"Code": str})

    for code in codes_list:
        to_enter = df_ot.loc[df_ot["Code"] == code].loc[df_ot["Added_to_Sum"].isnull()]
        if not to_enter.empty:
            sum_to_add = to_enter.sum(axis=0)["Koli4estvo_obshto"]

            fin_df = df_uni_ot.loc[df_uni_ot["Code"] == code]
            indx = df_uni_ot.loc[df_uni_ot["Code"] == code].index
            if not fin_df.empty:
                final_sum = fin_df.sum(axis=0)["Koli4estvo_obshto"] + sum_to_add
                df_uni_ot.at[indx, "Koli4estvo_obshto"] = final_sum
            else:
                new_row = {'Code': code, 'Koli4estvo_obshto': sum_to_add}
                df_uni_ot = df_uni_ot.append(new_row, ignore_index=True)
    df_ot.loc[df_ot['Added_to_Sum'].isnull(), 'Added_to_Sum'] = 'Da'
    df_ot.to_excel('suhranenie.xlsx', index=False)
    df_uni_ot.to_excel('unishtozhenie_drugi.xlsx', index=False)
    label['text'] = 'Информацията за унищожение\nе ъпдейтната'
    return None


def storage_import(k=1):
    destruction_filing()
    global web
    label['text'] = ''
    df = pd.read_excel("suhranenie.xlsx",
                       dtype={'EIK': str, 'Submitted': str})
    nepod = df.loc[df["Submitted"].isnull()]
    if nepod.empty:
        if k!=1:
            return None
        label['text'] = 'Всички карти са подадени и сумирани'
        tk.messagebox.showinfo(title=None,
                               message='Данните бяха подадени успешно (Събиране)')
        return None
    print('suhranenie')
    try:
        k = web
    except:
        web = webdriver.Chrome()
        web.maximize_window()

        url = "https://nwms.eea.government.bg/app/base/home"
        web.get(url)
        sleep(2)
        vhod = web.find_element_by_xpath(
            "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
        vhod.click()
        sleep(2)
        el_akt = web.find_element_by_xpath("/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
        el_akt.click()
        sleep(15)
        firm = None
        i = 0
        while firm == None and i < 6:
            try:
                firm = web.find_element_by_xpath(
                    '/html/body/app-root/app-auth-main-page/app-login-page/app-organization-selector/div/div/div[2]/div/button')
            except:
                sleep(2)
                web.close()
                web.get(url)
                sleep(2)
                vhod = web.find_element_by_xpath(
                    "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
                vhod.click()
                sleep(2)
                el_akt = web.find_element_by_xpath(
                    "/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
                el_akt.click()
                sleep(15)

        firm.click()

        sleep(5)
        otcheti = web.find_element_by_xpath(
            '/html/body/app-root/app-messages-main-page/div/div[2]/app-subheader/nav/div/ul/li[2]/a')
        otcheti.click()
        sleep(10)
        otchetni_knigi = web.find_element_by_xpath(
            '/html/body/app-root/app-reports-main-page/div/div[1]/app-sidebar-menu/nav/app-tree-view/ul/li[1]/button/div/div[2]')
        otchetni_knigi.click()
        sleep(5)
        web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
        sleep(2)
    try:
        tursene0 = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/div[2]/div/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
    except:
        sleep(1)
        tursene0 = web.find_element_by_xpath('/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/div[2]/div/app-fi-select-dropdown/div/div/ng-select/div/div/div[3]/input')
    tursene0.send_keys('събиране')
    sleep(2)
    transp0 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
    transp0.click()
    sleep(2)
    tursene = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-search-and-navigation-bar/ul/li[3]/button')
    tursene.click()
    sleep(2)
    trans = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-data-viewer-as-table/div[2]/table/tbody/tr[1]/td/div/button/i')
    trans.click()
    sleep(2)
    polu4en_otpaduk = web.find_element_by_xpath('/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/ul/li[2]/a')
    polu4en_otpaduk.click()
    sleep(10)

    for index, row in nepod.iterrows():
        code_enter = nepod.at[index, "Code"]
        eik_enter = str(nepod.at[index, "EIK"])
        amnt = str(nepod.at[index, "Koli4estvo_obshto"])
        card_date_0 = nepod.at[index, "Data"].strftime('%d/%m/%Y')
        card_date = datetime.strptime(card_date_0, '%d/%m/%Y')
        card_date = card_date.strftime('%d-%m-%Y')
        try:
            month_move = diff_month(card_date)
            card_date = date_str(card_date)
        except IndexError:
            label['text'] = "Невалиден формат на датата\n (форматът трябва изглежда така: 09/02/2021)"
            label.config(font='Roboto 16 italic bold')
            return 0
        try:
            web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
            sleep(3)
            x = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[3]/app-fi-select-dropdown/div/div/ng-select'
            inputc = web.find_element_by_xpath(x)
            inputc.click()
            sleep(2)
            uridi4esko_lice = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div[1]')
            sleep(2)
            uridi4esko_lice.click()
            code_inp = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[1]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            code1 = web.find_element_by_xpath(code_inp)
            code1.send_keys(code_enter)
            sleep(2)
            code2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            code2.click()

            eik_u = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            eik1 = web.find_element_by_xpath(eik_u)
            eik1.send_keys(eik_enter)
            sleep(4)
            eik2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(4)
            eik2.click()
            sleep(3)

            if today != card_date:
                cal = web.find_element_by_xpath(
                    '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[2]/app-fi-input-date/div/div/div[2]/button[1]')
                cal.click()
                sleep(2)
                if month_move >= 1:
                    sleep(3)
                    cal_mv = web.find_element_by_xpath(
                        '/html/body/ngb-datepicker/div[1]/ngb-datepicker-navigation/div[1]/button')
                    for x in range(month_move):
                        cal_mv.click()
                        sleep(1)
                # elif month_move >= 1:
                #     label['text'] = "Невалидна дата\n(датата е от преди повече от 1 месец)"
                #     web.close()
                #     return 0
                data_pos = web.find_element_by_css_selector(f"div[aria-label='{card_date}']")
                data_pos.click()
                sleep(2)

            lice = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[5]/app-fi-input-text/div/div/input')
            try:
                lice.send_keys(names_eik[eik_enter])
            except KeyError:
                # label['text'] = "В names.txt файла липсва \nимето срещу това ЕИК"
                tk.messagebox.showinfo(title=None,
                                       message=f'В names1.xlsx файла липсва \nимето срещу това ЕИК:\n{eik_enter}')
                return 0

            koli4estvo = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[6]/app-fi-input-float/div/div/input')
            koli4estvo.send_keys(amnt)
            sleep(2)
            koli4estvo.send_keys(Keys.PAGE_DOWN)

            proizhod = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[7]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            pr = web.find_element_by_xpath(proizhod)
            pr.send_keys('извън')
            sleep(3)
            pr2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            pr2.click()
            sleep(2)
            if code_enter == '16 03 05*' or code_enter == '16 03 03*':
                osn_text = ' '
                op_text = ' '
            elif code_enter == '15 01 10*':
                osn_text = 'От дейността на фирмата'
                op_text = 'Празни опаковки'
            elif code_enter == '20 01 21*':
                osn_text = 'От периодична промяна на луминисцентните лампи'
                op_text = 'Луминисцентни лампи'
            else:
                osn_text = 'Здравеопазване'
                op_text = 'Клинични отпадъци'
            osnovanie = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[8]/app-fi-input-long-text/div/div/textarea')
            osnovanie.send_keys(osn_text)
            sleep(2)
            opisanie = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[9]/app-fi-input-long-text/div/div/textarea')
            opisanie.send_keys(op_text)
            zapis = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[10]/div[2]/button')
            sleep(5)
            if not zapis.is_enabled():
                tk.messagebox.showerror(title='Грешка',
                                        message=f'Възникна проблем при \nподаването на картата')
                return 0
            sleep(5)
            zapis.click()
            sleep(3)
            df.at[index, "Submitted"] = 'Podadeno'
            df.to_excel('suhranenie.xlsx', index=False)
            web.find_element_by_tag_name('body').click()
            web.find_element_by_tag_name('body').send_keys(Keys.HOME)
            sleep(2)
            ### Mahni Posle
        except Exception as e:
            tk.messagebox.showerror(title='Грешка',
                                    message=f'Възникна проблем при \nподаването на картата')
            print(e)
            return 0
    if k==1:
        web.close()
        # df.to_excel('suhranenie.xlsx', index=False)
        label['text'] = 'Всичкo e подаденo и сумиранo (Събиране)'
        tk.messagebox.showinfo(title=None,
                               message='Данните бяха подадени успешно (Събиране)')
        return None
    else:
        try:
            sleep(2)
            nazad = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/div[1]/button')
            nazad.click()
        except:
            pass
    return None

def unishtozhenie_ml(k=1):
    r1 = twopointeight()
    if r1 == 'problem':
        return 0
    df = pd.read_excel("unishtozhenie_ML.xlsx",
                       dtype={'EIK': str, 'Submitted': str, 'Submitted_Tretiran': str})
    nepod = df.loc[df["Submitted_Tretiran"].isnull()]
    if nepod.empty:
        label['text'] = 'Всички карти вече са подадени и сумирани'
        tk.messagebox.showinfo(title=None,
                               message='Няма неподадени карти')
        return None
    global web
    print('unishto')
    try:
        k = web
    except:
        web = webdriver.Chrome()
        web.maximize_window()

        url = "https://nwms.eea.government.bg/app/base/home"
        web.get(url)
        sleep(2)
        vhod = web.find_element_by_xpath(
            "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
        vhod.click()
        sleep(2)
        el_akt = web.find_element_by_xpath("/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
        el_akt.click()
        sleep(15)
        firm = None
        i = 0
        while firm == None and i < 6:
            try:
                firm = web.find_element_by_xpath(
                    '/html/body/app-root/app-auth-main-page/app-login-page/app-organization-selector/div/div/div[2]/div/button')
            except:
                sleep(2)
                web.close()
                web.get(url)
                sleep(2)
                vhod = web.find_element_by_xpath(
                    "/html/body/app-root/app-home-main-page/app-home-page/div/div[2]/div/div[2]/div/div[1]/div[1]")
                vhod.click()
                sleep(2)
                el_akt = web.find_element_by_xpath(
                    "/html/body/div[3]/div/div[1]/div/div/div/div/div[2]/ul/li/button")
                el_akt.click()
                sleep(15)

        firm.click()
        sleep(5)
        otcheti = web.find_element_by_xpath(
            '/html/body/app-root/app-messages-main-page/div/div[2]/app-subheader/nav/div/ul/li[2]/a')
        otcheti.click()
        sleep(10)
        otchetni_knigi = web.find_element_by_xpath(
            '/html/body/app-root/app-reports-main-page/div/div[1]/app-sidebar-menu/nav/app-tree-view/ul/li[1]/button/div/div[2]')
        otchetni_knigi.click()
        sleep(5)
        web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
        sleep(2)
    try:
        tursene0 = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/div[2]/div/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
    except:
        sleep(1)
        tursene0= web.find_element_by_xpath('/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/div[2]/div/app-fi-select-dropdown/div/div/ng-select/div/div/div[3]/input')
    tursene0.send_keys('ополз')
    sleep(2)
    transp0 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
    transp0.click()
    sleep(2)
    tursene = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-search-and-navigation-bar/ul/li[3]/button')
    tursene.click()
    sleep(2)
    opolz = web.find_element_by_xpath(
        '/html/body/app-root/app-reports-main-page/div/div[2]/app-reports-wrapper-page/div/ngb-tabset/div/div/div/app-reports/app-data-viewer-as-table/div[2]/table/tbody/tr[1]/td/div/button/i')
    opolz.click()
    sleep(10)

    for index, row in nepod.iterrows():
        code_enter = nepod.at[index, "Code"]
        eik_enter = str(nepod.at[index, "EIK"])
        amnt = str(nepod.at[index, "Koli4estvo_obshto"])
        card_date_0 = nepod.at[index, "Data"].strftime('%d/%m/%Y')
        card_date = datetime.strptime(card_date_0, '%d/%m/%Y')
        card_date = card_date.strftime('%d-%m-%Y')
        try:
            month_move = diff_month(card_date)
            card_date = date_str(card_date)
        except IndexError:
            label['text'] = "Невалиден формат на датата\n (форматът трябва изглежда така: 09/02/2021)"
            label.config(font='Roboto 16 italic bold')
            return 0
        try:
            sleep(2)
            polu4en_otpaduk = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/ul/li[2]/a')
            polu4en_otpaduk.click()
            web.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
            sleep(3)
            x = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[3]/app-fi-select-dropdown/div/div/ng-select'
            inputc = web.find_element_by_xpath(x)
            inputc.click()
            sleep(2)
            uridi4esko_lice = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div[1]')
            sleep(2)
            uridi4esko_lice.click()
            sleep(2)
            code_inp = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[1]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            code1 = web.find_element_by_xpath(code_inp)
            code1.send_keys(code_enter)
            sleep(2)
            code2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            code2.click()

            eik_u = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            eik1 = web.find_element_by_xpath(eik_u)
            eik1.send_keys(eik_enter)
            sleep(4)
            eik2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(4)
            eik2.click()
            sleep(3)

            if today != card_date:
                cal = web.find_element_by_xpath(
                    '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[2]/app-fi-input-date/div/div/div[2]/button[1]')
                cal.click()
                sleep(2)
                if month_move >= 1:
                    sleep(3)
                    cal_mv = web.find_element_by_xpath(
                        '/html/body/ngb-datepicker/div[1]/ngb-datepicker-navigation/div[1]/button')
                    for x in range(month_move):
                        cal_mv.click()
                        sleep(1)
                # elif month_move >= 1:
                #     label['text'] = "Невалидна дата\n(датата е от преди повече от 1 месец)"
                #     web.close()
                #     return 0
                data_pos = web.find_element_by_css_selector(f"div[aria-label='{card_date}']")
                data_pos.click()
                sleep(2)

            lice = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[5]/app-fi-input-text/div/div/input')
            try:
                lice.send_keys(names_eik[eik_enter])
            except KeyError:
                # label['text'] = "В names.txt файла липсва \nимето срещу това ЕИК"
                tk.messagebox.showerror(title='Грешка',
                                        message=f'В names.txt файла липсва \nимето срещу това ЕИК\n{eik_enter}')
                return 0

            koli4estvo = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[8]/app-fi-input-float/div/div/input')
            koli4estvo.send_keys(amnt)
            sleep(2)
            koli4estvo.send_keys(Keys.PAGE_DOWN)

            proizhod = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[6]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            pr = web.find_element_by_xpath(proizhod)
            pr.send_keys('извън')
            sleep(3)
            pr2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            pr2.click()
            sleep(2)
            if code_enter == '16 03 05*' or code_enter == '16 03 03*':
                osn_text = ' '
                op_text = ' '
            elif code_enter == '15 01 10*':
                osn_text = 'От дейността на фирмата'
                op_text = 'Празни опаковки'
            elif code_enter == '20 01 21*':
                osn_text = 'От периодична промяна на луминисцентните лампи'
                op_text = 'Луминисцентни лампи'
            else:
                osn_text = 'Здравеопазване'
                op_text = 'Клинични отпадъци'
            osnovanie = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[7]/app-fi-input-long-text/div/div/textarea')
            osnovanie.send_keys(osn_text)
            sleep(8)
            zapis = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[9]/div[2]/button')
            sleep(5)
            if not zapis.is_enabled():
                tk.messagebox.showerror(title='Грешка',
                                        message=f'Възникна проблем при \nподаването на картата')
                return 0
            sleep(3)
            zapis.click()
            df.at[index, "Submitted"] = 'Podadeno'
            df.to_excel('unishtozhenie_ML.xlsx', index=False)
            sleep(5)
            tretiran = web.find_element_by_xpath('/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/ul/li[3]/a')
            tretiran.click()
            sleep(5)
            code_inp = '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[1]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input'
            code1 = web.find_element_by_xpath(code_inp)
            code1.send_keys(code_enter)
            sleep(2)
            code2 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            sleep(2)
            code2.click()
            sleep(3)

            if today != card_date:
                cal = web.find_element_by_xpath(
                    '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[2]/app-fi-input-date/div/div/div[2]/button[1]')
                cal.click()
                sleep(2)
                if month_move >= 1:
                    sleep(3)
                    cal_mv = web.find_element_by_xpath(
                        '/html/body/ngb-datepicker/div[1]/ngb-datepicker-navigation/div[1]/button')
                    for x in range(month_move):
                        cal_mv.click()
                        sleep(1)
                # elif month_move >= 1:
                #     label['text'] = "Невалидна дата\n(датата е от преди повече от 1 месец)"
                #     web.close()
                #     return 0
                data_pos = web.find_element_by_css_selector(f"div[aria-label='{card_date}']")
                data_pos.click()
                sleep(2)
            sleep(2)
            deinost_op = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[3]/app-fi-select-dropdown/div/div/ng-select/div/div/div[2]/input')
            deinost_op.click()
            sleep(2)
            deinost_op09 = web.find_element_by_xpath('/html/body/ng-dropdown-panel/div[2]/div[2]/div')
            deinost_op09.click()
            sleep(2)
            deinost_op2 = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[4]/app-fi-input-long-text/div/div/textarea')
            deinost_op2.send_keys('Физико-химично третиране, чрез автоклавиране')
            kol = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[5]/app-fi-input-float/div/div/input')
            kol.send_keys(amnt)
            sleep(8)
            zapis = web.find_element_by_xpath(
                '/html/body/ngb-modal-window/div/div/div[2]/ngb-tabset/div/div/form/div[6]/div[2]/button')
            sleep(5)
            if not zapis.is_enabled():
                tk.messagebox.showerror(title='Грешка',
                                        message=f'Възникна проблем при \nподаването на картата')
                return 0
            sleep(5)
            zapis.click()
            sleep(2)
            df.at[index, "Submitted_Tretiran"] = 'Podadeno'
            df.to_excel('unishtozhenie_ML.xlsx', index=False)


        except Exception as e:
            tk.messagebox.showerror(title='Грешка',
                                    message=f'Възникна проблем при \nподаването на картата')
            print(e)
            return 0
    web.close()
    # df.to_excel('unishtozhenie_ML.xlsx', index=False)
    destruction_filing()
    label['text'] = 'Всичкo e подаденo\nуспешно'
    if k==1:
        tk.messagebox.showinfo(title=None,
                           message='Данните бяха подадени успешно (Унищожение)')
    else:
        tk.messagebox.showinfo(title=None,
                               message='Данните бяха подадени успешно (Всичко)')
    return None

def new_arch():
    df = pd.read_excel("individual_info.xlsx",
                       dtype={"EIK_Tovarodatel": str, "EIK_Polu4atel": str, "Added_to_Sum": str, 'Submitted': str})
    df_suh = pd.read_excel("suhranenie.xlsx",
                           dtype={"EIK": str, "Added_to_Sum": str, 'Submitted': str})
    not_submitted = df.loc[df["Added_to_Sum"].isnull()].loc[df["EIK_Polu4atel"] != '112106418']

    list_dates = not_submitted["Data"].tolist()
    list_dates = list(set(list_dates))

    for date in list_dates:
        date_f = not_submitted.loc[not_submitted["Data"] == date]
        for code in codes_list:
            to_enter = date_f.loc[date_f["Code"] == code]
            if not to_enter.empty:
                sum_to_add = to_enter.sum(axis=0)["Koli4estvo"]
                fin_df = df_suh.loc[df_suh["Data"] == date].loc[df_suh["Code"] == code]
                indx = df_suh.loc[df_suh["Data"] == date].loc[df_suh["Code"] == code].index
                if not fin_df.empty:
                    final_sum = fin_df.sum(axis=0)["Koli4estvo_obshto"] + sum_to_add
                    df_suh.at[indx, "Koli4estvo_obshto"] = final_sum
                else:
                    new_row = {'EIK': '112106418', 'Koli4estvo_obshto': sum_to_add, 'Code': code, 'Data': date,
                               'Added_to_Sum': '', "Submitted": ""}
                    df_suh = df_suh.append(new_row, ignore_index=True)

    df.loc[df['Added_to_Sum'].isnull()].loc[df['EIK_Polu4atel'] != '112106418', 'Added_to_Sum'] = 'Da'
    df.to_excel('individual_info.xlsx', index=False)
    df_suh.to_excel('suhranenie.xlsx', index=False)

    df = pd.read_excel("individual_info.xlsx",
                       dtype={"EIK_Tovarodatel": str, "EIK_Polu4atel": str, "Added_to_Sum": str, 'Submitted': str})
    df_suh = pd.read_excel("unishtozhenie_ML.xlsx",
                           dtype={"EIK": str, "Added_to_Sum": str, 'Submitted': str})
    not_submitted = df.loc[df["Added_to_Sum"].isnull()].loc[df["EIK_Polu4atel"] == '112106418']

    list_dates = not_submitted["Data"].tolist()
    list_dates = list(set(list_dates))

    for date in list_dates:
        date_f = not_submitted.loc[not_submitted["Data"] == date]
        for code in codes_list:
            to_enter = date_f.loc[date_f["Code"] == code]
            if not to_enter.empty:
                sum_to_add = to_enter.sum(axis=0)["Koli4estvo"]
                fin_df = df_suh.loc[df_suh["Data"] == date].loc[df_suh["Code"] == code]
                indx = df_suh.loc[df_suh["Data"] == date].loc[df_suh["Code"] == code].index
                if not fin_df.empty:
                    final_sum = fin_df.sum(axis=0)["Koli4estvo_obshto"] + sum_to_add
                    df_suh.at[indx, "Koli4estvo_obshto"] = final_sum
                else:
                    new_row = {'EIK': '112106418', 'Koli4estvo_obshto': sum_to_add, 'Code': code, 'Data': date,
                               "Added_to_Sum": "", "Submitted": ""}
                    df_suh = df_suh.append(new_row, ignore_index=True)

    df.loc[df['Added_to_Sum'].isnull()].loc[df['EIK_Polu4atel'] == '112106418', 'Added_to_Sum'] = 'Da'
    df.loc[df["Added_to_Sum"].isnull(), 'Added_to_Sum'] = "Da"
    df.to_excel('individual_info.xlsx', index=False)
    df_suh.to_excel('unishtozhenie_ML.xlsx', index=False)

    label['text'] = 'Информацията за съхранение\nе ъпдейтната'
    return None

def all_ffs():
    import_stuff(k=2)
    storage_import(k=2)
    unishtozhenie_ml(k=2)

def arch_both():
    new_arch()
    destruction_filing()

root = tk.Tk()

wght_factor = root.winfo_screenwidth() / 3840
height_factor = root.winfo_screenheight() / 2160
HEIGHT = 1500 * height_factor
WIDTH = 1700 * wght_factor

root.title('МЛ-България')

ttk.Style().configure("TButton", padding=6, relief="flat",
                      background="#ccc", font=('Roboto', '15', 'bold'), justify='center')

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()

frame = tk.Frame(root, bg="#ccc")
frame.place(relx=0.5, rely=0.02, relwidth=0.9, relheight=0.15, anchor='n')

button = ttk.Button(frame, text='Подай\n (Индивидуални)',
                    command=lambda: import_stuff())
button.place(relx=0, relheight=1, relwidth=0.35)

button_arch = ttk.Button(frame, text='Обнови сумите\n (Архивирай)',
                         command=lambda: arch_both())
button_arch.place(relx=0.35, relheight=1, relwidth=0.30)

button_sum = ttk.Button(frame, text='Подай\n (Съхранение)',
                        command=lambda: storage_import())
button_sum.place(relx=0.65, relheight=1, relwidth=0.35)

lower_frame = tk.Frame(root, bd=10)
lower_frame.place(relx=0.5, rely=0.18, relwidth=0.885, relheight=0.19, anchor='n')

label = tk.Label(lower_frame, font='Roboto 18 italic bold', anchor='center', justify='center', bd=4, bg='white')
label.config(bd=8, relief='ridge', fg='#19a3e3')
label.place(relwidth=1, relheight=1)

fin_frame = tk.Frame(root, bd=10, highlightthickness=2)
fin_frame.config(highlightbackground='grey', highlightcolor='grey')
fin_frame.place(relx=0.5, rely=0.39, relwidth=0.885, relheight=0.58, anchor='n')

T = tk.Text(fin_frame, height=2, width=30)
T.insert(tk.END, "Товарополучател: ")
T.config(font='Roboto, 17', bd=0, bg='#f2f2f2', state='disabled')
T.place(relx=0.02, relwidth=0.4, rely=0.025, relheight=0.15)

variable = tk.StringVar(fin_frame)
variable.set("ПУДООС")
entry_EIK = tk.OptionMenu(fin_frame, variable, "Екосейф", "ПУДООС")
entry_EIK.config(font='Roboto, 16', anchor='c')
entry_EIK.place(relx=0.5, relwidth=0.5, rely=0, relheight=0.15)
entry_EIK["menu"].config(font='Roboto, 16')

Cd = tk.Text(fin_frame, height=2, width=30)
Cd.insert(tk.END, "Код: ")
Cd.config(font='Roboto, 17', bd=0, bg='#f2f2f2', state='disabled')
Cd.place(relx=0.02, relwidth=0.4, rely=0.225, relheight=0.15)

variable_cd = tk.StringVar(fin_frame)
variable_cd.set(codes_list[0])  # default value

codes_entry = tk.OptionMenu(fin_frame, variable_cd, *codes_list)
codes_entry.config(font='Roboto, 16', anchor='c')
codes_entry['menu'].config(font='Roboto, 16')
codes_entry.place(relx=0.5, relwidth=0.5, rely=0.2, relheight=0.15)

Dt = tk.Text(fin_frame, height=2, width=30)
Dt.insert(tk.END, "Дата: ")
Dt.config(font='Roboto, 17', bd=0, bg='#f2f2f2', state='disabled')
Dt.place(relx=0.02, relwidth=0.13, rely=0.43, relheight=0.15)

cal = DateEntry(fin_frame, date_pattern='d-m-yyyy')
cal.place(relx=0.5, relwidth=0.5, rely=0.4, relheight=0.15)
cal.config(font='Roboto, 17', justify='center')

Kol = tk.Text(fin_frame, height=2, width=30)
Kol.insert(tk.END, "Количество (т.) : ")
Kol.config(font='Roboto, 17', bd=0, bg='#f2f2f2', state='disabled')
Kol.place(relx=0.02, relwidth=0.4, rely=0.625, relheight=0.15)

entry_kol = tk.Entry(fin_frame, font='Roboto, 20', justify='center')  # .configure(state='disabled'/'normal)
entry_kol.place(relx=0.5, relwidth=0.5, rely=0.6, relheight=0.15)

s = ttk.Style()
s.configure('my.TButton', font=('Roboto', 13, 'bold'), justify='center')

button_sum_ml = ttk.Button(fin_frame, style='my.TButton', text='Унищожение\n(МЛ)',
                           command=lambda: unishtozhenie_ml())

button_sum_ml.place(relx=0, rely=0.8, relheight=0.2, relwidth=0.24)

button_sum = ttk.Button(fin_frame, style='my.TButton', text='Унищожение\n(други)',
                        command=lambda: unishtozhenie(variable.get(), variable_cd.get(), entry_kol.get()))

button_sum.place(relx=0.25, rely=0.8, relheight=0.2, relwidth=0.24)

button_all = ttk.Button(fin_frame, text='Пусни\nвсичко',
                        command=lambda: all_ffs())

button_all.place(relx=0.5, rely=0.8, relheight=0.2, relwidth=0.5)

root.mainloop()
