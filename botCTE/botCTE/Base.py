import requests
import threading
import numpy as np
import pandas as pd
import datetime as dt
from json import loads
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkcalendar import DateEntry
from ttkthemes import ThemedStyle
import bot
import pywintypes
import win32api
from win32ctypes.pywin32 import pywintypes
from functions import *


def cte_list():
    def get_address_name(add):
        add_lenght = len(add)
        count = 0
        add_total = ''
        for i in range(add_lenght):
            add_name = address.loc[address['id'] == add[i], 'trading_name'].values.item()
            add_branch = address.loc[address['id'] == add[i], 'branch'].values.item()
            count += 1
            if count != add_lenght:
                space = "\n"
            else:
                space = ''
            if add_lenght == 1:
                listing = str(add_name)
            else:
                listing = str(count) + '. ' + str(add_name)
            add_total += listing + ' - ' + add_branch + space
        return add_total

    def get_address_meta(add):
        add_lenght = len(add)
        count = 0
        add_total = ''
        for i in range(add_lenght):
            add_street = address.loc[address['id'] == add[i], 'street'].values.item()
            add_num = address.loc[address['id'] == add[i], 'number'].values.item()
            add_dist = address.loc[address['id'] == add[i], 'neighborhood'].values.item()
            add_city = address.loc[address['id'] == add[i], 'cityIDAddress.name'].values.item()
            add_state = address.loc[address['id'] == add[i], 'state'].values.item()
            add_cep = address.loc[address['id'] == add[i], 'cep'].values.item()
            count += 1
            if count != add_lenght:
                space = "\n"
            else:
                space = ''
            if add_lenght == 1:
                listing = ''
            else:
                listing = str(count) + '. '
            add_total += listing + add_street.title() + ', ' + add_num + ' - ' + add_dist.title() + ', ' \
                         + add_city.title() + ' - ' + add_state + ', ' + add_cep + space
        return add_total

    def get_address_cnpj_listed(add_list):
        add_lenght = len(add_list)
        count = 0
        add_total = ''
        for add in add_list:
            cnpj = address.loc[address['id'] == add, 'cnpj_cpf'].values.item()
            count += 1
            if count != add_lenght:
                space = "\n"
            else:
                space = ''
            if add_lenght == 1:
                listing = ''
            else:
                listing = str(count) + '. '
            add_total += listing + cnpj + space
        return add_total

    def get_address_cnpj(add_list):
        _cnpj_list = []
        for add in add_list:
            cnpj = address.loc[address['id'] == add, 'cnpj_cpf'].values.item()
            _cnpj_list.append(cnpj)
        return _cnpj_list

    def get_address_city(add):
        add_lenght = len(add)
        add_total = []
        count = 0
        cities = ''
        for i in range(add_lenght):
            add_city = address.loc[address['id'] == add[i], 'cityIDAddress.name'].values.item()
            add_total.append(add_city)
        add_total = np.unique(add_total)
        for p in add_total:
            if count == 0:
                cities = p
            elif count == len(add_total) - 1:
                cities += ', ' + p
            else:
                cities += ', ' + p
            count += 1
        return cities

    def get_address_city_listed(add_list):
        _cities_list = []
        for add in add_list:
            add_city = address.loc[address['id'] == add, 'cityIDAddress.name'].values.item()
            _cities_list.append(add_city)
        cities_list = np.unique(_cities_list)
        return cities_list

    def get_collector(col):
        col_name = collector.loc[collector['id'] == col, 'trading_name'].values.item()
        return col_name

    now = dt.datetime.now()
    now_date = dt.datetime.strftime(cal.get_date(), '%d-%m-%Y')
    now = dt.datetime.strftime(now, "%H-%M")

    di = dt.datetime.strftime(cal.get_date(), '%d/%m/%Y')
    df = dt.datetime.strftime(cal.get_date(), '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    di_temp = di_dt - dt.timedelta(days=5)
    df_temp = df_dt + dt.timedelta(days=5)

    di = dt.datetime.strftime(di_temp, '%d/%m/%Y')
    df = dt.datetime.strftime(df_temp, '%d/%m/%Y')

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    collector = r.request_public('https://transportebiologico.com.br/api/public/collector')
    services_ongoing = r.request_public('https://transportebiologico.com.br/api/public/service')
    services_finalized = r.post_public(
        f'https://transportebiologico.com.br/api/public/service/finalized/?startFilter={di}&endFilter={df}')

    sv = pd.concat([services_ongoing, services_finalized], ignore_index=True)
    sv = sv.loc[sv['is_business'] == False]

    df_dt += dt.timedelta(days=1)

    sv['collectDateTime'] = pd.to_datetime(sv['serviceIDRequested.collect_date']) - dt.timedelta(hours=3)
    sv['collectDateTime'] = sv['collectDateTime'].dt.tz_localize(None)

    sv.drop(sv[sv['collectDateTime'] < di_dt].index, inplace=True)
    sv.drop(sv[sv['collectDateTime'] > df_dt].index, inplace=True)
    sv.to_excel('ServicesAPI.xlsx', index=True)
    sv = pd.concat([sv[sv['cte_loglife'].isnull()], sv[sv['cte_loglife'] == 'nan']], ignore_index=True)
    sv.drop(sv[sv['customerIDService.emission_type'] == 'NF'].index, inplace=True)
    sv.drop(sv[sv['customerIDService.trading_firstname'] == 'LOGLIFE'].index, inplace=True)
    sv.drop(sv[sv['serviceIDRequested.budgetIDService.price'] == 0].index, inplace=True)
    sv['origCityList'] = sv['serviceIDRequested.source_address_id'].map(get_address_city_listed)
    sv['destCityList'] = sv['serviceIDRequested.destination_address_id'].map(get_address_city_listed)
    sv['origCity'] = sv['serviceIDRequested.source_address_id'].map(get_address_city)
    sv['destCity'] = sv['serviceIDRequested.destination_address_id'].map(get_address_city)

    sv.drop(sv[
                (sv['origCityList'].str.len() == 1) &
                (sv['destCityList'].str.len() == 1) &
                (sv['origCity'] == sv['destCity'])
                ].index, inplace=True)

    sv.sort_values(
        by="protocol", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last'
    )

    report = pd.DataFrame(columns=[])

    report['PROTOCOLO'] = sv['protocol']
    report['CLIENTE'] = sv['customerIDService.trading_firstname']
    report['ETAPA'] = np.select(
        condlist=[
            sv['step'] == 'availableService',
            sv['step'] == 'toAllocateService',
            sv['step'] == 'toDeliveryService',
            sv['step'] == 'deliveringService',
            sv['step'] == 'toLandingService',
            sv['step'] == 'landingService',
            sv['step'] == 'toBoardValidate',
            sv['step'] == 'toCollectService',
            sv['step'] == 'collectingService',
            sv['step'] == 'toBoardService',
            sv['step'] == 'boardingService',
            sv['step'] == 'finishedService'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO', 'AGUARDANDO ALOCAÇÃO', 'EM ROTA DE ENTREGA', 'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA', 'DESEMBARCANDO', 'VALIDAR EMBARQUE', 'AGENDADO', 'COLETANDO',
            'EM ROTA DE EMBARQUE', 'EMBARCANDO SERVIÇO', 'FINALIZADO'],
        default=0
    )
    report['DATA COLETA'] = sv['collectDateTime'].dt.strftime(date_format='%d/%m/%Y')
    report['PREÇO TRANSPORTE'] = sv['serviceIDRequested.budgetIDService.price']
    report['PREÇO KG EXTRA'] = sv['serviceIDRequested.budgetIDService.price_kg_extra']
    report['NOME REMETENTE'] = sv['serviceIDRequested.source_address_id'].map(get_address_name)
    report['CIDADE ORIGEM'] = sv['origCity']
    report['ENDEREÇO REMETENTE'] = sv['serviceIDRequested.source_address_id'].map(get_address_meta)
    report['CNPJ/CPF REMETENTE'] = sv['serviceIDRequested.source_address_id'].map(get_address_cnpj_listed)
    report['COLETADOR ORIGEM'] = sv['serviceIDRequested.source_collector_id'].map(get_collector)
    report['NOME DESTINATÁRIO'] = sv['serviceIDRequested.destination_address_id'].map(get_address_name)
    report['CIDADE DESTINO'] = sv['destCity']
    report['ENDEREÇO DESTINATÁRIO'] = sv['serviceIDRequested.destination_address_id'].map(get_address_meta)
    report['CNPJ/CPF DESTINATÁRIO'] = sv['serviceIDRequested.destination_address_id'].map(get_address_cnpj_listed)
    report['COLETADOR DESTINO'] = sv['serviceIDRequested.destination_collector_id'].map(get_collector)

    cte_path = folderpath.get()
    cte_path = cte_path.replace('/', '\\')

    report.to_excel(
        f'{cte_path}\\Lista CTE-{now_date}-{now}.xlsx',
        index=False
    )

    print("Relatório exportado!")

    report_date = dt.datetime.strftime(dt.datetime.now(), '%d/%m/%Y')

    csv_report = pd.DataFrame(columns=['Protocolo', 'CTE Loglife', 'Data Emissão CTE'])
    csv_report['Protocolo'] = sv['protocol']
    csv_report['Data Emissão CTE'] = report_date

    bot_cte = bot.Bot()

    # bot_cte.open_bsoft(path=filename.get(), login=login.get(), password=password.get())
    current_row = 0

    for protocol in sv['protocol']:

        excel_file = f'{cte_path}\\Lista CTE-{now_date}-{now}.xlsx'
        csv_file = f'{cte_path}\\Upload-{now_date}-{now}.csv'

        tomador_cnpj = sv.loc[sv['protocol'] == protocol, 'customerIDService.cnpj_cpf'].values.item()
        source_add = sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.source_address_id'].values.item()
        destination_add = sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.destination_address_id'].values.item()
        valor = str(sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.budgetIDService.price'].values.item())
        uf1 = address.loc[address['id'] == source_add[0], 'state'].values.item()
        uf2 = address.loc[address['id'] == destination_add[0], 'state'].values.item()
        cnpj_remetente = get_address_cnpj(source_add)
        cnpj_destinatario = get_address_cnpj(destination_add)
        uf_rem = uf_base.loc[uf_base['Estado'] == uf1, 'UF'].values.item()
        uf_dest = uf_base.loc[uf_base['Estado'] == uf2, 'UF'].values.item()
        icms_obs = uf_base.loc[uf_base['Estado'] == uf_rem, 'Info'].values.item()
        aliq = aliquota_base.loc[aliquota_base['UF'] == uf_rem, uf_dest].values.item()
        obs_text = f'Protocolo {protocol} - {icms_obs}'

        if uf_rem != "MG":
            aliq_text = float(aliq)*float(valor)*0.008
            aliq_text = "{:0.2f}".format(aliq_text)
            obs_text = obs_text.replace('#', aliq_text)
            aliq = "0"

        if tomador_cnpj in cnpj_remetente:
            tomador = "Remetente"
            cnpj_remetente = [tomador_cnpj]
        elif tomador_cnpj in cnpj_destinatario:
            tomador = "Destinatário"
            cnpj_destinatario = [tomador_cnpj]
        else:
            tomador = "Outro"

        bot_cte.action(cnpj_sender=cnpj_remetente,
                       cnpj_receiver=cnpj_destinatario,
                       payer=tomador,
                       payer_cnpj=tomador_cnpj)
        bot_cte.part3_normal()
        bot_cte.part4(
            tax=str(aliq),
            uf=uf_rem,
            icms_text=obs_text,
            price=valor
        )

        cte_llm = int(bot_cte.get_clipboard())
        report.at[report.index[current_row], 'CTE LOGLIFE'] = cte_llm
        csv_report.at[csv_report.index[current_row], 'Protocolo'] = protocol
        csv_report.at[csv_report.index[current_row], 'CTE Loglife'] = cte_llm

        report.to_excel(
            excel_file,
            index=False
        )

        csv_report = csv_report.astype(str)
        csv_report = csv_report.replace(to_replace="\.0+$", value="", regex=True)

        csv_report.to_csv(
            csv_file,
            index=False,
        )

        post_private('https://transportebiologico.com.br/api/uploads/cte-loglife', csv_file)

        current_row += 1


def cte_complimentary(unique=False):

    def get_address_name(add):
        add_lenght = len(add)
        count = 0
        add_total = ''
        for i in range(add_lenght):
            add_name = address.loc[address['id'] == add[i], 'trading_name'].values.item()
            add_branch = address.loc[address['id'] == add[i], 'branch'].values.item()
            count += 1
            if count != add_lenght:
                space = "\n"
            else:
                space = ''
            if add_lenght == 1:
                listing = str(add_name)
            else:
                listing = str(count) + '. ' + str(add_name)
            add_total += listing + ' - ' + add_branch + space
        return add_total

    # def get_address_cnpj(add_list):
    #     _cnpj_list = []
    #     for add in add_list:
    #         cnpj = address.loc[address['id'] == add, 'cnpj_cpf'].values.item()
    #         _cnpj_list.append(cnpj)
    #     return _cnpj_list

    def get_address_city(add):
        add_lenght = len(add)
        add_total = []
        count = 0
        cities = ''
        for i in range(add_lenght):
            add_city = address.loc[address['id'] == add[i], 'cityIDAddress.name'].values.item()
            add_total.append(add_city)
        add_total = np.unique(add_total)
        for w in add_total:
            if count == 0:
                cities = w
            elif count == len(add_total) - 1:
                cities += ', ' + w
            else:
                cities += ', ' + w
            count += 1
        return cities

    def get_address_city_listed(add_list):
        _cities_list = []
        for add in add_list:
            add_city = address.loc[address['id'] == add, 'cityIDAddress.name'].values.item()
            _cities_list.append(add_city)
        cities_list = np.unique(_cities_list)
        return cities_list

    def materials_value(mat_list, price):
        price_list = [2, 5, 10, 10, 12.5, 12.5, 0, price, 15]
        total_price = 0
        for material, mat_price in zip(mat_list, price_list):
            total_price += material * mat_price
        return total_price

    def materials_description(mat_list):
        mat_description = ""
        count = 0
        description_list = ['Embalagem secundária', 'Gelox', 'Isopor 3L', 'Terciária 3L', 'Isopor 7L', 'Terciária 8L',
                            'Caixa térmica', 'Gelo Seco', 'Pote 1L']
        for material in mat_list:
            if material == 0:
                count += 1
                pass
            else:
                mat_description += f'{material} {description_list[count]}\n'
                count += 1
        mat_description = mat_description.strip("\n")
        return mat_description

    def get_material_price(provider_id):
        try:
            dry_ice_price = provider.loc[provider['id'] == provider_id, 'material_price'].values.item()
        except ValueError:
            dry_ice_price = 0
        return dry_ice_price

    now = dt.datetime.now()
    now_date = dt.datetime.strftime(cal.get_date(), '%d-%m-%Y')
    now = dt.datetime.strftime(now, "%H-%M")

    di = dt.datetime.strftime(cal1.get_date(), '%d/%m/%Y')
    df = dt.datetime.strftime(cal2.get_date(), '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    di_temp = di_dt - dt.timedelta(days=5)
    df_temp = df_dt + dt.timedelta(days=5)

    di = dt.datetime.strftime(di_temp, '%d/%m/%Y')
    df = dt.datetime.strftime(df_temp, '%d/%m/%Y')

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    services_ongoing = r.request_public('https://transportebiologico.com.br/api/public/service')
    services_finalized = r.post_public(
        f'https://transportebiologico.com.br/api/public/service/finalized/?startFilter={di}&endFilter={df}')

    sv = pd.concat([services_ongoing, services_finalized], ignore_index=True)

    sv.to_excel('Debug.xlsx', index=False)

    provider = r.request_private('https://transportebiologico.com.br/api/provider')

    df_dt += dt.timedelta(days=1)

    sv['collectDateTime'] = pd.to_datetime(sv['serviceIDRequested.collect_date']) - dt.timedelta(hours=3)
    sv['collectDateTime'] = sv['collectDateTime'].dt.tz_localize(None)

    sv.drop(sv[sv['collectDateTime'] < di_dt].index, inplace=True)
    sv.drop(sv[sv['collectDateTime'] > df_dt].index, inplace=True)
    if unique is False:
        sv = pd.concat([sv[sv['cte_complementary'].isnull()], sv[sv['cte_complementary'] == 'nan']], ignore_index=True)
    sv.drop(sv[sv['customerIDService.trading_firstname'] == "BIOGEN"].index, inplace=True)
    sv.drop(sv[sv['customerIDService.trading_firstname'] == "HPV - HOSPITAL MOINHOS DE VENTO"].index, inplace=True)
    sv = sv.loc[(sv['cte_loglife'].notnull()) &
                (sv['cte_loglife'] != 'nan') &
                (sv['step'] != 'toCollectService') &
                (sv['step'] != 'collectingService') &
                (sv['step'] != 'toBoardService') &
                (sv['step'] != 'boardingService')]
    # # Client list FILTER START
    # sv = sv.loc[
    #     (sv['customerIDService.trading_firstname'] == "CERBA-LCA") |
    #     (sv['customerIDService.trading_firstname'] == "PROVET") |
    #     (sv['customerIDService.trading_firstname'] == "FLOW") |
    #     (sv['customerIDService.trading_firstname'] == "ALCHEMYPET MEDICINA DIAGNÓSTICA VETERINÁRIA LTDA.") |
    #     (sv['customerIDService.trading_firstname'] == "LEMOS LABORATORIOS") |
    #     (sv['customerIDService.trading_firstname'] == "LABORATÓRIO KTZ") |
    #     (sv['customerIDService.trading_firstname'] == "NORDD PATOLOGIA")
    # ]
    # # Client list FILTER END
    sv.to_excel('ServicesAPI.xlsx', index=True)
    sv['origCityList'] = sv['serviceIDRequested.source_address_id'].map(get_address_city_listed)
    sv['destCityList'] = sv['serviceIDRequested.destination_address_id'].map(get_address_city_listed)
    sv['origCity'] = sv['serviceIDRequested.source_address_id'].map(get_address_city)
    sv['destCity'] = sv['serviceIDRequested.destination_address_id'].map(get_address_city)

    sv.sort_values(
        by="protocol", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last'
    )

    # Materials

    mt = pd.DataFrame(columns=[])

    sv['priceDryIce'] = sv['serviceIDRequested.provider_id'].map(get_material_price)

    materials = ['embalagem_secundaria',
                 'gelox',
                 'isopor3l',
                 'terciaria3l',
                 'isopor7l',
                 'terciaria8l',
                 'caixa_termica',
                 'gelo_seco',
                 'embalagem_secundaria_pote_1l']

    for mat in materials:
        mt[f'{mat}Extra'] = sv[f'serviceIDRequested.{mat}'] - sv[f'serviceIDRequested.budgetIDService.{mat}']
        mt[f'{mat}Extra'] = np.where(mt[f'{mat}Extra'] < 0, 0, mt[f'{mat}Extra'])

    extra_list = []

    for rows in mt.itertuples():
        extram_list = [rows.embalagem_secundariaExtra,
                       rows.geloxExtra,
                       rows.isopor3lExtra,
                       rows.terciaria3lExtra,
                       rows.isopor7lExtra,
                       rows.terciaria8lExtra,
                       rows.caixa_termicaExtra,
                       rows.gelo_secoExtra,
                       rows.embalagem_secundaria_pote_1lExtra]

        extra_list.append(extram_list)

    sv = sv.assign(extraMaterials=extra_list)

    sv['PROTOCOLO'] = sv['protocol']
    sv['CLIENTE'] = sv['customerIDService.trading_firstname']
    sv['ETAPA'] = np.select(
        condlist=[
            sv['step'] == 'availableService',
            sv['step'] == 'toAllocateService',
            sv['step'] == 'toDeliveryService',
            sv['step'] == 'deliveringService',
            sv['step'] == 'toLandingService',
            sv['step'] == 'landingService',
            sv['step'] == 'toBoardValidate',
            sv['step'] == 'toCollectService',
            sv['step'] == 'collectingService',
            sv['step'] == 'toBoardService',
            sv['step'] == 'boardingService',
            sv['step'] == 'finishedService'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO', 'AGUARDANDO ALOCAÇÃO', 'EM ROTA DE ENTREGA', 'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA', 'DESEMBARCANDO', 'VALIDAR EMBARQUE', 'AGENDADO', 'COLETANDO',
            'EM ROTA DE EMBARQUE', 'EMBARCANDO SERVIÇO', 'FINALIZADO'],
        default=0
    )
    sv['TIPO DE SERVIÇO'] = sv['serviceIDRequested.service_type']
    sv['DATA COLETA'] = sv['collectDateTime'].dt.strftime(date_format='%d/%m/%Y')
    sv['CIDADE ORIGEM'] = sv['origCity']
    sv['CTE LOGLIFE'] = sv['cte_loglife']
    sv['VALOR DO ORÇAMENTO'] = sv['serviceIDRequested.budgetIDService.price']
    sv['FRANQUIA'] = sv['serviceIDRequested.budgetIDService.franchising']
    sv['PESO TAXADO NO SERVIÇO'] = sv['serviceIDBoard'].str[0].map(
        lambda x: x.get('taxed_weight', np.nan) if isinstance(x, dict) else 0)
    sv['VOLUME EMBARQUE'] = sv['serviceIDBoard'].str[0].map(
        lambda x: x.get('board_volume', np.nan) if isinstance(x, dict) else 0)
    sv['VALOR KG EXTRA'] = sv['serviceIDRequested.budgetIDService.price_kg_extra']
    sv['VALOR TOTAL KG EXTRA'] = np.ceil(
        (sv['PESO TAXADO NO SERVIÇO'] - sv['FRANQUIA'])
    ) * sv['VALOR KG EXTRA']
    sv['VALOR TOTAL KG EXTRA'] = np.where(
        sv['VALOR TOTAL KG EXTRA'] < 0, 0, sv['VALOR TOTAL KG EXTRA'])

    sv['QTD. END. ORIGEM NO SERVIÇO'] = sv['serviceIDRequested.source_address_id'].str.len()
    sv['QTD. END. ORIGEM NO ORÇAMENTO'] = sv['serviceIDRequested.budgetIDService.source_address_qty']
    sv['VALOR COLETA ADICIONAL'] = sv['serviceIDRequested.budgetIDService.price_add_collect']
    sv['VALOR TOTAL COLETAS ADICIONAIS'] = (sv['QTD. END. ORIGEM NO SERVIÇO']
                                                - sv['QTD. END. ORIGEM NO ORÇAMENTO'].astype(int))
    sv['VALOR TOTAL COLETAS ADICIONAIS'] = np.where(
        sv['VALOR TOTAL COLETAS ADICIONAIS'] < 0,
        0,
        sv['VALOR TOTAL COLETAS ADICIONAIS'] * sv['VALOR COLETA ADICIONAL']
    )

    sv['QTD. END. DESTINO NO SERVIÇO'] = sv['serviceIDRequested.destination_address_id'].str.len()
    sv['QTD. END. DESTINO NO ORÇAMENTO'] = sv['serviceIDRequested.budgetIDService.destination_address_qty']
    sv['VALOR ENTREGA ADICIONAL'] = sv['serviceIDRequested.budgetIDService.price_add_delivery']
    sv['VALOR TOTAL ENTREGAS ADICIONAIS'] = (sv['QTD. END. DESTINO NO SERVIÇO']
                                                 - sv['QTD. END. DESTINO NO ORÇAMENTO'].astype(int))
    sv['VALOR TOTAL ENTREGAS ADICIONAIS'] = np.where(
        sv['VALOR TOTAL ENTREGAS ADICIONAIS'] < 0,
        0,
        sv['VALOR TOTAL ENTREGAS ADICIONAIS'] * sv['VALOR ENTREGA ADICIONAL']
    )
    sv['NOME REMETENTE'] = sv['serviceIDRequested.source_address_id'].map(get_address_name)
    sv['NOME DESTINATÁRIO'] = sv['serviceIDRequested.destination_address_id'].map(get_address_name)
    sv['MATERIAL EXTRA (CLIENTE)'] = sv['extraMaterials'].map(materials_description)
    sv['CUSTO TOTAL MATERIAL EXTRA (CLIENTE)'] = sv.apply(
        lambda x: materials_value(x.extraMaterials, x.priceDryIce), axis=1)
    sv['PREÇO COLETA SEM SUCESSO'] = sv['serviceIDRequested.budgetIDService.price_unsuccessful_collect']
    sv['VALOR TOTAL DO EXTRA'] = sv['CUSTO TOTAL MATERIAL EXTRA (CLIENTE)'] + sv['VALOR TOTAL COLETAS ADICIONAIS'] \
                                 + sv['VALOR TOTAL ENTREGAS ADICIONAIS'] + sv['VALOR TOTAL KG EXTRA']
    sv['OBSERVAÇÕES'] = sv['serviceIDRequested.observation']

    sv = sv.loc[sv['VALOR TOTAL DO EXTRA'] != 0]

    report = sv[[
        'PROTOCOLO',
        'CLIENTE',
        'ETAPA',
        'TIPO DE SERVIÇO',
        'DATA COLETA',
        'CIDADE ORIGEM',
        'CTE LOGLIFE',
        'VALOR DO ORÇAMENTO',
        'FRANQUIA',
        'PESO TAXADO NO SERVIÇO',
        'VOLUME EMBARQUE',
        'VALOR KG EXTRA',
        'VALOR TOTAL KG EXTRA',
        'QTD. END. ORIGEM NO SERVIÇO',
        'QTD. END. ORIGEM NO ORÇAMENTO',
        'VALOR COLETA ADICIONAL',
        'VALOR TOTAL COLETAS ADICIONAIS',
        'QTD. END. DESTINO NO SERVIÇO',
        'QTD. END. DESTINO NO ORÇAMENTO',
        'VALOR ENTREGA ADICIONAL',
        'VALOR TOTAL ENTREGAS ADICIONAIS',
        'NOME REMETENTE',
        'NOME DESTINATÁRIO',
        'MATERIAL EXTRA (CLIENTE)',
        'CUSTO TOTAL MATERIAL EXTRA (CLIENTE)',
        'VALOR TOTAL DO EXTRA',
        'OBSERVAÇÕES'
    ]].copy()

    cte_comp_path = folderpath2.get()
    cte_comp_path = cte_comp_path.replace('/', '\\')

    csv_report = pd.DataFrame(columns=['Protocolo', 'CTE Complementar', 'Data Emissão CTE'])

    if unique is True:
        protocol_entry = cte_cs.get()
        protocol_list = protocol_entry.split(";")
        integer_map = map(int, protocol_list)
        protocol_list = list(integer_map)
        report = report[report['PROTOCOLO'].isin(protocol_list)]
        csv_report['Protocolo'] = protocol_list
    else:
        protocol_list = sv['protocol'].to_list()
        csv_report['Protocolo'] = sv['protocol']

    report_date = dt.datetime.strftime(dt.datetime.now(), '%d/%m/%Y')
    csv_report['Data Emissão CTE'] = report_date

    report.to_excel(
        f'{cte_comp_path}\\CTE complementar-{now_date}-{now}.xlsx',
        index=False
    )

    print("Relatório exportado!")

    bot_cte = bot.Bot()

    # bot_cte.open_bsoft(path=filename.get(), login=login.get(), password=password.get())
    current_row = 0

    for protocol in protocol_list:

        # tomador_cnpj = sv.loc[sv['protocol'] == protocol, 'customerIDService.cnpj_cpf'].values.item()
        source_add = sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.source_address_id'].values.item()
        destination_add = sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.destination_address_id'].values.item()
        valor = str(sv.loc[sv['protocol'] == protocol, 'VALOR TOTAL DO EXTRA'].values.item())
        uf1 = address.loc[address['id'] == source_add[0], 'state'].values.item()
        uf2 = address.loc[address['id'] == destination_add[0], 'state'].values.item()
        # cnpj_remetente = get_address_cnpj(source_add)
        # cnpj_destinatario = get_address_cnpj(destination_add)
        uf_rem = uf_base.loc[uf_base['Estado'] == uf1, 'UF'].values.item()
        uf_dest = uf_base.loc[uf_base['Estado'] == uf2, 'UF'].values.item()
        icms_obs = uf_base.loc[uf_base['Estado'] == uf_rem, 'Info'].values.item()
        aliq = aliquota_base.loc[aliquota_base['UF'] == uf_rem, uf_dest].values.item()
        cte = sv.loc[sv['protocol'] == protocol, 'cte_loglife'].values.item()
        obs_text = f'Protocolo {protocol} - {icms_obs}'
        if uf_rem != "MG":
            aliq_text = float(aliq)*float(valor)*0.008
            aliq_text = "{:0.2f}".format(aliq_text)
            icms_obs = icms_obs.replace('#', aliq_text)
            aliq = "0"

        # if tomador_cnpj in cnpj_remetente:
        #     tomador = "Remetente"
        #     cnpj_remetente = [tomador_cnpj]
        # elif tomador_cnpj in cnpj_destinatario:
        #     tomador = "Destinatário"
        #     cnpj_destinatario = [tomador_cnpj]
        # else:
        #     tomador = "Outro"

        # bot_cte.action(
        #     cnpj_sender=cnpj_remetente,
        #     cnpj_receiver=cnpj_destinatario,
        #     payer=tomador,
        #     payer_cnpj=tomador_cnpj
        # )

        bot_cte.part3_complimentary(
            cte=cte
        )
        bot_cte.part4(
            tax=str(aliq),
            icms_text=icms_obs,
            uf=uf_rem,
            price=valor,
            complimentary=True
        )

        cte_llm_complimentary = int(bot_cte.get_clipboard())
        report.at[report.index[current_row], 'CTE Complementar'] = cte_llm_complimentary
        csv_report.at[csv_report.index[current_row], 'Protocolo'] = protocol
        csv_report.at[csv_report.index[current_row], 'CTE Complementar'] = cte_llm_complimentary

        report.to_excel(
            f'{cte_comp_path}\\CTE complementar-{now_date}-{now}.xlsx',
            index=False
        )

        csv_report = csv_report.astype(str)
        csv_report = csv_report.replace(to_replace="\.0+$", value="", regex=True)

        csv_report.to_csv(
            f'{cte_comp_path}\\Upload complementar-{now_date}-{now}.csv',
            index=False,
            encoding='utf-8'
        )

        current_row += 1


def cte_unique():
    def get_address_cnpj(add_list):
        _cnpj_list = []
        for add in add_list:
            cnpj = address.loc[address['id'] == add, 'cnpj_cpf'].values.item()
            _cnpj_list.append(cnpj)
        return _cnpj_list

    def get_collector_cnpj(col):
        _cnpj_list = []
        cnpj = collector.loc[collector['id'] == col, 'cnpj'].values.item()
        _cnpj_list.append(cnpj)
        return _cnpj_list

    di = cal.get_date()
    df = di

    di = dt.datetime.strftime(di, '%d/%m/%Y')
    df = dt.datetime.strftime(df, '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    di_temp = di_dt - dt.timedelta(days=5)
    df_temp = df_dt + dt.timedelta(days=5)

    di = dt.datetime.strftime(di_temp, '%d/%m/%Y')
    df = dt.datetime.strftime(df_temp, '%d/%m/%Y')

    # Requesting data from API

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    collector = r.request_public('https://transportebiologico.com.br/api/public/collector')
    services_ongoing = r.request_public('https://transportebiologico.com.br/api/public/service')
    services_finalized = r.post_public(
        f'https://transportebiologico.com.br/api/public/service/finalized/?startFilter={di}&endFilter={df}')

    sv = pd.concat([services_ongoing, services_finalized], ignore_index=True)

    df_dt += dt.timedelta(days=1)

    sv['collectDateTime'] = pd.to_datetime(sv['serviceIDRequested.collect_date']) - dt.timedelta(hours=3)
    sv['collectDateTime'] = sv['collectDateTime'].dt.tz_localize(None)

    # sv.drop(sv[sv['collectDateTime'] < di_dt].index, inplace=True)
    # sv.drop(sv[sv['collectDateTime'] > df_dt].index, inplace=True)

    protocol_entry = cte_s.get()
    protocol_list = protocol_entry.split(";")

    csv_report = pd.DataFrame(columns=['Protocolo', 'CTE Loglife', 'nan'])
    csv_report['Protocolo'] = sv['protocol']
    csv_report['nan'] = sv['protocol'] - sv['protocol']

    bot_cte = bot.Bot()

    current_row = 0

    for protocol in protocol_list:

        protocol = int(protocol)
        tomador_cnpj = sv.loc[sv['protocol'] == protocol, 'customerIDService.cnpj_cpf'].values.item()
        source_add = sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.source_address_id'].values.item()
        destination_add = sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.destination_address_id'].values.item()
        uf1 = address.loc[address['id'] == source_add[0], 'state'].values.item()
        uf2 = address.loc[address['id'] == destination_add[0], 'state'].values.item()
        uf_rem = uf_base.loc[uf_base['Estado'] == uf1, 'UF'].values.item()
        uf_dest = uf_base.loc[uf_base['Estado'] == uf2, 'UF'].values.item()
        icms_obs = uf_base.loc[uf_base['Estado'] == uf_rem, 'Info'].values.item()
        aliq = aliquota_base.loc[aliquota_base['UF'] == uf_rem, uf_dest].values.item()
        if uf_rem != "MG":
            aliq_text = float(aliq)*5*0.008
            aliq_text = "{:0.2f}".format(aliq_text)
            icms_obs = icms_obs.replace('#', aliq_text)
            aliq = "0"

        if cte_type.get() == 0:
            valor = str(sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.budgetIDService.price'].values.item())
            cnpj_remetente = get_address_cnpj(source_add)
            cnpj_destinatario = get_address_cnpj(destination_add)
            tipo_cte = None
            vols = None
            if tomador_cnpj in cnpj_remetente:
                tomador = "Remetente"
                cnpj_remetente = [tomador_cnpj]
            elif tomador_cnpj in cnpj_destinatario:
                tomador = "Destinatário"
                cnpj_destinatario = [tomador_cnpj]
            else:
                tomador = "Outro"
        else:
            source_collector = sv.loc[
                sv['protocol'] == protocol, 'serviceIDRequested.source_collector_id'
            ].values.item()
            dest_collector = sv.loc[
                sv['protocol'] == protocol, 'serviceIDRequested.destination_collector_id'
            ].values.item()
            valor = "5,00"
            cnpj_remetente = get_collector_cnpj(source_collector)
            cnpj_destinatario = get_collector_cnpj(dest_collector)
            tipo_cte = 1
            vols = volumes.get()
            if cnpj_remetente[0] in [
                "17.062.517/0001-08", "17.062.517/0002-99"
            ] or cnpj_remetente in [
                "17.062.517/0001-08", "17.062.517/0002-99"
            ]:
                tomador = "Destinatário"
            else:
                tomador = "Remetente"

        bot_cte.action(
            cnpj_sender=cnpj_remetente,
            cnpj_receiver=cnpj_destinatario,
            payer=tomador,
            payer_cnpj=tomador_cnpj
        )
        bot_cte.part3_normal(
            cte_instance=tipo_cte,
            volumes=vols
        )
        bot_cte.part4(
            tax=str(aliq),
            uf=uf_rem,
            icms_text=icms_obs,
            price=valor
        )

        if cte_type.get() == 0:
            cte_llm = int(bot_cte.get_clipboard())

            csv_report.at[csv_report.index[current_row], 'Protocolo'] = protocol
            csv_report.at[csv_report.index[current_row], 'CTE Loglife'] = cte_llm

            csv_report = csv_report.astype(str)
            csv_report = csv_report.replace(to_replace="\.0+$", value="", regex=True)

            csv_report.to_csv(
                f'{cte_path}\\Upload-{now_date}-{now}.csv',
                index=False
            )

            current_row += 1


r = RequestDataFrame()

header = {"xtoken": "myqhF6Nbzx"}
uf_base = pd.read_excel('Complementares.xlsx', sheet_name='Plan1')
aliquota_base = pd.read_excel('Alíquota.xlsx', sheet_name='Planilha1')

# Create Object
root = Tk()
root.title('CTe LogLife')
root.geometry("")
root.resizable(False, False)
root.iconbitmap('my_icon.ico')
thread_0 = Start(root)
thread_1 = Start(root)

# Setting tabs

tabs = ttk.Notebook(root)
tab1 = ttk.Frame(tabs)
tab2 = ttk.Frame(tabs)
tab3 = ttk.Frame(tabs)
tab4 = ttk.Frame(tabs)

tabs.add(tab1, text='CTe')
tabs.add(tab2, text='CTe complementar')
tabs.add(tab4, text='Pastas Relatório')
tabs.add(tab3, text='BSoft')


tabs.pack(expand=1, fill="both")

tab2_frame = Frame(tab2, pady=22)
tab2_frame.pack()

tab2_frame2 = Frame(tab2)
tab2_frame2.pack()

tab4_frame = Frame(tab4)
tab4_frame.pack()

# Set calendar

today = dt.datetime.today()

dia = today.day
mes = today.month
ano = today.year

style = ThemedStyle(root)
style.theme_use('breeze')

cal_config = {'selectmode': 'day',
              'day': dia,
              'month': mes,
              'year': ano,
              'locale': 'pt_BR',
              'firstweekday': 'sunday',
              'showweeknumbers': False,
              'bordercolor': "white",
              'background': "white",
              'disabledbackground': "white",
              'headersbackground': "white",
              'normalbackground': "white",
              'normalforeground': 'black',
              'headersforeground': 'black',
              'selectbackground': '#00a5e7',
              'selectforeground': 'white',
              'weekendbackground': 'white',
              'weekendforeground': 'black',
              'othermonthforeground': 'black',
              'othermonthbackground': '#E8E8E8',
              'othermonthweforeground': 'black',
              'othermonthwebackground': '#E8E8E8',
              'foreground': "black"}

cal = DateEntry(tab1, **cal_config)
cal1 = DateEntry(tab2_frame, **cal_config)
cal2 = DateEntry(tab2_frame, **cal_config)

cal.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")
cal1.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")
cal2.grid(column=1, row=1, padx=30, pady=10, sticky="E, W")

# Read file name

filename = StringVar()

try:
    with open('filename.txt') as m:
        text = m.read()
    lines = text.split('\n')
    filename.set(lines[0])
except FileNotFoundError:
    filename.set('Bsoft Web.exe')

filename_Label = ttk.Label(tab3, width=20, text=filename.get(), wraplength=140)
filename_Label.grid(column=1, row=0, padx=5, pady=5, ipadx=8)

# Read folder path

folderpath = StringVar()

try:
    with open('folderpath.txt') as m:
        text = m.read()
    lines = text.split('\n')
    folderpath.set(lines[0])
except FileNotFoundError:
    folderpath.set('Pasta para CTe normal')

folder_Label = ttk.Label(tab4_frame, width=20, text=folderpath.get(), wraplength=140)
folder_Label.grid(column=1, row=0, sticky="E, W", padx=20, pady=10)

folderpath2 = StringVar()

try:
    with open('folderpath2.txt') as m:
        text = m.read()
    lines = text.split('\n')
    folderpath2.set(lines[0])
except FileNotFoundError:
    folderpath2.set('Pasta para CTe complementar')

folder_Label2 = ttk.Label(tab4_frame, width=20, text=folderpath2.get(), wraplength=140)
folder_Label2.grid(column=1, row=1, sticky="E, W", padx=20, pady=10)

browse1 = Browse(filename_Label)
browse2 = Browse(folder_Label)
browse3 = Browse(folder_Label2)

# Add Buttons
ttk.Button(tab1,
           text="Emitir lista de CTe",
           command=lambda: thread_0.start_thread(
               cte_list, progressbar
           )).grid(column=2, row=0, padx=10, pady=15, sticky="N, S, E, W")

ttk.Button(tab1,
           text="Emitir",
           command=lambda: thread_0.start_thread(
               cte_unique, progressbar
           )).grid(column=2, row=1, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(tab2_frame,
           text="Emitir",
           command=lambda: thread_1.start_thread(
               cte_complimentary, progressbar2
           )).grid(column=2, row=0, rowspan=2, padx=10, pady=15, ipady=15, sticky=" E, W")

ttk.Button(tab2_frame2,
           text="Emitir",
           command=lambda: thread_1.start_thread(
               cte_complimentary, progressbar2, arguments=[True]
           )).grid(column=2, row=0, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(tab3,
           text="Procurar",
           command=lambda: browse1.browse_exe(filename, 'filename.txt', master=tab3,
                                                label_config={'wraplength': 140, 'width': 20},
                                                grid_config={'column': 1, 'row': 0, 'sticky': "E, W", 'padx': 20,
                                                             'pady': 10})
           ).grid(column=2, row=0, padx=5, pady=5, sticky="E, W")
ttk.Button(tab3, text="Salvar").grid(column=2, row=1, padx=5, pady=5, rowspan=2, ipady=13, sticky="E, W")

ttk.Button(tab4_frame,
           text="Procurar",
           command=lambda: browse2.browse_folder(folderpath, 'folderpath.txt', master=tab4_frame,
                                                 label_config={'wraplength': 140, 'width': 20},
                                                 grid_config={'column': 1, 'row': 0, 'sticky': "E, W", 'padx': 20,
                                                              'pady': 10})
           ).grid(column=2, row=0, padx=5, pady=5, sticky="E, W")

ttk.Button(tab4_frame,
           text="Procurar",
           command=lambda: browse3.browse_folder(folderpath2, 'folderpath2.txt', master=tab4_frame,
                                                 label_config={'wraplength': 140, 'width': 20},
                                                 grid_config={'column': 1, 'row': 1, 'sticky': "E, W", 'padx': 20,
                                                              'pady': 10})
           ).grid(column=2, row=1, padx=5, pady=5, sticky="E, W")


# Labels

ttk.Label(tab1, text="Data da Emissão:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
ttk.Label(tab1, text="CTe avulso:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')
ttk.Label(tab1, text="Vols:").grid(column=1, row=4, padx=10, pady=10, sticky='W, E')

ttk.Label(tab2_frame, text="Data Inicial:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
ttk.Label(tab2_frame, text="Data Final:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')

ttk.Label(tab2_frame2, text="CTe avulso:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')

ttk.Label(tab3, text="Nome do Arquivo:").grid(column=0, row=0, padx=5, pady=5, ipadx=10)
ttk.Label(tab3, text='Usuário:').grid(column=0, row=1, padx=5, pady=5, ipadx=10)
ttk.Label(tab3, text='Senha:').grid(column=0, row=2, padx=5, pady=5, ipadx=10)

ttk.Label(tab4_frame, text="CTe normal:").grid(column=0, row=0, padx=5, pady=5)
ttk.Label(tab4_frame, text="CTe complementar:").grid(column=0, row=1, padx=5, pady=5)

# Adding Entry texts

cte_s = StringVar()

ttk.Entry(tab1,
          textvariable=cte_s, width=14).grid(column=1, row=1, padx=10, pady=10)

cte_cs = StringVar()

ttk.Entry(tab2_frame2,
          textvariable=cte_cs, width=14).grid(column=1, row=0, padx=10, pady=10)

login = StringVar()
password = StringVar()

ttk.Entry(tab3,
          textvariable=login).grid(column=1, row=1, padx=5, pady=10, sticky='E, W')
ttk.Entry(tab3,
          textvariable=password).grid(column=1, row=2, padx=5, pady=5, sticky='E, W')

volumes = StringVar()
ttk.Entry(tab1,
          textvariable=volumes, width=8).grid(column=1, row=4, pady=5, ipadx=1)
volumes.set('1')

# Radio Buttons

cte_type = IntVar()

ttk.Radiobutton(tab1, text="Normal", value=0, variable=cte_type).grid(column=0, row=4, ipadx=18)
ttk.Radiobutton(tab1, text="Simbólico", value=1, variable=cte_type).grid(column=0, row=5, ipadx=10)

# Progress Bar

progressbar = ttk.Progressbar(tab1, mode='indeterminate')
progressbar.grid(column=0, row=6, sticky='W, E', columnspan=3)

progressbar2 = ttk.Progressbar(tab2, mode='indeterminate')
progressbar2.pack(side=BOTTOM, fill='x')

# Auto resize tabs

tab1.rowconfigure(2, weight=2)
tab1.columnconfigure(0, weight=1)
tab1.columnconfigure(1, weight=1)
tab1.columnconfigure(2, weight=1)

tab3.rowconfigure(0, weight=1)
tab3.columnconfigure(0, weight=1)
tab3.columnconfigure(1, weight=1)
tab3.columnconfigure(2, weight=1)

# Execute Tkinter
root.mainloop()
