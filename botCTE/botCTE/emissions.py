import datetime as dt
import os
import time

import numpy as np

import bot
from functions import *


def cte_list(start_date, folderpath, cte_folder, root):
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
    now_date = dt.datetime.strftime(start_date, '%d-%m-%Y')
    now = dt.datetime.strftime(now, "%H-%M")

    di = dt.datetime.strftime(start_date, '%d/%m/%Y')
    df = dt.datetime.strftime(start_date, '%d/%m/%Y')

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
    sv = pd.concat([sv[sv['cte_loglife'].isnull()],
                    sv[sv['cte_loglife'] == 'nan'],
                    sv[sv['cte_loglife'] == '-']], ignore_index=True)
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

    cte_path = folderpath
    cte_path = cte_path.replace('/', '\\')

    excel_file = f'{cte_path}\\Lista CTE-{now_date}-{now}.xlsx'
    csv_file = f'{cte_path}\\Upload-{now_date}-{now}.csv'
    csv_associate = f'{cte_path}\\Associar-{now_date}-{now}.csv'

    cte_folder_path = cte_folder.replace('/', '\\')

    report.to_excel(excel_file, index=False)

    print("Relatório exportado!")

    bot_cte = bot.Bot()

    # bot_cte.open_bsoft(path=filename.get(), login=login.get(), password=password.get())
    current_row = 0

    for protocol in sv['protocol']:

        report_date = dt.datetime.strftime(dt.datetime.now(), '%d/%m/%Y')

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
            aliq_text = float(aliq) * float(valor) * 0.008
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

        csv_report = pd.DataFrame({
            'Protocolo': [int(protocol)],
            'CTE Loglife': [cte_llm],
            'Data Emissão CTE': [report_date]
        })

        cte_file = f'{str(cte_llm).zfill(8)}.pdf'

        cte_csv = pd.DataFrame({
            'Protocolo': [int(protocol)],
            'Arquivo PDF': [cte_file],
        })

        report.to_excel(
            excel_file,
            index=False
        )

        csv_report = csv_report.astype(str)
        csv_report = csv_report.replace(to_replace="\.0+$", value="", regex=True)

        csv_report.to_csv(csv_file, index=False, encoding='utf-8')

        cte_csv.to_csv(csv_associate, index=False, encoding='utf-8')

        r.post_file('https://transportebiologico.com.br/api/uploads/cte-loglife', csv_file)

        while True:
            try:
                r.post_file("https://transportebiologico.com.br/api/pdf",
                            f'{cte_folder_path}\\{cte_file}',
                            upload_type="CTE LOGLIFE",
                            file_format="application/pdf",
                            file_type="pdf_files")
                break
            except FileNotFoundError:
                time.sleep(0.5)
                continue

        r.post_file('https://transportebiologico.com.br/api/pdf/associate',
                    csv_associate,
                    upload_type="CTE LOGLIFE")

        os.remove(csv_associate)

        current_row += 1

    confirmation_pop_up(root)


def cte_complimentary(start_date, final_date, cte_comp_path, cte_folder, root, unique=False, cte_cs=""):
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
        price_list = [5, 5, 10, 10, 12.5, 12.5, 0, price, 10]
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
    now_date = dt.datetime.strftime(start_date, '%d-%m-%Y')
    now = dt.datetime.strftime(now, "%H-%M")

    di = dt.datetime.strftime(start_date, '%d/%m/%Y')
    df = dt.datetime.strftime(final_date, '%d/%m/%Y')

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
    # sv.drop(sv[sv['customerIDService.trading_firstname'] == "HPV - HOSPITAL MOINHOS DE VENTO"].index, inplace=True)
    sv = sv.loc[(sv['cte_loglife'].notnull()) &
                (sv['cte_loglife'] != 'nan') &
                (sv['cte_loglife'] != '-') &
                (sv['step'] != 'toCollectService') &
                (sv['step'] != 'collectingService') &
                (sv['step'] != 'toBoardService') &
                (sv['step'] != 'boardingService') &
                (sv['step'] != 'toBoardValidate')]
    # # Client list FILTER START.
    # sv = sv.loc[
    #     (sv['customerIDService.trading_firstname'] == "CERBA-LCA") |
    #     (sv['customerIDService.trading_firstname'] == "PROVET") |
    #     (sv['customerIDService.trading_firstname'] == "FLOW") |
    #     (sv['customerIDService.trading_firstname'] == "ALCHEMYPET MEDICINA DIAGNÓSTICA VETERINÁRIA LTDA.") |
    #     (sv['customerIDService.trading_firstname'] == "LEMOS LABORATORIOS") |
    #     (sv['customerIDService.trading_firstname'] == "LABORATÓRIO KTZ") |
    #     (sv['customerIDService.trading_firstname'] == "NORDD PATOLOGIA")
    # ]
    # # Client list FILTER END.
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

    cte_comp_path = cte_comp_path.replace('/', '\\')

    if unique is True:
        protocol_entry = cte_cs
        protocol_list = protocol_entry.split(";")
        integer_map = map(int, protocol_list)
        protocol_list = list(integer_map)
        report = report[report['PROTOCOLO'].isin(protocol_list)]
    else:
        protocol_list = sv['protocol'].to_list()

    excel_file = f'{cte_comp_path}\\CTE complementar-{now_date}-{now}.xlsx'
    csv_file = f'{cte_comp_path}\\Upload complementar-{now_date}-{now}.csv'
    csv_associate = f'{cte_comp_path}\\Associar complementar-{now_date}-{now}.csv'

    cte_folder_path = cte_folder.replace('/', '\\')

    report.to_excel(excel_file, index=False)

    print("Relatório exportado!")

    bot_cte = bot.Bot()

    # bot_cte.open_bsoft(path=filename.get(), login=login.get(), password=password.get())
    current_row = 0

    for protocol in protocol_list:

        report_date = dt.datetime.strftime(dt.datetime.now(), '%d/%m/%Y')

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
            aliq_text = float(aliq) * float(valor) * 0.008
            aliq_text = "{:0.2f}".format(aliq_text)
            obs_text = obs_text.replace('#', aliq_text)
            aliq = "0"

        bot_cte.part3_complimentary(
            cte=cte
        )

        bot_cte.part4(
            tax=str(aliq),
            icms_text=obs_text,
            uf=uf_rem,
            price=valor,
            complimentary=True
        )

        cte_llm_complimentary = int(bot_cte.get_clipboard())
        report.at[report.index[current_row], 'CTE Complementar'] = cte_llm_complimentary

        cte_file = f'{str(cte_llm_complimentary).zfill(8)}.pdf'

        cte_csv = pd.DataFrame({
            'Protocolo': [protocol],
            'Arquivo PDF': [cte_file],
        })

        report.to_excel(excel_file, index=False)

        csv_report = pd.DataFrame({
            'Protocolo': [int(protocol)],
            'CTE Complementar': [cte_llm_complimentary],
            'Data Emissão CTE': [report_date]
        })

        csv_report = csv_report.astype(str)
        csv_report = csv_report.replace(to_replace="\.0+$", value="", regex=True)

        csv_report.to_csv(csv_file, index=False, encoding='utf-8')

        cte_csv.to_csv(csv_associate, index=False, encoding='utf-8')

        first_response = r.post_file('https://transportebiologico.com.br/api/uploads/cte-complementary', csv_file)

        while True:
            try:
                second_response = r.post_file("https://transportebiologico.com.br/api/pdf",
                                              f'{cte_folder_path}\\{cte_file}',
                                              upload_type="CTE COMPLEMENTAR",
                                              file_format="application/pdf",
                                              file_type="pdf_files")
                break
            except FileNotFoundError:
                time.sleep(0.5)
                second_response = 0
                continue

        third_response = r.post_file('https://transportebiologico.com.br/api/pdf/associate',
                                     csv_associate,
                                     upload_type="CTE COMPLEMENTAR")

        csv_report.to_csv(csv_file, index=False, encoding='utf-8')

        print(first_response, second_response, third_response)
        print(first_response.text, second_response.text, third_response.text)

        current_row += 1

    confirmation_pop_up(root)


def cte_unique(cal_date, cte_path, cte_folder_path, cte_type, cte_s, volumes, root):
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

    def get_collector(col):
        col_name = collector.loc[collector['id'] == col, 'trading_name'].values.item()
        return col_name

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

    di = cal_date
    df = di

    now = dt.datetime.now()
    now_date = dt.datetime.strftime(di, '%d-%m-%Y')
    now = dt.datetime.strftime(now, "%H-%M")

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

    sv['origCity'] = sv['serviceIDRequested.source_address_id'].map(get_address_city)
    sv['destCity'] = sv['serviceIDRequested.destination_address_id'].map(get_address_city)

    # sv.drop(sv[sv['collectDateTime'] < di_dt].index, inplace=True)
    # sv.drop(sv[sv['collectDateTime'] > df_dt].index, inplace=True)

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

    cte_path = cte_path.replace('/', '\\')

    if cte_type == 0:
        excel_file = f'{cte_path}\\Lista CTE-{now_date}-{now}.xlsx'
    else:
        excel_file = f'{cte_path}\\CTe Simbólico-{now_date}-{now}.xlsx'

    csv_file = f'{cte_path}\\Upload-{now_date}-{now}.csv'
    csv_associate = f'{cte_path}\\Associar-{now_date}-{now}.csv'

    cte_folder_path = cte_folder_path.replace('/', '\\')

    protocol_entry = cte_s
    protocol_list = protocol_entry.split(";")

    report = report[report['PROTOCOLO'].isin([int(x) for x in protocol_list])]
    report.to_excel(excel_file, index=False)

    bot_cte = bot.Bot()

    current_row = 0

    for protocol in protocol_list:

        report_date = dt.datetime.strftime(dt.datetime.now(), '%d/%m/%Y')

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
            aliq_text = float(aliq) * 5 * 0.008
            aliq_text = "{:0.2f}".format(aliq_text)
            icms_obs = icms_obs.replace('#', aliq_text)
            aliq = "0"

        if cte_type == 0:
            valor = str(sv.loc[sv['protocol'] == protocol, 'serviceIDRequested.budgetIDService.price'].values.item())
            cnpj_remetente = get_address_cnpj(source_add)
            cnpj_destinatario = get_address_cnpj(destination_add)
            tipo_cte = None
            vols = volumes
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
            vols = volumes
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

        cte_llm = int(bot_cte.get_clipboard())

        cte_file = f'{str(cte_llm).zfill(8)}.pdf'

        report.at[report.index[current_row], 'CTE LOGLIFE'] = cte_llm
        report.to_excel(excel_file, index=False)

        csv_report = pd.DataFrame({
            'Protocolo': [int(protocol)],
            'CTE Loglife': [cte_llm],
            'Data Emissão CTE': [report_date]
        })

        csv_report = csv_report.astype(str)
        csv_report = csv_report.replace(to_replace="\.0+$", value="", regex=True)

        csv_report.to_csv(csv_file, index=False)

        associate = pd.DataFrame({
            'Protocolo': [int(protocol)],
            'Arquivo PDF': [cte_file],
        })

        associate.to_csv(csv_associate, index=False, encoding='utf-8')

        while True:
            try:
                r.post_file("https://transportebiologico.com.br/api/pdf",
                            f'{cte_folder_path}\\{cte_file}',
                            upload_type="CTE LOGLIFE",
                            file_format="application/pdf",
                            file_type="pdf_files")
                break
            except FileNotFoundError:
                time.sleep(0.5)
                continue

        r.post_file('https://transportebiologico.com.br/api/pdf/associate',
                    csv_associate,
                    upload_type="CTE LOGLIFE")

        os.remove(csv_file)
        os.remove(csv_associate)

        if cte_type == 1:
            os.remove(excel_file)

        current_row += 1

    confirmation_pop_up(root)


r = RequestDataFrame()

uf_base = pd.read_excel('Complementares.xlsx', sheet_name='Plan1')
aliquota_base = pd.read_excel('Alíquota.xlsx', sheet_name='Planilha1')
