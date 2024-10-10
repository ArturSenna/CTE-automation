import threading
from json import loads
from tkinter import filedialog as fd
from tkinter import Toplevel
from tkinter import ttk
from tkinter import *
import csv
import pandas as pd
import numpy as np
import requests


# import xlwings as xw
# from pythoncom import CoInitialize


class Start:

    def __init__(self, root_master):
        self.submit_thread = None
        self.master = root_master

    def start_thread(self, target_function, progress_bar_func=None, arguments=()):

        def check_thread():
            if self.submit_thread.is_alive():
                self.master.after(20, check_thread)
            else:
                progress_bar_func.stop()

        self.submit_thread = threading.Thread(target=target_function, args=arguments)
        self.submit_thread.daemon = True
        self.submit_thread.start()
        if progress_bar_func is not None:
            progress_bar_func.start()
            self.master.after(20, check_thread)


class Browse:

    def __init__(self, label_variable):
        self.label_variable = label_variable

    def browse_files(self, filename_variable=None, archive_name=None, master=None,
                     label_config=None, grid_config=None):
        filetypes = (
            ('Arquivos Excel', '*.xlsx'),
            ('Excel habilitado para macro', '*.xlsm')
        )

        file_name = fd.askopenfilename(
            title='Selecione o arquivo',
            initialdir='cd',
            filetypes=filetypes
        )

        with open(f"{archive_name}", 'w') as w:
            w.write(file_name)
        filename_variable.set(file_name)

        if self.label_variable is not None:
            self.label_variable.destroy()
            self.label_variable = ttk.Label(master, text=filename_variable.get(), **label_config)
            self.label_variable.grid(**grid_config)

    def browse_folder(self, folder_variable=None, archive_name='folderpath.txt', master=None,
                      label_config=None, grid_config=None):

        folder_path = fd.askdirectory(
            title='Selecione a pasta',
            initialdir='cd',
        )

        with open(archive_name, 'w') as w:
            w.write(folder_path)
        folder_variable.set(folder_path)

        if self.label_variable is not None:
            self.label_variable.destroy()
            self.label_variable = ttk.Label(master, text=folder_variable.get(), **label_config)
            self.label_variable.grid(**grid_config)

    def browse_exe(self, filename_variable=None, archive_name=None, master=None,
                   label_config=None, grid_config=None):
        filetypes = (('Arquivo executável', '*.exe'),)

        file_name = fd.askopenfilename(
            title='Selecione o arquivo',
            initialdir='cd',
            filetypes=filetypes
        )

        with open(f"{archive_name}", 'w') as w:
            w.write(file_name)
        filename_variable.set(file_name)

        if self.label_variable is not None:
            self.label_variable.destroy()
            self.label_variable = ttk.Label(master, text=filename_variable.get(), **label_config)
            self.label_variable.grid(**grid_config)


class RequestDataFrame:

    def __init__(self):
        self.headers = {"xtoken": "myqhF6Nbzx"}
        details = {"email": "artursenna@loglifelogistica.com.br", "password": "A3928024854c#"}
        key = requests.post('https://transportebiologico.com.br/api/sessions', json=details)
        key_json = loads(key.text)
        self.auth = {"authorization": key_json['token']}
        self.file_auth = {"authorization": key_json['token'],
                          'Content-Type': 'multipart/form-data',
                          'boundary': '----WebKitFormBoundaryKUpFTtNyFBN'}

    def request_public(self, link):
        response = requests.get(link, headers=self.headers)
        response_json = loads(response.text)
        dataframe = pd.json_normalize(response_json)

        return dataframe

    def post_public(self, link):
        response = requests.post(link, headers=self.headers)
        response_json = loads(response.text)
        dataframe = pd.json_normalize(response_json)

        return dataframe

    def request_private(self, link):
        response = requests.get(link, headers=self.auth)
        response_json = loads(response.text)
        dataframe = pd.json_normalize(response_json)

        return dataframe

    def post_file(self, link, file, file_type="csv_file", file_format="text/csv", upload_type=None):

        payload = {
            file_type: (
                file,
                open(file, "rb"),
                file_format,
                {"Expires": "0"},
            )
        }

        if upload_type is not None:
            upload_type = {
                "type": upload_type
            }

        response = requests.post(url=link, headers=self.auth, files=payload, data=upload_type)

        return response


# def export_to_excel(df, excel_name, sheet="Planilha1", clear_range="A1:A1", autofit=True, change_header=True,
#                     start_write=None, clear_filters=False):
#     app = xw.App(visible=False)
#     wb = xw.Book(f'{excel_name}')
#     ws = wb.sheets[f'{sheet}']
#     app.kill()
#     if wb.sheets[f'{sheet}'].api.AutoFilter:
#         wb.sheets[f'{sheet}'].api.AutoFilter.ShowAllData()
#     elif clear_filters:
#         wb.sheets[f'{sheet}'].api.AutoFilter.ShowAllData()
#     ws.range(clear_range).clear_contents()
#
#     if start_write is None:
#
#         if change_header:
#             start_write = "A1"
#             header_config = 1
#         else:
#             start_write = "A2"
#             header_config = 0
#     else:
#
#         if change_header:
#             header_config = 1
#         else:
#             header_config = 0
#
#     # Inserção do DataFrame na planilha
#     ws[f"{start_write}"].options(pd.DataFrame, header=header_config, index=False, expand='table').value = df
#
#     if autofit:
#         ws.autofit('r')
#
#
# def clear_data(filename, *sheet):
#     file_name = filename.replace('/', '\\')
#
#     CoInitialize()
#
#     for value in sheet:
#         app = xw.App(visible=False)
#         wb = xw.Book(f"{file_name}")
#         terms = value.split(';')
#         ws = wb.sheets[f'{terms[0]}']
#         app.kill()
#         if wb.sheets[f'{terms[0]}'].api.AutoFilter:
#             wb.sheets[f'{terms[0]}'].api.AutoFilter.ShowAllData()
#         ws.range(terms[1]).clear_contents()


def confirmation_pop_up(root, text):
    pop = Toplevel(root)
    pop.iconbitmap('my_icon.ico')
    pop.attributes('-topmost', 'true')
    pop.geometry("")
    pop.title("Confirmação")
    thread_pop = Start(pop)
    ttk.Label(pop, text=text).pack(padx=20, pady=10)
    ttk.Button(
        pop,
        text="OK",
        command=lambda: thread_pop.start_thread(lambda: [pop.destroy()])
    ).pack()


def post_file():
    r = RequestDataFrame()

    protocol = 95475
    cte_file = 'Minuta.pdf'
    cte = 123456789
    data_emissao = '18/10/2023'

    cte_csv = pd.DataFrame({
        'Protocolo': [protocol],
        'Arquivo PDF': [cte_file],
    })

    upload_cte_csv = pd.DataFrame({
        'Protocolo': [protocol],
        'CTE Loglife': [cte],
        'Data Emissão CTE': [data_emissao]
    })

    cte_csv.to_csv('AssociarPDF.csv', index=False, encoding='utf-8')
    upload_cte_csv.to_csv('UploadCTE.csv', index=False,  encoding='utf-8')

    cte_upload = r.post_file('https://transportebiologico.com.br/api/uploads/cte-loglife', 'UploadCTE.csv')

    pdf_upload = r.post_file("https://transportebiologico.com.br/api/pdf",
                             cte_file,
                             upload_type="CTE LOGLIFE",
                             file_format="application/pdf",
                             file_type="pdf_files")

    associate_pdf = r.post_file('https://transportebiologico.com.br/api/pdf/associate',
                                'AssociarPDF.csv',
                                upload_type="CTE LOGLIFE")

    print(pdf_upload.text, pdf_upload)
    print(associate_pdf.text, associate_pdf)
    print(cte_upload.text, cte_upload)


if __name__ == "__main__":
    post_file()
