def reiniciar():
    if caminho == 3:
        exit()
    print('\033[33mType E in the question below if you want to exit\033[m')
    clean = input(
        'The program will be restarted. Do you want to clean the workspace? [Y/N]').upper()

    while True:
        if clean == 'Y':
            os.system('cls')
            subprocess.call(
                [sys.executable, os.path.realpath(__file__)] + sys.argv[1:])
            exit()
        elif clean == 'N':
            subprocess.call(
                [sys.executable, os.path.realpath(__file__)] + sys.argv[1:])
            exit()
        elif clean == 'E':
            exit()
        else:
            print('\033[31mINVALID OPTION\033[m')
            clean = input(
                'Type only Y for YES or N for NO: [Y/N]').upper()


def popup_html(row):
    """--> Função para criar uma tabela estilizada para exibir as informações da(s) parcela(s) do(s) talhão(ões)"""
    i = row
    name = df_arvore['STATE'].iloc[i]
    hdom = df_arvore['HDOM'].iloc[i]
    diameter = df_arvore['DIAMETER'].iloc[i]
    codetalhao = df_arvore['CODETALHAO'].iloc[i]

    left_col_color = "#28d4bd"
    right_col_color = "#22b534"

    html = """<!DOCTYPE html>
<html>
<head>
<h4 style="margin-bottom:10"; width="200px">{}</h4>""".format(name) + """
</head>
    <table style="height: 80px; width: 200px;">
<tbody>
<tr>
<td style="padding-left: 4px; background-color: """ + left_col_color + """;"><span style="color: #000;">HDOM</span></td>
<td style="width: 100px; color: #FFF; padding-left: 4px; background-color: """ + right_col_color + """;">{}</td>""".format(
        hdom) + """
</tr>
<tr>
<td style="padding-left: 4px; background-color: """ + left_col_color + """;"><span style="color: #000;">DIAMETER</span></td>
<td style="width: 100px; color: #FFF; padding-left: 4px; background-color: """ + right_col_color + """;">{}</td>""".format(
        diameter) + """
</tr>
<tr>
<td style="padding-left: 4px; background-color: """ + left_col_color + """;"><span style="color: #000;">CODETALHAO</span></td>
<td style="width: 100px; padding-left: 4px; color: #FFF; background-color: """ + right_col_color + """;">{}</td>""".format(
        codetalhao) + """
</tr>
</tbody>
</table>
</html>
"""
    return html


while True:
    try:
        import os
        import sys
        import utm
        import math
        import branca
        import folium
        import tkinter
        import subprocess
        import webbrowser
        import sqlalchemy
        import numpy as np
        import pandas as pd
        from numpy import exp
        from tqdm import tqdm
        from tqdm import trange
        import tkinter.filedialog
        from datetime import datetime
        from sqlalchemy.exc import IntegrityError, DatabaseError

        while True:
            print('What to do?')
            caminho = int(input('INSERT[1] UPDATE[2] EXIT[3]: '))

            if caminho == 3:
                exit()
            elif caminho == 2:
                print('Function unavailable! Do not have a db to update.')
            elif caminho == 1:
                print('IMPORTING EXCEL FILES')

                excel_to_merge = list()
                arv_excel = pd.DataFrame()
                arv = pd.DataFrame()

                tkinter.Tk()
                # Remember to change initialdir to YOUR dir
                selected_files = tkinter.filedialog.askopenfilenames(
                    initialdir=r"C:\Users\tress\Documents\My Documents\Meus Projetos\CalculoDeInventario"
                               r"\Medições de Inventário", title='Selecione os arquivos que deseja processar',
                    filetypes=(('All types', '*.*'), ('Excel files', '.xlsx')))

                for file in selected_files:
                    if file.endswith('.xlsx') or file.endswith('.xls'):
                        excel_to_merge.append(
                            pd.read_excel(fr'{file}', sheet_name=0))

                for excel_file in excel_to_merge:
                    arv_excel = pd.concat(
                        [arv_excel, excel_file], ignore_index=True)

                if arv_excel.isnull().values.any():
                    print('\033[31mERROR: Excel file with blank columns/rows!')
                    reiniciar()

                arv_excel['CODETALHAO'] = arv_excel['CODETALHAO'].apply(
                    '{0:0>4}'.format)

                arv_excel['DATE'] = pd.to_datetime(
                    arv_excel['DATE'], dayfirst=True)

                chave_talhao = arv_excel.apply(lambda ct: '%s%s%s%s' % (ct['CODETALHAO'], '/', ct['CODE'], ct[
                    'STATE']), axis=1)
                arv_excel.insert(0, "KEYTREE", chave_talhao,
                                 allow_duplicates=False)

                arv_excel.to_excel(
                    r"C:\Users\tress\Documents\My Documents\Meus Projetos\CalculoDeInventario/Arvore.xlsx")

                break
            else:
                print('\033[31mINVALID OPTION\033[m')

        df_arvore = pd.DataFrame(arv_excel)

        df_arvore.columns = df_arvore.columns.str.upper()

        df_arvore['KEYTREE'] = df_arvore.apply(lambda ctn: '%s%s%s%s%s' % (ctn['PARC'], '-', ctn['CODETALHAO'],
                                                                           '/', ctn['KEYTREE']), axis=1)
        print(df_arvore)
        print('CALCULANDO HDOM')

        for y in tqdm(df_arvore.KEYTREE.unique()):
            dfmask = df_arvore['KEYTREE'] == y
            dadosarvchave = df_arvore[dfmask]
            dadosarvfiltrado = pd.DataFrame(dadosarvchave)
            div_num = 0

            if 'DT' not in dadosarvfiltrado['SITUATION'].values:
                DIAMETERs = list(dadosarvfiltrado.DIAMETER.values)
                four_majors = []

                for i in range(4):
                    div_num += 1
                    if len(DIAMETERs) == 0:
                        break
                    max_DIAMETER_index = DIAMETERs.index(max(DIAMETERs))
                    four_majors.append(max(DIAMETERs))
                    DIAMETERs.remove(DIAMETERs[max_DIAMETER_index])

                if len(DIAMETERs) == 0:
                    df_arvore.loc[df_arvore.KEYTREE == y, 'HDOM'] = (
                        sum(four_majors) / div_num).__round__(2)
                else:
                    df_arvore.loc[df_arvore.KEYTREE == y, 'HDOM'] = (
                        sum(four_majors) / 4).__round__(2)
            else:
                res = dadosarvfiltrado.loc[df_arvore.SITUATION == 'DT', 'HEIGHT'].mean().round(
                    2)
                df_arvore.loc[df_arvore.KEYTREE == y, 'HDOM'] = res

        df_arvore = df_arvore.replace([np.inf, -np.inf], 0)
        df_arvore = df_arvore.fillna(0)

        print('CALCULANDO COLUNAS')

        for i in trange(len(df_arvore)):
            g = math.pi * (df_arvore.DIAMETER.iloc[i] ** 2) / 40000
            df_arvore.loc[df_arvore.index == i, 'G'] = g

            if df_arvore.DIAMETER.iloc[i] == 0:
                df_arvore.loc[df_arvore.index == i, 'TREESHAPE'] = 0
            else:
                treeshape = (math.pi * 20)
                df_arvore.loc[df_arvore.index == i,
                              'TREESHAPE'] = treeshape

        df_arvore.to_excel(
            r"C:\Users\tress\Documents\My Documents\Meus Projetos\CalculoDeInventario\Result.xlsx")

        unidades_convertidas = utm.to_latlon(
            df_arvore.LAT.iloc[0], df_arvore.LON.iloc[0], 23, 'K')
        mapa = folium.Map(location=unidades_convertidas,
                          zoom_start=10, control_scale=True)

        folium.TileLayer(tiles='https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}', attr='Google',
                         name='Google Earth Satellite', overlay=True, control=True).add_to(mapa)

        for i in range(len(df_arvore)):
            unidades_convertidas = utm.to_latlon(
                df_arvore.LAT.iloc[i], df_arvore.LON.iloc[i], 23, 'K')
            html = popup_html(i)
            iframe = branca.element.IFrame(html=html, width=510, height=280)
            popup = folium.Popup(folium.Html(html, script=True), max_width=500)
            folium.Marker(location=unidades_convertidas, popup=popup, icon=folium.Icon(color="green", icon='info-sign'),
                          tooltip='Click Me').add_to(mapa)

        mapa.save('mapa.html'), webbrowser.open('mapa.html')

    except (DatabaseError) as e:
        if 'ORA-00904' in str(e):
            print('\033[31mERROR: "NAN" encontrado\n'
                  'Possível Causa: Talhão Dividido/Cortado\n'
                  f'Descrição: {str(e)}\033[m')
        else:
            print(f'\033[31mERROR: {str(e)}\033[m')
    except Exception as e:
        print(f'\033[31mERROR: {e}\033[m')

    reiniciar()
