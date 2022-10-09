# The variables names, functions, etc are in portuguese!
# This is just an idea of what the code really are. It is non functional and several lines were removed
# for contractual and ethical issues
def reiniciar():
    if caminho == 3:
        exit()
    print('\033[33mDigite F na pergunta abaixo caso queira fechar o programa\033[m')
    clean = input(
        'O programa será reiniciado. Deseja limpar a área de trabalho? [S/N]').upper()

    while True:
        if clean == 'S':
            os.system('cls')
            subprocess.call(
                [sys.executable, os.path.realpath(__file__)] + sys.argv[1:])
            exit()
        elif clean == 'N':
            subprocess.call(
                [sys.executable, os.path.realpath(__file__)] + sys.argv[1:])
            exit()
        elif clean == 'F':
            print('\033[mFINALIZADO')
            exit()
        else:
            print('\033[31mOPÇÃO INVÁLIDA\033[m')
            clean = input(
                'Digite apenas S para SIM ou N para NÃO: [S/N]').upper()


def popup_html(row):
    """--> Função para criar uma tabela estilizada para exibir as informações da(s) parcela(s) do(s) talhão(ões)"""
    i = row
    NAME = calcmedicao['NOME'].iloc[i]
    volume = calcmedicao['VOLUME'].iloc[i]
    dapm = calcmedicao['DIAMETRO'].iloc[i]
    altm = calcmedicao['ALTURA'].iloc[i]

    left_col_color = "#28d4bd"
    right_col_color = "#22b534"

    html = """<!DOCTYPE html>
<html>
<head>
<h4 style="margin-bottom:10"; width="200px">{}</h4>""".format(nome) + """
</head>
    <table style="height: 80px; width: 200px;">
<tbody>
<tr>
<td style="padding-left: 4px; background-color: """ + left_col_color + """;"><span style="color: #000;">Volume</span></td>
<td style="width: 100px; color: #FFF; padding-left: 4px; background-color: """ + right_col_color + """;">{}</td>""".format(
        volume) + """
</tr>
<tr>
<td style="padding-left: 4px; background-color: """ + left_col_color + """;"><span style="color: #000;">DIAMETRO</span></td>
<td style="width: 100px; color: #FFF; padding-left: 4px; background-color: """ + right_col_color + """;">{}</td>""".format(
        diametro) + """
</tr>
<tr>
<td style="padding-left: 4px; background-color: """ + left_col_color + """;"><span style="color: #000;">ALTURA</span></td>
<td style="width: 100px; padding-left: 4px; color: #FFF; background-color: """ + right_col_color + """;">{}</td>""".format(
        altura) + """
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

        engine_inv = sqlalchemy.create_engine(
            "oracle+cx_oracle://dbName:password@db")
        dados_existentes = 0

        while True:
            print('O QUE DESEJA FAZER?')
            caminho = int(input('INSERT[1] UPDATE[2] SAIR[3]: '))

            if caminho == 3:
                print('\033[0mFINALIZADO')
                exit()
            elif caminho == 2:
                dados_existentes = 1
                print('OS DADOS SERÃO PROCESSADOS E ATUALIZADOS:')
                print('\033[33mUtilize espaço para a separação; * para filtrar tudo; Não se preocupe com maiúsculas e '
                      'minúsculas\nExemplo --> ESTADO: Mg | CODE: 0001 0378 | CODTALHAO: *\033[m')

                while True:
                    while True:
                        estado = input('SELECIONE O ESTADO SP[1] OU MG[2]: ')
                        if estado == '1':
                            estado = 'SP'
                            break
                        elif estado == '2':
                            estado = 'MG'
                            break
                        else:
                            print('\033[31mOPÇÃO INVÁLIDA\033[m')

                    horto = input('CODE: ').split(' ')
                    talhao = input('CODTALHAO: ').upper().split(' ')
                    sql_query = f"""query"""

                    if '*' in estado:
                        print('\033[31m* NÃO permitido no campo ESTADO\033[m')
                    elif '*' in talhao and '*' in horto:
                        print(
                            f'\033[mATUALIZANDO: {estado} | {horto} | {talhao}\033[m')
                        sql_query += f"""query"""
                        df_sql = pd.read_sql_query(sql_query, con=engine_inv)
                        break
                    elif '*' in talhao:
                        print(
                            f'\033[mATUALIZANDO: {estado} | {horto} | {talhao}\033[m')
                        sql_query += f"""query"""
                        df_sql = pd.read_sql_query(sql_query, con=engine_inv)
                        break
                    elif '*' in horto:
                        print(
                            f'\033[mATUALIZANDO: {estado} | {horto} | {talhao}\033[m')
                        sql_query += f"""query"""
                        df_sql = pd.read_sql_query(sql_query, con=engine_inv)
                        break
                    else:
                        print(
                            f'\033[mATUALIZANDO: {estado} | {horto} | {talhao}\033[m')
                        sql_query += f"""query"""
                        df_sql = pd.read_sql_query(sql_query, con=engine_inv)
                        break

                df_sql = df_sql.rename(columns={'rename some columns'})
                df_sql.columns = df_sql.columns.str.upper()
                df_sql['DAP'] = df_sql['DAP'].round(2)

                # Export df to Excel
                df_sql.to_excel(r"Somewhere")
                # Import from Excel
                arv = pd.read_excel(r"Some Excel file", sheet_name=0)

                arv['CODTALHAO'] = arv['CODTALHAO'].apply('{0:0>4}'.format)

                chave_talhao = arv.apply(lambda ct: '%s%s%s%s' % (
                    ct['CODTALHAO'], '/', ct['CODE'], ct['STAT']), axis=1)
                arv.insert(0, "KEY", chave_talhao, allow_duplicates=False)
                break
            elif caminho == 1:
                print('OS DADOS SERÃO PROCESSADOS E INSERIDOS:')
                print('IMPORTANDO ARQUIVO EXCEL')

                excel_to_merge = list()
                arv_excel = pd.DataFrame()
                arv = pd.DataFrame()

                tkinter.Tk()
                selected_files = tkinter.filedialog.askopenfilenames(
                    initialdir=r"Somewhere", title='Selecione os arquivos que deseja processar',
                    filetypes=(('All types', '*.*'), ('Excel files', '.xlsx')))

                for file in selected_files:
                    if file.endswith('.xlsx') or file.endswith('.xls'):
                        excel_to_merge.append(
                            pd.read_excel(fr'{file}', sheet_name=0))

                for excel_file in excel_to_merge:
                    arv_excel = pd.concat(
                        [arv_excel, excel_file], ignore_index=True)

                if arv_excel.isnull().values.any():
                    print(
                        '\033[31mERRO: Arquivo Excel importado apresenta colunas/linhas em branco!')
                    reiniciar()

                arv_excel['CODTALHAO'] = arv_excel['CODTALHAO'].apply(
                    '{0:0>4}'.format)

                arv_excel['DATAMEDICAO'] = pd.to_datetime(
                    arv_excel['DATAMEDICAO'], dayfirst=True)

                arv_excel.to_excel(r"Somewhere")

                sqlcols = pd.DataFrame()
                sqlverificacao = pd.DataFrame()

                print('VERIFICANDO EXISTÊNCIA DOS DADOS NO BANCO')

                try:
                    verificacao = pd.read_sql_query("query", con=engine_inv)

                    chavesql = verificacao.apply(
                        lambda cs: '%s%s%s%s' % (cs['PARC'], cs['codtalhao_arv'], cs[TE_arv'],
                                                 cs['CODE_arv']), axis=1)
                    sqlverificacao.insert(
                        0, "CHAVESQL", chavesql, allow_duplicates=False)

                    chaveexcel = arv_excel.apply(
                        lambda ce: '%s%s%s%s' % (ce['PARC'], ce['CODTALHAO'], 'STATE'],
                                                 ce['CODE']), axis=1)
                    sqlverificacao.insert(
                        1, "CHAVEEXCEL", chaveexcel, allow_duplicates = False)

                    if set(sqlverificacao['CHAVESQL']).intersection(set(sqlverificacao['CHAVEEXCEL'])):
                        print(
                            '\033[31mATENÇÃO! Dados já inseridos no Banco de Dados!\033[m')
                        while True:
                            pergunta=str(
                                input('Deseja continuar mesmo assim? [S/N]')).upper()
                            if pergunta == 'S':
                                print('\033[mCONTINUANDO...')
                                break
                            elif pergunta == 'N':
                                print('\033[mFINALIZADO')
                                reiniciar()
                            else:
                                print(
                                    '\033[31mOPÇÃO INVÁLIDA!\033[m Digite apenas S ou N')
                except ValueError:
                    pass

                print('CONSULTANDO COLUNAS AUXILIARES DO SQL')

                chave_talhao=arv_excel.apply(lambda ct: '%s%s%s%s' % (
                    ct['CODTALHAO'], '/', ct['CODE'], ct['STAT']), axis=1)
                arv_excel.insert(0, "KEY", chave_talhao,
                                 allow_duplicates=False)

                arv_to_merge= list()

                for y in tqdm(arv_excel.KEY.unique()):
                    dfmask= arv_excel['KEY'] == y
                    dadosarvchave= arv_excel[dfmask]
                    dadosarvfiltrado= pd.DataFrame(dadosarvchave)
                    dadosarvfiltrado= dadosarvfiltrado.reset_index(drop=True)

                    datacomp= pd.read_sql_query("query", con=engine_inv)

                    datacomp= pd.concat([datacomp] * len(dadosarvfiltrado), ignore_index=True)
                    arv_to_merge.append(
                        pd.concat([dadosarvfiltrado, datacomp], axis=1))

                for file in arv_to_merge:
                    arv= pd.concat([arv, file], ignore_index=True)
                arv.drop(['STATE', 'CODE', 'cd_talhao',
                         'cd_clone_semente'], axis=1, ace=True)
                break
            else:
                print('\033[31mOPÇÃO INVÁLIDA\033[m')

        df_arvore= pd.DataFrame(arv)

        df_arvore.columns= df_arvore.columns.str.upper()

        df_arvore['KEY']= df_arvore.apply(lambda ctn: '%s%s%s%s%s' % (ctn['PARC'], '-', ctn['ID_TALHAO'],
                                                                               '/', ctn['KEY']), axis=1)
        print('CALCULANDO HDOM')

        for y in tqdm(df_arvore.KEY.unique()):
            dfmask= df_arvore['KEY'] == y
            dadosarvchave= df_arvore[dfmask]
            dadosarvfiltrado= pd.DataFrame(dadosarvchave)
            div_num= 0

            if 'DT' not in dadosarvfiltrado['SITUATION'].values:
                daps= list(dadosarvfiltrado.DAP.values)
                four_majors= []

                for i in range(4):
                    div_num += 1
                    if len(daps) == 0:
                        break
                    max_dap_index= daps.index(max(daps))
                    four_majors.append(max(daps))
                    daps.remove(daps[max_dap_index])

                if len(daps) == 0:
                    df_arvore.loc[df_arvore.KEY == y, 'HDOM']= (sum(four_majors) / div_num).__round__(2)
                else:
                    df_arvore.loc[df_arvore.KEY == y, 'HDOM']= (sum(four_majors) / 4).__round__(2)
            else:
                res= dadosarvfiltrado.loc[df_arvore.SITUATION == 'DT', 'H'].mean().round(2)
                df_arvore.loc[df_arvore.KEY == y, 'HDOM']= res

        df_arvore= df_arvore.replace([np.inf, -np.inf], 0)
        df_arvore= df_arvore.fillna(0)

        print('CALCULANDO NUMMEDICAO E ULTIMAMEDICAO')

        df_nummedicao= pd.DataFrame(columns=['KEY'])
        df_nummedicao['KEY']= df_arvore.KEY.unique()

        cursor= engine_inv.connect()

        for i in tqdm(df_nummedicao.KEY.unique()):
            dfmask= df_arvore['KEY'] == i
            dadosarvchave= df_arvore[dfmask]
            dadosarvfiltrado= pd.DataFrame(dadosarvchave)

            invmed= pd.read_sql_query("query", con=engine_inv)

            if pd.isnull(invmed.nummedicao.max()) or pd.isnull(invmed.datarealizada.max()):
                df_arvore.loc[df_arvore.KEY == i, 'NUMMEDICAO']= 1

            elif invmed.datarealizada.max() == dadosarvfiltrado.DATAMEDICAO.max():
                df_arvore.loc[df_arvore.KEY == i, 'NUMMEDICAO']= invmed.nummedicao.max()

            elif invmed.datarealizada.max() < dadosarvfiltrado.DATAMEDICAO.max():
                df_arvore.loc[df_arvore.KEY == i, 'NUMMEDICAO']= invmed.nummedicao.max() + 1

                update_ultmed= f"""UPDATE"""
                cursor.execute(update_ultmed)

            elif invmed.datarealizada.max() > dadosarvfiltrado.DATAMEDICAO.max():
                df_arvore.loc[df_arvore.KEY == i, 'NUMMEDICAO']= 900

            if pd.isnull(invmed.datarealizada.max()):
                df_arvore.loc[df_arvore.KEY == i, 'ULTIMAMEDICAO']= 1

            elif invmed.datarealizada.max() > dadosarvfiltrado.DATAMEDICAO.max():
                df_arvore.loc[df_arvore.KEY == i, 'ULTIMAMEDICAO']= 0

            else:
                df_arvore.loc[df_arvore.KEY == i, 'ULTIMAMEDICAO']= 1

        cursor.close()

        print('CALCULANDO COLUNAS')

        for i in trange(len(df_arvore)):
            query_hest= pd.read_sql_query(f"query", con=engine_inv)

            y= float(query_hest.values)
            df_arvore.loc[df_arvore.index == i, 'HEST']= y

            if df_arvore.DAP.iloc[i] == 0:
                df_arvore.loc[df_arvore.index == i, 'HEST_P']= 0
            elif df_arvore.H.iloc[i] == 0:
                query_hest_p= pd.read_sql_query("query", con=engine_inv)

                x= float(query_hest_p.values)
                df_arvore.loc[df_arvore.index == i, 'HEST_P']= x
            else:
                df_arvore.loc[df_arvore.index == i, 'HEST_P']= df_arvore.H.iloc[i]

            g= math.pi * (df_arvore.DAP.iloc[i] ** 2) / 40000
            df_arvore.loc[df_arvore.index == i, 'G']= g

            if df_arvore.HEST.iloc[i] != 0:
                query_vttvind= pd.read_sql_query("query", con=engine_inv)

                y= float(query_vttvind.values)
                df_arvore.loc[df_arvore.index == i, 'VTTVIND']= y
            else:
                df_arvore.loc[df_arvore.index == i, 'VTTVIND']= 0

            if df_arvore.DAP.iloc[i] == 0:
                df_arvore.loc[df_arvore.index == i, 'FATORFORMAARV']= 0
            else:
                volume_cilindro= (math.pi * '''some_calcs''' * df_arvore.HEST.iloc[i])
                fatorformaarv= df_arvore.VTTVIND.iloc[i] / volume_cilindro
                df_arvore.loc[df_arvore.index == i, 'FATORFORMAARV']= fatorformaarv

        df_arvore= df_arvore.rename(columns={'rename some columns'})
        df_arvore.to_excel(r"Somewhere")

        print('\033[0mCRIANDO TABELA EM NÍVEL PARCELA')
        calcmedicao= pd.DataFrame(columns=['KEY'])
        df_parcela= pd.DataFrame()

        calcmedicao['KEY']= df_arvore.KEY.unique()
        df_parcela['KEY']= df_arvore.KEY.unique()

        for i in tqdm(calcmedicao.KEY.unique()):
            dfmask= df_arvore['KEY'] == i
            dadosarvchave= df_arvore[dfmask]
            dadosarvfiltrado= pd.DataFrame(dadosarvchave)

            maxnumarv= dadosarvfiltrado['SITUATION'].loc[dadosarvfiltrado['SITUATION'] != 'F'][
                dadosarvfiltrado['NUMARVORE'] < 200].count()

            maxnumarvF= dadosarvfiltrado['SITUATION'].loc[dadosarvfiltrado['SITUATION'] != 'F'].count()

            calcmedicao.loc[
                calcmedicao.KEY == i, 'STATE_ARV']= dadosarvfiltrado.STATE.unique()

            calcmedicao.loc[
                calcmedicao.KEY == i, 'CODTALHAO_ARV']= dadosarvfiltrado.CODTALHAO_ARV.unique()

            # Other calculations...

        calcmedicao= calcmedicao.replace([np.inf, -np.inf], 0)
        calcmedicao= calcmedicao.fillna(0)

        for y in trange(len(calcmedicao.KEY.unique())):
            datamed= datetime.date(calcmedicao['DATA_INV'].iloc[y])
            datainirot= datetime.date(calcmedicao['DATE'].iloc[y])
            datenow= datetime.now().date()

            id_inv= datamed - datainirot
            daysInv= float(id_inv.days) / 365

            id_atual= datenow - datainirot
            daysAtual= float(id_atual.days) / 365

            calcmedicao.loc[calcmedicao.index == y, 'ID_INV']= daysInv.__round__(2)
            calcmedicao.loc[calcmedicao.index == y, 'ID_INV_MEDIA']= daysInv.__round__(0)
            calcmedicao.loc[calcmedicao.index == y, 'ID_ATUAL']= daysAtual.__round__(2)
            calcmedicao.loc[calcmedicao.index == y, 'ANOREF']= datenow.year

        for i in trange(len(calcmedicao)):
            try:
                query_proj= pd.read_sql_query("query", con=engine_ip)

                query_proj= pd.read_sql_query("query", con=engine_ip)

                query_PRED= pd.read_sql_query("query", con=engine_ip)

                PRED= float(query_PRED.values)
                calcmedicao.loc[calcmedicao.KEY == calcmedicao.KEY.iloc[i], 'VOLPRED']= PRED

                IMA= (float(query_proj.values) + calcmedicao.VTANP.iloc[i]) / 7
                calcmedicao.loc[calcmedicao.KEY == calcmedicao.KEY.iloc[i], 'IMA']= IMA

                if calcmedicao.NUMMEDICAO.iloc[i] - 1 == 0:
                    calcmedicao.loc[calcmedicao.KEY == calcmedicao.KEY.iloc[i], 'PROJ']= 0
                else:
                    query_sql= pd.read_sql_query(f'''query''', con=engine_inv)

                    query_proj= pd.read_sql_query("query", con=engine_ip)

                    PROJ= float(query_proj.values)
                    calcmedicao.loc[calcmedicao.KEY == calcmedicao.KEY.iloc[i], 'PROJ']= PROJ
            except Exception as e:
                print(str(f'\033[31mERRO: {str(e)}\033[m'))

        calcmedicao.to_excel(r"Somewhere)

        print('\033[0mCRIANDO TABELA EM NÍVEL TALHÃO')

        calcmedicao['KEY']= calcmedicao['KEY'].str.split('-').str[1]

        calcrotacao= pd.DataFrame(columns=['KEY'])
        calcrotacao['KEY']= calcmedicao.KEY.unique()

        for i in tqdm(calcrotacao.KEY.unique()):
            dfmask= calcmedicao['KEY'] == i
            dadosparccod= calcmedicao[dfmask]
            dadosparcfiltrado= pd.DataFrame(dadosparccod)

            calcrotacao.loc[
                calcrotacao.KEY == i, 'STATE']= dadosparcfiltrado.STATE.unique()
            calcrotacao.loc[calcrotacao.KEY == i, 'CODE']= dadosparcfiltrado.CODE_ARV.unique()
            calcrotacao.loc[calcrotacao.KEY == i, 'NAME']= dadosparcfiltrado.NAME.unique()
            calcrotacao.loc[calcrotacao.KEY == i, 'FOREST']= dadosparcfiltrado.FOREST.unique()

        calcrotacao.to_excel(r"Somewhere")

        print('ENVIANDO PARA O SQL:')

        cursor= engine_inv.connect()

        if dados_existentes == 0:
            print('INSERINDO INVPARCELA')
            for i in trange(len(df_parcela)):
                try:
                    sql= f"""INSERT INTO"""
                    cursor.execute(sql)
                except IntegrityError as e:
                    if 'ORA-00001' in str(e):
                        pass
                    else:
                        print(f'\033[31mLinha: {i} NÃO inserida na tabela invparcela!\n'
                              f'Motivo: {str(e)}\n')
                        cursor.close()
                        reiniciar()
                except Exception as e:
                    print('\n\033[31mERRO: ', e)
                    reiniciar()

            print('INSERINDO INVMEDICAO')
            for i in trange(len(calcmedicao)):
                if calcmedicao.DATAPREVISTA.iloc[i] == 0:
                    try:
                        sql= f"""INSERT INTO"""
                        cursor.execute(sql)
                    except IntegrityError as e:
                        if 'ORA-00001' in str(e):
                            pass
                        else:
                            print(f'\033[31mLinha: {i} NÃO inserida na tabela invmedicao!\n'
                                  f'Motivo: {str(e)}\n')
                            cursor.close()
                            reiniciar()
                    except Exception as e:
                        print('\n\033[31mERRO: ', e)
                        reiniciar()
                else:
                    try:
                        sql= f"""INSERT INTO"""
                        cursor.execute(sql)
                    except IntegrityError as e:
                        if 'ORA-00001' in str(e):
                            pass
                        else:
                            print(f'\033[31mLinha: {i} NÃO inserida na tabela invmedicao!\n'
                                  f'Motivo: {str(e)}\n')
                            cursor.close()
                            reiniciar()
                    except Exception as e:
                        print('\n\033[31mERRO: ', e)
                        reiniciar()

            print('INSERINDO INVCALCMEDICAO')
            for i in trange(len(calcmedicao)):
                try:
                    sql= f"""INSERT INTO"""
                    cursor.execute(sql)
                except IntegrityError as e:
                    if 'ORA-00001' in str(e):
                        pass
                    else:
                        print(f'\033[31mLinha: {i} NÃO inserida na tabela invcalcmedicao!\n'
                              f'Motivo: {str(e)}\n')
                        cursor.close()
                        reiniciar()
                except Exception as e:
                    print('\n\033[31mERRO: ', e)
                    reiniciar()

            print('INSERINDO INVDADOARVORE')
            for i in trange(len(df_arvore)):
                try:
                    sql= f"""INSERT INTO"""
                    cursor.execute(sql)
                except IntegrityError as e:
                    if 'ORA-00001' in str(e):
                        pass
                    else:
                        print(f'\033[31mLinha: {i} NÃO inserida na tabela invdadoarvore!\n'
                              f'Motivo: {str(e)}\n')
                        cursor.close()
                        reiniciar()
                except Exception as e:
                    print('\n\033[31mERRO: ', e)
                    reiniciar()

            print('INSERINDO INVCALCROTACAO')
            for i in trange(len(calcrotacao)):
                try:
                    sql= f"""INSERT INTO"""
                    cursor.execute(sql)
                except IntegrityError as e:
                    if 'ORA-00001' in str(e):
                        try:
                            update_sql= f"""UPDATE"""
                            cursor.execute(update_sql)
                        except Exception as e:
                            print(str(e))
                            print(f'\033[31mLinha {i} NÃO ATUALIZADA!\033[m')
                    else:
                        print(f'\033[31mLinha: {i} NÃO inserida na tabela invcalcrotacao!\n'
                              f'Motivo: {str(e)}\n')
                        cursor.close()
                        reiniciar()
                except Exception as e:
                    print('\n\033[31mERROR: ', e)
                    reiniciar()
        else:
            print('ATUALIZANDO INVMEDICAO')
            for i in trange(len(calcmedicao)):
                try:
                    update_sql= f"""UPDATE"""
                    cursor.execute(update_sql)
                except Exception as e:
                    print(str(e))
                    print(f'\033[31mLinha {i} NÃO ATUALIZADA!\033[m')

            print('ATUALIZANDO INVCALCMEDICAO')
            for i in trange(len(calcmedicao)):
                try:
                    update_sql = f"""UPDATE"""
                    cursor.execute(update_sql)
                except Exception as e:
                    print(str(e))
                    print(f'\033[31mLinha {i} NÃO ATUALIZADA!\033[m')

            print('ATUALIZANDO INVDADOARVORE')

            for i in trange(len(df_arvore)):
                try:
                    update_sql= f"""UPDATE"""
                    cursor.execute(update_sql)
                except Exception as e:
                    print(str(e))
                    print(f'\033[31mLinha {i} NÃO ATUALIZADA!\033[m')

            print('ATUALIZANDO INVCALCROTACAO')
            for i in trange(len(calcrotacao)):
                try:
                    update_sql= f"""UPDATE"""
                    cursor.execute(update_sql)
                except Exception as e:
                    print(str(e))
                    print(f'\033[31mLinha {i} NÃO ATUALIZADA!\033[m')

        cursor.close()

        # convert UTM to WGS84
        unidades_convertidas= utm.to_latlon(calcmedicao.LONGITUDE.iloc[0], calcmedicao.LATITUDE.iloc[0], 23, 'K')
        # Set the map location
        mapa = folium.Map(location =unidades_convertidas, zoom_start=10, control_scale=True)
        # Connect with ArcGIS server
        folium.TileLayer(tiles ='https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}', attr='Google',
                         name='Google Earth Satellite', overlay=True, control=True).add_to(mapa)

        # for se repetirá pela número de longitudes únicas que houver em df_parcela
        for i in range(len(calcmedicao)):
            # Convert longitude and latitude to WGS84
            unidades_convertidas = utm.to_latlon(calcmedicao.LONGITUDE.iloc[i], calcmedicao.LATITUDE.iloc[i], 23, 'K')
            # HTML code to create table
            html = popup_html(i)
            # Table size
            iframe = branca.element.IFrame(html=html, width=510, height=280)
            # Popup creation
            popup = folium.Popup(folium.Html(html, script=True), max_width=500)
            # Add marker to the map locations
            folium.Marker(location=unidades_convertidas, popup=popup, icon=folium.Icon(color="green", icon='info-sign'),
                          tooltip='Click Me').add_to(mapa)

        # Save the map and open it
        mapa.save('mapa.html'), webbrowser.open('mapa.html')
    except (DatabaseError, ValueError) as e:
        if 'ORA-00904' in str(e):
            print('\033[31mERROR: "NAN" encontrado\n'
                  'Possível Causa: Talhão Dividido/Cortado\n'
                  f'Descrição: {str(e)}\033[m')
        else:
            print(f'\033[31mERROR: {str(e)}\033[m')
    except Exception as e:
        print(f'\033[31mERROR: {e}\033[m')

    reiniciar()
