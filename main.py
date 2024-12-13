from datetime import datetime, time, timedelta
import subprocess
import pandas as pd
import mysql.connector
import win32com.client as win32
import pymysql
import tkinter as tk
from pathlib import Path
import os
import time
import subprocess
import sys
from dotenv import load_dotenv

load_dotenv()

# Configurações do pandas para exibir todas as colunas e largura ajustada
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.max_colwidth', None)

hj = datetime.today().date()
ontem = (datetime.now() - timedelta(days=1)).date()
amanha = hj + timedelta(days=1)
data_mysql_hj = hj.strftime('%Y-%m-%d')
data_mysql_am = amanha.strftime('%Y-%m-%d')

data_formatada_hj = hj.strftime('%d/%m/%Y')
data_formatada_am = amanha.strftime('%d/%m/%Y')

def apagar_arquivo():
    time.sleep(5)
    pasta_downloads = str(Path.home() / "Downloads")
    
    nome_arquivo = "APP Solicita Preço.xlsx"  
    
    caminho_arquivo = os.path.join(pasta_downloads, nome_arquivo)
    
    if os.path.exists(caminho_arquivo):
        os.remove(caminho_arquivo)
        print(f"O arquivo {nome_arquivo} foi apagado com sucesso.")
    else:
        print(f"O arquivo {nome_arquivo} não foi encontrado na pasta de downloads.")

def apagar_arquivo2():
    caminho_arquivo = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\dados_resumidos.xlsx'
    
    if os.path.exists(caminho_arquivo):
        try:
            os.remove(caminho_arquivo)
            print(f"Arquivo {caminho_arquivo} apagado com sucesso!")
        except Exception as e:
            print(f"Erro ao apagar o arquivo: {e}")
    else:
        print(f"O arquivo {caminho_arquivo} não foi encontrado.")

def atualizar_variaveis(botao_id):
    global caminho, aba, qloja, lojaexcel
    if botao_id == "001":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_001.xlsx'
        aba = "BD_001_SEGUNDA_ITENS"
        qloja = "SED,001"
    elif botao_id == "002":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_002.xlsx'
        aba = "BD_002_SEGUNDA_ITENS"
        qloja = "SED,002"
    elif botao_id == "003":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,003"
    elif botao_id == "004":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_004.xlsx'
        aba = "BD_004_SEGUNDA_ITENS"
        qloja = "SED,004"
    elif botao_id == "005":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_005.xlsx'
        aba = "BD_005_SEGUNDA_ITENS"
        qloja = "SED,005"
    elif botao_id == "006":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_006.xlsx'
        aba = "BD_006_SEGUNDA_ITENS"
        qloja = "SED,006"
    elif botao_id == "007":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_007.xlsx'
        aba = "BD_007_SEGUNDA_ITENS"
        qloja = "SED,007"
    elif botao_id == "008":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_008.xlsx'
        aba = "BD_008_SEGUNDA_ITENS"
        qloja = "SED,008"
    elif botao_id == "009":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_009.xlsx'
        aba = "BD_009_SEGUNDA_ITENS"
        qloja = "SED,009"
    elif botao_id == "010":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_010.xlsx'
        aba = "BD_010_SEGUNDA_ITENS"
        qloja = "SED,010"
    elif botao_id == "011":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_011.xlsx'
        aba = "BD_011_SEGUNDA_ITENS"
        qloja = "SED,011"
    elif botao_id == "012":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_012.xlsx'
        aba = "BD_012_SEGUNDA_ITENS"
        qloja = "SED,012"
    elif botao_id == "013":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_013.xlsx'
        aba = "BD_013_SEGUNDA_ITENS"
        qloja = "SED,013"
    elif botao_id == "014":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_014.xlsx'
        aba = "BD_014_SEGUNDA_ITENS"
        qloja = "SED,014"
    elif botao_id == "015":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_015.xlsx'
        aba = "BD_015_SEGUNDA_ITENS"
        qloja = "SED,015"
    elif botao_id == "016":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_016.xlsx'
        aba = "BD_016_SEGUNDA_ITENS"
        qloja = "SED,016"
    elif botao_id == "017":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_017.xlsx'
        aba = "BD_017_SEGUNDA_ITENS"
        qloja = "SED,017"
    elif botao_id == "001 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,001"
        lojaexcel = "1"
        app()
    elif botao_id == "002 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,002"
        lojaexcel = "2"
        app()
    elif botao_id == "003 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,003"
        lojaexcel = "3"
        app()
    elif botao_id == "004 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,004"
        lojaexcel = "4"
        app()
    elif botao_id == "005 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,005"
        lojaexcel = "5"
        app()
    elif botao_id == "006 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,006"
        lojaexcel = "6"
        app()
    elif botao_id == "007 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,007"
        lojaexcel = "7"
        app()
    elif botao_id == "008 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,008"
        lojaexcel = "8"
        app()
    elif botao_id == "009 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,009"
        lojaexcel = "9"
        app()
    elif botao_id == "010 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,010"
        lojaexcel = "10"
        app()
    elif botao_id == "011 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,011"
        lojaexcel = "11"
        app()
    elif botao_id == "012 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,012"
        lojaexcel = "12"
        app()
    elif botao_id == "013 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,013"
        lojaexcel = "13"
        app()
    elif botao_id == "014 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,014"
        lojaexcel = "14"
        app()
    elif botao_id == "015 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,015"
        lojaexcel = "15"
        app()
    elif botao_id == "016 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,016"
        lojaexcel = "16"
        app()
    elif botao_id == "017 App":
        caminho = r'\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\BD_SOLICITACOES_LOJAS\BD_SOLICITACOES_003_APP.xlsx'
        aba = "BD_003_SEGUNDA_ITENS"
        qloja = "SED,017"
        lojaexcel = "17"
        app()
    
    print(f"\nLoja selecionada {botao_id}")
    if botao_id == "001 App" or botao_id == "002 App" or botao_id == "003 App" or botao_id == "004 App" or botao_id == "005 App" or botao_id == "006 App" or botao_id == "007 App" or botao_id == "008 App" or botao_id == "009 App" or botao_id == "010 App" or botao_id == "011 App" or botao_id == "012 App" or botao_id == "013 App" or botao_id == "014 App" or botao_id == "015 App" or botao_id == "016 App" or botao_id == "017 App":

        processar_excel()
        att_plan()
        horti_inserir()
        hoje_inserir()
        sai_amanha_inserir()
        apagar_arquivo()
        apagar_arquivo2()
    else:
        att_plan()
        horti_inserir()
        hoje_inserir()
        sai_amanha_inserir()

def conectar_mysql():   
    try:
        conn = pymysql.connect(
    host=os.getenv('host'),
    user=os.getenv('user'),
    password=os.getenv('password'),
    database=os.getenv('database')
    )

        return conn
    except mysql.connector.Error as err:
        print(f"Erro ao se conectar ao MySQL: {err}")
        return None

def att_plan():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    
    excel.Visible = True
    
    try:
        workbook = excel.Workbooks.Open(caminho)
        excel.DisplayAlerts = False
        workbook.Save()
    
    except Exception as e:
        print(f"Erro: {e}")
    
    finally:
        workbook.Close(SaveChanges=True)
        excel.Quit()

def horti_inserir():
    try:
        df_segunda = pd.read_excel(
            caminho,
            sheet_name=aba
        )
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return
    if not any(df_segunda['finalidade'] == 'X'):
        print("Não contém itens do horti.")
        return
    connection = conectar_mysql()
    if not connection:
        return
    try:
        cursor = connection.cursor()
        insert_promocoes = """
            INSERT INTO prc_promocoes (
                Descricao, DataInicio, DataFim, Codigo, Observacoes, OkBdc, Lojas,
                DataFimCompra, DataInicioCompra, spack, Limite, packs, 
                finalidade_padrao, seloutfixocalculado, removeoferta, hora
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(
            insert_promocoes,
            (
                'EXC HORTI', data_mysql_hj, data_mysql_hj, '', '', 'N',
                qloja, '', '', '0', '', '', 'X', '0', '0', '00:00:00'
            )
        )
        connection.commit()

        cursor.execute("SELECT LAST_INSERT_ID();")
        codigo_gerado = cursor.fetchone()[0]
        print(f"Código gerado: {codigo_gerado}")

        df_horti = df_segunda[df_segunda['e_horti'] == 'Sim']
        df_horti = df_horti.drop_duplicates(subset=['CODIGOINT'])
        
        dados_para_inserir = []

        for _, row in df_horti.iterrows():

            linha = [codigo_gerado]

            codigo_int = str(int(row['CODIGOINT'])).zfill(7)
            linha.append(codigo_int)

            for col in df_horti.columns[2:39]:
                valor = row[col] if pd.notna(row[col]) else None
                linha.append(valor)

            dados_para_inserir.append(linha)

        try:
            insert_itens = """
    INSERT INTO prc_promocaoitens (
        CodPromocao, CODIGOINT, QtdEstimada, VlrVenda, CODIGOFORNEC, VALCOMPRA,
        custocor, VlrVendaNormal, midia, Local, OkBdc, Pr_bonificacao,
        tppromocao, qtdgatilho, codproddesconto, PrFinalDesconto,
        semprepack, Etiqueta, margempromo, UltAtu, fatorapuracaodesconto,
        avisoqtdultrapassada, finalidade, VendaEstimada, ppack, dtvenc,
        formapgto, formapagto, vlrss, tpapuracao, selinaut,
        selinselaut, blqprompalm_it, prom_midia_elet, removeoferta,
        limite, precoclube, margemprecoclube, qtdemb
    ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
        %s, %s, %s, %s, %s, %s, %s, %s, %s
    )
"""
            cursor.executemany(insert_itens, dados_para_inserir)
            connection.commit()
            print("Dados inseridos em prc_promocaoitens com sucesso!")
        except mysql.connector.Error as err:
            print(f"Erro MySQL: {err}")
            connection.rollback()
        except Exception as e:
            print(f"Erro inesperado: {e}")

    finally:
            cursor.close()
            connection.close()

def hoje_inserir():
    try:
        df_segunda = pd.read_excel(
            caminho,
            sheet_name=aba
        )
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return

    if not any((df_segunda['finalidade'] == 'V') & (df_segunda['sai_hoje'] == 'Sim')):
        print("Não contém itens para sair hoje.")
        return

    connection = conectar_mysql()
    if not connection:
        return

    try:
        cursor = connection.cursor()
        insert_promocoes = """
            INSERT INTO prc_promocoes (
                Descricao, DataInicio, DataFim, Codigo, Observacoes, OkBdc, Lojas,
                DataFimCompra, DataInicioCompra, spack, Limite, packs, 
                finalidade_padrao, seloutfixocalculado, removeoferta, hora
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(
            insert_promocoes,
            (
                'DATA CURTA', data_mysql_hj, data_mysql_hj, '', '', 'N',
                qloja, '', '', '0', '', '', 'V', '0', '0', '00:00:00'
            )
        )
        connection.commit()

        cursor.execute("SELECT LAST_INSERT_ID();")
        codigo_gerado = cursor.fetchone()[0]
        print(f"Código gerado: {codigo_gerado}")

        df_hoje = df_segunda[(df_segunda['finalidade'] == 'V') & (df_segunda['sai_hoje'] == 'Sim')]
        df_hoje = df_hoje.drop_duplicates(subset=['CODIGOINT'])

        dados_para_inserir = []

        for _, row in df_hoje.iterrows():
            
            linha = [codigo_gerado]

            codigo_int = str(int(row['CODIGOINT'])).zfill(7)
            linha.append(codigo_int)

            for col in df_hoje.columns[2:39]:
                valor = row[col] if pd.notna(row[col]) else None
                linha.append(valor)

            dados_para_inserir.append(linha)

        try:
            insert_itens = """
                INSERT INTO prc_promocaoitens (
                    CodPromocao, CODIGOINT, QtdEstimada, VlrVenda, CODIGOFORNEC, VALCOMPRA,
                    custocor, VlrVendaNormal, midia, Local, OkBdc, Pr_bonificacao,
                    tppromocao, qtdgatilho, codproddesconto, PrFinalDesconto,
                    semprepack, Etiqueta, margempromo, UltAtu, fatorapuracaodesconto,
                    avisoqtdultrapassada, finalidade, VendaEstimada, ppack, dtvenc,
                    formapgto, formapagto, vlrss, tpapuracao, selinaut,
                    selinselaut, blqprompalm_it, prom_midia_elet, removeoferta,
                    limite, precoclube, margemprecoclube, qtdemb
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                    %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
            """
            cursor.executemany(insert_itens, dados_para_inserir)
            connection.commit()
            print("Dados inseridos em prc_promocaoitens com sucesso!")
        except mysql.connector.Error as err:
            print(f"Erro MySQL: {err}")
            connection.rollback()
        except Exception as e:
            print(f"Erro inesperado: {e}")

    finally:
        cursor.close()
        connection.close()

def sai_amanha_inserir():
    try:
        df_segunda = pd.read_excel(
            caminho,
            sheet_name=aba
        )
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return

    if not any((df_segunda['finalidade'] == 'V') & (df_segunda['sai_hoje'] == 'Nao')):
        print("Não contém itens para sair com mais de um dia.")
        return

    connection = conectar_mysql()
    if not connection:
        return

    try:
        cursor = connection.cursor()
        insert_promocoes = """
            INSERT INTO prc_promocoes (
                Descricao, DataInicio, DataFim, Codigo, Observacoes, OkBdc, Lojas,
                DataFimCompra, DataInicioCompra, spack, Limite, packs, 
                finalidade_padrao, seloutfixocalculado, removeoferta, hora
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(
            insert_promocoes,
            (
                'DATA CURTA', data_mysql_hj, data_mysql_am, '', '', 'N',
                qloja, '', '', '0', '', '', 'V', '0', '0', '00:00:00'
            )
        )
        connection.commit()

        cursor.execute("SELECT LAST_INSERT_ID();")
        codigo_gerado = cursor.fetchone()[0]
        print(f"Código gerado: {codigo_gerado}")

        df_hoje = df_segunda[(df_segunda['finalidade'] == 'V') & (df_segunda['sai_hoje'] == 'Nao')]
        df_hoje = df_hoje.drop_duplicates(subset=['CODIGOINT'])

        dados_para_inserir = []

        for _, row in df_hoje.iterrows():

            linha = [codigo_gerado]

            codigo_int = str(int(row['CODIGOINT'])).zfill(7)
            linha.append(codigo_int)

            for col in df_hoje.columns[2:39]:
                valor = row[col] if pd.notna(row[col]) else None
                linha.append(valor)

            dados_para_inserir.append(linha)

        try:
            insert_itens = """
                INSERT INTO prc_promocaoitens (
                    CodPromocao, CODIGOINT, QtdEstimada, VlrVenda, CODIGOFORNEC, VALCOMPRA,
                    custocor, VlrVendaNormal, midia, Local, OkBdc, Pr_bonificacao,
                    tppromocao, qtdgatilho, codproddesconto, PrFinalDesconto,
                    semprepack, Etiqueta, margempromo, UltAtu, fatorapuracaodesconto,
                    avisoqtdultrapassada, finalidade, VendaEstimada, ppack, dtvenc,
                    formapgto, formapagto, vlrss, tpapuracao, selinaut,
                    selinselaut, blqprompalm_it, prom_midia_elet, removeoferta,
                    limite, precoclube, margemprecoclube, qtdemb
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                    %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
            """
            cursor.executemany(insert_itens, dados_para_inserir)
            connection.commit()
            print("Dados inseridos em prc_promocaoitens com sucesso!")
        except mysql.connector.Error as err:
            print(f"Erro MySQL: {err}")
            connection.rollback()
        except Exception as e:
            print(f"Erro inesperado: {e}")

    finally:
        cursor.close()
        connection.close()

def app():
    link_download = os.getenv('link')
    
    subprocess.run(['start', 'chrome', link_download], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def painel():
    
    icon_path = resource_path("icon.ico")

    root = tk.Tk()
    root.title("Painel DATA CURTA")
    root.geometry("650x700")
    root.configure(bg="#0078A1")

    root.iconbitmap(icon_path)

    frame = tk.Frame(root, bg="#0078A1")
    frame.pack(pady=20, padx=20, expand=True)

    botao_ids = [f"{i:03}" for i in range(1, 18)]
    botao_ids_app = [f"{i:03} App" for i in range(1, 18)]


    def estilo_botao(botao, tamanho="normal"):

        largura = 12 if tamanho == "normal" else 10
        altura = 2 if tamanho == "normal" else 1
        
        botao.config(
            width=largura,
            height=altura,
            font=("Arial", 12),
            relief="flat",
            bg="white",
            fg="black",
            activebackground="#e0e0e0",
            activeforeground="black",
            highlightbackground="#005F7A",
            highlightthickness=2,
            padx=5,
            pady=5
        )

        botao.bind("<Enter>", lambda e: botao.config(bg="#f0f0f0"))
        botao.bind("<Leave>", lambda e: botao.config(bg="white"))

    frame1 = tk.Frame(frame, bg="#0078A1")
    frame1.grid(row=0, column=0, padx=20, pady=10)

    frame2 = tk.Frame(frame, bg="#0078A1")
    frame2.grid(row=0, column=1, padx=20, pady=10)

    row_idx = 0
    col_idx = 0
    for botao_id in botao_ids:
        botao = tk.Button(frame1, text=botao_id, command=lambda id=botao_id: atualizar_variaveis(id))
        
        estilo_botao(botao, tamanho="pequeno")
        
        botao.grid(row=row_idx, column=col_idx, padx=10, pady=10)

        col_idx += 1
        if col_idx == 2:
            col_idx = 0
            row_idx += 1

    row_idx = 0
    col_idx = 0
    for botao_id in botao_ids_app:
        botao = tk.Button(frame2, text=botao_id, command=lambda id=botao_id: atualizar_variaveis(id))
        
        estilo_botao(botao, tamanho="pequeno")
        
        botao.grid(row=row_idx, column=col_idx, padx=10, pady=10)

        col_idx += 1
        if col_idx == 2:
            col_idx = 0
            row_idx += 1

    root.mainloop()

def obter_caminho_downloads():
    caminho_usuario = Path(os.getenv('USERPROFILE'))
    caminho_downloads = caminho_usuario / "Downloads"
    print(f"Caminho para a pasta Downloads: {caminho_downloads}")
    return caminho_downloads

def processar_excel():
    time.sleep(2)
    try:
        caminho_entrada = Path.home() / "Downloads" / "APP Solicita Preço.xlsx"
        
        print(f"Caminho do arquivo Excel: {caminho_entrada}")
        if not caminho_entrada.exists():
            print(f"Arquivo não encontrado: {caminho_entrada}")
            return

        df = pd.read_excel(caminho_entrada, sheet_name='bd_solicitação')

        df_filtrado = df[df['Id_Loja'] == int(lojaexcel)].copy()

        if df_filtrado.empty:
            print("Nenhum dado encontrado para a loja especificada.")
            return

        df_filtrado.loc[:, 'Data'] = pd.to_datetime(df_filtrado['Data'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        df_filtrado.loc[:, 'Data_Hora'] = pd.to_datetime(df_filtrado['Data_Hora'], format='%H:%M:%S', errors='coerce')

        df_filtrado.dropna(subset=['Data', 'Data_Hora'], inplace=True)

        hoje = datetime.today().date()
        ont = hoje - timedelta(days=1)

        dados_resumidos = []
        produtos_vistos = set()

        for _, row in df_filtrado.iterrows():
            produto = row['Id_Produto']

            if produto in produtos_vistos:
                continue

            if row['Data'].date() == hoje:
                dados_resumidos.append([produto, row['Quantidade'], row['Validade']])
                produtos_vistos.add(produto)

            elif row['Data'].date() == ont:
                hora = row['Data_Hora'].time()
                if hora >= datetime.strptime("12:00", "%H:%M").time():
                    dados_resumidos.append([produto, row['Quantidade'], row['Validade']])
                    produtos_vistos.add(produto)

        if not dados_resumidos:
            print("Não há dados que atendem às condições.")
            return
            
        df_resumido = pd.DataFrame(dados_resumidos, columns=['Id_Produto', 'Quantidade', 'Validade'])

        df_resumido['Validade'] = pd.to_datetime(df_resumido['Validade']).dt.strftime('%d/%m/%Y')

        caminho_saida = Path(r"\\192.168.1.243\samba\Metas\INTELIGENCIA\EDUARDO\dados_resumidos.xlsx")

        df_resumido.to_excel(caminho_saida, index=False)

        print(f"Arquivo salvo com sucesso em: {caminho_saida}")

    except Exception as e:
        print(f"Erro ao processar o Excel: {e}")

painel()