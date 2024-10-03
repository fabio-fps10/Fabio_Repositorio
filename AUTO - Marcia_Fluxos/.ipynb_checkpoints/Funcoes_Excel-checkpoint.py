import pandas as pd
import os
import locale
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill
from openpyxl import load_workbook


def read_arquivo(mes_atual):
    df_fluxo_cru = pd.read_excel(f"V:\\Vendas\\Acompanhamento Venda Semanais\\{mes_atual}_2024\\FLUXOS.xlsx")
    return df_fluxo_cru

#### Função Main

def FUNCAO_AUTO_FULXO_MARCIA (mes_atual):
    df_fluxo = read_arquivo(mes_atual)    
    return df_fluxo

