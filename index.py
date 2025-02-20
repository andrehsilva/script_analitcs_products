# -*- coding: utf-8 -*-
"""analytics_por_produtos_V3.ipynb

Automatically generated by Colab.

Original file is located at
    https://colab.research.google.com/drive/1jaddl6mQXzG1lgxG7KtChtC0JzKPIEeY

Leitura e tratamento da planilha: LEX-Usuarios_com_suas_turmas
"""

!pip install xlsxwriter

"""Import + paths"""

import pandas as pd
from datetime import datetime, date
import numpy as np
import time
import xlsxwriter
import re
import glob

today = date.today().strftime('%d-%m-%Y')

caminho_input = '/content/drive/MyDrive/Colab Notebooks/analytics-squadlex/input'
caminho_output = '/content/drive/MyDrive/Colab Notebooks/analytics-squadlex/output'

"""Funções"""

def up(data):
    data.columns = data.columns.str.upper()
    for columns in data.columns:
        if data[columns].dtype == 'object': # Check if the column is of object type (often used for strings)
            data[columns] = data[columns].str.upper() # Only apply .str.upper() to string columns
    return data

def lower(data):
    data.columns = data.columns.str.lower()
    for columns in data.columns:
        data[columns] = data[columns].str.lower()
    return data

def remover_caracteres_especiais(texto):
    texto_sem_especiais = re.sub(r'[^\w\s]', '', texto)
    return texto_sem_especiais

def carregar_e_concatenar_csv(caminho_input, padrao_arquivo):
    # Encontra os arquivos que correspondem ao padrão especificado
    arquivos = glob.glob(f'{caminho_input}/{padrao_arquivo}*.csv')

    if arquivos:
        # Lê e concatena todos os DataFrames em um único DataFrame
        df_concatenado = pd.concat([pd.read_csv(arquivo, sep=';', low_memory=False) for arquivo in arquivos], ignore_index=True)
        return df_concatenado
    else:
        print("Nenhum arquivo CSV encontrado.")
        return pd.DataFrame()  # Retorna um DataFrame vazio se nenhum arquivo for encontrado

"""Listas"""

modulo_comunicacao = ['BINÓCULO - MÓDULO DE COMUNICAÇÃO 2024','HELLO CAROLINA PATRÍCIO - 2024','HELLO PUERI CANDANGUINHO - 2024','HELLO PUERI DOMUS - 2024','HELLO SPHERE - 2024']
educacross = ['EDUCACROSS 2024','EDUCACROSS AZ 2024','EDUCACROSS HIGH FIVE 2024']
produtos = ['ÁRVORE','EDUCACROSS','EDUCACROSS HIGH FIVE','UBBU','BINÓCULO BY TELLME','HELLO CAROLINA PATRÍCIO','HELLO PUERI CANDANGUINHO','HELLO PUERI DOMUS', 'HELLO SPHERE','TINDIN','SCHOOL GUARDIAN','SCHOLASTIC LP + LPL BUNDLE','COMPANION SITE EARLY YEARS - STUDENT','SCHOLASTIC PR1ME MATHEMATICS', 'SCHOLASTIC EARLY BIRD','SCHOLASTIC LITERACY PRO COMPREHENSION SKILLS TEST','SCHOLASTIC BOOKFLIX', 'TINDIN',]

"""TABELAS AUXILIARES"""

grade = pd.read_excel(f'{caminho_input}/grade.xlsx')
#grade

"""Leitura e tratamento da planilha: LEX-Usuarios_com_suas_turmas"""

usuarios = carregar_e_concatenar_csv(caminho_input,'LEX-Usuarios_com_suas_turmas' )

up(usuarios)
usuarios = usuarios.query('TYPE == "SÉRIE" and SCHOOLYEAR == "2024"')
usuarios = usuarios[['ID','NAME','PERFIL','SCHOOLID','SCHOOLNAME','SCHOOLTYPE','CLASSID','TYPE','GRADE']]

usuarios = pd.merge(usuarios,grade, how='left', on=['GRADE'])

df_usuarios = usuarios.copy()
#df_usuarios.head()

turmas = carregar_e_concatenar_csv(caminho_input,'LEX-Turma_X_Produtos' )

up(turmas)
turmas = turmas.drop(0).reset_index().drop(columns=['index'])
turmas = turmas[['TENANTID','SCHOOLID','GROUPID','PRODUCT']]
turmas = turmas.rename(columns={'GROUPID':'CLASSID'})

df_turmas = turmas.copy()
#df_turmas.head()

df_usuarios_turmas = pd.merge(df_usuarios, df_turmas, how='left', on=['SCHOOLID','CLASSID'])
#df_usuarios_turmas.head()

df_usuarios_alunos = df_usuarios_turmas[df_usuarios_turmas['PERFIL']=='ALUNO']
#tipos_produtos = ['ÁRVORE', 'EDUCACROSS','BINÓCULO BY TELLME','PLATAFORMA AZ','REDAÇÃO','MY LIFE','TINDIN', 'UBBU','HELLO CAROLINA PATRÍCIO', 'SCHOOL GUARDIAN','EDUCACROSS HIGH FIVE','HELLO PUERI DOMUS', 'HELLO SPHERE']
tipos_produtos = ['ÁRVORE', 'EDUCACROSS','BINÓCULO','PLATAFORMA AZ','TINDIN', 'UBBU','SCHOOL GUARDIAN']

df_usuarios_alunos['PRODUCT'] = df_usuarios_alunos['PRODUCT'].replace(['BINÓCULO BY TELLME','HELLO CAROLINA PATRÍCIO','HELLO PUERI DOMUS', 'HELLO SPHERE'], 'BINÓCULO')
# Select the 'PRODUCT' column before applying the replace function.

df_usuarios_alunos.head()

cliente = pd.read_csv(f'{caminho_input}/grupo_cliente.csv', sep=',', low_memory=False, on_bad_lines='skip')
cliente = cliente[['Nome','Grupo','CPF ou CNPJ','Tipo de Cliente']]
# Assuming 'up' is a function defined elsewhere
up(cliente)
cliente = cliente.rename(columns={'NOME':'SCHOOLNAME'})
cliente = cliente[cliente['TIPO DE CLIENTE']=='PJ']
cliente.to_excel(f'{caminho_output}/cliente.xlsx', index=False)

# prompt: fazer um merge de cliente e df_usuarios_alunos quero apenas manter o GRUPO

df_merged = pd.merge(df_usuarios_alunos, cliente[['SCHOOLNAME', 'GRUPO']], on='SCHOOLNAME', how='left')
df_merged.head(2)

df_f2 = df_merged[['SCHOOLNAME','SCHOOLTYPE','SEG']]
df_f2 = df_f2.loc[df_f2['SEG'] == 'EFII']
df_f2 = df_f2.drop_duplicates()
#df_f2.to_excel(f'{caminho_output}/df_f2.xlsx', index=False)

# Loop para criar os dataframes filtrados para cada perfil
for produto in tipos_produtos:
    # Filtra o dataframe principal pelo perfil atual
    df_filtrado_produto = df_usuarios_alunos[df_usuarios_alunos['PRODUCT'] == produto].groupby(['SCHOOLNAME', 'SCHOOLTYPE', 'GRADE'], as_index=False)['PRODUCT'].count()
    # Filtra o dataframe principal pelo perfil atual
    df_filtrado_produto_escola = df_usuarios_alunos[df_usuarios_alunos['PRODUCT'] == produto].groupby(['SCHOOLNAME', 'SCHOOLTYPE'], as_index=False)['PRODUCT'].count()
    #df_filtrado_produto_soma = df_usuarios_alunos[df_usuarios_alunos['PRODUCT'] == produto].groupby(['SCHOOLNAME', 'SCHOOLTYPE'], as_index=False)['PRODUCT'].count()
    df_filtrado_produto_schooltype = df_usuarios_alunos[df_usuarios_alunos['PRODUCT'] == produto].groupby(['SCHOOLTYPE'], as_index=False)['PRODUCT'].count()
    with pd.ExcelWriter(f'{caminho_output}/{today}_{produto}.xlsx') as writer:
        df_filtrado_produto.to_excel(writer, sheet_name=f'{produto}_GRADE', index=False)
        #df_filtrado_produto_soma.to_excel(writer, sheet_name=f'{produto}', index=False)
        df_filtrado_produto_schooltype.to_excel(writer, sheet_name=f'{produto}_TYPE', index=False)
        df_filtrado_produto_escola.to_excel(writer, sheet_name=f'{produto}_SCHOOL', index=False)

################
