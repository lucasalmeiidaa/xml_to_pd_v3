#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xmltodict
import pandas as pd
from os.path import join
from os import listdir
import numpy as np


# In[31]:


# with open('bases/v3/13201105326555000141550030000226881202053343.xml') as fd:
#     doc = xmltodict.parse(fd.read())


# In[105]:


# #ide
# data_emissao = doc['nfeProc']['NFe']['infNFe']['ide']['dhEmi'][:10]
# nota_fiscal = doc['nfeProc']['NFe']['infNFe']['ide']['nNF']

# #emit
# emissor = doc['nfeProc']['NFe']['infNFe']['emit']['xNome']

# #dest
# destinatario = doc['nfeProc']['NFe']['infNFe']['dest']['xNome']

# #Prod
# lst_prods = []
# lst_ncm = []
# lst_cst = []
# lst_cfop = []
# lst_quantidade = []
# lst_valor = []

# for prod in doc['nfeProc']['NFe']['infNFe']['det']:
#     #nome_produto
#     produto = prod['prod']['xProd']
#     lst_prods.append(produto)
    
#     #ncm
#     ncm = prod['prod']['NCM']
#     lst_ncm.append(ncm)
    
#     #CST
#     cst = prod['imposto']['ICMS']['ICMS51']['CST']
#     lst_cst.append(cst)
    
#     #CFOP
#     cfop = prod['prod']['CFOP']
#     lst_cfop.append(cfop)
    
#     #Quantidade
#     quantidade = prod['prod']['qCom']
#     lst_quantidade.append(quantidade)
    
#     #Valor
#     valor = prod['prod']['vProd']
#     lst_valor.append(valor)


# In[132]:


def busca_xml_v3(doc):
    
    #pai: ide
    
    ##filho: data_emissao
    data_emissao = doc['nfeProc']['NFe']['infNFe']['ide']['dhEmi'][:10]
    
    ##filho: nota_fiscal
    nota_fiscal = doc['nfeProc']['NFe']['infNFe']['ide']['nNF']
    
    
    #pai: emit
    
    ##filho: emissor
    emissor = doc['nfeProc']['NFe']['infNFe']['emit']['xNome']
    
    
    #pai: dest
    
    ##filho: destinatario
    destinatario = doc['nfeProc']['NFe']['infNFe']['dest']['xNome']
    
    
    #pai: prod
    lst_prods = []
    lst_ncm = []
    lst_cst = []
    lst_cfop = []
    lst_quantidade = []
    lst_valor = []

    for prod in doc['nfeProc']['NFe']['infNFe']['det']:
        ##filho: nome_produto
        try:
            produto = prod['prod']['xProd']
            lst_prods.append(produto)
        except:
            produto = np.nan
            
        ##filho: ncm
        try:
            ncm = prod['prod']['NCM']
            lst_ncm.append(ncm)
        except:
            ncm = np.nan
            
        ##filho: CST
        try:
            cst = prod['imposto']['ICMS']['ICMS51']['CST']
            lst_cst.append(cst)
        except:
            cst = np.nan
            
        ##filho: CFOP
        try:
            cfop = prod['prod']['CFOP']
            lst_cfop.append(cfop)
        except:
            cfop = np.nan
            
        ##filho: Quantidade
        try:
            quantidade = prod['prod']['qCom']
            lst_quantidade.append(quantidade)
        except:
            quantidade = np.nan
            
        ##filho: Valor
        try:
            valor = prod['prod']['vProd']
            lst_valor.append(valor)
        except:
            valor = np.nan
            
            
    #Criando o dataframe
    ##Colunas criadas com lista
    df = pd.DataFrame()
    df['Produto'] = lst_prods
    df['NCM'] = lst_ncm
    df['CST'] = lst_cst
    df['CFOP'] = lst_cfop
    df['Quantidade'] = lst_quantidade
    df['Valor'] = lst_valor
    
    ##Colunas com valores fixos em todas as linhas
    df['data_emissao'] = data_emissao
    df['data_emissao'] = pd.to_datetime(df['data_emissao'], format='%Y-%m-%d')
    df['nota_fiscal'] = nota_fiscal
    df['EMISSOR'] = emissor
    df['REMETENTE/Destinatario'] = destinatario
    
    #Reordenando df
    df = df[['data_emissao', 'nota_fiscal', 'EMISSOR' ,'REMETENTE/Destinatario',
            'Produto', 'NCM', 'CST', 'CFOP', 'Quantidade', 'Valor']]
    
    # Transformando colunas str -> float
    df['Quantidade'] = df['Quantidade'].astype(float, errors='ignore')
    df['Valor'] = df['Valor'].astype(float, errors='ignore')
    
    return df


# In[118]:


path_xml = 'bases'

files_xml = []

for file in listdir(path_xml):
    files_xml.append(file)


# In[133]:


df = pd.DataFrame()

writer = pd.ExcelWriter('nfs.xlsx', engine='xlsxwriter')

for file in files_xml:
    # Abrindo xml
    with open(join(path_xml, file)) as fd:
        doc = xmltodict.parse(fd.read())
    
    df_aux = busca_xml_v3(doc)
   
    # Transformando colunas em float
#     df['desc_total'] = df['desc_total'].astype(float)
#     df['quantidade'] = df['quantidade'].astype(float)
#     df['valor_total'] = df['valor_total'].astype(float)
    
#     # Fazendo operações de cálculo
#     df['Coluna1'] = df['desc_total'] / df['quantidade']
#     df['Coluna2'] = df['valor_total'] - df['desc_total']
    
    # Exportando para excel
    df = pd.concat([df, df_aux])
    df.to_excel(writer, index=False)
    
    
writer.save()

