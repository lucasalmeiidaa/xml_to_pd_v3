#!/usr/bin/env python
# coding: utf-8

# In[39]:


import xmltodict
import pandas as pd
from os.path import join
from os import listdir
import numpy as np


# In[50]:


#with open('bases/13201105326555000141550030000226881202053343.xml') as fd:
#    doc = xmltodict.parse(fd.read())


# In[51]:


#doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['CEP']


# In[40]:


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


# In[60]:


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


# In[61]:


def busca_xml_v3_dados_pessoais(doc):
    
    lst_cpfs = []
    lst_nomes = []
    lst_logradouros = []
    lst_numeros = []
    lst_bairro = []
    lst_cmun = []
    lst_xmun = []
    lst_uf = []
    lst_cep = []
    

    for prod in doc['nfeProc']['NFe']['infNFe']['det']:
           #pai: dest
    
        ##filho: CPF
        try:
            cpf = doc['nfeProc']['NFe']['infNFe']['dest']['CPF']
        except:
            cpf = np.nan

        ##filho: nome
        try:
            nome = doc['nfeProc']['NFe']['infNFe']['dest']['xNome']
        except:
            nome = np.nan

        #pai: dest/enderDest

        ##filho: logradouro
        try:
            logr = doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['xLgr']
        except:
            logr = np.nan

        ##filho: numero
        try:
            numero = doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['xLgr']
        except:
            numero = np.nan

        ##filho: bairro
        try:
            bairro = doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['xBairro']
        except:
            bairro = np.nan
        ##filho: cmunicipio
        try:
            cmun = doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['cMun']
        except:
            cmun = np.nan

        ##filho: xmunicipio
        try:
            xmun = doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['xMun']
        except:
            xmun = np.nan

        ##filho: UF
        try:
            uf = doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['UF']
        except:
            uf = np.nan

        ##filho: CEP
        try:
            cep = doc['nfeProc']['NFe']['infNFe']['dest']['enderDest']['CEP']
        except:
            cep = np.nan
            
        
        lst_cpfs.append(cpf)
        lst_nomes.append(nome)
        lst_logradouros.append(logr)
        lst_numeros.append(numero)
        lst_bairro.append(bairro)
        lst_cmun.append(cmun)
        lst_xmun.append(xmun)
        lst_uf.append(uf)
        lst_cep.append(cep)
    
    #Criando o dataframe
    ##Colunas criadas com lista
    df = pd.DataFrame()
    df['CPF'] = lst_cpfs
    df['Nome'] = lst_nomes
    df['logradouro'] = lst_logradouros
    df['Numero'] = lst_numeros
    df['Bairro'] = lst_bairro
    df['cMun'] = lst_cmun
    df['xMun'] = lst_xmun
    df['UF'] = lst_uf
    df['CEP'] = lst_uf
    
    return df


# In[62]:


path_xml = 'bases'

files_xml = []

for file in listdir(path_xml):
    files_xml.append(file)


# In[63]:


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


# In[64]:


df = pd.DataFrame()

writer = pd.ExcelWriter('informações.xlsx', engine='xlsxwriter')

for file in files_xml:
    # Abrindo xml
    with open(join(path_xml, file)) as fd:
        doc = xmltodict.parse(fd.read())
    
    df_aux = busca_xml_v3_dados_pessoais(doc)
    
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

