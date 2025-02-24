# %% [markdown]
# # Converter as tabelas CT2 e SC7 para o padrão do Accountfy
# 
# **Passo a passo:**
# 1. A exportação das tabelas do Protheus estão sendo realizadas através do caminho: *Consultas/Cadastros/Genéricos*
# <br>
# <br>
# 2. Protheus: Extrair a tabela **SC7 (Pedidos de compra)** para que seja possível identificar aqueles pedidos que já foram aprovados (consumiram o orçamento) mas não foram lançados ainda.<br>
# Filtro: Dt. Entrega = Mês atual<br>
# Dicionário: Marca todos<br>
# Formato: CSV separado por ponto e vírgula<br>
# Nome: *SA2.csv*
# <br>
# <br>
# 3. Protheus: Extrair a tabela **CT2 (Lançamentos contábeis)** para que seja possível identificar todos os lançamentos já realizados.<br>
# Filtro: Data Lcto = Mês atual<br>
# Dicionário: Marca todos<br>
# Formato: CSV separado por ponto e vírgula<br>
# Nome: *CT2.csv*
# <br>
# <br>
# 
# **Passos adicionais a serem executados de tempo em tempo:**
# 
# 1. Protheus: Extrair a tabela **SA2 (Fornecedores)** para que seja possível completar o nome do fornecedor através do código da SC7.<br>
# Filtro: Não se aplica<br>
# Dicionário: Exportar somente as colunas "Codigo" e "Razao Social"<br>
# Formato: CSV separado por ponto e vírgula<br>
# Nome: *SA2.csv*
# <br>
# <br>
# 2. Accountfy: Extrair o **plano de contas do Accountfy**.<br>
# Nome: *Accountfy - Plano de contas - Tecadi.xlsx*
# 

# %% [markdown]
# ## Configurações

# %%
# Considerar pedidos "Aprovados"?
sc7_aprovados = True

# Considerar pedidos "Em aprovação"?
sc7_em_aprovacao = True

# Caminho dos arquivos
plano_filename = 'Accountfy - Plano de contas - Tecadi.xlsx'
ct2_filename = 'CT2.csv'
sa2_filename = 'SA2.csv'
sc7_filename = 'SC7.csv'
ct2_filename_dagnoni = 'CT2_Dagnoni.csv'
parametros_rateio_patrimonial_filename = 'Parametros_rateio_patrimonial.xlsx'
parametros_rateio_corporativo_filename = 'Parametros_rateio_corporativo.xlsx'
ajustes_gerenciais_filename = 'Ajustes_gerenciais.xlsx'

# Alíquotas de impostos
aliquota_pis = 0.0165
aliquota_cofins = 0.076
aliquota_iss = {
   "103":   0.03,
   "105":   0.05,
   "107":   0.03,
   "108":   0.02,
   "109":   0.02,
   "114":   0.025,
   "115":   0.02
}

# Nome das filiais assim como no Accountfy
filiais = {
   "101":   "101 - Corporativo",
   "103":   "103 - CD Itajaí (Salseiros)",
   "105":   "105 - CD Curitiba",
   "107":   "107 - CD Itajaí (Itaipava)",
   "107TR": "107 - TR Itajaí (Itaipava)",
   "108":   "108 - CD Navegantes",
   "109":   "109 - CD Cajamar",
   "110":   "110 - TR Paranaguá",
   "111":   "111 - TR Santa Cruz do Sul",
   "112":   "112 - TR Rio Grande",
   "113":   "113 - TR Santos",
   "114":   "114 - CD São José dos Pinhais",
   "115":   "115 - CD Fazenda Rio Grande",
   "ADM": "Patrimonial - Dagnoni e Kenig"
}

# Quais contas devem ser desconsideradas?
def contas_desconsideradas():
    return ('1', '2')

# Quais filiais devem ser desconsideradas?
filiais_desconsideradas = []

# Importa todas as bibliotecas necessárias
import pandas as pd
from datetime import datetime
import pytz
import os
import shutil

# %% [markdown]
# ## Converter os lançamentos da CT2 (Lançamentos contábeis)

# %%
def process_ct2(filename):
    # Ler CT2 pulando 2 primeiras linhas
    df = pd.read_csv(filename, sep=';', encoding='latin-1', quotechar='"', skiprows=2, low_memory=False)
    
    # Conversões iniciais
    df['Valor'] = pd.to_numeric(df['Valor'].str.replace(',', '.'), errors='coerce')
    df['C Custo Deb'] = pd.to_numeric(df['C Custo Deb'], errors='coerce')
    df['C Custo Crd'] = pd.to_numeric(df['C Custo Crd'], errors='coerce')
    
    # Tratar centros de custo
    df['C Custo Deb'] = df['C Custo Deb'].apply(lambda x: f'0{int(x)}' if pd.notnull(x) else '')
    df['C Custo Crd'] = df['C Custo Crd'].apply(lambda x: f'0{int(x)}' if pd.notnull(x) else '')
    
    # Converter contas para string
    df['Cta Debito'] = df['Cta Debito'].astype(str)
    df['Cta Credito'] = df['Cta Credito'].astype(str)
    
    # Limpar contra partidas
    df.loc[~df['Cta Debito'].str.startswith(('3', '4', '5', '6', '7', '8', '9')), 'Cta Debito'] = ''
    df.loc[~df['Cta Credito'].str.startswith(('3', '4', '5', '6', '7', '8', '9')), 'Cta Credito'] = ''
    
    # Criar D/C
    df['D/C'] = ''
    df.loc[df['Cta Debito'] != '', 'D/C'] = 'D'
    df.loc[df['Cta Credito'] != '', 'D/C'] = 'C'
    
    # Criar Conta
    df['Conta'] = df['Cta Debito'].where(df['Cta Debito'] != '', df['Cta Credito'])
    df['Conta'] = df['Conta'].astype(str).str.replace('.0', '')
    df = df[(df['Conta'] != 'nan') & (df['Conta'] != '')]
    
    # Criar Centro de custo
    df['Centro de custo'] = df['C Custo Deb'].where(df['C Custo Deb'] != '', df['C Custo Crd'])
    
    # Tratar histórico e observações
    df['Obs'] = df['Hist Lanc'].str.extract(r'( - .+)$')[0].str.replace(' - ', '', regex=False)
    df['Hist Lanc'] = df['Hist Lanc'].str.replace(r' - .+$', '', regex=True)
    df['Obs'] = df['Obs'].fillna('')
    df['Hist Lanc'] = df['Hist Lanc'].str.replace(' - ', '')

    # Ajustar histórico para lançamentos do RH
    df.loc[df['Rotina'] == 'CTBA500', 'Obs'] = 'RH / Folha de pagamento'
    
    return df

# Processar arquivos
df_tecadi = process_ct2(ct2_filename)
df_dagnoni = process_ct2(ct2_filename_dagnoni)

# Processar filiais Tecadi
df_tecadi = df_tecadi.rename(columns={'Filial Orig': 'Cod filial'})

# Processar filiais Dagnoni
df_dagnoni = df_dagnoni.rename(columns={'Filial Orig': 'Cod filial'})
df_dagnoni['Centro de custo'] = '999999' # A Dagnoni não tem centro de custo
df_dagnoni['Cod filial'] = 'ADM'

# Concatenar preservando as filiais
df = pd.concat([df_tecadi, df_dagnoni], ignore_index=True)

# %% [markdown]
# ## Converter os lançamentos da SC7 (Pedidos de compra)
# 
# **Atenção:** Executar apenas se for necessário trazer esses lançamentos, do contrário passar para o próximo passo.

# %%
# Ler SA2 pulando 2 primeiras linhas
sa2 = pd.read_csv(sa2_filename, sep=';', encoding='latin-1', quotechar='"', skiprows=2, low_memory=False)

# Ler SC7 pulando 2 primeiras linhas
sc7 = pd.read_csv(sc7_filename, sep=';', encoding='latin-1', quotechar='"', skiprows=2, low_memory=False)

# Remove linhas com dados na coluna extra e reseta o índice
#sa2 = sa2[sa2['Unnamed: 2'].isna()].reset_index(drop=True)

# Remove zeros à esquerda do código do fornecedor
sa2['Codigo'] = sa2['Codigo'].str.lstrip('0')

# Filtrar na SC7 apenas aprovados e não encerrados
sc7_aprovados_df = sc7[
    (sc7['Ped. Encerr.'] != 'E') &
    (sc7['Resid. Elim.'] != 'S') &
    (sc7['Status'] == 'Aprovado')
]

# Filtrar na SC7 apenas em aprovação e não encerrados
sc7_em_aprovacao_df = sc7[
    (sc7['Ped. Encerr.'] != 'E') &
    (sc7['Resid. Elim.'] != 'S') &
    (sc7['Status'] == 'B')
]

# Remove zeros à esquerda do código do fornecedor
sc7_aprovados_df['Fornecedor'] = sc7_aprovados_df['Fornecedor'].astype(str).str.lstrip('0')
sc7_em_aprovacao_df['Fornecedor'] = sc7_em_aprovacao_df['Fornecedor'].astype(str).str.lstrip('0')

# Define quais as TES que tomam crédito de Pis e Cofins
def tes_credito():
    return ('001', '002', '01A', '01B', '01C', '01D', '01E', '01F', '01G', '01H',
        '01N', '01O', '01P', '01Q', '01R', '01S', '01X', '02A', '02C', '02H',
        '02I', '040', '04D', '051', '052', '053', '054', '055', '05D', '060',
        '061', '063', '064', '066', '067', '068', '069', '070', '071', '072',
        '073', '074', '075', '076', '079', '07D', '080', '081', '082', '083',
        '084', '085', '086', '087', '088', '08A', '090', '091', '092', '094',
        '095', '096', '097', '098', '105', '130', '133', '209', '216', '217',
        '218', '48B', '48C')

# Função para formatar número do pedido
def formatar_pedido(filial, numero):
   return f"{filial}/{str(numero).zfill(6)}"

# Criar lançamentos de débito da SC7 (aprovados)
if sc7_aprovados:
    if not sc7_aprovados_df.empty:  # Verifica se o DataFrame não está vazio
        df_sc7_aprovados = pd.DataFrame({
            'Conta': sc7_aprovados_df['Cta Contabil'],
            'Valor': sc7_aprovados_df['Vlr.Total'].str.replace(',', '.').astype(float),
            'D/C': 'D',
            'Hist Lanc': 'Pedido ' + sc7_aprovados_df.apply(lambda x: formatar_pedido(x['Filial'], x['Numero PC']), axis=1) + ' aprovado e não recebido.',
            'Data Lcto': sc7_aprovados_df['Dt. Entrega'],
            'Centro de custo': sc7_aprovados_df['Centro Custo'].apply(lambda x: f'0{int(x)}' if pd.notnull(x) else ''),
            'Cod filial': sc7_aprovados_df['Filial'],
            'Obs': sc7_aprovados_df['Fornecedor'].map(dict(zip(sa2['Codigo'], sa2['Razao Social']))).fillna('Fornecedor não encontrado')
        })

        # Criar lançamentos de crédito para Pis e Cofins (aprovados)
        df_sc7_piscofins_aprovados = df_sc7_aprovados[sc7_aprovados_df['Tipo Entrada'].isin(tes_credito())].copy()
        df_sc7_piscofins_aprovados['D/C'] = 'C'
        df_sc7_piscofins_aprovados['Valor'] = (df_sc7_piscofins_aprovados['Valor'] * 0.0925).round(2)
        df_sc7_piscofins_aprovados = df_sc7_piscofins_aprovados[df_sc7_piscofins_aprovados['Valor'] != 0]
        df_sc7_piscofins_aprovados['Hist Lanc'] = 'Créd. de Pis e Cofins ref. pedido ' + sc7_aprovados_df[sc7_aprovados_df['Tipo Entrada'].isin(tes_credito())].apply(lambda x: formatar_pedido(x['Filial'], x['Numero PC']), axis=1) + ' aprovado e não recebido.'

        # Tratar conta contábil e centro de custo (aprovados)
        for df in [df_sc7_aprovados, df_sc7_piscofins_aprovados]:
            df['Conta'] = df['Conta'].astype(str)
    else:
        # Se o DataFrame estiver vazio, cria um DataFrame vazio com as colunas necessárias
        df_sc7_aprovados = pd.DataFrame(columns=['Conta', 'Valor', 'D/C', 'Hist Lanc', 'Data Lcto', 'Centro de custo', 'Cod filial', 'Obs'])
        df_sc7_piscofins_aprovados = pd.DataFrame(columns=['Conta', 'Valor', 'D/C', 'Hist Lanc', 'Data Lcto', 'Centro de custo', 'Cod filial', 'Obs'])

# Criar lançamentos de débito da SC7 (em aprovação)
if sc7_em_aprovacao:
    if not sc7_em_aprovacao_df.empty:  # Verifica se o DataFrame não está vazio
        df_sc7_em_aprovacao = pd.DataFrame({
            'Conta': sc7_em_aprovacao_df['Cta Contabil'],
            'Valor': sc7_em_aprovacao_df['Vlr.Total'].str.replace(',', '.').astype(float),
            'D/C': 'D',
            'Hist Lanc': 'Pedido ' + sc7_em_aprovacao_df.apply(lambda x: formatar_pedido(x['Filial'], x['Numero PC']), axis=1) + ' em aprovação.',
            'Data Lcto': sc7_em_aprovacao_df['Dt. Entrega'],
            'Centro de custo': sc7_em_aprovacao_df['Centro Custo'].apply(lambda x: f'0{int(x)}' if pd.notnull(x) else ''),
            'Cod filial': sc7_em_aprovacao_df['Filial'],
            'Obs': sc7_em_aprovacao_df['Fornecedor'].map(dict(zip(sa2['Codigo'], sa2['Razao Social']))).fillna('Fornecedor não encontrado')
        })

        # Criar lançamentos de crédito para Pis e Cofins (em aprovação)
        df_sc7_piscofins_em_aprovacao = df_sc7_em_aprovacao[sc7_em_aprovacao_df['Tipo Entrada'].isin(tes_credito())].copy()
        df_sc7_piscofins_em_aprovacao['D/C'] = 'C'
        df_sc7_piscofins_em_aprovacao['Valor'] = (df_sc7_piscofins_em_aprovacao['Valor'] * 0.0925).round(2)
        df_sc7_piscofins_em_aprovacao = df_sc7_piscofins_em_aprovacao[df_sc7_piscofins_em_aprovacao['Valor'] != 0]
        df_sc7_piscofins_em_aprovacao['Hist Lanc'] = 'Créd. de Pis e Cofins ref. pedido ' + sc7_em_aprovacao_df[sc7_em_aprovacao_df['Tipo Entrada'].isin(tes_credito())].apply(lambda x: formatar_pedido(x['Filial'], x['Numero PC']), axis=1) + ' em aprovação.'

        # Tratar conta contábil e centro de custo (em aprovação)
        for df in [df_sc7_em_aprovacao, df_sc7_piscofins_em_aprovacao]:
            df['Conta'] = df['Conta'].astype(str)
    else:
        # Se o DataFrame estiver vazio, cria um DataFrame vazio com as colunas necessárias
        df_sc7_em_aprovacao = pd.DataFrame(columns=['Conta', 'Valor', 'D/C', 'Hist Lanc', 'Data Lcto', 'Centro de custo', 'Cod filial', 'Obs'])
        df_sc7_piscofins_em_aprovacao = pd.DataFrame(columns=['Conta', 'Valor', 'D/C', 'Hist Lanc', 'Data Lcto', 'Centro de custo', 'Cod Filial', 'Obs'])

# Concatenar os dados da SC7 com DataFrame principal
df_list = [df_tecadi, df_dagnoni]
if sc7_aprovados:
    df_list.extend([df_sc7_aprovados, df_sc7_piscofins_aprovados])
if sc7_em_aprovacao:
    df_list.extend([df_sc7_em_aprovacao, df_sc7_piscofins_em_aprovacao])

df_list = [df for df in df_list if not df.empty]
df = pd.concat(df_list, ignore_index=True, copy=True)

# %% [markdown]
# ## Gerar os lançamentos de ajustes gerenciais e os ajustes de contas contábeis e centros de custo

# %%
# Data mais recente
data_mais_recente = df['Data Lcto'].max()

# Ajustes gerenciais
try:
    ajustes = pd.read_excel(ajustes_gerenciais_filename)
    
    if not ajustes.empty:
        soma_debitos = ajustes[ajustes['D/C'] == 'D']['Valor'].sum()
        soma_creditos = ajustes[ajustes['D/C'] == 'C']['Valor'].sum()
        
        if soma_debitos != soma_creditos:
            resposta = input(f"ATENÇÃO: Diferença de {abs(soma_debitos - soma_creditos):.2f} entre débitos e créditos. Continuar? (S/N): ")
            if resposta.upper() != 'S':
                raise SystemExit("Processo interrompido pelo usuário.")
        
        ajustes['Conta'] = ajustes['Conta'].astype(str)
        ajustes['Centro de custo'] = ajustes['Centro de custo'].astype(str)
        ajustes['Filial Orig'] = ajustes['Filial Orig'].astype(str)
        ajustes['Centro de custo'] = ajustes['Centro de custo'].apply(lambda x: f'0{int(x)}' if pd.notnull(x) else '')        
        
        ajustes_df = pd.DataFrame({
            'Conta': ajustes['Conta'],
            'Valor': ajustes['Valor'],
            'D/C': ajustes['D/C'],
            'Hist Lanc': '(Ajuste gerencial) ' + ajustes['Hist Lanc'],
            'Data Lcto': data_mais_recente,
            'Centro de custo': ajustes['Centro de custo'],
            'Cod filial': ajustes['Filial Orig'],
            'Obs': ajustes.get('Obs', '')
        })
        
        df = pd.concat([df, ajustes_df], ignore_index=True)
        
except FileNotFoundError:
    pass

# Migrar os lançamentos das contas que começam com 7, 8 ou 9 para a filial 101, exceto quando for ADM
df.loc[(df['Conta'].str.startswith(('7', '8', '9'))) & (df['Cod filial'] != 'ADM'), 'Cod filial'] = '101'

# %% [markdown]
# ## Gerar o rateio da patrimonial (Dagnoni) nas unidades operacionais

# %%
# Ler arquivo de parâmetros de rateio 
rateio_patrimonial = pd.read_excel(parametros_rateio_patrimonial_filename)

# Extrair mês e ano da data_mais_recente
data_ref = pd.to_datetime(data_mais_recente)
coluna_ref = pd.to_datetime(f"{data_ref.year}-{data_ref.month}-01")

# Calcular percentuais de rateio
valores_mes = rateio_patrimonial[coluna_ref].fillna(0)
total_mes = valores_mes[rateio_patrimonial['Cod filial'] == 'TOTAL'].values[0]
percentuais = valores_mes / total_mes

# Soma todos os lançamentos das contas 3, 4, 5, 6, 7, 8 e 9 da filial ADM
lancamentos_adm = df[
   (df['Cod filial'] == 'ADM') & 
   (df['Conta'].str.match(r'^[3456789]')) &
   (df['Valor'] != 0)
]
valor_total = sum(lancamentos_adm['Valor'].where(lancamentos_adm['D/C'] == 'D', -lancamentos_adm['Valor']))

# Soma todos os lançamentos das contas de depreciação
contas_depreciacao= ['6101010231', '5201010115', '5101010112', '6101010110']
lancamentos_adm_2 = df[
   (df['Cod filial'] == 'ADM') & 
   (df['Conta'].isin(contas_depreciacao)) &
   (df['Valor'] != 0)
]

valor_total_2 = sum(lancamentos_adm_2['Valor'].where(lancamentos_adm_2['D/C'] == 'D', -lancamentos_adm_2['Valor']))

# Soma todos os lançamentos das contas 7, 8 e 9 da filial ADM
lancamentos_adm_3 = df[
    (df['Cod filial'] == 'ADM') & 
    (df['Conta'].str.match(r'^[789]')) &
    (df['Valor'] != 0)
]
valor_total_3 = sum(lancamentos_adm_3['Valor'].where(lancamentos_adm_3['D/C'] == 'D', -lancamentos_adm_3['Valor']))

novos_lancamentos = []

# Processar ambos os conjuntos
for _, row in rateio_patrimonial.iterrows():
   if row['Cod filial'] not in ['TOTAL', 'ADM']:
       # Lançamentos na 5301010901 - Resultado da equivalência patrimonial (Dagnoni)
       valor_rateio = valor_total * percentuais[_]
       if valor_rateio != 0:
           novos_lancamentos.extend([
               {
                   'Conta': '5301010901',
                   'Valor': abs(valor_rateio),
                   'D/C': 'D' if valor_rateio > 0 else 'C',
                   'Hist Lanc': f'(Rateio patrimonial) Rateio da patrimonial para a filial {row["Filial"]}',
                   'Data Lcto': data_mais_recente,
                   'Centro de custo': '999999',
                   'Cod filial': row['Cod filial'],
                   'Obs': 'Lançamento automático'
               }
           ])
       
       # Lançamentos na 2303010998 - ( - ) Depreciação / Amortização (Rateio patrimonial)
       valor_rateio_2 = valor_total_2 * percentuais[_]
       if valor_rateio_2 != 0:
           novos_lancamentos.extend([
               {
                   'Conta': '2303010998',
                   'Valor': abs(valor_rateio_2),
                   'D/C': 'D' if valor_rateio_2 > 0 else 'C',
                   'Hist Lanc': f'(Rateio patrimonial) Rateio da patrimonial para a filial {row["Filial"]}',
                   'Data Lcto': data_mais_recente,
                   'Centro de custo': '999999',
                   'Cod filial': row['Cod filial'],
                   'Obs': 'Lançamento automático'
               }
           ])
        
        # Lançamentos na 2303010997 - ( - ) Resultado financeiro / IR / CSLL (Rateio patrimonial)
       valor_rateio_3 = valor_total_3 * percentuais[_]
       if valor_rateio_3 != 0:
           novos_lancamentos.extend([
               {
                   'Conta': '2303010997',
                   'Valor': abs(valor_rateio_3),
                   'D/C': 'D' if valor_rateio_3 > 0 else 'C',
                   'Hist Lanc': f'(Rateio patrimonial) Rateio da patrimonial para a filial {row["Filial"]}',
                   'Data Lcto': data_mais_recente,
                   'Centro de custo': '999999',
                   'Cod filial': row['Cod filial'],
                   'Obs': 'Lançamento automático'
               }
           ])

# Adicionar novos lançamentos ao DataFrame principal
if novos_lancamentos:
   df_rateio = pd.DataFrame(novos_lancamentos)
   df = pd.concat([df, df_rateio], ignore_index=True)

# %% [markdown]
# ## Gerar o rateio do corporativo nas filiais

# %%
# Ler arquivo de parâmetros de rateio 
rateio_corporativo = pd.read_excel(parametros_rateio_corporativo_filename)

# Extrair mês e ano da data_mais_recente
data_ref = pd.to_datetime(data_mais_recente)
coluna_ref = pd.to_datetime(f"{data_ref.year}-{data_ref.month}-01")

# Busca o percentual de rateio para cada filial
valores_mes = rateio_corporativo[coluna_ref].fillna(0)
percentuais = valores_mes

# Soma todos os lançamentos das contas 3, 4, 5, 6 da filial 101
lancamentos_corporativo = df[
   (df['Cod filial'].astype(str) == '101') & 
   (df['Conta'].str.match(r'^[3456]'))
]
valor_total = sum(lancamentos_corporativo['Valor'].where(lancamentos_corporativo['D/C'] == 'D', -lancamentos_corporativo['Valor']))

# Soma todos os lançamentos das contas de depreciação
contas_depreciacao= ['6101010231', '5201010115', '5101010112']
lancamentos_corporativo_2 = df[
   (df['Cod filial'].astype(str) == '101') & 
   (df['Conta'].isin(['6101010231', '5201010115', '5101010112']))
]

valor_total_2 = sum(lancamentos_corporativo_2['Valor'].where(lancamentos_corporativo_2['D/C'] == 'D', -lancamentos_corporativo_2['Valor']))

novos_lancamentos = []

# Processar ambos os conjuntos
for _, row in rateio_corporativo.iterrows():
   if row['Cod filial'] not in ['TOTAL', '101']:
       # Lançamentos na 5301010902 - Despesas corporativas
       valor_rateio = valor_total * percentuais[_]
       if valor_rateio != 0:
           novos_lancamentos.extend([
               {
                   'Conta': '5301010902',
                   'Valor': abs(valor_rateio),
                   'D/C': 'D' if valor_rateio > 0 else 'C',
                   'Hist Lanc': f'(Rateio corporativo) {(percentuais[_]*100):.2f}% - {row["Filial"]}',
                   'Data Lcto': data_mais_recente,
                   'Centro de custo': '999999',
                   'Cod filial': row['Cod filial'],
                   'Obs': 'Lançamento automático'
               },
               {
                   'Conta': '5301010902',
                   'Valor': abs(valor_rateio),
                   'D/C': 'C' if valor_rateio > 0 else 'D',
                   'Hist Lanc': f'(Rateio corporativo) {(percentuais[_]*100):.2f}% - {row["Filial"]}',
                   'Data Lcto': data_mais_recente,
                   'Centro de custo': '999999',
                   'Cod filial': '101',
                   'Obs': 'Lançamento automático'
               }
           ])
       
       # Lançamentos na 2303010996 - ( - ) Depreciação e amortização (Rateio corporativo)
       valor_rateio_2 = valor_total_2 * percentuais[_]
       if valor_rateio_2 != 0:
           novos_lancamentos.extend([
               {
                   'Conta': '2303010996',
                   'Valor': abs(valor_rateio_2),
                   'D/C': 'D' if valor_rateio_2 > 0 else 'C',
                   'Hist Lanc': f'(Rateio corporativo) {(percentuais[_]*100):.2f}% da depreciação do corporativo - {row["Filial"]}',
                   'Data Lcto': data_mais_recente,
                   'Centro de custo': '999999',
                   'Cod filial': row['Cod filial'],
                   'Obs': 'Lançamento automático'
               },
               {
                   'Conta': '2303010996',
                   'Valor': abs(valor_rateio_2),
                   'D/C': 'C' if valor_rateio_2 > 0 else 'D',
                   'Hist Lanc': f'(Rateio corporativo) {(percentuais[_]*100):.2f}% da depreciação do corporativo - {row["Filial"]}',
                   'Data Lcto': data_mais_recente,
                   'Centro de custo': '999999',
                   'Cod filial': '101',
                   'Obs': 'Lançamento automático'
               }
           ])

# Adicionar novos lançamentos ao DataFrame principal
if novos_lancamentos:
   df_rateio = pd.DataFrame(novos_lancamentos)
   df = pd.concat([df, df_rateio], ignore_index=True)

# %% [markdown]
# ## Gerar a transferência das receitas e custos para a 107TR e recalcular o ISS / Pis / Cofins

# %%
# Quais filiais devem ser consideradas para transferir as contas de receita e que iniciam com 52?
filiais_transferencia = ['103', '105', '107', '108', '109', '114', '115']

# Migrar os lançamentos das contas 52 para a filial 107TR
df.loc[(df['Cod filial'].astype(str).isin(filiais_transferencia)) & 
      (df['Conta'].astype(str).str.startswith('52')), 'Cod filial'] = '107TR'

# Migrar os lançamentos das contas de ICMS e Crédito pró-cargas
df.loc[(df['Cod filial'].astype(str).isin(filiais_transferencia)) & 
      (df['Conta'].isin(['4101010206', '4101010201'])), 'Cod filial'] = '107TR'

# Quais contas devem ser transferidas para a filial 107TR?
contas_transferencia = ['3101010104', '3101010105']

# Calcular impostos após transferência
novos_lancamentos = []

for filial in filiais_transferencia:
   # Identificar saldo original das contas 3 antes da transferência
   mask_conta_3 = (
       (df['Cod filial'].astype(str) == filial) & 
       (df['Conta'].astype(str).str.startswith('3'))
   )
   saldo_original = df[mask_conta_3]['Valor'].where(
       df[mask_conta_3]['D/C'] == 'C', 
       -df[mask_conta_3]['Valor']
   ).sum()
   
   # Identificar valor que será transferido
   mask_transferencia = (
       (df['Cod filial'].astype(str) == filial) & 
       (df['Conta'].isin(contas_transferencia))
   )
   valor_transferir = df[mask_transferencia]['Valor'].where(
       df[mask_transferencia]['D/C'] == 'C',
       -df[mask_transferencia]['Valor']
   ).sum()
   
   # Transferir contas para 107TR
   df.loc[mask_transferencia, 'Cod filial'] = '107TR'
   
   # Calcular novos impostos
   saldo_remanescente = saldo_original - valor_transferir
   novo_pis = abs(saldo_remanescente * aliquota_pis)
   novo_cofins = abs(saldo_remanescente * aliquota_cofins)
   novo_iss = abs(saldo_remanescente * aliquota_iss[filial])
   
   # Processar cada imposto
   impostos = {
       '4101010202': {'valor': novo_pis, 'nome': 'Pis'},
       '4101010203': {'valor': novo_cofins, 'nome': 'Cofins'},
       '4101010204': {'valor': novo_iss, 'nome': 'ISS'}
   }
   
   for conta, info in impostos.items():
       # Remover lançamentos antigos
       mask_imposto = (
           (df['Cod filial'].astype(str) == filial) & 
           (df['Conta'] == conta)
       )
       imposto_atual = df[mask_imposto]['Valor'].where(
           df[mask_imposto]['D/C'] == 'D',
           -df[mask_imposto]['Valor']
       ).sum()
       df = df[~mask_imposto]
       
       # Criar lançamento do novo imposto
       if info['valor'] > 0:
           novos_lancamentos.append({
               'Conta': conta,
               'Valor': info['valor'],
               'D/C': 'D',
               'Hist Lanc': f'(Recálculo dos impostos) Recálculo do {info["nome"]}',
               'Data Lcto': data_mais_recente,
               'Centro de custo': '999999',
               'Cod filial': filial,
               'Obs': 'Lançamento automático'
           })
       
       # Transferir diferença para 107TR
       diferenca = imposto_atual - info['valor']
       if diferenca != 0:
           novos_lancamentos.append({
               'Conta': conta,
               'Valor': abs(diferenca),
               'D/C': 'D' if diferenca > 0 else 'C',
               'Hist Lanc': f'(Recálculo dos impostos) {info["nome"]} da filial {filiais[filial]}',
               'Data Lcto': data_mais_recente,
               'Centro de custo': '999999',
               'Cod filial': '107TR',
               'Obs': 'Lançamento automático'
           })

if novos_lancamentos:
   df_novos = pd.DataFrame(novos_lancamentos)
   df = pd.concat([df, df_novos], ignore_index=True)

# %% [markdown]
# ## Gerar os lançamentos de zeramento da base

# %%
# Definir as contas específicas que devem ser mantidas mesmo começando com 1 ou 2
contas_excecao = ['2303010996', '2303010997', '2303010998', '2303010999']

# Manter todas as contas que NÃO começam com 1 ou 2 OU estão na lista de exceções
df = df[
    (~df['Conta'].str.startswith(('1', '2'))) |  # Não começa com 1 ou 2
    (df['Conta'].isin(contas_excecao))           # OU está na lista de exceções
]

## Gerar os lançamentos de zeramento
def create_zeramento_df(df_input):
    filial_df = df_input.copy()
    # Converte todos os códigos de filial para string
    filial_df['Cod filial'] = filial_df['Cod filial'].astype(str)
    
    saldos = filial_df.groupby('Cod filial').apply(
        lambda x: x[x['D/C'] == 'D']['Valor'].sum() - x[x['D/C'] == 'C']['Valor'].sum()
    ).reset_index()
    saldos.columns = ['Cod filial', 'Saldo']
    
    zeramentos = []
    for _, row in saldos.iterrows():
        if row['Saldo'] != 0:
            zeramentos.append({
                'Conta': '2303010999',
                'Valor': abs(row['Saldo']),
                'D/C': 'C' if row['Saldo'] > 0 else 'D',
                'Hist Lanc': 'Zeramento resultado contra passivo',
                'Data Lcto': data_mais_recente,
                'Centro de custo': '999999',
                'Cod filial': row['Cod filial'],
                'Obs': 'Lançamento automático'
            })
    
    return pd.DataFrame(zeramentos) if zeramentos else pd.DataFrame()

def debug_saldos(df_input):
    # Converte todos os códigos de filial para string antes de ordenar
    df_temp = df_input.copy()
    df_temp['Cod filial'] = df_temp['Cod filial'].astype(str)
    
    for filial in sorted(df_temp['Cod filial'].unique()):
        filial_df = df_input[df_input['Cod filial'].astype(str) == filial]
        debitos = filial_df[filial_df['D/C'] == 'D']['Valor'].sum()
        creditos = filial_df[filial_df['D/C'] == 'C']['Valor'].sum()
        saldo = debitos - creditos
        
        print(f"\nFilial {filial}:")
        print(f"Total Débitos: {debitos:,.2f}")
        print(f"Total Créditos: {creditos:,.2f}")
        print(f"Saldo: {saldo:,.2f}")

# Criar zeramentos
zeramento_df = create_zeramento_df(df)

# Concatenar somente se houver zeramentos
if not zeramento_df.empty:
    df = pd.concat([df, zeramento_df], ignore_index=True)

# %% [markdown]
# 
# ## Gerar o arquivo de importação

# %%
# Adicionar lógica para Centro de custo padrão para contas do grupo 3, 4, 7, 8 e 9 e valores vazios/NaN
df.loc[(df['Conta'].astype(str).str.match(r'^[34789]\d{9}$')) | 
       (df['Centro de custo'].isna()) | 
       (df['Centro de custo'] == ''), 'Centro de custo'] = '999999'

# DE-PARA das contas da ADM (Patrimonial)
de_para_contas = {
    '6101010110': '6101010231',
    '6101010101': '6101010213',
    '6101010201': '6101010301'
}

# Aplicar DE-PARA somente quando a Filial for "ADM"
df.loc[df['Cod filial'] == "ADM", 'Conta'] = df['Conta'].replace(de_para_contas)

# Ajustar DE-PARA das contas da ADM
#df['Conta'] = df['Conta'].replace('6101010110', '6101010231')
#df['Conta'] = df['Conta'].replace('6101010101', '6101010213')
#df['Conta'] = df['Conta'].replace('6101010201', '6101010301')

# Ler plano de contas pulando 3 primeiras linhas
plano_contas = pd.read_excel(plano_filename, sheet_name='CONTAS_CONTABEIS', skiprows=3)
plano_contas['Código da conta'] = plano_contas['Código da conta'].astype(str)

# Criar coluna "Nome da conta" e fazer o cruzamento com o plano de contas
conta_dict = dict(zip(plano_contas['Código da conta'], plano_contas['Nome da conta']))

# Preenche o nome da filial na coluna 'Filial' a partir do código da filial
df['Filial'] = df['Cod filial'].astype(str).map(filiais)

df['Nome da conta'] = df['Conta'].map(conta_dict)

# Define a ordem das colunas e quais permanecem
columns_to_keep = ['Conta', 'Nome da conta', 'Valor', 'D/C', 'Hist Lanc', 'Data Lcto', 'Centro de custo', 'Filial', 'Obs', 'Cod filial']
df = df[columns_to_keep]

# Gerar nome do arquivo com data atual
sao_paulo_tz = pytz.timezone('America/Sao_Paulo')
current_datetime = datetime.now(sao_paulo_tz).strftime('%Y%m%d_%Hh%M')

# Extrair mês e ano do dataframe
data_ref = pd.to_datetime(df['Data Lcto'].max())
month_year = data_ref.strftime('%Y%m')

# Criar estrutura de pastas
base_dir = 'Output'
month_dir = os.path.join(base_dir, month_year)
date_dir = os.path.join(month_dir, current_datetime)

# Criar diretórios
for dir_path in [base_dir, month_dir, date_dir]:
   if not os.path.exists(dir_path):
       os.makedirs(dir_path)

# Salvar arquivo de output
output_filename = f'{current_datetime}_importacao_accountfy.xlsx'
output_path = os.path.join(date_dir, output_filename)
df.to_excel(output_path, index=False)

# Mover arquivos de origem
files_to_move = [
   ct2_filename, 
   sc7_filename,
   ct2_filename_dagnoni
]

# Copiar arquivos de parâmetros
files_to_copy = [
   parametros_rateio_corporativo_filename,
   parametros_rateio_patrimonial_filename,
   sa2_filename,
   ajustes_gerenciais_filename
]

for file in files_to_move:
   if os.path.exists(file):
       shutil.move(file, date_dir)

for file in files_to_copy:
   if os.path.exists(file):
       shutil.copy2(file, date_dir)

#df.to_excel(output_filename, index=False, engine='xlsxwriter')


