# 🏢 Sistema de Integração Totvs Protheus para Accountfy (Lançamentos do razão)

Sistema em Python para processamento e integração de dados contábeis do Protheus para o Accountfy (multirazão contábil), realizando a conversão das tabelas **CT2 (Lançamentos Contábeis)** e **SC7 (Pedidos de Compra)**.

## 📌 Principais Funcionalidades

### 📥 Importação e Processamento de Dados
- Importa lançamentos contábeis da tabela CT2
- Processa pedidos de compra da tabela SC7 (aprovados e em aprovação)
- Carrega o nome do fornecedor através da tabela SA2
- Trata as particularidades da companhia

### 💰 Cálculos Financeiros
- Calcula créditos de PIS/COFINS para os pedidos da SC7, com alíquotas configuráveis
- Processa ISS com diferentes alíquotas por filial
- Gerencia lançamentos de depreciação
- Suporta zeramento automático de saldos contábeis

### 📊 Distribuição de Custos
- Implementa rateio de custos corporativos entre filiais
- Processa rateio da filial patrimonial para unidades operacionais
- Transfere contas de receita e custo para filial de transportes
- Recalcula ISS/PIS/COFINS após transferências

### ⚙️ Suporte a Configurações
- Nomes e códigos de filiais configuráveis
- Mapeamento de contas personalizável
- Alíquotas ajustáveis (PIS, COFINS, ISS)
- Suporte para ajustes contábeis manuais

## 🔧 Como Usar

### 1️⃣ Extração dos Arquivos no Protheus

Para garantir o correto funcionamento do processo, exporte os seguintes arquivos:

| Arquivo | Descrição | Formato | Obrigatório |
|---------|-----------|---------|-------------|
| CT2.csv | Lançamentos contábeis | CSV (;) | Sim |
| SC7.csv | Pedidos de compra | CSV (;) | Sim |
| SA2.csv | Fornecedores | CSV (;) | Opcional |
| Accountfy - Plano de contas - Tecadi.xlsx | Plano de contas | Excel | Sim |
| Parametros_rateio_patrimonial.xlsx | Parâmetros de rateio patrimonial | Excel | Sim |
| Parametros_rateio_corporativo.xlsx | Parâmetros de rateio corporativo | Excel | Sim |
| Ajustes_gerenciais.xlsx | Ajustes gerenciais | Excel | Não |

> ⚠️ **Atenção**: Os arquivos CSV devem estar separados por ponto e vírgula (;)

### 2️⃣ Execução do Notebook
1. Coloque todos os arquivos na mesma pasta do notebook
2. Execute as células sequencialmente
3. Os resultados serão gerados na pasta Output

## 📂 Estrutura do Projeto

```
📁 Importador-Accountfy/
├── 📄 Tecadi_para_Accountfy_LOCAL.ipynb    # Notebook principal
├── 📄 README.md                            # Este arquivo
├── 📂 Input/                               # Arquivos de entrada
│   ├── CT2.csv
│   ├── SC7.csv
│   └── SA2.csv
└── 📂 Output/                              # Arquivos processados
    └── AAAAMM/
        └── AAAAMMDD_HHhMM/
            ├── AAAAMMDD_HHhMM_importacao_accountfy.xlsx
            └── [Arquivos de origem e parâmetros]
```

## 🚀 Tecnologias Utilizadas
- 🐍 Python
- 📊 Pandas
- 🕒 Datetime
- 🌍 Pytz
- 📁 Os/Shutil

## 📢 Observações
- O notebook é atualizado periodicamente para se adequar às mudanças no Protheus e Accountfy
- Para melhor performance, certifique-se de que os arquivos CSV estão corretamente formatados
- Mantenha sempre backups dos arquivos de origem antes de executar o processo

---
📌 **Desenvolvido por:** Lucian Silva para Tecadi Operador Logístico 🚛📦
