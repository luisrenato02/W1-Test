import pandas as pd

# Caminho para o arquivo Excel
path = 'Modelo W1 Holdings - Teste técnico.xlsx'

# Leitura das planilhas
contratos = pd.read_excel(path, sheet_name='Contratos')
pagamentos = pd.read_excel(path, sheet_name='Pagamentos')
consultores = pd.read_excel(path, sheet_name='Estrutura comercial')
closers = pd.read_excel(path, sheet_name='Estrutura comercial closers')

# Merge das tabelas contratos e pagamentos
pacelas_closer_consultor = pd.merge(contratos, pagamentos, on='ID de negócio', how='inner')

# Função para calcular percentuais
def calc_percent(x, percent):
    return x * percent

# Remover a coluna 'Valor Holdings' e as colunas 'Unnamed'
pacelas_closer_consultor = pacelas_closer_consultor.drop(columns=['Valor Holdings'])
pacelas_closer_consultor = pacelas_closer_consultor.loc[:, ~pacelas_closer_consultor.columns.str.contains('^Unnamed')]

# Calcular comissões
pacelas_closer_consultor['Comissão Closer'] = pacelas_closer_consultor.apply(lambda x: calc_percent(x['Valor'], 0.2), axis=1)
pacelas_closer_consultor['Comissão Líder Closer'] = pacelas_closer_consultor.apply(lambda x: calc_percent(x['Valor'], 0.15), axis=1)
pacelas_closer_consultor['Comissão Consultor'] = pacelas_closer_consultor.apply(lambda x: calc_percent(x['Valor'], 0.1), axis=1)
pacelas_closer_consultor['Comissão Líder Consultor'] = pacelas_closer_consultor.apply(lambda x: calc_percent(x['Valor'], 0.05), axis=1)

# Adicionar coluna para mês e ano
pacelas_closer_consultor['Data de pagamento'] = pd.to_datetime(pacelas_closer_consultor['Data de pagamento'])
pacelas_closer_consultor['Ano-Mes'] = pacelas_closer_consultor['Data de pagamento'].dt.to_period('M')

# Merge com a tabela de closers
pacelas_closer_consultor = pd.merge(pacelas_closer_consultor, closers, on='Closer ID', how='inner')
# Merge com a tabela de Consultores
pacelas_closer_consultor = pd.merge(pacelas_closer_consultor, consultores, on='ID Consultor', how='inner')
# pacelas_consultor = pd.merge(pacelas_closer_consultor, consultores, on='ID Consultor', how='inner')
# print(pacelas_consultor)


# Agrupar por 'Ano-Mes' e 'Closer'
gasto_mensal_comissao = pacelas_closer_consultor.groupby(['Ano-Mes','ID de negócio', 'Closer', 'Consultor', 'Líder consultor']).agg({
    'Closer ID': 'first',
    'ID Consultor': 'first',
    'Valor': 'sum',
    'Comissão Closer': 'sum',
    'Comissão Líder Closer': 'sum',
    'Comissão Consultor': 'sum',
    'Comissão Líder Consultor': 'sum'
}).reset_index()

# Adicionar a coluna 'Comissão Total' ao DataFrame agrupado
gasto_mensal_comissao['Comissão Total'] = (
    gasto_mensal_comissao['Comissão Closer'] +
    gasto_mensal_comissao['Comissão Líder Closer'] +
    gasto_mensal_comissao['Comissão Consultor'] +
    gasto_mensal_comissao['Comissão Líder Consultor']
)

# calculando gasto total bruno
# bruno = gasto_mensal_comissao[gasto_mensal_comissao['Closer'] == 'Bruno']
# total_ganho_bruno =  (bruno['Comissão Closer'].sum() + gasto_mensal_comissao['Comissão Líder Closer'].sum())
# gasto_mensal_comissao.at[0, 'Comissão Total Bruno'] = 'R${:,.2f}'.format(total_ganho_bruno)



def somar_comissao(row):
    if row['Closer'] == 'Bruno':
        return row['Comissão Closer'] + row['Comissão Líder Closer']
    else:
        return row['Comissão Closer']

# gasto_mensal_comissao['Comissão Closer'] = gasto_mensal_comissao.apply(lambda row: somar_comissao(row), axis=1)

gasto_mensal_descrito = gasto_mensal_comissao.drop(columns=['Valor', 'Closer ID', 'ID Consultor'])

#Tabela somente gastos
somente_gasto_total = gasto_mensal_descrito.groupby(['Ano-Mes']).agg({
    'Comissão Total': 'sum'
}).reset_index()
somente_gasto_total['Comissão Total'] = somente_gasto_total['Comissão Total'].apply(lambda x: 'R${:,.2f}'.format(x))


#Inserindo R$ nos valores
format_columns = ['Comissão Closer', 'Comissão Consultor', 'Comissão Líder Consultor', 'Comissão Total', 'Comissão Líder Closer']
for col in format_columns:
    gasto_mensal_descrito[col] = gasto_mensal_descrito[col].map('R${:,.2f}'.format)

#somente gasto total


# Escrever o resultado para um novo arquivo Excel
somente_gasto_total.to_excel('Comissão Total Mensal.xlsx', index=False)
gasto_mensal_descrito.to_excel('Comissão Total Mensal Separada.xlsx', index=False)






