import pandas as pd

path = 'Modelo W1 Holdings - Teste técnico.xlsx'

contratos = pd.read_excel(path, sheet_name= 'Contratos')
estrutura = pd.read_excel(path, sheet_name= 'Estrutura comercial')
pagamentos = pd.read_excel(path, sheet_name= 'Pagamentos')
contratos_estruturas = pd.merge(contratos, estrutura, on='ID Consultor', how='inner')
parcelas_comissao = pd.merge(contratos_estruturas, pagamentos, on='ID de negócio', how='inner')

consultor = parcelas_comissao['Valor'] 
lider_consultor = parcelas_comissao['Valor'] * 0.05
closer = parcelas_comissao['Valor'] * 0.2
lider_closer = parcelas_comissao['Valor'] * 0.15



def calc_percent(x, percent):
    return x * percent



# comissao_calculada = parcelas_comissao.apply(calc_percent(parcelas_comissao['Valor'], 0.1), axis=1)

parcelas_comissao['Comissão Consultor'] = parcelas_comissao.apply(lambda x: calc_percent(x['Valor'], 0.1), axis = 1)
parcelas_comissao['Comissão Lider Consultor'] = parcelas_comissao.apply(lambda x: calc_percent(x['Valor'], 0.05), axis = 1)
parcelas_comissao['Comissão Closer'] = parcelas_comissao.apply(lambda x: calc_percent(x['Valor'], 0.2), axis = 1)
parcelas_comissao['Comissão Lider Closer'] = parcelas_comissao.apply(lambda x: calc_percent(x['Valor'], 0.15), axis = 1)
parcelas_comissao.to_excel('contratos_atualizado.xlsx', index=False)







