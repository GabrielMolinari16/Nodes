import pandas as pd

# Carregar o arquivo Excel
file_path = 'Cópia de DADOS JULIOS.xlsx'  # Substitua pelo caminho correto do arquivo
df = pd.read_excel(file_path, sheet_name='Plan2')

# Somar as quantidades de vendas por CNPJ e por coleção
soma_vendas = df.groupby(['CNPJ', 'COLECAO'])['QTDE'].sum().reset_index()

# Calcular a média de compra por CNPJ
media_vendas = soma_vendas.groupby('CNPJ')['QTDE'].mean().reset_index()

# Renomear as colunas para melhor entendimento
soma_vendas.rename(columns={'QTDE': 'SOMA_QTDE'}, inplace=True)
media_vendas.rename(columns={'QTDE': 'MEDIA_QTDE'}, inplace=True)

# Combinar as somas e médias em um único dataframe
resultado_final = pd.merge(soma_vendas, media_vendas, on='CNPJ')

# Salvar o resultado em um novo arquivo Excel
resultado_final.to_excel('resultado_final.xlsx', index=False)

print("Resultado salvo em 'resultado_final.xlsx'")
