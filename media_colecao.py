import pandas as pd

# Carregar o arquivo Excel
file_path = 'resultado_final.xlsx'  # Substitua pelo caminho correto do arquivo
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Somar as quantidades de vendas por coleção
total_por_colecao = df.groupby('COLECAO')['SOMA_QTDE'].transform('sum')

# Adicionar o total por coleção como uma nova coluna no DataFrame original
df['TOTAL_POR_COLECAO'] = total_por_colecao
df['%_PESO_COLECAO'] = df['SOMA_QTDE'] / df['TOTAL_POR_COLECAO']

# Agora você pode continuar com as operações desejadas ou salvar o resultado
df.to_excel('resultado_final.xlsx', index=False)

print("Resultado com total por coleção salvo em 'resultado_final.xlsx'")
