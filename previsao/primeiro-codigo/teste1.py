import pandas as pd

# Carregar o arquivo Excel
file_path = 'Cópia de DADOS JULIOS.xlsx'  # Substitua pelo caminho correto do arquivo
df = pd.read_excel(file_path, sheet_name='Plan2')

# Somar as quantidades de vendas por coleção
total_por_colecao = df.groupby('COLECAO')['QTDE'].sum().reset_index()
print(total_por_colecao)

# Renomear a coluna para melhor entendimento
total_por_colecao.rename(columns={'QTDE': 'TOTAL_VENDIDO'}, inplace=True)
print(total_por_colecao)


# Salvar o resultado em um novo arquivo Excel
total_por_colecao.to_excel('total_por_colecao.xlsx', index=False)


print("Total vendido por coleção salvo em 'total_por_colecao.xlsx'")
