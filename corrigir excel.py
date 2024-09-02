import pandas as pd
import xlrd
import xlwt

# Caminho para o arquivo Excel original
file_path = 'C:\\Teste\\Sedex_fraa_smart.xls'

# Caminho para o arquivo Excel corrigido
corrected_file_path = 'C:\\Teste\Sedex_fraa_smart_corrigida.xls'

# Carregar o arquivo Excel
xls = pd.ExcelFile(file_path)

# Criar um novo arquivo Excel usando xlwt
workbook = xlwt.Workbook()

# Processar cada planilha
for sheet_name in xls.sheet_names:
    # Carregar a planilha em um DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Localizar onde a coluna D (4ª coluna) é 301 e a coluna E (5ª coluna) é 0
    condition = (df.iloc[:, 3] == 301) & (df.iloc[:, 4] == 0)
    
    # Substituir o valor 0 por 500 na coluna E onde a condição é atendida
    df.loc[condition, df.columns[4]] = 500
    
    # Preencher a coluna C com valores em branco, exceto pelo cabeçalho
    df.iloc[:, 2] = df.iloc[:, 2].replace(pd.NA, "")
    
    # Adicionar a planilha corrigida ao novo arquivo Excel
    sheet = workbook.add_sheet(sheet_name)
    
    # Escrever o cabeçalho
    for col_idx, header in enumerate(df.columns):
        sheet.write(0, col_idx, header)
    
    # Escrever os dados da planilha
    for row_idx, row in enumerate(df.values, start=1):
        for col_idx, value in enumerate(row):
            # Se for a coluna C, garantir que não seja NaN
            if pd.isna(value) and col_idx == 2:
                value = ""
            sheet.write(row_idx, col_idx, value)

# Salvar o arquivo Excel corrigido
workbook.save(corrected_file_path)

print("Arquivo corrigido salvo como:", corrected_file_path)