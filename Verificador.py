import pandas as pd
import os

def encontrar_notas_fiscais_nao_presentes(caminho_arquivo_principal, caminho_arquivo_1, caminho_arquivo_2, coluna_chave):
    try:
        # Carregar arquivos em DataFrames
        df_principal = pd.read_excel(caminho_arquivo_principal)
        print("\nArquivo principal carregado...")

        df_1 = pd.read_excel(caminho_arquivo_1)
        print("Arquivo 1 carregado...")

        df_2 = pd.read_excel(caminho_arquivo_2)
        print("Arquivo 2 carregado...")

        # Verificar e remover colunas duplicadas nos DataFrames (se existirem)
        df_principal = df_principal.loc[:, ~df_principal.columns.duplicated()]
        df_1 = df_1.loc[:, ~df_1.columns.duplicated()]
        df_2 = df_2.loc[:, ~df_2.columns.duplicated()]

        # Combinar os arquivos 1 e 2 em um único DataFrame para comparação
        df_comparacao = pd.concat([df_1, df_2]).drop_duplicates(subset=[coluna_chave])  # Remove duplicatas da coluna-chave
        print("Arquivo de comparação criado...")

        # Identificar os números de nota fiscal que estão no arquivo de comparação mas NÃO estão no arquivo principal
        notas_fiscais_nao_presentes = df_comparacao[~df_comparacao[coluna_chave].isin(df_principal[coluna_chave])]

        print("\nNotas fiscais não presentes no arquivo principal:")
        print(notas_fiscais_nao_presentes)

        # Verifica se o arquivo já existe e exclui, para garantir que ele seja refeito
        arquivo_saida = 'C:/Users/fiscal/.vscode/Verificador de planilhas/notas_nao_presentes.xlsx'
        if os.path.exists(arquivo_saida):
            os.remove(arquivo_saida)  # Remove o arquivo existente
            print(f"\nArquivo {arquivo_saida} removido para atualização...")

        # Salvar o resultado em um novo arquivo Excel
        notas_fiscais_nao_presentes.to_excel(arquivo_saida, index=False)
        print(f"\n\nAs notas fiscais não presentes foram salvas em '{arquivo_saida}'.\n\n")

    except Exception as e:
        print(f"Ocorreu um erro ao processar os arquivos: {e}")

# Definindo os caminhos dos arquivos
caminho_arquivo_principal = 'C:/Users/fiscal/.vscode/Verificador de planilhas/Relatoriobsoft.xlsx'
caminho_arquivo_1 = 'C:/Users/fiscal/.vscode/Verificador de planilhas/nsdocs 20_03_2025 11_49_53.xlsx'
caminho_arquivo_2 = 'C:/Users/fiscal/.vscode/Verificador de planilhas/gissonline084CC7A6155D00113BB16D0A4E4085BD.xlsx'

# Definindo o nome da coluna principal de comparação    
coluna_chave = 'NUMERO DA NOTA'

# Chama a função para encontrar as notas fiscais não presentes
encontrar_notas_fiscais_nao_presentes(caminho_arquivo_principal, caminho_arquivo_1, caminho_arquivo_2, coluna_chave)
