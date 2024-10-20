import pandas as pd


def transformar_x_cols_para_uma(df_tabela_info, primeira_coluna, qtd_cols, novo_nome_coluna_uniao):
    # Pega a posição numerica da "primeira_coluna". Não a primeira coluna da tabela
    # mas sim a primeira coluna que deve entrar na unificação de colunas
    posicao_primeira_coluna = df_tabela_info.columns.get_loc(primeira_coluna)

    # Pega o nome de todas as colunas de devem entrar na unificação de colunas
    # O ".columns" retorna o nome de todas as colunas do Dataframe
    # O loop for é utilizado para rodar os nomes das colunas a partir da posição da "primeira_coluna" até o limite estabelecido (qtd_cols)
    colunas_nome_array = [df_tabela_info.columns[posicao_primeira_coluna + i] for i in range(qtd_cols)]

    colunas_cortadas_df_array = []
    for coluna_nome in colunas_nome_array:
        # Cria um Dataframe cópia possuindo as colunas "coluna_nome" (que vai mudando devido ao loop for), 'NOTA' e 'FILME'
        coluna_cortada_df = df_tabela_info[[coluna_nome, 'NOTA', 'FILME']].copy()
        # Troca o nome de "coluna_nome" para "novo_nome_coluna_uniao" (ou seja, padroniza o nome da primeira coluna)
        coluna_cortada_df.rename(columns={coluna_nome: novo_nome_coluna_uniao}, inplace=True)

        # Adiciona esse Dataframe criado ao array
        colunas_cortadas_df_array.append(coluna_cortada_df)


    # Unifica todos os Dataframes presentes em "colunas_cortadas_df_array"
    # Isso pode ser feito porque o nome de todas as colunas são iguais
    # Por isso o rename foi feito todas as vezes em uma das colunas
    df_colunas_juntadas = pd.concat(colunas_cortadas_df_array, axis = 0, ignore_index = True)

    return df_colunas_juntadas


