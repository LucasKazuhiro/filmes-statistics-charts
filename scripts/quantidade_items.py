import pandas as pd


def calc_quantidade_items(ws_stats, df_tabela_info, coluna, end_item):
    ws_stats[end_item].value = None

    # Vê se algum valor em df_tabela_info[coluna] é diferente de '-'
    # Caso seja diferente de '-', o retorno é TRUE e aquele valor em questão entrará para o novo "df_tabela_info"
    # Caso seka igual a '-', o retorno é FALSE e aquele valor em questão NÃO ENTRARÁ para o novo "df_tabela_info"
    df_tabela_info = df_tabela_info[df_tabela_info[coluna] != '-']
    # Conta a quantidade de valoers UNICOS contidos em uma coluna X (valor informado pela variável "coluna")
    qtd_unica_item = df_tabela_info[coluna].nunique()
    # Imprime o valor no endereço informado por parametro
    ws_stats[end_item].value = qtd_unica_item


    