import pandas as pd
from openpyxl.styles import numbers


def calc_prox_filme_seq(ws_stats, df_tabela_info, end_prox_filme_seq):
    ws_stats[end_prox_filme_seq].value = None

    # Vê se algum valor em df_tabela_info['FRANQUIA'] é diferente de '-'
    # Caso seja diferente de '-', o retorno é TRUE e aquele valor em questão entrará para o "df_franquia_ano_sort"
    # Caso seka igual a '-', o retorno é FALSE e aquele valor em questão NÃO ENTRARÁ para o "df_franquia_ano_sort"
    df_franquia_ano_sort = df_tabela_info[df_tabela_info['FRANQUIA'] != '-']

    # Organiza por Franquia em ordem alfabética, e depois suborganiza por Ano em ordem crescente
    df_franquia_ano_sort = df_franquia_ano_sort.sort_values(by=['FRANQUIA', 'ANO'])

    # Categoriza por 'FRANQUIA' e depois calcula a diferença entre os anos
    # Armazena os resultados em uma nova coluna chamada 'PROX_FILME_SEQ_TEMPO'
    df_franquia_ano_sort['PROX_FILME_SEQ_TEMPO'] = df_franquia_ano_sort.groupby('FRANQUIA')['ANO'].diff()

    # Faz a média das diferenças de tempo entre o lançamento dos filmes
    prox_filme_seq_tempo_media = "{:.2f}".format(df_franquia_ano_sort['PROX_FILME_SEQ_TEMPO'].mean())


    # Imprime 'Prox. Filme Seq. (anos)'
    ws_stats[end_prox_filme_seq].number_format = numbers.FORMAT_NUMBER_00
    ws_stats[end_prox_filme_seq].value = float(prox_filme_seq_tempo_media)


