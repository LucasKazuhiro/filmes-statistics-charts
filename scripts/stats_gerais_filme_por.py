import pandas as pd
from openpyxl.styles import numbers


'''
ws_stats                ---> Worksheet de Stats (referência a guia "Stats" do Excel)
df_tabela_info          ---> Dataframe da tabela de informações contida na guia "Tabela" do Excel
coluna                  ---> Coluna no Datagrame onde deverá ser feito o groupby (agrupar por ANO, DIRETOR, DISTRIBUIDORA e etc)
end_opcoes              ---> Endereço (na guia Stats) para Opções (gênero ou classificação indicativa)
end_nota_media          ---> Endereço (na guia Stats) para Maior Nota Média por XXXXX
end_qtd                 ---> Endereço (na guia Stats) para Quantidade de Filmes por XXXXX
end_porcentagem         ---> Endereço (na guia Stats) para Porcentagem
end_barra               ---> Endereço (na guia Stats) para Grafico de Barra 
num_max_print           ---> Numero de linha que será impresso no Excel para compor o Stats Gerais de Filmes por XXXXX
'''

def calc_stats_gerais_filme_por(ws_stats, df_tabela_info, coluna, end_opcoes, end_nota_media, end_qtd, end_porcentagem, end_barra, num_max_print):

    # Remove todas as informações presentes nos Rankings (apagar dados), e para de apagar na primeira célula sem valor
    ranking_linha_atual = 0
    while(ws_stats[end_opcoes].offset(row=ranking_linha_atual).value is not None):
        ws_stats[end_opcoes].offset(row=ranking_linha_atual).value = None
        ws_stats[end_nota_media].offset(row=ranking_linha_atual).value = None
        ws_stats[end_qtd].offset(row=ranking_linha_atual).value = None
        ws_stats[end_porcentagem].offset(row=ranking_linha_atual).value = None
        ws_stats[end_barra].offset(row=ranking_linha_atual).value = None

        ranking_linha_atual+=1


    # Transforma os valores da coluna key (informada por 'coluna') em String (para evitar erro no sort_values)
    df_tabela_info[coluna] = df_tabela_info[coluna].astype(str)

    # Cria dataframe agrupando por uma chave X (informada por 'coluna')
    df_agrupar = df_tabela_info.groupby(coluna)

    # Conta a quantidade de filmes por categoria do agrupamento
    df_qtd_filme_por = df_agrupar['FILME'].count()
    # Calcula a média da nota por categoria do agrupamento
    df_nota_media_por = df_agrupar['NOTA'].mean().apply(lambda x: f"{x:.2f}")

    # Junta os dois dataframes em um único. A chave X utilizada no agrupamento entra no dataframe da união
    df_opcoes = pd.DataFrame({
        'QTD_FILME_POR': df_qtd_filme_por,
        'NOTA_MEDIA_POR': df_nota_media_por
    }).reset_index() # reset_index() cria um novo index (basicamente numeros começando em 0)
                     # Já o antigo index (informado por "coluna") vira um coluna comum deixando de ser um index

    # Vê se algum valor em df_opcoes[coluna] é diferente de '-'
    # Caso seja diferente de '-', o retorno é TRUE e aquele valor em questão entrará para o novo "df_opcoes"
    # Caso seka igual a '-', o retorno é FALSE e aquele valor em questão NÃO ENTRARÁ para o novo "df_opcoes"
    df_opcoes = df_opcoes[df_opcoes[coluna] != '-']

    # Organiza o dataframe por "coluna" (valor informado nos parametros) em ordem alfabética
    df_opcoes.sort_values(by=coluna, ascending=True).reset_index(drop=True) # drop=True pega o atual index e descarta ele, criando um novo index


    # Calcula a quantidade total de filmes assistidos que possuem algum(a) X (coluna)
    total_qtd_filme_por = df_opcoes['QTD_FILME_POR'].sum()


    qtd_linhas_descer = 0
    # O 'itertuples' ajuda a iterar sobre um dataframe
    for linha_df_opcao in df_opcoes.itertuples(index=False):
        if(qtd_linhas_descer < num_max_print):
            # Imprime os valores de "coluna" (keys)
            ws_stats[end_opcoes].offset(row=qtd_linhas_descer).value = getattr(linha_df_opcao, coluna)

            # Imprime a Nota Média ("Avg") por X
            ws_stats[end_nota_media].offset(row=qtd_linhas_descer).number_format = numbers.FORMAT_NUMBER_00
            ws_stats[end_nota_media].offset(row=qtd_linhas_descer).value = float(linha_df_opcao.NOTA_MEDIA_POR)

            # Imprime a Quantidade de filmes ("Qtd") por X
            ws_stats[end_qtd].offset(row=qtd_linhas_descer).value = linha_df_opcao.QTD_FILME_POR

            # Calcula o "Porcent"
            opcao_porcentagem = "{:.2f}".format((linha_df_opcao.QTD_FILME_POR * 100) / total_qtd_filme_por)
            # Imprime o "Porcent"
            ws_stats[end_porcentagem].offset(row=qtd_linhas_descer).value = "[ " + opcao_porcentagem.replace('.', ',') + "% ]"

            # Imprime os valores das Barras
            ws_stats[end_barra].offset(row=qtd_linhas_descer).number_format = numbers.FORMAT_NUMBER_00
            ws_stats[end_barra].offset(row=qtd_linhas_descer).value = float(opcao_porcentagem)

            qtd_linhas_descer+=1

        else:
            # Sair do loop while
            break;