import pandas as pd


'''
ws_stats                ---> Worksheet de Stats (referência a guia "Stats" do Excel)
df_tabela_info          ---> Dataframe da tabela de informações contida na guia "Tabela" do Excel
coluna                  ---> Coluna no Datagrame onde deverá ser feito o groupby (agrupar por ANO, DIRETOR, DISTRIBUIDORA e etc)
qtd_filme_minimo        ---> Quantidade minima de filmes para entrar no Ranking
end_mais_filmes         ---> Endereço (na guia Stats) para Mais Filmes por XXXXX
end_maior_nota_media    ---> Endereço (na guia Stats) para Maior Nota Média por XXXXX
num_max_print           ---> Numero de linha que será impresso no Excel para compor o Ranking
xaxis_label_angulacao   ---> Angulação (inclinação) dos labels do Eixo X do gráfico gerado
'''

def calc_mais_maior(ws_stats, df_tabela_info, coluna, qtd_filme_minimo, end_mais_filmes, end_maior_nota_media, num_max_print, xaxis_label_angulacao):

    # Remove todas as informações presentes nos Rankings (apagar dados), e para de apagar na primeira célula sem valor
    ranking_linha_atual = 0
    while(ws_stats[end_mais_filmes].offset(row=ranking_linha_atual).value is not None):
        ws_stats[end_mais_filmes].offset(row=ranking_linha_atual).value = None
        ws_stats[end_maior_nota_media].offset(row=ranking_linha_atual).value = None
        ranking_linha_atual+=1


    # Cria dataframe agrupando por uma chave X (informada por 'coluna')
    df_agrupar = df_tabela_info.groupby(coluna)

    # Conta a quantidade de filmes por categoria do agrupamento
    df_mais_filme_por = df_agrupar['FILME'].count()
    # Calcula a média da nota por categoria do agrupamento
    df_maior_nota_media_por = df_agrupar['NOTA'].mean()


    # Junta os dois dataframes em um único. A chave X utilizada no agrupamento entra no dataframe da união
    df_mais_maior = pd.DataFrame({
        'Mais_Filmes_por': df_mais_filme_por,
        'Maior_Nota_Media_por': df_maior_nota_media_por
    }).reset_index() # reset_index() cria um novo index (basicamente numeros começando em 0)
                     # Já o antigo index (informado por "coluna") vira um coluna comum deixando de ser um index

    # Vê se algum valor em df_mais_maior[coluna] é diferente de '-'
    # Caso seja diferente de '-', o retorno é TRUE e aquele valor em questão entrará para o novo "df_mais_maior"
    # Caso seka igual a '-', o retorno é FALSE e aquele valor em questão NÃO ENTRARÁ para o novo "df_mais_maior"
    df_mais_maior = df_mais_maior[df_mais_maior[coluna] != '-']




    if(not df_tabela_info.empty):
        from scripts import grafico_linha
        # Cria um gráfico com os valores 'Mais_Filmes_por' e 'Maior_Nota_Media_por'
        print(f"Gerando grafico interativo de {coluna}")
        grafico_linha.gerar_grafico_linha(df_mais_maior, coluna, xaxis_label_angulacao)




    
    # Cria um novo dataframe organizado por "Mais_Filmes_por" (quantidade de filmes) em ordem decrescente
    df_mais = df_mais_maior.sort_values(by='Mais_Filmes_por', ascending=False).reset_index(drop=True) # drop=True pega o atual index e descarta ele, criando um novo index
    # Cria um novo dataframe organizado por "Maior_Nota_Media_por" em ordem decrescente
    df_maior = df_mais_maior.sort_values(by='Maior_Nota_Media_por', ascending=False).reset_index(drop=True) # drop=True pega o atual index e descarta ele, criando um novo index


    rodar_mais_maior = 0
    rodar_mais_maior_interno = 0
    # Loop para imprimir X linhas do ranking na tela
    while rodar_mais_maior < num_max_print and rodar_mais_maior < len(df_mais):
        # Define o numero minimo de filmes para que o item apareça no Ranking
        if(df_mais['Mais_Filmes_por'][rodar_mais_maior] >= qtd_filme_minimo):
            # Formata a String que será impressa no Excel
            string_mais_maior = "{}.  {} [{} Filmes]   ~{} pts".format(
                rodar_mais_maior + 1,
                df_mais[coluna][rodar_mais_maior],
                df_mais['Mais_Filmes_por'][rodar_mais_maior],
                "{:.2f}".format(df_mais['Maior_Nota_Media_por'][rodar_mais_maior]).replace('.', ',')
            )

            # Imprime "Mais Filmes por"
            ws_stats[end_mais_filmes].offset(row=rodar_mais_maior_interno).value = string_mais_maior

            rodar_mais_maior_interno+=1
        
        else:
            # Para o loop while
            break

        rodar_mais_maior+=1


    rodar_mais_maior = 0
    rodar_mais_maior_interno = 0
    # Loop para imprimir X linhas do ranking na tela
    while rodar_mais_maior_interno < num_max_print and rodar_mais_maior < len(df_maior):
        # Define o numero minimo de filmes para que o item apareça no Ranking
        if(df_maior['Mais_Filmes_por'][rodar_mais_maior] >= qtd_filme_minimo):
            # Formata a String que será impressa no Excel
            string_mais_maior = "{}.  {} [{} pts]   ~{} Filmes".format(
                rodar_mais_maior_interno + 1,
                df_maior[coluna][rodar_mais_maior],
                 "{:.2f}".format(df_maior['Maior_Nota_Media_por'][rodar_mais_maior]).replace('.', ','),
                df_maior['Mais_Filmes_por'][rodar_mais_maior]
            )

            # Imprime "Maior Nota Média por"
            ws_stats[end_maior_nota_media].offset(row=rodar_mais_maior_interno).value = string_mais_maior

            rodar_mais_maior_interno+=1

        rodar_mais_maior+=1


