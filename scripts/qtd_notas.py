import pandas as pd
from openpyxl.styles import numbers


def calc_qtd_notas(ws_stats, df_tabela_info, end_assistidos, end_nota_media, end_notas, end_nota_qtd, end_nota_porcentagem, end_nota_barra, barra_prog_meta, barra_prog_porcentagem):
    # Retorna a quantidade de filmes por nota. As notas estão em ordem decrescente (Nota 10 = 3 filmes, Nota 9 = 7 filmes etc)
    qtd_filmes_nota_dict = df_tabela_info.groupby('NOTA').size().sort_index(ascending=False).to_dict()


    # Apaga o valor
    ws_stats[end_assistidos].value = None
    # Calcula "Assistidos"
    qtd_filmes = sum(qtd_filmes_nota_dict.values())
    # Imprime "Assistidos"
    ws_stats[end_assistidos].value = qtd_filmes


    # Apaga o valor
    ws_stats[end_nota_media].value = None
    # Calcula a "Nota Média"
    nota_X_qtdfilme_array = []
    for nota, qtd_filme_por_nota in qtd_filmes_nota_dict.items():
        nota_X_qtdfilme_array.append(nota * qtd_filme_por_nota);

    if(qtd_filmes > 0):
        nota_media = sum(nota_X_qtdfilme_array) / qtd_filmes
        nota_media = float("{:.2f}".format(nota_media))
    else:
        nota_media = 0

    # Imprime "Nota Média"
    ws_stats[end_nota_media].value = nota_media



    # Pega toda as Notas (10, 9, 8, 7 ... 3, 2, 1) e coloca em um array
    notas_list = []
    rodar_notas = 0
    while(rodar_notas < 10):
        nota = ws_stats[end_notas].offset(row=rodar_notas).value
        nota = nota[5:]
        # Vai descendo as células até que se o valor dela for None (null), saia (break) do loop
        if(nota is None):
            break
        # Transforma a Nota encontrada em uma String e salva no array
        notas_list.append(str(nota))
        rodar_notas+=1


    qtd_linhas_descer = 0
    while(ws_stats[end_nota_qtd].offset(row=qtd_linhas_descer).value is not None):
        ws_stats[end_nota_qtd].offset(row=qtd_linhas_descer).value = None;
        ws_stats[end_nota_porcentagem].offset(row=qtd_linhas_descer).value = None;
        ws_stats[end_nota_barra].offset(row=qtd_linhas_descer).value = None;

        qtd_linhas_descer += 1



    # Loop de iteração sobre um dicionario
    for nota, qtd_filme_por_nota in qtd_filmes_nota_dict.items():
        try:
            # notas_list.index()    ----> Pega o valor de str(nota) e procura ele em notas_list (que é um array) e informa seu index 
            qtd_linhas_descer = notas_list.index(str(int(nota)))
        except ValueError:
            # Caso não ache o valor em notas_list, o retorno é -1
            qtd_linhas_descer = -1


        if(qtd_linhas_descer != -1):
            # Imprime a Quantidade de Filmes por Nota
            ws_stats[end_nota_qtd].offset(row=qtd_linhas_descer).value = qtd_filme_por_nota

            # Calcula o "Porcent"
            nota_porcentagem = "{:.2f}".format((qtd_filme_por_nota * 100) / qtd_filmes)
            # Imprime o "Porcent"
            ws_stats[end_nota_porcentagem].offset(row=qtd_linhas_descer).value = "[ " + nota_porcentagem.replace('.', ',') +"% ]"

            # Imprime os valores das Barras
            ws_stats[end_nota_barra].offset(row=qtd_linhas_descer).number_format = numbers.FORMAT_NUMBER_00
            ws_stats[end_nota_barra].offset(row=qtd_linhas_descer).value = float(nota_porcentagem)



    ws_stats[barra_prog_porcentagem].value = None

    # Verifica se existe algum texto de Meta
    if(ws_stats[barra_prog_meta].value != None):
        index_dois_pontos = ws_stats[barra_prog_meta].value.find(':')
        # Verifica se existe ':' no texto de Meta
        if(index_dois_pontos != -1):
            # Retorna o valor da Meta pegando tudo após 2 caracteres de ':'
            meta_qtd_filmes =  int(ws_stats[barra_prog_meta].value[index_dois_pontos+2:])

            # Calcula a Porcentagem de Progresso
            prog_porcentagem = (qtd_filmes * 100) / meta_qtd_filmes
            # Imprime a Porcentagem de Progesso
            ws_stats[barra_prog_porcentagem].value = prog_porcentagem


