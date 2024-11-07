'''

    PARA PERMITIR QUE O SCRIPT FUNCIONE, EXECUTAR OS SEGUINTES COMANDOS NO CMD:
    pip install openpyxl
    pip install pandas
    pip install plotly

'''


from openpyxl import load_workbook
import pandas as pd
import time

from scripts import enderecos


# Salvando o tempo em que o script começou a ser executado
tempo_inicial_script = time.time()


print("\n\n\nIniciando execucao do Script!\n")
print("Carregando Workbook e Worksheets...")
# Carrega o Workbook MovieList.xlsx em um variável
wb_movielist = load_workbook(filename = 'MovieList.xlsx')

# Carrega as guias (worksheets) em variáveis
ws_stats = wb_movielist["Stats"]

print("Criando dataframe com os dados...")
# Cria um dataframe a partir da tabela presente na guia "Tabela" e salva em uma variável
df_tabela_info = pd.read_excel('MovieList.xlsx', 'Tabela', skiprows = 4, usecols = 'C:O')
# Limpa as linhas vazias do "df_tabela_info"
df_tabela_info = df_tabela_info.dropna(how='all')



'''
Preenche os valores em:
    Barra de Progresso
        Calcula a porcentagem e imprime na barra

    Quantidade de Notas
        Filmes     Porcent     Barras de Progresso

    Estatisticas Gerais
        Nota Média     Assistidos
'''
print("Calculando 'Nota Media' e 'Assistidos'...")
print("Preenchendo os dados em 'Quantidade de Notas'...")
print("Atualizando barra de progesso...")
from scripts import qtd_notas
qtd_notas.calc_qtd_notas(ws_stats, df_tabela_info, enderecos.end_assistidos, enderecos.end_nota_media, enderecos.end_notas, enderecos.end_nota_qtd, enderecos.end_nota_porcentagem, enderecos.end_nota_barra, enderecos.barra_prog_meta, enderecos.barra_prog_porcentagem)



'''
Preenche os valores em:
    Estatisticas Gerais
        Tempo Visto (dias)     Tempo Médio
'''
print("Calculando 'Tempo Visto (dias)' e 'Tempo Medio'...")
from scripts import tempo_visto_medio
tempo_visto_medio.calc_tempo_visto_medio(ws_stats, df_tabela_info, enderecos.end_tempo_assistido, enderecos.end_tempo_medio)



'''
Preenche os valores em:
    Estatisticas Gerais
        Prox. Filme Seq. (anos)
'''
print("Calculando 'Prox. Filme Seq. (anos)'...")
from scripts import prox_filme_seq
prox_filme_seq.calc_prox_filme_seq(ws_stats, df_tabela_info, enderecos.end_prox_filme_seq)



'''
Junta X colunas em apenas um Dataframe 
    O Dataframe retornado possui:
        Chave (valor informado nos parametros em 'novo_nome_coluna_uniao')
        NOTA
        FILME    
'''
print("Juntado colunas para efetuar calculos...")
from scripts import x_cols_para_uma
df_genero = x_cols_para_uma.transformar_x_cols_para_uma(df_tabela_info, "GENERO_1", 2, "GENERO")
df_pais_origem = x_cols_para_uma.transformar_x_cols_para_uma(df_tabela_info, "PAIS_ORIGEM_1", 3, "PAIS_ORIGEM")



'''
Preenche os valores em:
    Estatisticas Gerais
        Diretores     Pais de Origem     Franquias     Distribuidoras
'''
print("Calculando 'Diretores', 'Pais de Origem', 'Franquias' e 'Distribuidores'...")
from scripts import quantidade_items
quantidade_items.calc_quantidade_items(ws_stats, df_tabela_info, "DIRETOR", enderecos.end_diretores)
quantidade_items.calc_quantidade_items(ws_stats, df_pais_origem, "PAIS_ORIGEM", enderecos.end_pais_origem)
quantidade_items.calc_quantidade_items(ws_stats, df_tabela_info, "FRANQUIA", enderecos.end_franquias)
quantidade_items.calc_quantidade_items(ws_stats, df_tabela_info, "DISTRIBUIDORA", enderecos.end_distribuidoras)



'''
Preenche os valores em:
    Stats Gerais de Filmes por XXXXX
        Avg     Qtd.     Porcent.
'''
from scripts import stats_gerais_filme_por
print("Preenchendo 'Stats Gerais de Filmes por Genero'...")
stats_gerais_filme_por.calc_stats_gerais_filme_por(ws_stats, df_genero, "GENERO", enderecos.end_generos, enderecos.end_genero_nota_media, enderecos.end_genero_qtd, enderecos.end_genero_porcentagem, enderecos.end_genero_barra, 17)
print("Preenchendo 'Stats Gerais de Filmes por Class.Ind.'...")
stats_gerais_filme_por.calc_stats_gerais_filme_por(ws_stats, df_tabela_info, "CLASS_IND", enderecos.end_classinds, enderecos.end_classind_nota_media, enderecos.end_classind_qtd, enderecos.end_classind_porcentagem, enderecos.end_classind_barra, 6)



'''
Preenche os valores em:
    Mais Filmes por XXXXX
        Imprime ranking

    Maior Nota Média por XXXXX
        Imprime ranking
'''
from scripts import mais_maior
# Mais_Média por Gênero
print("Preenchendo ranking 'MaisMaior' de Genero...")
mais_maior.calc_mais_maior(ws_stats, df_genero, "GENERO", 1, enderecos.end_mais_filme_genero, enderecos.end_maior_nota_media_genero, 10, 30)

# Mais_Média por Classificação Indicativa
print("Preenchendo ranking 'MaisMaior' de Class. Ind...")
mais_maior.calc_mais_maior(ws_stats, df_tabela_info, "CLASS_IND", 1, enderecos.end_mais_filme_classind, enderecos.end_maior_nota_media_classind, 6, 0)

# Mais_Média por Ano
print("Preenchendo ranking 'MaisMaior' de Ano...")
mais_maior.calc_mais_maior(ws_stats, df_tabela_info, "ANO", 5, enderecos.end_mais_filme_ano, enderecos.end_maior_nota_media_ano, 7, 0)

# Mais_Média por Diretor
print("Preenchendo ranking 'MaisMaior' de Diretor...")
mais_maior.calc_mais_maior(ws_stats, df_tabela_info, "DIRETOR", 3, enderecos.end_mais_filme_diretor, enderecos.end_maior_nota_media_diretor, 7, 30)

# Mais_Média por Franquia
print("Preenchendo ranking 'MaisMaior' de Franquia...")
mais_maior.calc_mais_maior(ws_stats, df_tabela_info, "FRANQUIA", 3, enderecos.end_mais_filme_franquia, enderecos.end_maior_nota_media_franquia, 7, 30)

# Mais_Média por País de Origem
print("Preenchendo ranking 'MaisMaior' de Pais de Origem...")
mais_maior.calc_mais_maior(ws_stats, df_pais_origem, "PAIS_ORIGEM", 5, enderecos.end_mais_filme_pais, enderecos.end_maior_nota_media_pais, 7, 30)

# Mais_Média por Distribuidora
print("Preenchendo ranking 'MaisMaior' de Distribuidora...")
mais_maior.calc_mais_maior(ws_stats, df_tabela_info, "DISTRIBUIDORA", 5, enderecos.end_mais_filme_distribuidora, enderecos.end_maior_nota_media_distribuidora, 7, 30)








# Salva as alterações efetudas no workbook
print("Salvando as alteracoes feitas no Excel...")
wb_movielist.save('MovieList.xlsx')



# Salvando o tempo em que o script terminou de ser executado
tempo_final_script = time.time()

# Calculando tempo...
tempo_execucao_script = "{:.4f}".format(tempo_final_script - tempo_inicial_script)

# Imprime tempo de execução na tela
print(f"\n\nTempo de Execucao: {tempo_execucao_script} segundos\n\n\n\n\n")




