import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import math


def gerar_grafico_linha(df_mais_maior, coluna, xaxis_label_angulacao):

    # Para df_mais_maior[coluna] com muitos valores Strings (ou 'object' no Pandas)
    # é necessario remover algumas linhas para não poluir o gráfico
    # Se o numero total de linhas for maior que 'df_tipo_string_limite', o dataframe sera cortado
    df_tipo_string_limite = 25

    # Verifica se a coluna principal é do tipo String (object)
    if(df_mais_maior[coluna].dtype == object):
        # Verifica se o total de linhas é maior que 'df_tipo_string_limite'
        if(df_mais_maior[coluna].count() > df_tipo_string_limite):
            # Organiza o dataframe por 'Mais_Filmes_por' em ordem decrescente, e depois por 'Maior_Nota_Media_por' em ordem decrescente
            # inplace=True é para salvar as modificações no próprio dataframe (ao invés de criar uma cópia)
            df_mais_maior.sort_values(by=['Mais_Filmes_por', 'Maior_Nota_Media_por'], ascending=[False, False], inplace=True) 
            # Seleciona todas as linhas até o indice 'df_tipo_string_limite - 1' (excludente)
            df_mais_maior = df_mais_maior.iloc[:df_tipo_string_limite].copy()
            # Organiza o dataframe por 'coluna' em ordem alfabetica
            df_mais_maior.sort_values(by=coluna, ascending=True, inplace=True) 
            # Reseta o index do dataframe
            df_mais_maior.reset_index(drop=True, inplace=True)




    # Definição das cores usadas para personalizar o gráfico
    cor_linhas_grid = "#261C2F"
    cor_plot_fundo = "#101010"


    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(      # Adiciona um "trace" (traço, risco, rastro) do tipo "Scatter" no plot 
        go.Scatter(
            x=df_mais_maior[coluna],                # Define os valores do Eixo X
            y=df_mais_maior['Mais_Filmes_por'],     # Define os valores do Eixo Y principal
            name="Quantidade de Filmes",            # Define o nome do Eixo X presente na Legenda 
            mode='lines+markers',                   # Define o modo do simbolo usado no grafico (linhas && marcadores/bolinhas)
            marker=dict(size=4),                    # Define o tamanho dos marcadores/bolinhas
            line=dict(width=0.7),                   # Define a espessura das linhas
        ),
        secondary_y = False                         # Aplica todas as alterações no Eixo Y principal
    )

    fig.add_trace(
        go.Scatter(
            x=df_mais_maior[coluna], 
            y=df_mais_maior['Maior_Nota_Media_por'],    # Define os valores do Eixo Y secundário
            name="Nota Média",
            mode='lines+markers',
            marker=dict(size=4),
            line=dict(width=0.7),   
        ),
        secondary_y = True                             # Aplica todas as alterações no Eixo Y secundário
    )



    # Adiciona +2 ao maior valor em df_mais_maior['Mais_Filmes_por']
    y_axis_left_range = df_mais_maior['Mais_Filmes_por'].max() + 2
    # Divide o numero conseguido por 10, arredonda para cima, e depois multiplica novamente por 10
    # Exemplo: 
    #           valor máximo = 16
    #           16 + 2 = 18
    #           18/10 = 1.8
    #           2   (arredonda para cima)
    #           2 * 10 = 20 
    # Esse calculo serve para encontrar o primeiro valor divisivel por 10 mais proximo do valor máximo de df_mais_maior['Mais_Filmes_por']
    # Se o valor máximo é 16, o valor divisivel por 10 mais próximo é 20
    y_axis_left_range = math.ceil(y_axis_left_range / 10) * 10

    # Define o dtick
    # Se o range maximo for 20 (calculado no exemplo anterior)
    # Entao o dtick sera 20/10 = 2
    y_axis_left_dtick = y_axis_left_range / 10

    # Dessa forma cria-se um intervalo de DEZ VALORES que vai de 0 a 20, pulando de 2 em 2
    # Intervalo = [0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20] ----> 10 valores ao todo
    # O '0' é definido por min
    # O '20' foi calculado no primeiro exemplo
    # O '2' foi calculado no segundo exemplo
    # Tudo isso para alinhas o Eixo Y principal (quantidade de filmes) com o Eixo Y secundário (que vai de 0 a 10, são as notas das médias)



    fig.update_layout(
        title_text=f"Quantidade de filmes e média de nota por {coluna}",    # Define o titulo do plot (grafico)
        font=dict(color='white'),                                           # Define a cor de todas as palavras
        
        xaxis=dict(
            title=coluna,
            gridcolor=cor_linhas_grid,          # Define a cor da linha do grid do Eixo X
            tickangle=xaxis_label_angulacao,    # Define o grau de inclinação das labels do Eixo X
            tickfont=dict(size=10)              # Define o tamanho dos labels do Eixo X
        ),

        yaxis=dict(
            title='Quantidade de Filmes',       # Define o nome da label do Eixo Y principal
            gridcolor=cor_linhas_grid,          # Define a cor da linha do grid do Eixo Y principal
            range=[0, y_axis_left_range],       # Define o range do Eixo Y principal
            dtick = y_axis_left_dtick,          # Define o passo dos valores
            zeroline=False                      # Remove a linha de origem do Eixo Y principal
        ),

        yaxis2=dict(
            title='Nota Média', # Define o nome da label do Eixo Y secundário
            showgrid=False,     # O yaxis2 (Eixo Y secundário não mostrará as linhas de grid)
            range=[0, 10],      # Define o range do Eixo Y secundário
            dtick=1,            # Define o passo dos valores (aqui vão de 1 em 1 [0, 1, 2, 3, 4 etc...])
            zeroline=False      # Remove a linha de origem do Eixo Y secundário
        ),
        
        legend=dict(
            y=1.17,     # Posição vertical da Legenda
            x=0.70      # Posição horizontal da Legenda
        ),

        plot_bgcolor=cor_plot_fundo,     # Cor de fundo da área do gráfico (plot)
        paper_bgcolor=cor_plot_fundo,    # Cor de fundo da área de fora

        margin=dict(l=70, r=0, t=100, b=0),      # Define as margens do plot
    )





    # Passa todo o plot (gráfico) para um arquivo HTML.
    # E baixa ele na pasta contendo o arquivo "movielist-stats.py"
    # O valor entre "()" é o nome desse novo arquivo criado
    fig.write_html(f"{coluna}_MaisMaior.html")     


