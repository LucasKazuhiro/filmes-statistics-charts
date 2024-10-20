import pandas as pd


def calc_tempo_visto_medio(ws_stats, df_tabela_info, end_tempo_assistido, end_tempo_medio):
    ws_stats[end_tempo_assistido].value = None
    ws_stats[end_tempo_medio].value = None

    if(not df_tabela_info.empty):
        tempo_assistido_hora_array = []
        tempo_assistido_min_array = []

        rodar_df = 0;
        while rodar_df < df_tabela_info.shape[0]:
            # Salva os valores das horas dos filmes (Ex: 2,34 ---> nesse caso salva o 2, ou seja, 2 horas)
            tempo_assistido_hora_array.append(int(df_tabela_info['DURACAO'][rodar_df]))
            # Salva os valores dos minutos dos filmes (Ex: 2,34 --> nesse caso salva o 0,34 ou seja, 34 minutos)
            tempo_assistido_min_array.append(float("{:.2f}".format(df_tabela_info['DURACAO'][rodar_df] % 1)))

            rodar_df+=1;

        # Calcula "Tempo Visto (dias)"
        tempo_assistido_dias = float("{:.2f}".format((sum(tempo_assistido_hora_array) + (sum(tempo_assistido_min_array) / 0.6)) / 24.0))
        # Imprime "Tempo Visto (dias)"
        ws_stats[end_tempo_assistido].value = tempo_assistido_dias


        # Calcula "Tempo Médio"
        tempo_medio_horas = float("{:.2f}".format((tempo_assistido_dias * 24) / df_tabela_info.shape[0]))
        # Imprime "Tempo Médio"
        ws_stats[end_tempo_medio].value = tempo_medio_horas

    else:
        ws_stats[end_tempo_assistido].value = 0
        ws_stats[end_tempo_medio].value = 0