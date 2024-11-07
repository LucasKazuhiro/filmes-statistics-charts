# Sobre o projeto
Este simples projeto foi feito com o intuíto de estudar um pouco de manipulação de Dataframes com [Pandas](https://pandas.pydata.org/), arquivos Excel com [Openpyxl](https://openpyxl.readthedocs.io/), e criação de gráficos com [Plotly](https://plotly.com/).\
\
Os scripts têm como objetivo gerar gráficos html e tabelas no excel a respeito dos dados presentes na tabela, ajudando a análisar diversas informações a respeito dos filmes registrados.


---


# Guias da Planilha
## Nota
Guia para fazer anotações das notas. Não possui relação alguma com os Scripts, serve apensa para organização geral.

## Tabela
É o local onde deverá ser colocado as `informações dos filmes assistidos`. Essa tabela possui as seguintes colunas:

| Nome da Coluna           | Significado                                                         | Valores / Exemplos                                            |                       
|:-------------------------|:--------------------------------------------------------------------|:--------------------------------------------------------------|
| **FRANQUIA**             | O nome comum pelo qual um filme poderá ser reconhecido.             | Avengers                                                      | 
| **NOTA**                 | A sua nota para o filme (valor inteiro)                             | 1, 2, 3, 4, 5, 6, 7, 8, 9, 10                                 |         
| **FILME**                | Nome do filme                                                       | Interstellar, Little Women, Back to the Future ...            |
| **ANO**                  | Ano de estreia                                                      | ... 1965, 2012, 2015, 2020, 2023...                           |         
| **DURACAO**              | Tempo do filme **(Ex: 2,32 -> Duas horas e trinta e dois minutos)** | ... 1,40 \| 1,49 \| 2,00 \| 2,54 \| ...                       |
| **GENERO_1 (e 2)**       | Gênero (genre) do filme                                             | Ação, Aventura, Drama, Comédia, Thriller ...                  |
| **DIRETOR**              | Nome do diretor                                                     | Christopher Nolan, Matthew Vaughn, David Fincher ...          |
| **CLASS_IND**            | Classificação de idade                                              | Livre, 10, 12, 14, 16, 18                                     |
| **PAIS_ORIGEM_1 (2, 3)** | Nome do país de origem                                              | Brasil, Estados Unidos, Espanha, França ...                   |                
| **DISTRIBUIDIRA**        | Empresa distribuidora do filme                                      | Warner Bros. Pictures, Sony Pictures Releasing, Lionsgate ... |


## Stats
É o local onde as `tabelas e rankings de estatísticas` aparecerão após executar o Script. Eles serão gerados com textos simples nas próprias células do Excel.\
Aqui você **não deverá escrever nada**, pois diversas células possuem seus dados apagados para gerarem novas informações.


---


# Pré requisitos
## Python
Ter o python instalado. Veja [Python Download](https://www.python.org/downloads/) para as instruções de instalação.


## Instalar pacotes
Para saber mais sobre instalações de pacotes no python, acesse [Installing Python Modules](https://docs.python.org/3/installing/index.html)
1. Baixe o Pandas:
```
pip install pandas
```
2. Baixe o Plotly:
```
pip install plotly
```
3. Baixe o Openpyxl
```
pip install openpyxl
```


---
# Execução do Script

> [!WARNING]  
> Os seguintes comandos devem ser executados com o arquivo Excel **FECHADO**

Siga o passo a passo para gerar os gráficos no Excel e HTML.

1. Inicialize o CMD na `pasta raiz` do projeto (local onde está localizado o arquivo `movielist-stats.py`)
2. Digite o seguinte comando:
```
py movielist-stats.py
```
3. Acesse o arquivo Excel para visualizar as tabelas, ou abra os arquivos HTML (gerados no próprio root) para visualizar os gráficos.


<br><br><br><br><br>

---
# Imagens do Projeto
## Rankings/Tabelas no Excel
![Stats Gerais](/readme-images/stats_gerais.png)
![Stats Gerais de Classificação Indicativa](/readme-images/stats_gerais_classind.png)
![Stats Gerais de Genero](/readme-images/stats_gerais_genero.png)
![Ranking de Distribuidora](/readme-images/mais_maior_distribuidora.png)
![Ranking de Franquia](/readme-images/mais_maior_franquia.png)

## Gráficos HTML
![Grafico de Ano](/readme-images/grafico_ano.png)
![Grafico de Franquia](/readme-images/grafico_franquia.png)
![Grafico de Genero](/readme-images/grafico_genero.png)
![Grafico de Distribuidora](/readme-images/grafico_Distribuidora.png)