import pandas as pd
import plotly.express as px
from plotly.subplots import make_subplots
import dash
from dash import dcc, html
from dash.dependencies import Input, Output

# Cores personalizadas
cor1 = '#006DC3'  # azul escuro
cor2 = '#4287f5'  # azul claro
cor3 = '#8c9900'  # verde claro
cor4 = '#56a8af'  # verde

# Função para carregar os dados do Excel
def carregar_dados_excel():
    tabela = pd.read_excel("inventario.xlsx")
    return tabela

# Função para criar os gráficos
def criar_graficos(tabela):
    # Calcular a contagem de desktops
    contagem_desktop = tabela["Tipo_Equipamento:"].value_counts()

    # Criar o gráfico de barras para a contagem de desktops
    fig1 = px.bar(x=contagem_desktop.index, y=contagem_desktop.values, title="Contagem de Desktops e Notebooks",
                  color_discrete_sequence=[cor1])

    # Calcular a contagem de PDA e tablets
    dados_pda = tabela["MODELO"].value_counts()

    # Criar o gráfico de barras para a contagem de PDA e tablets
    fig2 = px.bar(x=dados_pda.index, y=dados_pda.values, title='Contagem de PDA e Tablet',
                  color_discrete_sequence=[cor2])

    # Calcular a contagem de impressoras
    dados_impressora = tabela["MODELO_impressora"].value_counts()

    # Criar o gráfico de barras para a contagem de impressoras
    fig3 = px.bar(x=dados_impressora.index, y=dados_impressora.values, title='Contagem de Impressoras',
                  color_discrete_sequence=[cor3])

    # Calcular a contagem de smartphones
    dados_smartphone = tabela["MARCA_smartphone"].value_counts()

    # Criar o gráfico de barras para a contagem de smartphones
    fig4 = px.bar(x=dados_smartphone.index, y=dados_smartphone.values, title='Contagem de Smartphones',
                  color_discrete_sequence=[cor4])

    return fig1, fig2, fig3, fig4

# Inicializar o Dash app
app = dash.Dash(__name__)

# Layout da aplicação
app.layout = html.Div([
    html.H1('Controle de Estoque'),
    html.Div(id='graphs-container'),
    dcc.Interval(
        id='interval-component',
        interval=60*1000,  # Intervalo de atualização em milissegundos (60 segundos)
        n_intervals=0
    )
])

# Callback para atualizar os gráficos
@app.callback(Output('graphs-container', 'children'), [Input('interval-component', 'n_intervals')])
def update_graphs(n):
    # Carregar os dados do Excel
    tabela = carregar_dados_excel()
    
    # Criar os gráficos
    fig1, fig2, fig3, fig4 = criar_graficos(tabela)
    
    # Retornar os gráficos como componentes dcc.Graph
    return [
        dcc.Graph(figure=fig1),
        dcc.Graph(figure=fig2),
        dcc.Graph(figure=fig3),
        dcc.Graph(figure=fig4)
    ]

if __name__ == '__main__':
    app.run_server(debug=True)
