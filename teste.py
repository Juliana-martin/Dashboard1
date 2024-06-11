import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import dash
from dash import html
from dash import dcc
from dash.dependencies import Input, Output
from dash_bootstrap_templates import ThemeSwitchAIO
import warnings

cor1 = '#006DC3' #azul escuro
cor2 = '#4287f5' #azul claro
cor3 = '#8c9900' #verde claro 
cor4 = '#56a8af' #verde


# Carregar os dados do arquivo Excel
tabela = pd.read_excel("./inventario.xlsx")

# Calcular a contagem de desktops
contagem_desktop = tabela["Tipo_Equipamento:"].value_counts()

# Criar o gráfico de barras para a contagem de desktops
fig1 = px.bar(x=contagem_desktop.index, y=contagem_desktop.values, title="Contagem de Desktops e Notebooks",
              color_discrete_sequence= [cor2])

# Definir o nome dos eixos
fig1.update_xaxes(title_text="Tipo de Equipamento")
fig1.update_yaxes(title_text="Quantidade")


#______________________________________________________________________

# Carregar os dados da planilha "PDA e Tablet"
nome_planilha = "PDA e Tablet"
dados_pda = pd.read_excel("./inventario.xlsx", sheet_name=nome_planilha)

# Calcular a contagem de PDA e tablets
contagem_pda = dados_pda["MODELO"].value_counts()

# Criar o gráfico de barras para a contagem de PDA e tablets
fig2 = px.bar(x=contagem_pda.index, y=contagem_pda.values, title='Contagem de PDA e Tablet',
              color_discrete_sequence= [cor2])

#______________________________________________________________________


nome_planilha_impressora = "Impressoras"
dados_impressora = pd.read_excel("./inventario.xlsx", sheet_name=nome_planilha_impressora)

# Calcular a contagem de impressoras
contagem_impressora = dados_impressora["MODELO"].value_counts()

# Criar o gráfico de barras para a contagem de impressoras
fig3 = px.bar(x=contagem_impressora.index, y=contagem_impressora.values, title='Contagem de Impressoras', 
              color_discrete_sequence= [cor2])


#__________________________________________________________

smart = "Smartphone"
phone = pd.read_excel("./inventario.xlsx", sheet_name=smart)
contagem_smart = phone["MARCA"].value_counts()

fig4 = px.bar(x=contagem_smart.index, y=contagem_smart.values, title="SmartPhone",
              color_discrete_sequence=[cor2])





app = dash.Dash(__name__, external_stylesheets=[ThemeSwitchAIO])
server = app.server

# Layout do aplicativo
app.layout = html.Div(children=[
    html.H1("Controle de Estoque", style={'textAlign': 'center'}),

    html.Div([
        dcc.Graph(figure=fig1),
        dcc.Graph(figure=fig2),
    ], style={'display': 'flex'}),

    html.Div([
        dcc.Graph(figure=fig3),
        dcc.Graph(figure=fig4),
    ], style={'display': 'flex'})
])

# Executar o aplicativo
if __name__ == '__main__':
    app.run_server(debug=True)