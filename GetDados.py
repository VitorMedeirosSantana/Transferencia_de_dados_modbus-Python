import struct
from datetime import datetime
import pandas as pd
from pymodbus.client.serial import ModbusSerialClient
import dash
from dash import dcc, html, Input, Output, State, dash_table, no_update
import dash_bootstrap_components as dbc
import serial.tools.list_ports
from dash.exceptions import PreventUpdate
import math
import io
import time
import xlsxwriter

# Inicialização do app Dash
app = dash.Dash(__name__, suppress_callback_exceptions=True,
                external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

# Variáveis globais
piranometer_client = None
cr1000_client = None
last_update_time = 0
update_interval = 5  # segundos

historical_data = pd.DataFrame(columns=[
    'timestamp', 'irradiance', 'voltage_out', 'angle_x', 'angle_y',
    'cr1000_1', 'cr1000_2', 'cr1000_3'
])


# Função para listar portas seriais
def get_available_ports():
    return [port.device for port in serial.tools.list_ports.comports()]


# Função de conversão de registros para float
def concat_16bits_to_float(reg1, reg2):
    int_32bit = (reg1 << 16) | reg2
    return struct.unpack('!f', struct.pack('!I', int_32bit))[0]


# Função para interpretar valores do CR1000 via Modbus
def interpret_cr1000_values(registers, active_ports):
    """Interpreta os registros considerando apenas as portas ativas"""
    if not registers or len(registers) != 6:
        print(f"Dados incompletos do CR1000. Recebidos: {len(registers) if registers else 0}/6 registros")
        return [float('nan')] * 3

    # Converte para pares hexadecimais
    hex_values = [f"{reg:04X}" for reg in registers]

    values = []
    for i in range(0, 6, 2):
        channel_num = (i // 2) + 1
        if channel_num in active_ports:
            hex_pair = f"{hex_values[i]}{hex_values[i + 1]}"
            try:
                value = struct.unpack('>f', bytes.fromhex(hex_pair))[0]
                values.append(value)
            except:
                values.append(float('nan'))
        else:
            values.append(float('nan'))

    return values[:3]  # Retorna sempre 3 valores (NaN para portas inativas)


# Layout do aplicativo
app.layout = dbc.Container(
    fluid=True,
    style={
        "display": "flex",
        "flex-direction": "column",
        "height": "100vh",
        "padding": "20px",
        "background": "#006d68",
        "color": "white"
    },
    children=[
        dcc.Location(id='url', refresh=False),
        dcc.Interval(id='update-interval', interval=update_interval * 1000, disabled=True),
        dcc.Store(id='connection-store'),
        dcc.Store(id='data-store'),

        # Cabeçalho
        dbc.Row(
            [
                dbc.Col(
                    html.Img(
                        src="assets/Marca_Branca.png",
                        style={"height": "8rem", "max-width": "100%", "object-fit": "contain"}
                    ),
                    width="auto",
                ),
                dbc.Col(
                    html.H1(
                        children='Calibração de Piranômetros',
                        style={
                            'color': '#FFFFFF',
                            'font-size': 'clamp(2rem, 6vw, 4rem)',
                            'margin': '0',
                            'text-align': 'center',
                            'font-weight': 'bold'
                        }
                    ),
                    width="auto",
                )
            ],
            justify="center",
            align="center",
            className="w-100 mb-4",
        ),

        # Linha 1: Configuração Modbus - Piranômetro
        dbc.Row(
            dbc.Card(
                [
                    dbc.CardHeader("Configuração Modbus - Piranômetro", className="bg-primary text-white"),
                    dbc.CardBody(
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Label("Porta Serial"),
                                        dcc.Dropdown(
                                            id='piranometer-port-dropdown',
                                            options=[{'label': port, 'value': port} for port in get_available_ports()],
                                            value=get_available_ports()[0] if get_available_ports() else None,
                                        ),
                                    ],
                                    md=2
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("Baudrate"),
                                        dbc.Input(id='piranometer-baudrate', type='number', value=19200),
                                    ],
                                    md=2
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("Paridade"),
                                        dbc.Select(
                                            id='piranometer-parity',
                                            options=[
                                                {'label': 'Nenhuma', 'value': 'N'},
                                                {'label': 'Par', 'value': 'E'},
                                                {'label': 'Ímpar', 'value': 'O'}
                                            ],
                                            value='E'
                                        ),
                                    ],
                                    md=2
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("ID do Escravo"),
                                        dbc.Input(id='piranometer-slave-id', type='number', value=32),
                                    ],
                                    md=2
                                ),
                            ],
                            className="g-2",
                        )
                    ),
                ],
                className="mb-3"
            )
        ),

        # Linha 2: Configuração Modbus - CR1000
        dbc.Row(
            dbc.Card(
                [
                    dbc.CardHeader("Configuração Modbus - CR1000", className="bg-primary text-white"),
                    dbc.CardBody(
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Label("Porta Serial"),
                                        dcc.Dropdown(
                                            id='cr1000-port-dropdown',
                                            options=[{'label': port, 'value': port} for port in get_available_ports()],
                                            value=None,
                                        ),
                                    ],
                                    md=2
                                ),
                                # Adicione este componente na seção de Configuração Modbus - CR1000 (dentro do CardBody)
                                dbc.Col(
                                    [
                                        dbc.Label("Portas Ativas"),
                                        dbc.Checklist(
                                            id='active-ports',
                                            options=[
                                                {"label": "Canal 1", "value": 1},
                                                {"label": "Canal 2", "value": 2},
                                                {"label": "Canal 3", "value": 3},
                                            ],
                                            value=[1],  # Canal 1 ativo por padrão
                                            inline=True,
                                        ),
                                    ],
                                    md=2
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("Baudrate"),
                                        dbc.Input(id='cr1000-baudrate', type='number', value=9600),
                                    ],
                                    md=2
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("Paridade"),
                                        dbc.Select(
                                            id='cr1000-parity',
                                            options=[
                                                {'label': 'Nenhuma', 'value': 'N'},
                                                {'label': 'Par', 'value': 'E'},
                                                {'label': 'Ímpar', 'value': 'O'}
                                            ],
                                            value='N'
                                        ),
                                    ],
                                    md=2
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("ID do Escravo"),
                                        dbc.Input(id='cr1000-slave-id', type='number', value=1),
                                    ],
                                    md=2
                                ),
                            ],
                            className="g-2",
                        )
                    ),
                ],
                className="mb-3"
            )
        ),

        # Linha 3: Botões de conexão
        dbc.Row(
            [
                dbc.Col(
                    dbc.Button("Conectar Piranômetro", id='connect-piranometer-btn', color="success", className="me-2"),
                    width="auto"),
                dbc.Col(dbc.Button("Conectar CR1000", id='connect-cr1000-btn', color="success", className="me-2"),
                        width="auto"),
                dbc.Col(dbc.Button("Desconectar Tudo", id='disconnect-all-btn', color="danger", disabled=True),
                        width="auto"),
            ],
            className="mb-4 justify-content-center",
        ),

        # Linha 4: Cards de dados
        dbc.Row(
            id='data-display-row',
            style={'display': 'none'},
            children=[
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Gráficos", className="bg-primary text-white"),
                            dbc.CardBody(
                                dcc.Graph(
                                    id='data-graph',
                                    config={'displayModeBar': True},
                                    style={'height': '400px'}
                                )
                            )
                        ],
                        style={"height": "100%"}
                    ),
                    md=8
                ),
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Dados do Piranômetro", className="bg-primary text-white"),
                            dbc.CardBody(
                                dash_table.DataTable(
                                    id='piranometer-table',
                                    columns=[
                                        {"name": "Parâmetro", "id": "parameter"},
                                        {"name": "Valor", "id": "value"},
                                        {"name": "Unidade", "id": "unit"}
                                    ],
                                    style_table={'height': '300px', 'overflowY': 'auto'},
                                    style_cell={'textAlign': 'left', 'padding': '8px', 'color': 'black'},
                                )
                            )
                        ],
                        style={"height": "100%"}
                    ),
                    md=4
                )
            ],
            className="g-3"
        ),

        # Linha 5: Dados CR1000
        dbc.Row(
            id='cr1000-data-row',
            style={'display': 'none'},
            children=[
                dbc.Col(
                    dbc.Card(
                        [
                            dbc.CardHeader("Dados CR1000 - Canais", className="bg-primary text-white"),
                            dbc.CardBody(
                                [
                                    dash_table.DataTable(
                                        id='cr1000-table',
                                        columns=[
                                            {"name": "Canal", "id": "channel"},
                                            {"name": "Valor", "id": "value"},
                                            {"name": "Timestamp", "id": "timestamp"}
                                        ],
                                        style_table={'height': '200px', 'overflowY': 'auto'},
                                        style_cell={'textAlign': 'left', 'padding': '8px', 'color': 'black'},
                                    ),
                                    dbc.Button("Exportar Dados", id='export-btn', color="info", className="mt-3"),
                                    dcc.Download(id="download-excel")
                                ]
                            )
                        ]
                    )
                )
            ],
            className="mt-3"
        ),

        dbc.Alert(id='connection-status', className="mt-3", is_open=False)
    ]
)


# Callback para gerenciar conexões
@app.callback(
    [Output('connection-status', 'children'),
     Output('connection-status', 'color'),
     Output('connection-status', 'is_open'),
     Output('connect-piranometer-btn', 'disabled'),
     Output('connect-cr1000-btn', 'disabled'),
     Output('disconnect-all-btn', 'disabled'),
     Output('data-display-row', 'style'),
     Output('cr1000-data-row', 'style'),
     Output('connection-store', 'data'),
     Output('update-interval', 'disabled')],
    [Input('connect-piranometer-btn', 'n_clicks'),
     Input('connect-cr1000-btn', 'n_clicks'),
     Input('disconnect-all-btn', 'n_clicks')],
    [State('piranometer-port-dropdown', 'value'),
     State('piranometer-baudrate', 'value'),
     State('piranometer-parity', 'value'),
     State('piranometer-slave-id', 'value'),
     State('cr1000-port-dropdown', 'value'),
     State('cr1000-baudrate', 'value'),
     State('cr1000-parity', 'value'),
     State('cr1000-slave-id', 'value')],
    prevent_initial_call=True
)
def manage_connections(piranometer_clicks, cr1000_clicks, disconnect_clicks,
                       piranometer_port, piranometer_baud, piranometer_parity, piranometer_slave,
                       cr1000_port, cr1000_baud, cr1000_parity, cr1000_slave):
    ctx = dash.callback_context
    if not ctx.triggered:
        raise PreventUpdate

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    global piranometer_client, cr1000_client

    if button_id == 'connect-piranometer-btn':
        try:
            piranometer_client = ModbusSerialClient(
                port=piranometer_port,
                baudrate=piranometer_baud,
                parity=piranometer_parity,
                stopbits=1,
                bytesize=8,
                timeout=2.0
            )

            if piranometer_client.connect():
                return (
                    "Piranômetro conectado com sucesso!", "success", True,
                    True, False, False,
                    {'display': 'flex'}, {'display': 'flex'},
                    {'piranometer_connected': True, 'piranometer_port': piranometer_port,
                     'piranometer_slave': piranometer_slave,
                     'cr1000_connected': cr1000_client is not None, 'cr1000_slave': cr1000_slave},
                    False
                )
            else:
                raise Exception("Falha na conexão física com piranômetro")

        except Exception as e:
            return (
                f"Erro na conexão Piranômetro: {str(e)}", "danger", True,
                False, False, False,
                {'display': 'none'}, {'display': 'none'},
                None, True
            )

    elif button_id == 'connect-cr1000-btn':
        try:
            cr1000_client = ModbusSerialClient(
                port=cr1000_port,
                baudrate=cr1000_baud,
                parity=cr1000_parity,
                stopbits=1,
                bytesize=8,
                timeout=3.0
            )

            if cr1000_client.connect():
                # Testar comunicação lendo registros
                response = cr1000_client.read_holding_registers(
                    address=0,
                    count=6,
                    slave=cr1000_slave
                )

                if not response.isError():
                    return (
                        "CR1000 conectado com sucesso!", "success", True,
                        False, True, False,
                        {'display': 'flex'}, {'display': 'flex'},
                        {'piranometer_connected': piranometer_client is not None, 'piranometer_port': piranometer_port,
                         'piranometer_slave': piranometer_slave,
                         'cr1000_connected': True, 'cr1000_slave': cr1000_slave},
                        False
                    )
                else:
                    cr1000_client.close()
                    cr1000_client = None
                    raise Exception("Falha na comunicação com CR1000 - Resposta inválida")
            else:
                raise Exception("Falha na conexão física com CR1000")

        except Exception as e:
            return (
                f"Erro na conexão CR1000: {str(e)}", "danger", True,
                False, False, False,
                {'display': 'none'}, {'display': 'none'},
                None, True
            )

    elif button_id == 'disconnect-all-btn':
        if piranometer_client:
            piranometer_client.close()
            piranometer_client = None
        if cr1000_client:
            cr1000_client.close()
            cr1000_client = None
        return (
            "Todos dispositivos desconectados", "warning", True,
            False, False, True,
            {'display': 'none'}, {'display': 'none'},
            None, True
        )


# Callback para atualizar dados
@app.callback(
    [Output('piranometer-table', 'data'),
     Output('cr1000-table', 'data'),
     Output('data-graph', 'figure'),
     Output('data-store', 'data')],
    [Input('update-interval', 'n_intervals'),
     Input('active-ports', 'value')],
    [State('connection-store', 'data')],

    prevent_initial_call=True
)
def update_data(n_intervals, active_ports, connection_data):
    global historical_data



    if not connection_data:
        print("Sem dados de conexão!")
        raise PreventUpdate

    # 1. Obter dados do piranômetro
    piranometer_data = {}
    if piranometer_client and connection_data.get('piranometer_connected'):
        try:
            response = piranometer_client.read_holding_registers(
                address=0, count=29, slave=connection_data['piranometer_slave']
            )

            if response.isError():
                print("Erro na resposta do piranômetro!")
            else:
                print(f"Dados piranômetro: {response.registers[:4]}...")  # Debug
                piranometer_data = {
                    'irradiance': concat_16bits_to_float(response.registers[2], response.registers[3]),
                    'voltage_out': concat_16bits_to_float(response.registers[20], response.registers[21]),
                    'angle_x': concat_16bits_to_float(response.registers[14], response.registers[15]),
                    'angle_y': concat_16bits_to_float(response.registers[16], response.registers[17])
                }
        except Exception as e:
            print(f"Erro no piranômetro: {str(e)}")

    # 2. Obter dados do CR1000
    cr1000_values = [float('nan')] * 3
    if cr1000_client and connection_data.get('cr1000_connected'):
        try:
            response = cr1000_client.read_holding_registers(
                address=0, count=6, slave=connection_data['cr1000_slave']
            )

            if response.isError():
                print("Erro na resposta do CR1000!")
            else:
                print(f"Dados CR1000: {response.registers}")  # Debug
                cr1000_values = interpret_cr1000_values(response.registers, active_ports)
        except Exception as e:
            print(f"Erro no CR1000: {str(e)}")

    # 3. Preparar dados para tabelas
    piranometer_table = [
        {"parameter": "Irradiância solar", "value": f"{piranometer_data.get('irradiance', float('nan')):.2f}",
         "unit": "W/m²"},
        {"parameter": "Tensão de saída", "value": f"{piranometer_data.get('voltage_out', float('nan')):.4f}",
         "unit": "mV"},
        {"parameter": "Inclinação X", "value": f"{piranometer_data.get('angle_x', float('nan')):.2f}", "unit": "°"},
        {"parameter": "Inclinação Y", "value": f"{piranometer_data.get('angle_y', float('nan')):.2f}", "unit": "°"}
    ]

    cr1000_table = []
    for i in active_ports:
        val = cr1000_values[i - 1]
        cr1000_table.append({
            'channel': f'Canal {i}',
            'value': f"{val:.4f}" if not math.isnan(val) else "NaN",
            'timestamp': datetime.now().strftime('%H:%M:%S')
        })

    # 4. Atualizar histórico
    new_row = {
        'timestamp': datetime.now(),
        **piranometer_data,
        'cr1000_1': cr1000_values[0],
        'cr1000_2': cr1000_values[1],
        'cr1000_3': cr1000_values[2]
    }
    historical_data = pd.concat([historical_data, pd.DataFrame([new_row])], ignore_index=True)

    # 5. Criar gráfico
    fig = create_figure(historical_data, active_ports)

    return piranometer_table, cr1000_table, fig, new_row


def create_figure(data, active_ports):
    fig_data = [{
        'x': data['timestamp'],
        'y': data['irradiance'],
        'type': 'line',
        'name': 'Irradiância (W/m²)',
        'yaxis': 'y1'
    }]

    colors = ['red', 'green', 'purple']
    for i in active_ports:
        fig_data.append({
            'x': data['timestamp'],
            'y': data[f'cr1000_{i}'],
            'type': 'line',
            'name': f'Canal {i}',
            'yaxis': 'y2',
            'line': {'color': colors[i - 1]}
        })

    return {
        'data': fig_data,
        'layout': {
            'title': 'Dados em Tempo Real',
            'xaxis': {'title': 'Tempo'},
            'yaxis': {'title': 'Irradiância (W/m²)', 'side': 'left'},
            'yaxis2': {'title': 'Valores CR1000', 'side': 'right', 'overlaying': 'y'}
        }
    }


@app.callback(
    Output("download-excel", "data"),
    Input("export-btn", "n_clicks"),
    prevent_initial_call=True
)
def export_to_excel(n_clicks):
    if not n_clicks:
        raise PreventUpdate

    try:
        # Usar a variável global diretamente
        global historical_data
        df = historical_data.copy()

        # Converter timestamp se necessário
        if 'timestamp' in df.columns:
            df['timestamp'] = pd.to_datetime(df['timestamp']).dt.strftime('%Y-%m-%d %H:%M:%S')

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return dcc.send_bytes(output.getvalue(), filename="dados_piranometro.xlsx")

    except Exception as e:
        print(f"Erro na exportação: {str(e)}")
        return no_update

if __name__ == '__main__':
    app.run_server(debug=True)