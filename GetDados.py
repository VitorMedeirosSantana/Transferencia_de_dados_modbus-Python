import struct
import time
import pandas as pd
from pymodbus.client.serial import ModbusSerialClient
import dash
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import serial.tools.list_ports
from dash.exceptions import PreventUpdate

# Inicialização do app Dash
app = dash.Dash(__name__, suppress_callback_exceptions=True,
                external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

# Variáveis globais
client = None
last_update_time = 0
update_interval = 5  # segundos


# Função para listar portas seriais
def get_available_ports():
    return [port.device for port in serial.tools.list_ports.comports()]


# Função de conversão de registros para float
def concat_16bits_to_float(reg1, reg2):
    int_32bit = (reg1 << 16) | reg2
    return struct.unpack('!f', struct.pack('!I', int_32bit))[0]


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

        # Cabeçalho
        dbc.Row(
            [
                dbc.Col(
                    html.Img(
                        src="assets/Marca_Branca.png",
                        style={
                            "height": "8rem",
                            "max-width": "100%",
                            "object-fit": "contain",
                        }
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

        # Seção de configuração
        dbc.Card(
            [
                dbc.CardHeader("Configuração da Comunicação", className="bg-primary text-white"),
                dbc.CardBody(
                    [
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Label("Porta Serial"),
                                        dcc.Dropdown(
                                            id='serial-port-dropdown',
                                            options=[{'label': port, 'value': port} for port in get_available_ports()],
                                            value=get_available_ports()[0] if get_available_ports() else None,
                                            placeholder="Selecione a porta serial"
                                        ),
                                    ],
                                    md=6
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("Baudrate"),
                                        dbc.Input(
                                            id='baudrate-input',
                                            type='number',
                                            value=19200,
                                            min=9600,
                                            step=9600
                                        ),
                                    ],
                                    md=6
                                ),
                            ],
                            className="mb-3"
                        ),
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Label("Paridade"),
                                        dbc.Select(
                                            id='parity-select',
                                            options=[
                                                {'label': 'Nenhuma', 'value': 'N'},
                                                {'label': 'Par', 'value': 'E'},
                                                {'label': 'Ímpar', 'value': 'O'}
                                            ],
                                            value='E'
                                        ),
                                    ],
                                    md=4
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("Stop Bits"),
                                        dbc.Input(
                                            id='stopbits-input',
                                            type='number',
                                            value=1,
                                            min=1,
                                            max=2,
                                            step=1
                                        ),
                                    ],
                                    md=4
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("Tamanho do Byte"),
                                        dbc.Input(
                                            id='bytesize-input',
                                            type='number',
                                            value=8,
                                            min=5,
                                            max=8,
                                            step=1
                                        ),
                                    ],
                                    md=4
                                ),
                            ],
                            className="mb-3"
                        ),
                        dbc.Row(
                            [
                                dbc.Col(
                                    [
                                        dbc.Label("Timeout (segundos)"),
                                        dbc.Input(
                                            id='timeout-input',
                                            type='number',
                                            value=2.0,
                                            min=0.1,
                                            step=0.1
                                        ),
                                    ],
                                    md=6
                                ),
                                dbc.Col(
                                    [
                                        dbc.Label("ID do Escravo Modbus"),
                                        dbc.Input(
                                            id='slave-id-input',
                                            type='number',
                                            value=32,
                                            min=1,
                                            max=247
                                        ),
                                    ],
                                    md=6
                                ),
                            ]
                        ),
                        dbc.Row(
                            [
                                dbc.Col(
                                    dbc.Button("Conectar", id='connect-btn', color="success", className="mt-3"),
                                    width=6
                                ),
                                dbc.Col(
                                    dbc.Button("Desconectar", id='disconnect-btn', color="danger", className="mt-3",
                                               disabled=True),
                                    width=6
                                )
                            ],
                            className="mb-3"
                        ),
                        dbc.Alert(id='connection-status', className="mt-3", is_open=False)
                    ]
                )
            ],
            className="mb-4"
        ),

        # Seção de dados
        dbc.Card(
            [
                dbc.CardHeader("Dados do Piranômetro", className="bg-primary text-white"),
                dbc.CardBody(
                    [
                        dash_table.DataTable(
                            id='data-table',
                            columns=[
                                {"name": "Parâmetro", "id": "parameter"},
                                {"name": "Valor", "id": "value"},
                                {"name": "Unidade", "id": "unit"}
                            ],
                            style_table={'overflowX': 'auto'},
                            style_cell={
                                'textAlign': 'left',
                                'padding': '10px',
                                'whiteSpace': 'normal',
                                'height': 'auto',
                                'color': 'black'
                            },
                            style_header={
                                'backgroundColor': 'rgb(230, 230, 230)',
                                'fontWeight': 'bold'
                            },
                        ),
                        dbc.Row(
                            dbc.Col(
                                dbc.Button("Exportar para Excel", id='export-btn', color="info", className="mt-3"),
                                width=12
                            )
                        ),
                        dcc.Download(id="download-excel")
                    ]
                )
            ],
            id='data-card',
            style={'display': 'none'}
        )
    ]
)


# Callbacks
@app.callback(
    [Output('connection-status', 'children'),
     Output('connection-status', 'color'),
     Output('connection-status', 'is_open'),
     Output('connect-btn', 'disabled'),
     Output('disconnect-btn', 'disabled'),
     Output('data-card', 'style'),
     Output('connection-store', 'data'),
     Output('update-interval', 'disabled')],
    [Input('connect-btn', 'n_clicks'),
     Input('disconnect-btn', 'n_clicks')],
    [State('serial-port-dropdown', 'value'),
     State('baudrate-input', 'value'),
     State('parity-select', 'value'),
     State('stopbits-input', 'value'),
     State('bytesize-input', 'value'),
     State('timeout-input', 'value'),
     State('slave-id-input', 'value'),
     State('connection-store', 'data')],
    prevent_initial_call=True
)
def manage_connection(connect_clicks, disconnect_clicks, port, baudrate, parity, stopbits, bytesize, timeout, slave_id,
                      connection_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        raise PreventUpdate

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    global client

    if button_id == 'connect-btn':
        try:
            client = ModbusSerialClient(
                port=port,
                baudrate=baudrate,
                parity=parity,
                stopbits=stopbits,
                bytesize=bytesize,
                timeout=timeout
            )

            if client.connect():
                # Teste de conexão lendo um registro
                response = client.read_holding_registers(address=0, count=2, slave=slave_id)
                if response.isError():
                    raise Exception("Falha na leitura inicial")

                return (
                    "Conectado com sucesso!", "success", True,
                    True, False,
                    {'display': 'block'},
                    {'port': port, 'slave_id': slave_id},
                    False
                )
            else:
                raise Exception("Falha na conexão física")

        except Exception as e:
            return (
                f"Erro na conexão: {str(e)}", "danger", True,
                False, True,
                {'display': 'none'},
                None,
                True
            )

    elif button_id == 'disconnect-btn':
        if client:
            client.close()
            client = None
        return (
            "Desconectado com sucesso", "warning", True,
            False, True,
            {'display': 'none'},
            None,
            True
        )


@app.callback(
    [Output('data-table', 'data'),
     Output('download-excel', 'data')],
    [Input('update-interval', 'n_intervals'),
     Input('export-btn', 'n_clicks')],
    [State('connection-store', 'data')],
    prevent_initial_call=True
)
def update_data(n_intervals, export_clicks, connection_data):
    ctx = dash.callback_context
    if not ctx.triggered or not connection_data or not client:
        raise PreventUpdate

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    data = []

    try:
        response = client.read_holding_registers(address=0, count=29, slave=connection_data['slave_id'])

        if response.isError():
            raise Exception("Erro na leitura dos registros")

        registros = [
            ("Modelo do transmissor", 0, ""),
            ("Irradiância solar ajustada", 2, "W/m²"),
            ("Temperatura do sensor Pt100", 8, "°C"),
            ("Ângulo de inclinação no eixo X", 14, "°"),
            ("Ângulo de inclinação no eixo Y", 16, "°"),
            ("Irradiância solar bruta", 18, "W/m²"),
            ("Tensão de saída do sensor", 20, "mV"),
            ("Temperatura interna", 22, "°C"),
            ("Umidade interna", 24, "% RH"),
            ("Alerta de umidade interna", 26, ""),
            ("Alerta de aquecimento da cúpula", 28, ""),
        ]

        for nome, indice, unidade in registros:
            if indice in [0, 26, 28]:  # Registros de 16 bits
                reg1 = response.registers[indice]
                if indice == 0:
                    value = str(reg1)
                else:
                    value = "Anormal" if reg1 else "Normal"
            else:  # Registros de 32 bits
                reg1 = response.registers[indice]
                reg2 = response.registers[indice + 1]
                valor = concat_16bits_to_float(reg1, reg2)
                if indice == 20:
                    value = f"{valor:.4f}"
                else:
                    value = f"{valor:.2f}"

            data.append({
                "parameter": nome,
                "value": value,
                "unit": unidade
            })

        if button_id == 'export-btn':
            df = pd.DataFrame([{
                "Parâmetro": item["parameter"],
                "Valor": item["value"],
                "Unidade": item["unit"],
                "Timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
            } for item in data])

            return data, dcc.send_data_frame(df.to_excel, "dados_piranometro.xlsx", index=False)

        return data, None

    except Exception as e:
        print(f"Erro: {str(e)}")
        raise PreventUpdate


if __name__ == '__main__':
    app.run_server(debug=True)