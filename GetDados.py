import struct
import time
import pandas as pd
from pymodbus.client.serial import ModbusSerialClient

# Função para obter entrada do usuário com um valor padrão
def get_input(prompt, default=None, type_cast=str):
    value = input(f"{prompt} (Padrão: {default}): ").strip()
    return type_cast(value) if value else default

# Conversão de 16 para 32 bits
def concat_16bits_to_float(reg1, reg2):
    int_32bit = (reg1 << 16) | reg2
    return struct.unpack('!f', struct.pack('!I', int_32bit))[0]

# Solicitar configurações ao usuário
serial_port = get_input("Digite a porta serial", "COM5")
baudrate = get_input("Digite o baudrate", 19200, int)
parity = get_input("Digite a paridade (N/E/O)", "E").upper()
stopbits = get_input("Digite o número de stop bits", 1, int)
bytesize = get_input("Digite o tamanho do byte", 8, int)
timeout = get_input("Digite o timeout (segundos)", 2, float)
slave_id = get_input("Digite o ID do escravo Modbus", 32, int)

# Criar o cliente Modbus configurando a porta serial
client = ModbusSerialClient(
    port=serial_port,
    baudrate=baudrate,
    parity=parity,
    stopbits=stopbits,
    bytesize=bytesize,
    timeout=timeout
)

print("\nTentando conectar...")
if client.connect():
    print("Conectado ao equipamento!")
    
    # Criar lista para armazenar os dados
    dados = []

    try:
        while True:  # Loop contínuo até interrupção manual
            response = client.read_holding_registers(address=0, count=29, slave=slave_id)

            if response.isError():
                print("Erro na leitura:", response)
            else:
                linha_dados = {"Timestamp": time.strftime("%Y-%m-%d %H:%M:%S")}
                
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
                    ("Alerta de umidade interna", 26, "Anormal" if response.registers[26] else "Normal"),
                    ("Alerta de aquecimento da cúpula", 28, "Anormal" if response.registers[28] else "Normal"),
                ]
                
                for nome, indice, unidade in registros:
                    if indice in [0, 26, 28]:  # Registros de 16 bits
                        reg1 = response.registers[indice]
                        linha_dados[nome] = f"{reg1} {unidade}"
                    else:  # Registros de 32 bits
                        reg1 = response.registers[indice]
                        reg2 = response.registers[indice + 1]
                        valor = concat_16bits_to_float(reg1, reg2)
                        if indice == 20:
                            linha_dados[nome] = f"{valor:.4f} {unidade}"
                        else:
                            linha_dados[nome] = f"{valor:.2f} {unidade}"

                dados.append(linha_dados)

                # Criar e salvar os dados no Excel
                df = pd.DataFrame(dados)
                df.to_excel("dados_modbus.xlsx", index=False)
                
                print(f"Dados salvos! Última leitura: {linha_dados}")

            time.sleep(5)  # Intervalo de 5 segundos entre leituras

    except KeyboardInterrupt:
        print("\nEncerrando programa...")
        client.close()
else:
    print("Falha na conexão!")
