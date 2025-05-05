from pymodbus.client import ModbusSerialClient
import struct
import math
import time
from datetime import datetime

CONFIG = {
    'port': 'COM5',
    'baudrate': 9600,
    'parity': 'N',
    'stopbits': 1,
    'bytesize': 8,
    'timeout': 3,
    'slave_id': 1,
    'start_address': 0,
    'register_count': 6
}


def conectar_modbus():
    print(f"Conectando ao CR1000 na porta {CONFIG['port']}...")
    try:
        client = ModbusSerialClient(
            port=CONFIG['port'],
            baudrate=CONFIG['baudrate'],
            parity=CONFIG['parity'],
            stopbits=CONFIG['stopbits'],
            bytesize=CONFIG['bytesize'],
            timeout=CONFIG['timeout']
        )
        if client.connect():
            print("✅ Conexão Modbus estabelecida com sucesso!")
            return client
        print("❌ Falha ao conectar")
        return None
    except Exception as e:
        print(f"❌ Erro na conexão: {str(e)}")
        return None


def ler_registros(client):
    """Função segura para leitura de registros"""
    try:
        response = client.read_holding_registers(
            address=CONFIG['start_address'],
            count=CONFIG['register_count'],
            slave=CONFIG['slave_id']
        )
        if response.isError():
            print(f"❌ Erro Modbus: {response}")
            return None
        return response.registers
    except Exception as e:
        print(f"❌ Erro na leitura: {str(e)}")
        return None


def interpretar_valores(registros):
    """Interpreta os registros considerando NaN e valores válidos"""
    if not registros or len(registros) != 6:
        print(f"❌ Dados incompletos. Recebidos: {len(registros) if registros else 0}/6 registros")
        return None

    # Converte para pares hexadecimais
    hex_values = [f"{reg:04X}" for reg in registros]
    print(f"Dados brutos (hex): {hex_values}")

    # Verifica padrões de NaN (FFC0 0000 é um padrão comum para NaN)
    valores = []
    for i in range(0, 6, 2):
        hex_pair = f"{hex_values[i]}{hex_values[i + 1]}"

        if hex_pair == "FFC00000":  # Padrão IEEE 754 para NaN
            valores.append(float('nan'))
        else:
            try:
                # Converte big endian (modbus padrão)
                valor = struct.unpack('>f', bytes.fromhex(hex_pair))[0]
                valores.append(valor)
            except:
                valores.append(float('nan'))

    return valores


def main():
    client = conectar_modbus()
    if not client:
        return

    try:
        while True:
            print("\n" + "=" * 50)
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

            registros = ler_registros(client)
            if registros:
                valores = interpretar_valores(registros)
                if valores:
                    print("\nValores convertidos:")
                    for i, val in enumerate(valores, 1):
                        if math.isnan(val):
                            print(f"Entrada {i}: NaN (valor inválido)")
                        else:
                            print(f"Entrada {i}: {val:.10f}")  # 10 casas decimais

            time.sleep(1)

    except KeyboardInterrupt:
        print("\nLeitura encerrada pelo usuário")
    finally:
        client.close()
        print("Conexão encerrada")


if __name__ == "__main__":
    main()