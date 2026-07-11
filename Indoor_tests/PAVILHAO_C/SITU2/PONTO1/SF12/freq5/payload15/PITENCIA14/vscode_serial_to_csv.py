import serial
import csv
import time
import re
from datetime import datetime
from openpyxl import Workbook

# =========================================================
# CONFIGURAÇÃO
# =========================================================
PORTA = "COM3"          # ajustar conforme necessário
BAUDRATE = 115200       # igual ao Serial.begin(...)
TIMEOUT = 1

timestamp_base = datetime.now().strftime("%Y%m%d_%H%M%S")
NOME_CSV = f"log_serial_parse_{timestamp_base}.csv"
NOME_XLSX = f"log_serial_parse_{timestamp_base}.xlsx"

# =========================================================
# FUNÇÃO DE PARSING
# =========================================================
def extrair_campos(linha):
    """
    Extrai da linha:
    - mensagem
    - RSSI
    - SNR
    - L
    - d_est

    Exemplo esperado:
    Recebido: ola | RSSI: -72.30 dBm | SNR: 8.50 dB | L: 85.30 dB | d_est: 123.45 m
    """

    resultado = {
        "mensagem": "",
        "RSSI_dBm": None,
        "SNR_dB": None,
        "L_dB": None,
        "d_est_m": None
    }

    # mensagem
    m_msg = re.search(r"Recebido:\s*(.*?)\s*(?=\|\s*RSSI:|$)", linha)
    if m_msg:
        resultado["mensagem"] = m_msg.group(1).strip()

    # RSSI
    m_rssi = re.search(r"RSSI:\s*([-+]?\d+(?:\.\d+)?)\s*dBm", linha)
    if m_rssi:
        resultado["RSSI_dBm"] = float(m_rssi.group(1))

    # SNR
    m_snr = re.search(r"SNR:\s*([-+]?\d+(?:\.\d+)?)\s*dB", linha)
    if m_snr:
        resultado["SNR_dB"] = float(m_snr.group(1))

    # L
    m_l = re.search(r"\bL:\s*([-+]?\d+(?:\.\d+)?)\s*dB", linha)
    if m_l:
        resultado["L_dB"] = float(m_l.group(1))

    # distância estimada
    m_d = re.search(r"d_est:\s*([-+]?\d+(?:\.\d+)?)\s*m", linha)
    if m_d:
        resultado["d_est_m"] = float(m_d.group(1))

    return resultado

# =========================================================
# FUNÇÃO PARA AJUSTAR LARGURA DAS COLUNAS XLSX
# =========================================================
def ajustar_largura_colunas(ws):
    larguras = {
        "A": 24,  # timestamp
        "B": 10,  # porta
        "C": 10,  # baudrate
        "D": 80,  # linha_raw
        "E": 30,  # mensagem
        "F": 12,  # RSSI
        "G": 12,  # SNR
        "H": 12,  # L
        "I": 12   # d_est
    }

    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura

# =========================================================
# FUNÇÃO PRINCIPAL
# =========================================================
def main():
    print("==============================================")
    print(" Leitura série + parsing + CSV + XLSX")
    print("==============================================")
    print(f"Porta     : {PORTA}")
    print(f"Baudrate  : {BAUDRATE}")
    print(f"CSV       : {NOME_CSV}")
    print(f"XLSX      : {NOME_XLSX}")
    print("Prima Ctrl+C para terminar.\n")

    # abrir porta série
    try:
        ser = serial.Serial(PORTA, BAUDRATE, timeout=TIMEOUT)
        time.sleep(2)  # estabilização/reset da board
        print(f"[OK] Ligação aberta em {PORTA}.\n")
    except serial.SerialException as e:
        print(f"[ERRO] Não foi possível abrir a porta série: {e}")
        return

    # preparar XLSX
    wb = Workbook()
    ws = wb.active
    ws.title = "log_serial"

    cabecalho = [
        "timestamp",
        "porta",
        "baudrate",
        "linha_raw",
        "mensagem",
        "RSSI_dBm",
        "SNR_dB",
        "L_dB",
        "d_est_m"
    ]
    ws.append(cabecalho)
    ajustar_largura_colunas(ws)

    try:
        with open(NOME_CSV, mode="w", newline="", encoding="utf-8") as f_csv:
            writer = csv.writer(f_csv)
            writer.writerow(cabecalho)

            contador_linhas = 0

            while True:
                if ser.in_waiting > 0:
                    linha = ser.readline().decode("utf-8", errors="ignore").strip()

                    if not linha:
                        continue

                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

                    # mostrar linha bruta
                    print(linha)

                    # extrair campos
                    dados = extrair_campos(linha)

                    linha_saida = [
                        timestamp,
                        PORTA,
                        BAUDRATE,
                        linha,
                        dados["mensagem"],
                        dados["RSSI_dBm"],
                        dados["SNR_dB"],
                        dados["L_dB"],
                        dados["d_est_m"]
                    ]

                    # escrever CSV
                    writer.writerow(linha_saida)
                    f_csv.flush()

                    # escrever XLSX
                    ws.append(linha_saida)
                    contador_linhas += 1

                    # guardar XLSX periodicamente para não perder dados
                    if contador_linhas % 10 == 0:
                        wb.save(NOME_XLSX)

    except KeyboardInterrupt:
        print("\n[INFO] Execução terminada pelo utilizador.")

    finally:
        ser.close()
        wb.save(NOME_XLSX)
        print("[INFO] Porta série fechada.")
        print(f"[INFO] CSV guardado em : {NOME_CSV}")
        print(f"[INFO] XLSX guardado em: {NOME_XLSX}")

# =========================================================
# EXECUÇÃO
# =========================================================
if __name__ == "__main__":
    main()3