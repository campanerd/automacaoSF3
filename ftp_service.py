import paramiko
import os
from dotenv import load_dotenv
import datetime
load_dotenv()


from pathlib import Path
from dotenv import load_dotenv

def down_ocorrencias():
    
    BASE_DIR = Path(__file__).resolve().parent
    load_dotenv(BASE_DIR / ".env")

    DOWNLOADS_DIR = BASE_DIR / "src" / "downloads"
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)



    # Dados do servidor
    SFTP_HOST = os.getenv("FTP_HOST")
    SFTP_USER = os.getenv("FTP_USER")
    SFTP_PASS = os.getenv("FTP_PASS")
    SFTP_PORT = int(os.getenv("FTP_PORT", 22))

    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março",
        4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro",
        10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    data_ref = datetime.date.today()

    mes = meses_pt[data_ref.month]
    ano = data_ref.year

    # caminho
    ARQUIVO_REMOTO = f"/arquivos/OCORRENCIAS SAC/Ocorrências - {mes} {ano}.xls"
    ARQUIVO_LOCAL = DOWNLOADS_DIR / "ocorrenciasSF3.xls"

    print("Arquivo remoto:", ARQUIVO_REMOTO)

    print("HOST:", SFTP_HOST)
    print("PORT:", SFTP_PORT, type(SFTP_PORT))
    print("USER:", SFTP_USER)
    print("PASS:", "OK" if SFTP_PASS else None)


    try:
        # Conexão SSH
        transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
        transport.connect(username=SFTP_USER, password=SFTP_PASS)

        sftp = paramiko.SFTPClient.from_transport(transport)

        # Download
        sftp.get(ARQUIVO_REMOTO, str(ARQUIVO_LOCAL))


        sftp.close()
        transport.close()

        print("Arquivo baixado com sucesso via SFTP!")

    except Exception as e:
        print("Erro ao baixar arquivo via SFTP")
        print(e)
