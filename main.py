from pathlib import Path
from filtre_service import filter_ocorrencias
from ftp_service import down_ocorrencias
from email_service import generate_html, send_email
import schedule
import time
from datetime import datetime


def main():
    print(f"Executando em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    down_ocorrencias()

    df, anexo = filter_ocorrencias()

    html = generate_html(df)

    # checa se email foi enviado correto
    send_email(html, anexo)

    # limpando aquivos downloads
    downloads = Path.cwd() / "src" / "downloads"

    for arquivo in downloads.glob("*"):
        try:
            arquivo.unlink()
            print(f"Removido: {arquivo.name}")
        except Exception as e:
            print(f"Erro ao remover {arquivo.name}: {e}")


if __name__ == "__main__":
    schedule.every().monday.at("19:00").do(main)
    schedule.every().tuesday.at("19:00").do(main)
    schedule.every().wednesday.at("19:00").do(main)
    schedule.every().thursday.at("19:00").do(main)
    schedule.every().friday.at("19:00").do(main)

    print("Agendador iniciado. Aguardando execução às 19:00...")
    
    while True:
        schedule.run_pending()
        time.sleep(60)
