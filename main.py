from pathlib import Path
from filtre_service import filter_ocorrencias
from ftp_service import down_ocorrencias
from email_service import generate_html, send_email


def main():
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
    main()
