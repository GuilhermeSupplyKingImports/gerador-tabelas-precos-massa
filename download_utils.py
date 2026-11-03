"""Criação e validação de pacotes ZIP para download no Protheus."""

from io import BytesIO
from zipfile import ZipFile, BadZipFile, ZIP_DEFLATED


def build_zip_checked(files: dict[str, bytes]) -> bytes:
    """Gera ZIP e verifica se todos os arquivos são recuperáveis sem diferenças."""
    if not files:
        raise ValueError("Nenhum arquivo foi selecionado para o pacote ZIP.")
    bio = BytesIO()
    with ZipFile(bio, mode="w", compression=ZIP_DEFLATED, allowZip64=True) as archive:
        for name, content in files.items():
            if not isinstance(content, bytes):
                raise TypeError(f"O arquivo {name} não foi gerado como bytes.")
            archive.writestr(name, content)
    payload = bio.getvalue()
    try:
        with ZipFile(BytesIO(payload), "r") as archive:
            if archive.testzip() is not None:
                raise ValueError("Falha de integridade no ZIP.")
            if set(archive.namelist()) != set(files):
                raise ValueError("O ZIP não contém todos os arquivos selecionados.")
            for name, expected in files.items():
                if archive.read(name) != expected:
                    raise ValueError(f"O arquivo {name} está diferente no pacote ZIP.")
    except (BadZipFile, OSError) as exc:
        raise ValueError("O pacote ZIP gerado é inválido.") from exc
    return payload



def browser_download_link(content: bytes, filename: str, mime: str = "application/zip") -> str:
    """Link de download direto, sem utilizar a rota temporária /media do Streamlit."""
    from base64 import b64encode
    from html import escape
    if not isinstance(content, bytes) or not content:
        raise ValueError("Arquivo de download vazio ou inválido.")
    if not filename or "/" in filename or "\\" in filename:
        raise ValueError("Nome de download inválido.")
    # Use apenas para pacotes moderados: data URIs duplicam parte do payload
    # no HTML/DOM do navegador e não são adequadas para arquivos enormes.
    uri = b64encode(content).decode("ascii")
    safe_name = escape(filename, quote=True)
    safe_mime = escape(mime, quote=True)
    return (
        '<a download="' + safe_name + '" href="data:' + safe_mime + ';base64,' + uri + '" '
        'style="display:inline-block;background:#ff4b4b;color:white;'
        'padding:0.65rem 1.1rem;border-radius:0.5rem;text-decoration:none;'
        'font-weight:600">⬇ Baixar ZIP diretamente no navegador</a>'
    )
