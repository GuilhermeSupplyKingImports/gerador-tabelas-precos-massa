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
