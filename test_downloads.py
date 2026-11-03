"""Regressão: CSV de fontes existentes e integridade dos pacotes ZIP."""
from datetime import date
from io import BytesIO
from zipfile import ZipFile

from openpyxl import Workbook

from common_pricing_flexible import processar_precificacao_comum_flexivel
from download_utils import build_zip_checked
from processor import CommonPricingConfig, build_common_protheus_csv, build_common_report_excel


def test_exporta_csv_001_e_007_mesmo_quando_sao_fontes():
    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "Custo Médio", "Preço 001", "Preço 007"])
    ws.append([68, 100, 180, 130])
    entrada = BytesIO()
    wb.save(entrada)
    entrada.seek(0)

    resultado = processar_precificacao_comum_flexivel(
        entrada,
        config=CommonPricingConfig(desconto_007_sobre_001=0.16),
        data_alteracao=date(2026, 11, 3),
        tabelas_existentes=["001", "007"],
        tabelas_para_gerar=["001", "007", "004", "012", "013"],
    )
    assert resultado.analise.iloc[0]["Preço Tabela 001"] == 180
    assert resultado.analise.iloc[0]["Preço Tabela 007"] == 130
    assert resultado.analise.iloc[0]["Origem Preço 007"] == "INFORMADO NA BASE"
    files = {f"{t}.csv": build_common_protheus_csv(resultado, t) for t in ("001", "007", "004", "012", "013")}
    files["relatorio.xlsx"] = build_common_report_excel(resultado)
    files["log.txt"] = resultado.log_text.encode("utf-8")
    for table, price in (("001", "180.00"), ("007", "130.00")):
        assert b"00068" in files[f"{table}.csv"]
        assert price.encode("ascii") in files[f"{table}.csv"]

    payload = build_zip_checked(files)
    with ZipFile(BytesIO(payload)) as archive:
        assert set(archive.namelist()) == set(files)
        assert all(archive.read(name) == data for name, data in files.items())


def test_zip_vazio_falha_com_erro_explicito():
    import pytest
    with pytest.raises(ValueError, match="Nenhum arquivo"):
        build_zip_checked({})
