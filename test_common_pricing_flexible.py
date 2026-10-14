from datetime import date
from io import BytesIO

from openpyxl import Workbook

from common_pricing_flexible import processar_precificacao_comum_flexivel
from processor import CommonPricingConfig


def _xlsx_bytes(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def test_fluxo_historico_somente_001_calcula_007():
    arquivo = _xlsx_bytes(
        ["SKU", "Custo Médio", "Preço 001"],
        [[68, 100, 200]],
    )
    result = processar_precificacao_comum_flexivel(
        arquivo,
        config=CommonPricingConfig(desconto_007_sobre_001=0.16, acrescimo_013_sobre_007=0.12),
        data_alteracao=date(2026, 10, 14),
        tabelas_existentes=["001"],
        tabelas_para_gerar=["001", "004", "012", "007", "013"],
    )

    row = result.analise.iloc[0]
    assert row["SKU"] == "00068"
    assert row["Preço Tabela 001"] == 200.00
    assert row["Preço Tabela 007"] == 168.00
    assert row["Origem Preço 007"] == "CALCULADO DA 001"
    assert row["Preço Tabela 013"] == 188.16
    assert row["Preço Tabela 004"] == 42.00
    assert row["Preço Tabela 012"] == 42.00


def test_001_e_007_preserva_007_informada_e_gera_restante():
    arquivo = _xlsx_bytes(
        ["SKU", "Custo Médio", "Preço 001", "Preço 007"],
        [[68, 100, 150, 100]],
    )
    result = processar_precificacao_comum_flexivel(
        arquivo,
        config=CommonPricingConfig(desconto_007_sobre_001=0.16, acrescimo_013_sobre_007=0.12),
        data_alteracao=date(2026, 10, 14),
        tabelas_existentes=["001", "007"],
        tabelas_para_gerar=["004", "012", "013"],
    )

    row = result.analise.iloc[0]
    assert row["Preço 007 Teórico pela 001"] == 126.00
    assert row["Preço Tabela 007"] == 100.00
    assert row["Origem Preço 007"] == "INFORMADO NA BASE"
    assert row["Diferença 007 Informada vs Teórica"] == -26.00
    assert row["Preço Tabela 013"] == 112.00
    assert row["Preço Tabela 004"] == 25.00
    assert row["Preço Tabela 012"] == 25.00
    assert result.saidas["004"].iloc[0]["SKU"] == "00068"
    assert result.saidas["013"].iloc[0]["Preço"] == "112.00"
