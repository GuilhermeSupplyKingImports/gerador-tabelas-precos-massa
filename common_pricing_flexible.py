from __future__ import annotations

from datetime import date, datetime
from typing import Iterable, Optional, Set

import pandas as pd
from openpyxl import load_workbook

from processor import (
    CommonPricingConfig,
    CommonPricingResult,
    build_common_log_txt_from_frames,
    calcular_preco_venda_interna,
    money_dot,
    normalize_sku,
    parse_money,
    validate_sku_format,
)


SOURCE_TABLES = ("001", "007")
ALL_TABLES = ("001", "004", "012", "007", "013")


def _normalizar_tabelas(values: Optional[Iterable[str]], default: Iterable[str]) -> Set[str]:
    normalized = {str(value).zfill(3) for value in (values if values is not None else default)}
    return {value for value in normalized if value in ALL_TABLES}


def _read_precificacao_flexivel(input_file, config: CommonPricingConfig, tabelas_existentes: Set[str]):
    """Lê A=SKU, B=Custo e, a partir de C, as tabelas-fonte selecionadas.

    Compatibilidade preservada:
    - somente 001: C=Preço 001 (layout histórico);
    - 001 + 007: C=Preço 001, D=Preço 007;
    - somente 007: C=Preço 007.
    """
    try:
        input_file.seek(0)
    except Exception:
        pass

    wb = load_workbook(input_file, read_only=True, data_only=True)
    sheet = wb.sheetnames[0]
    ws = wb[sheet]

    ordered_sources = [table for table in SOURCE_TABLES if table in tabelas_existentes]
    max_col = 2 + len(ordered_sources)
    rows = []
    log = []

    for excel_row, values in enumerate(
        ws.iter_rows(min_row=2, min_col=1, max_col=max_col, values_only=True),
        start=2,
    ):
        raw_sku = values[0] if len(values) > 0 else None
        raw_custo = values[1] if len(values) > 1 else None
        source_values = values[2:]

        sku, sev, msg = validate_sku_format(raw_sku)
        if sku is None and (raw_sku is None or str(raw_sku).strip().lower() in {"", "nan", "none", "null"}):
            continue

        if sev != "OK":
            log.append({
                "Tipo": "FORMATO_SKU_PRECIFICACAO_COMUM",
                "Severidade": sev,
                "SKU": sku or "",
                "Mensagem": "Formato de SKU na planilha de precificação comum exige revisão." if sev == "ATENÇÃO" else "Formato de SKU inválido na planilha de precificação comum.",
                "Detalhe": f"Linha Excel aproximada: {excel_row}; valor original: {raw_sku}; {msg}",
            })

        if not sku:
            continue

        record = {
            "SKU": sku,
            "CUSTO_MEDIO": parse_money(raw_custo),
            "PRECO_001": None,
            "PRECO_007": None,
            "SKU_ORIGINAL": str(raw_sku),
            "LINHA_EXCEL": excel_row,
        }
        for idx, table in enumerate(ordered_sources):
            raw_value = source_values[idx] if idx < len(source_values) else None
            record[f"PRECO_{table}"] = parse_money(raw_value)
        rows.append(record)

    try:
        wb.close()
    except Exception:
        pass

    df = pd.DataFrame(
        rows,
        columns=["SKU", "CUSTO_MEDIO", "PRECO_001", "PRECO_007", "SKU_ORIGINAL", "LINHA_EXCEL"],
    )
    if df.empty:
        return df, log, sheet

    dupes = df[df.duplicated("SKU", keep=False)]
    if not dupes.empty:
        for sku, grp in dupes.groupby("SKU"):
            custos = sorted({round(float(x), 2) for x in grp["CUSTO_MEDIO"].dropna().tolist()})
            p001 = sorted({round(float(x), 2) for x in grp["PRECO_001"].dropna().tolist()})
            p007 = sorted({round(float(x), 2) for x in grp["PRECO_007"].dropna().tolist()})
            log.append({
                "Tipo": "DUPLICIDADE_PRECIFICACAO_COMUM",
                "Severidade": "ATENÇÃO",
                "SKU": sku,
                "Mensagem": "SKU duplicado na planilha de Precificação Comum.",
                "Detalhe": (
                    f"Linhas: {grp['LINHA_EXCEL'].tolist()}; custos: {custos}; "
                    f"preços 001: {p001}; preços 007: {p007}. "
                    f"Usada a {config.usar_ocorrencia_duplicada} ocorrência."
                ),
            })
        keep = "last" if config.usar_ocorrencia_duplicada == "ultima" else "first"
        df = df.drop_duplicates("SKU", keep=keep)

    return df, log, sheet


def processar_precificacao_comum_flexivel(
    input_file,
    config: Optional[CommonPricingConfig] = None,
    data_alteracao: Optional[date] = None,
    tabelas_existentes: Optional[Iterable[str]] = None,
    tabelas_para_gerar: Optional[Iterable[str]] = None,
) -> CommonPricingResult:
    """Processa Precificação Comum aceitando 001 e/ou 007 como preços já existentes.

    Regra de autoridade:
    - preço informado na planilha sempre prevalece;
    - se 007 não foi informado, ele é calculado a partir da 001;
    - 013 sempre parte do preço final da 007;
    - 004/012 usam o menor valor entre custo/2 e preço final 007/4;
    - 001 nunca é inferida a partir da 007.
    """
    config = config or CommonPricingConfig()
    data_alteracao = data_alteracao or date.today()
    data_str = data_alteracao.strftime("%d/%m/%Y")
    fontes = _normalizar_tabelas(tabelas_existentes, {"001"}) & set(SOURCE_TABLES)
    gerar = _normalizar_tabelas(tabelas_para_gerar, ALL_TABLES)

    if not fontes:
        raise ValueError("Selecione pelo menos uma tabela-fonte existente: 001 e/ou 007.")

    base, log_rows, sheet = _read_precificacao_flexivel(input_file, config, fontes)
    records = []

    for _, row in base.iterrows():
        sku = row.get("SKU")
        custo = row.get("CUSTO_MEDIO")
        preco_001 = row.get("PRECO_001")
        preco_007_informado = row.get("PRECO_007")
        linha = row.get("LINHA_EXCEL")
        obs = []

        custo_valido = custo is not None and not pd.isna(custo) and float(custo) > 0
        preco_001_valido = preco_001 is not None and not pd.isna(preco_001) and float(preco_001) > 0
        preco_007_informado_valido = (
            preco_007_informado is not None
            and not pd.isna(preco_007_informado)
            and float(preco_007_informado) > 0
        )

        preco_001_calc = round(float(preco_001), 2) if preco_001_valido else None
        preco_007_teorico = (
            round(float(preco_001_calc) * (1 - config.desconto_007_sobre_001), 2)
            if preco_001_calc is not None
            else None
        )

        if "007" in fontes:
            if preco_007_informado_valido:
                preco_007_calc = round(float(preco_007_informado), 2)
                origem_007 = "INFORMADO NA BASE"
            else:
                preco_007_calc = None
                origem_007 = "007 INFORMADO INVÁLIDO"
                obs.append("Preço 007 informado inválido/zerado/ausente")
                if gerar & {"004", "007", "012", "013"}:
                    log_rows.append({
                        "Tipo": "PRECO_007_INVALIDO_PRECIFICACAO_COMUM",
                        "Severidade": "CRÍTICO",
                        "SKU": sku or "",
                        "Mensagem": "Preço 007 selecionado como fonte, mas inválido/zerado/ausente.",
                        "Detalhe": f"Linha Excel aproximada: {linha}; preço 007 informado: {preco_007_informado}. O app não substituiu esse valor pela 001 para evitar sobrescrever uma tabela existente.",
                    })
        else:
            preco_007_calc = preco_007_teorico
            origem_007 = "CALCULADO DA 001"

        if "001" in fontes and not preco_001_valido:
            obs.append("Preço 001 inválido/zerado/ausente")
            if gerar & {"001", "007", "013"} and "007" not in fontes:
                log_rows.append({
                    "Tipo": "PRECO_001_INVALIDO_PRECIFICACAO_COMUM",
                    "Severidade": "CRÍTICO",
                    "SKU": sku or "",
                    "Mensagem": "Preço 001 inválido, zerado ou ausente.",
                    "Detalhe": f"Linha Excel aproximada: {linha}; preço 001 informado: {preco_001}. Não é possível calcular 007/013 a partir da 001.",
                })

        if (gerar & {"004", "012"}) and not custo_valido:
            obs.append("Custo médio inválido/zerado/ausente")
            log_rows.append({
                "Tipo": "CUSTO_INVALIDO_PRECIFICACAO_COMUM",
                "Severidade": "ATENÇÃO" if preco_007_calc is not None else "CRÍTICO",
                "SKU": sku or "",
                "Mensagem": "Custo médio inválido, zerado ou ausente.",
                "Detalhe": f"Linha Excel aproximada: {linha}; custo informado: {custo}. Se houver 007 válida, 004/012 ainda podem usar 007/4.",
            })

        preco_013_calc = (
            round(float(preco_007_calc) * (1 + config.acrescimo_013_sobre_007), 2)
            if preco_007_calc is not None
            else None
        )
        preco_interno_original, preco_interno_007_div_4, preco_interno_final, criterio_preco_interno = calcular_preco_venda_interna(custo, preco_007_calc)

        diferenca_007 = None
        if preco_007_informado_valido and preco_007_teorico is not None:
            diferenca_007 = round(float(preco_007_calc) - float(preco_007_teorico), 2)

        records.append({
            "SKU": sku,
            "Custo Médio": round(float(custo), 2) if custo is not None and not pd.isna(custo) else None,
            "Preço 001 Informado": round(float(preco_001), 2) if preco_001 is not None and not pd.isna(preco_001) else None,
            "Preço 007 Informado": round(float(preco_007_informado), 2) if preco_007_informado is not None and not pd.isna(preco_007_informado) else None,
            "Preço 007 Teórico pela 001": preco_007_teorico,
            "Diferença 007 Informada vs Teórica": diferenca_007,
            "Origem Preço 007": origem_007,
            "Preço Tabela 001": preco_001_calc,
            "Preço Interno Original 004/012 (Custo/2)": preco_interno_original,
            "Preço Interno 007/4": preco_interno_007_div_4,
            "Preço Tabela 004": preco_interno_final,
            "Preço Tabela 012": preco_interno_final,
            "Critério 004/012": criterio_preco_interno,
            "Preço Tabela 007": preco_007_calc,
            "Preço Tabela 013": preco_013_calc,
            "Data Alteração": data_str,
            "Status": "OK" if not obs else "VERIFICAR",
            "Observação": "; ".join(obs),
        })

    analise = pd.DataFrame(records)

    def _saida(preco_col: str) -> pd.DataFrame:
        if analise.empty or preco_col not in analise.columns:
            return pd.DataFrame(columns=["SKU", "Preço", "Data"])
        df = analise[
            analise[preco_col].notna() & analise["SKU"].notna()
        ][["SKU", preco_col, "Data Alteração"]].copy()
        df.columns = ["SKU", "Preço", "Data"]
        df["SKU"] = df["SKU"].astype(str).apply(normalize_sku)
        df["Preço"] = df["Preço"].apply(money_dot)
        return df

    saidas = {
        str(config.tabela_001).zfill(3): _saida("Preço Tabela 001"),
        str(config.tabela_004).zfill(3): _saida("Preço Tabela 004"),
        str(config.tabela_012).zfill(3): _saida("Preço Tabela 012"),
        str(config.tabela_007).zfill(3): _saida("Preço Tabela 007"),
        str(config.tabela_013).zfill(3): _saida("Preço Tabela 013"),
    }

    resumo_dict = {
        "Data de geração": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Aba de origem lida": sheet,
        "Tabelas-fonte informadas": ", ".join(sorted(fontes)),
        "Tabelas solicitadas para geração": ", ".join(sorted(gerar)),
        "SKUs válidos na base": int(len(base)),
        "SKUs com custo médio válido": int((analise["Custo Médio"].notna() & (analise["Custo Médio"] > 0)).sum()) if not analise.empty else 0,
        "SKUs com preço 001 válido": int((analise["Preço 001 Informado"].notna() & (analise["Preço 001 Informado"] > 0)).sum()) if not analise.empty else 0,
        "SKUs com preço 007 informado válido": int((analise["Preço 007 Informado"].notna() & (analise["Preço 007 Informado"] > 0)).sum()) if not analise.empty else 0,
        "Linhas geradas tabela 001": int(len(saidas[str(config.tabela_001).zfill(3)])),
        "Linhas geradas tabela 004": int(len(saidas[str(config.tabela_004).zfill(3)])),
        "Linhas geradas tabela 012": int(len(saidas[str(config.tabela_012).zfill(3)])),
        "Linhas geradas tabela 007": int(len(saidas[str(config.tabela_007).zfill(3)])),
        "Linhas geradas tabela 013": int(len(saidas[str(config.tabela_013).zfill(3)])),
        "Política de autoridade": "preço informado prevalece; 007 só é calculada da 001 quando 007 não foi selecionada como fonte",
        "Política 004/012": "menor valor entre custo médio / 2 e preço final 007 / 4",
        "Itens 004/012 com preço por custo/2": int((analise["Critério 004/012"] == "CUSTO/2").sum()) if not analise.empty else 0,
        "Itens 004/012 com preço por 007/4": int((analise["Critério 004/012"] == "007/4").sum()) if not analise.empty else 0,
        "Desconto da 007 sobre a 001 (quando calculada)": f"{config.desconto_007_sobre_001:.2%}",
        "Acréscimo aplicado na 013 sobre a 007": f"{config.acrescimo_013_sobre_007:.2%}",
        "Intercorrências críticas": int(sum(1 for x in log_rows if str(x.get("Severidade", "")).upper() == "CRÍTICO")),
        "Intercorrências de atenção": int(sum(1 for x in log_rows if str(x.get("Severidade", "")).upper() == "ATENÇÃO")),
    }

    resumo = pd.DataFrame([{"Indicador": k, "Valor": v} for k, v in resumo_dict.items()])
    log = pd.DataFrame(log_rows, columns=["Tipo", "Severidade", "SKU", "Mensagem", "Detalhe"])
    log_text = build_common_log_txt_from_frames(resumo, log, saidas)

    return CommonPricingResult(
        base=base,
        analise=analise,
        log=log,
        resumo=resumo,
        saidas=saidas,
        resumo_dict=resumo_dict,
        log_text=log_text,
    )
