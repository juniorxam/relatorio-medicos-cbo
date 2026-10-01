import io
import pandas as pd
from openpyxl import load_workbook
import app


def source(rows):
    return pd.DataFrame([["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "OCUPACAO"]] + rows)


def test_only_three_exact_normalized_cargos_are_selected():
    df = app.clean(app.parse(source([
        ["MÉDICO", "Cardio", "1", "1", "Ana", "Cardiologia"],
        ["MEDICO - RQE", "Cardio", "2", "1", "Bia", "Cardiologia"],
        ["MEDICO CLINICO", "Clínica", "3", "1", "Caio", "Clínica médica"],
        ["MÉDICO VETERINÁRIO", "Outro", "4", "1", "Dora", "Veterinária"],
    ])))
    assert df["SERVIDOR"].tolist() == ["Ana", "Bia", "Caio"]


def test_summaries_include_specialties_by_sector():
    df = app.clean(app.parse(source([
        ["MÉDICO", "Cardio", "1", "1", "Ana", "Cardiologia"],
        ["MEDICO - RQE", "Cardio", "2", "1", "Bia", "Cardiologia"],
        ["MEDICO CLINICO", "Clínica", "3", "1", "Caio", "Clínica médica"],
    ])))
    sector = app.summary_sector(df)
    specialty = app.summary_specialty(df)
    assert sector.loc[sector.SETOR == "Cardio", "MEDICOS"].iloc[0] == 2
    assert "Cardiologia" in sector.loc[sector.SETOR == "Cardio", "ESPECIALIDADES"].iloc[0]
    assert len(specialty) == 2
    reference = app.reference_summary(df)
    assert reference.iloc[0]["TIPO_LINHA"] == "SETOR"
    assert set(reference["CONTRATO_TEMPORARIO"]) == {0}
    total = reference.iloc[-1]
    assert total["TIPO_LINHA"] == "TOTAL"
    assert total["TOTAL_GERAL"] == 3
    assert total["EFETIVOS"] == 0


def test_report_order_and_excel_pdf_exports():
    df = app.clean(app.parse(source([["MÉDICO", "Cardio", "1", "1", "Ana", "Cardiologia"]])))
    detail = app.report(df)
    assert detail.ORDEM.tolist() == [1]
    assert detail.SETOR.tolist() == ["Cardio"]
    workbook = load_workbook(io.BytesIO(app.excel({"Medicos": detail})))
    assert "Medicos" in workbook.sheetnames
    pdf = app.pdf(app.reference_summary(df))
    assert pdf.startswith(b"%PDF")


def test_read_raw_falls_back_to_text_when_excel_extension_is_misleading():
    content = b"CARGO;SETOR;NUMFUNC;NUMVINC\nMEDICO;Cardio;1;1\n"

    raw = app.read_raw(content, "planilha.xlsx")

    assert raw.shape == (2, 4)
    assert raw.iloc[1, 0] == "MEDICO"


def test_report_formats_numeric_fields_and_cpf():
    df = app.clean(app.parse(pd.DataFrame([
        ["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "OCUPACAO", "ESCALA", "CPF"],
        ["MÉDICO", "Cardio", "001234", "02", "Ana", "Cardiologia", "000350374", "5393676131"],
    ])))

    detail = app.report(df)

    assert int(detail.loc[0, "NUMFUNC"]) == 1234
    assert int(detail.loc[0, "NUMVINC"]) == 2
    assert int(detail.loc[0, "ESCALA"]) == 350374
    assert detail.loc[0, "CPF"] == "053.936.761-31"
    assert str(detail["NUMFUNC"].dtype) == "Int64"


def test_unique_servers_report_deduplicates_by_numfunc_and_numvinc():
    df = app.clean(app.parse(pd.DataFrame([
        ["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "OCUPACAO", "ESCALA", "CARGA HORARIA"],
        ["MÉDICO", "Cardio", "10", "1", "Ana", "Cardiologia", "100", "180"],
        ["MÉDICO", "Cardio", "10", "1", "Ana", "Cardiologia", "200", "180"],
        ["MÉDICO", "Cardio", "10", "2", "Ana", "Cardiologia", "300", "180"],
    ])))

    unique = app.unique_servers_report(df)

    assert unique[["NUMFUNC", "NUMVINC"]].astype(str).values.tolist() == [["10", "1"], ["10", "2"]]
    assert unique["ORDEM"].tolist() == [1, 2]
    assert list(unique.columns) == app.UNIQUE_SERVER_COLUMNS


def test_unique_servers_treats_excel_decimal_ids_as_the_same_key():
    df = app.clean(app.parse(pd.DataFrame([
        ["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "OCUPACAO"],
        ["MÉDICO", "Cardio", "10", "1", "Ana", "Cardiologia"],
        ["MÉDICO", "Cardio", "10.0", "1.0", "Ana", "Cardiologia"],
    ])))

    assert len(app.unique_servers_report(df)) == 1


def test_management_analysis_counts_unique_links_scales_and_hours():
    df = app.clean(app.parse(pd.DataFrame([
        ["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "OCUPACAO", "VINCULO", "CARGA HORARIA ESCALADA", "CARGA HORARIA"],
        ["MÉDICO", "Cardio", "10", "1", "Ana", "Cardiologia", "Concursado", "12", "24"],
        ["MÉDICO", "Cardio", "10", "1", "Ana", "Cardiologia", "Concursado", "12", "24"],
        ["MÉDICO", "Cardio", "11", "1", "Bia", "Cardiologia", "Contrato", "24", "24"],
    ])))

    base = app.management_base(df)
    sector = app.management_sector(df)

    assert len(base) == 2
    assert base[app.KEY_COLUMN].tolist() == ["10-1", "11-1"]
    assert base.loc[base.CHAVE_VINCULO == "10-1", "ESCALAS"].iloc[0] == 2
    assert sector.MEDICOS.iloc[0] == 2
    assert sector.ESCALAS.iloc[0] == 3
    assert sector.HORAS_ESCALADAS.iloc[0] == 48
    assert sector.CARGA_HORARIA.iloc[0] == 48


def test_unique_doctors_count_does_not_count_multiple_links_twice():
    df = app.clean(app.parse(pd.DataFrame([
        ["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "CPF", "OCUPACAO"],
        ["MÉDICO", "Cardio", "10", "1", "Ana", "11122233344", "Cardiologia"],
        ["MÉDICO", "Cardio", "10", "2", "Ana", "11122233344", "Cardiologia"],
        ["MÉDICO", "Cardio", "11", "1", "Bia", "55566677788", "Cardiologia"],
    ])))

    assert df[app.KEY_COLUMN].tolist() == ["10-1", "10-2", "11-1"]
    assert app.medicos_unicos_count(df) == 3
    assert app.management_sector(df).MEDICOS.iloc[0] == 3


def test_missing_numfunc_or_numvinc_gets_a_distinct_key_per_row():
    df = app.clean(app.parse(pd.DataFrame([
        ["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "CPF", "OCUPACAO"],
        ["MÉDICO", "Cardio", "", "1", "Ana", "11122233344", "Cardiologia"],
        ["MÉDICO", "Cardio", "", "2", "Ana", "11122233344", "Cardiologia"],
    ])))

    assert df[app.KEY_COLUMN].tolist() == ["SEM_CHAVE_1", "SEM_CHAVE_2"]
    assert app.medicos_unicos_count(df) == 2


def test_missing_composite_ids_are_not_collapsed_into_one_server():
    df = app.clean(app.parse(pd.DataFrame([
        ["CARGO", "SETOR", "NUMFUNC", "NUMVINC", "SERVIDOR", "OCUPACAO"],
        ["MÉDICO", "Cardio", "", "", "Ana", "Cardiologia"],
        ["MÉDICO", "Cardio", "", "", "Bia", "Cardiologia"],
    ])))

    assert len(app.unique_servers_report(df)) == 2
