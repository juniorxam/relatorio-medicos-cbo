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


def test_report_order_and_excel_pdf_exports():
    df = app.clean(app.parse(source([["MÉDICO", "Cardio", "1", "1", "Ana", "Cardiologia"]])))
    detail = app.report(df)
    assert detail.ORDEM.tolist() == [1]
    assert detail.SETOR.tolist() == ["Cardio"]
    workbook = load_workbook(io.BytesIO(app.excel({"Medicos": detail})))
    assert "Medicos" in workbook.sheetnames
    pdf = app.pdf(app.summary_sector(df), detail, app.summary_specialty(df))
    assert pdf.startswith(b"%PDF")


def test_read_raw_falls_back_to_text_when_excel_extension_is_misleading():
    content = b"CARGO;SETOR;NUMFUNC;NUMVINC\nMEDICO;Cardio;1;1\n"

    raw = app.read_raw(content, "planilha.xlsx")

    assert raw.shape == (2, 4)
    assert raw.iloc[1, 0] == "MEDICO"
