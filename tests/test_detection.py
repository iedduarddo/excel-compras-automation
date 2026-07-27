from openpyxl import Workbook

from src.excel.detection import WorkbookDetector
from src.settings import load_aliases


def test_detects_renamed_sheets_and_reordered_columns() -> None:
    workbook = Workbook()
    base = workbook.active
    base.title = "Dados Operacionais 2026"
    base.append(["Relatório interno"])
    base.append([])
    base.append(
        [
            "Prestador",
            "Data do Serviço",
            "ID da Solicitação",
            "Data da Solicitação",
            "Centro Custo",
            "Tipo Serviço",
            "Qtde",
            "Preço Unitário",
            "Encargos",
            "Status da Reserva",
            "Nível de Criticidade",
            "Situação do Cartão",
        ]
    )

    policies = workbook.create_sheet("Parâmetros Corporativos")
    policies.append(["Título"])
    policies.append([])
    policies.append([])
    policies.append(
        [
            "Mínimo de Dias",
            "Teto",
            "Categoria do Serviço",
        ]
    )

    responses = workbook.create_sheet("Painel Gerencial")
    responses.append(["Indicador", "Resultado", "Cálculo"])

    layout = WorkbookDetector(load_aliases()).detect(workbook)

    assert layout.base.title == "Dados Operacionais 2026"
    assert layout.base.header_row == 3
    assert layout.base.columns["request_id"] == 3
    assert layout.base.columns["supplier"] == 1
    assert layout.policies.title == "Parâmetros Corporativos"
    assert layout.policies.header_row == 4
    assert layout.policies.columns["min_lead_days"] == 1
    assert layout.responses.title == "Painel Gerencial"
    assert layout.responses.columns["answer"] == 2
