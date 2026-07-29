"""Exceções com mensagens adequadas para o usuário final."""


class AutomationError(Exception):
    """Erro esperado durante a execução da automação."""


class DetectionError(AutomationError):
    """A estrutura necessária não foi encontrada na planilha."""


class ValidationError(AutomationError):
    """A planilha contém valores inválidos ou insuficientes."""


class ExcelDesktopError(AutomationError):
    """O Microsoft Excel não conseguiu concluir a etapa nativa."""


class ExcelDesktopCleanupError(ExcelDesktopError):
    """O Excel concluiu ou iniciou a operação, mas não liberou os recursos."""
