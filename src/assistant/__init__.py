"""Assistente local orientado por pastas e comandos de texto."""

from src.assistant.commands import AssistantCommand, AssistantIntent, parse_command
from src.assistant.service import FolderAssistant
from src.assistant.universal import UniversalAction, UniversalAutomation
from src.assistant.workspace import AssistantWorkspace

__all__ = [
    "AssistantCommand",
    "AssistantIntent",
    "AssistantWorkspace",
    "FolderAssistant",
    "UniversalAction",
    "UniversalAutomation",
    "parse_command",
]
