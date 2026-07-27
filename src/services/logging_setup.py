"""Configuração centralizada dos logs."""

from __future__ import annotations

import logging
from pathlib import Path


def configure_logging(log_file: Path, verbose: bool = False) -> logging.Logger:
    """Cria log no terminal e em arquivo, sem duplicar handlers."""

    logger = logging.getLogger("excel_compras_automation")
    logger.setLevel(logging.DEBUG)
    for existing_handler in tuple(logger.handlers):
        logger.removeHandler(existing_handler)
        existing_handler.close()

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(levelname)s | %(message)s"))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger
