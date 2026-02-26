import logging
from pathlib import Path


def configurar_logger(nome_rpa: str, pasta_log: Path) -> logging.Logger:
    pasta_log.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger(nome_rpa)
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    file_handler = logging.FileHandler(
        pasta_log / f"{nome_rpa}.log",
        encoding="utf-8"
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger