import logging
import sys

_formatter = logging.Formatter("%(asctime)s | %(levelname)-8s | %(name)s | %(message)s", datefmt="%H:%M:%S")

_initialized = set()

def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if name not in _initialized:
        logger.setLevel(logging.INFO)
        _ch = logging.StreamHandler(sys.stdout)
        _ch.setFormatter(_formatter)
        logger.addHandler(_ch)
        _initialized.add(name)
    return logger
