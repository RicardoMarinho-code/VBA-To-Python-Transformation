import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_env(path):
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def _resolve(p):
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)


_load_env(os.path.join(BASE_DIR, ".env"))

WORKBOOK_PATH = _resolve(os.getenv("WORKBOOK_PATH", "planilha.xlsx"))
REPORT_FILE = _resolve(os.getenv("REPORT_FILE", "relatorio_depuracao.csv"))
