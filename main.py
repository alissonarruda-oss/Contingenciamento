import pandas as pd
import arrow
import warnings
from pathlib import Path

from services.processor import processar_aba
from services.path_logic import tree_search, exportar_consolidado, exportar_trabalhistas, exportar_outros, exportar_encerradas

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

_start_time = arrow.now()

def _elapsed() -> str:
    """Retorna o tempo decorrido desde o início da execução no formato HH:MM:SS:cs."""
    delta = arrow.now() - _start_time
    total_seconds = delta.total_seconds()
    h = int(total_seconds // 3600)
    m = int((total_seconds % 3600) // 60)
    s = int(total_seconds % 60)
    cs = int((total_seconds % 1) * 100)
    return f"{h:02}:{m:02}:{s:02}:{cs:02}"

path = Path("")
docs = tree_search(path, ".xlsx")
numDocs = len(docs)

for i, doc in enumerate(docs):
    print(f"Lendo Documento : {doc.name}")

    try:
        excel_pages = pd.read_excel(doc, sheet_name=None)
    except Exception as e:
        print(f"⚠️ Falha ao tentar abrir o arquivo {doc.name}: {e}")
        continue

    for name, page in excel_pages.items():
        processar_aba(name, page, doc)

    print(f"{'='*35 + ' '} {((i+1)/numDocs*100):.1f}% {' ' + '='*35} [{_elapsed()}]")

exportar_consolidado()
exportar_trabalhistas()
exportar_outros()
exportar_encerradas()

print(f"\n✅ Relatórios exportados com sucesso! Tempo total: {_elapsed()}")
