import pandas as pd
from pathlib import Path

HEADERS = {
    "BH(ATIVAS)": [
        "ESCRITÓRIO",
        "CNPJ",
        "PARTE AUTORA",
        "PARTE RÉ",
        "CPF/CNPJ RÉ",
        "NÚMERO DO PROCESSO",
        "UF",
        "COMARCA",
        "JUÍZO",
        "DATA DA AÇÃO",
        "TIPO DE AÇÃO",
        "ASSUNTO",
        "LIMINAR",
        "VALOR DA CAUSA",
        "VALOR DO RISCO ATUALIZADO",
        "PROBABILIDADE DE PERDA",
        "HONORÁRIOS",
        "DEPÓSITOS BARI",
        "DEPÓSITOS ATUALIZADOS - BARI",
        "DEPÓSITOS - AUTORA",
        "ÚLTIMO ANDAMENTO PROCESSUAL",
        "PRODUTO",
        "OBSERVAÇÃO",
    ],
    "BH(PASSIVAS)": [
        "ESCRITÓRIO",
        "CNPJ",
        "PARTE AUTORA",
        "PARTE RÉ",
        "NÚMERO DO PROCESSO",
        "UF",
        "COMARCA",
        "JUÍZO",
        "DATA DA AÇÃO",
        "TIPO DE AÇÃO",
        "ASSUNTO",
        "LIMINAR",
        "VALOR DA CAUSA",
        "VALOR DO RISCO ATUALIZADO",
        "PROBABILIDADE DE PERDA",
        "HONORÁRIOS",
        "DEPÓSITOS BARI",
        "DEPÓSITOS ATUALIZADOS - BARI",
        "DEPÓSITOS - CLIENTE",
        "ÚLTIMO ANDAMENTO PROCESSUAL",
        "PRODUTO",
        "OBSERVAÇÃO",
    ],
    "SEC(ATIVAS)": [
        "ESCRITÓRIO",
        "CNPJ",
        "PARTE AUTORA",
        "PARTE RÉ",
        "CPF/CNPJ RÉ",
        "CRI",
        "NÚMERO DO PROCESSO",
        "UF",
        "COMARCA",
        "JUÍZO",
        "DATA DA AÇÃO",
        "TIPO DE AÇÃO",
        "ASSUNTO",
        "LIMINAR",
        "VALOR DA CAUSA",
        "VALOR DO RISCO ATUALIZADO",
        "PROBABILIDADE DE PERDA",
        "HONORÁRIOS",
        "DEPÓSITOS BARI",
        "DEPÓSITOS ATUALIZADOS - BARI",
        "DEPÓSITOS - AUTORA",
        "ÚLTIMO ANDAMENTO PROCESSUAL",
        "PRODUTO",
        "OBSERVAÇÃO",
    ],
    "SEC(PASSIVAS)": [
        "ESCRITÓRIO",
        "CNPJ",
        "PARTE AUTORA",
        "PARTE RÉ",
        "CPF/CNPJ RÉ",
        "CRI",
        "NÚMERO DO PROCESSO",
        "UF",
        "COMARCA",
        "JUÍZO",
        "DATA DA AÇÃO",
        "TIPO DE AÇÃO",
        "ASSUNTO",
        "LIMINAR",
        "VALOR DA CAUSA",
        "VALOR DO RISCO ATUALIZADO",
        "PROBABILIDADE DE PERDA",
        "HONORÁRIOS",
        "DEPÓSITOS BARI",
        "DEPÓSITOS ATUALIZADOS - BARI",
        "DEPÓSITOS - CLIENTE",
        "ÚLTIMO ANDAMENTO PROCESSUAL",
        "PRODUTO",
        "OBSERVAÇÃO",
    ],
}

banco: dict[str, list[pd.DataFrame]] = {"ativas": [], "passivas": []}
hipo: dict[str, list[pd.DataFrame]] = {"ativas": [], "passivas": []}
sec: dict[str, list[pd.DataFrame]] = {"ativas": [], "passivas": []}
trabalhistas: dict[str, list[pd.DataFrame]] = {"BANCO": [], "SERVICE": [], "PROMOTORA": []}
outros: list[pd.DataFrame] = [].clear()


def salvar_aba(lista_dfs: list[pd.DataFrame], writer, nome_aba) -> None:
    if lista_dfs:
        pd.concat(lista_dfs, ignore_index=True).drop_duplicates().to_excel(writer, sheet_name=nome_aba, index=False)


def tree_search(path: Path, doc_type: str, is_root: bool = True) -> list[Path]:
    docs = []
    for p in path.iterdir():
        if p.is_dir():
            docs.extend(tree_search(p, doc_type, False))
        elif p.suffix == doc_type and not is_root:
            docs.append(p)
    if not len(docs) > 0:
        return []
    else:
        return docs
