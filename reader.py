import pandas as pd
from pathlib import Path
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

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


def salvar_aba(lista_dfs:list[pd.DataFrame], writer, nome_aba):
    if lista_dfs:
        pd.concat(lista_dfs, ignore_index=True).drop_duplicates().to_excel(writer, sheet_name=nome_aba, index=False)


def tree_search(path: Path, doc_type: str, is_root: bool = True):
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


path = Path("")
docs = tree_search(path, ".xlsx")

banco: dict[str, list[pd.DataFrame]] = {"ativas": [], "passivas": []}
hipo: dict[str, list[pd.DataFrame]] = {"ativas": [], "passivas": []}
sec: dict[str, list[pd.DataFrame]] = {"ativas": [], "passivas": []}
trabalhistas: dict[str, list[pd.DataFrame]] = {"BANCO": [], "SERVICE": [], "PROMOTORA": []}

for doc in docs:
    print(f"Documento : {doc}")
    excel_pages = pd.read_excel(doc, sheet_name=None)
    for name, page in excel_pages.items():
        print(name)
        page = page.loc[:, ~page.columns.duplicated()]
        try:
            if not (page.columns == "OBS.").any():
                columns_str = page.columns.str
                if not columns_str.contains("ENCERRAMENTO", case=False).any() and columns_str.contains("ESCRITÓRIO", case=False).any():
                    filtrados = page[page["ESCRITÓRIO"].notna() & page["PARTE AUTORA"].notna()]

                    if not columns_str.contains("DEPÓSITOS RECLAMANTE", case=False).any():
                        
                        print(page.columns)

                        ba = filtrados[filtrados["PARTE AUTORA"].str.contains("BANCO BARI", case=False, na=False)]
                        if not ba.empty:
                            ba = ba.rename(columns=dict(zip(ba.columns, HEADERS["BH(ATIVAS)"])))
                            banco["ativas"].append(ba)
                        bp = filtrados[filtrados["PARTE RÉ"].str.contains("BANCO", case=False, na=False)]
                        if not bp.empty:
                            bp = bp.rename(columns=dict(zip(bp.columns, HEADERS["BH(PASSIVAS)"])))
                            banco["passivas"].append(bp)
                        ha = filtrados[filtrados["PARTE AUTORA"].str.contains("HIPOTECÁRIA", case=False, na=False)]
                        if not ha.empty:
                            ha = ha.rename(columns=dict(zip(ha.columns, HEADERS["BH(ATIVAS)"])))
                            hipo["ativas"].append(ha)
                        hp = filtrados[filtrados["PARTE RÉ"].str.contains("HIPOTECÁRIA", case=False, na=False)]
                        if not hp.empty:
                            hp = hp.rename(columns=dict(zip(hp.columns, HEADERS["BH(PASSIVAS)"])))
                            hipo["passivas"].append(hp)
                        sa = filtrados[filtrados["PARTE AUTORA"].str.contains("SECURITIZADORA", case=False, na=False)]
                        if not sa.empty:
                            sa = sa.rename(columns=dict(zip(sa.columns, HEADERS["SEC(ATIVAS)"])))
                            sec["ativas"].append(sa)
                        sp = filtrados[filtrados["PARTE RÉ"].str.contains("SECURITIZADORA", case=False, na=False)]
                        if not sp.empty:
                            sp = sp.rename(columns=dict(zip(sp.columns, HEADERS["SEC(PASSIVAS)"])))
                            sec["passivas"].append(sp)
                    else:
                        tb = filtrados[filtrados["PARTE RÉ"].str.contains("BANCO", case=False, na=False)]
                        if not tb.empty:
                            trabalhistas["BANCO"].append(tb)
                        ts = filtrados[filtrados["PARTE RÉ"].str.contains("SERVICE", case=False, na=False)]
                        if not ts.empty:
                            trabalhistas["SERVICE"].append(ts)
                        tp = filtrados[filtrados["PARTE RÉ"].str.contains("PROMOTORA", case=False, na=False)]
                        if not tp.empty:
                            trabalhistas["PROMOTORA"].append(tp)
        except Exception as e:
            print(f"Erro na aba {name}: {e}")

with pd.ExcelWriter("FINAL - CONSOLIDADO - HIPOTECÁRIA_BANCO_SEC_.xlsx", engine="openpyxl") as wr:
    salvar_aba(banco["ativas"], wr, "Banco(Ativas)")
    salvar_aba(banco["passivas"], wr, "Banco(Passivas)")
    salvar_aba(hipo["ativas"], wr, "HIPO(Ativas)")
    salvar_aba(hipo["passivas"], wr, "HIPO(Passivas)")
    salvar_aba(sec["ativas"], wr, "SEC(Ativas)")
    salvar_aba(sec["passivas"], wr, "SEC(Passivas)")

with pd.ExcelWriter("TRABALHISTA - SERVICE E PROMOTORA.xlsx", engine="openpyxl") as wr:
    salvar_aba(trabalhistas["SERVICE"], wr, "AÇÕES TRABALHISTAS - SERVICE")
    salvar_aba(trabalhistas["PROMOTORA"], wr, "AÇÕES TRABALHISTAS - PROMOTORA")

salvar_aba(trabalhistas["BANCO"], "TRABALHISTA - BANCO E HIPO.xlsx", "AÇÕES TRABALHISTAS - BANCO")

print("Relatórios exportados com sucesso!")
