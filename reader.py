import pandas as pd
from pathlib import Path
import warnings
from utils import *

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

path = Path(".")
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
        if name.upper() in ["GRÁFICOS", "ANALYTICS", "DADOS", "TESTE"]:
            continue

        try:
            columns_str = page.columns.str
            if not columns_str.contains("ENCERRAMENTO", case=False).any() and columns_str.contains("ESCRITÓRIO", case=False).any():
                
                page_limpa = page.dropna(how="all")
                
                mask_essencial = page_limpa["ESCRITÓRIO"].notna() & page_limpa["PARTE AUTORA"].notna() & page_limpa["PARTE RÉ"].notna()
                
                filtrados = page_limpa[mask_essencial].copy()
                
                linhas_sem_essencial = page_limpa[~mask_essencial].copy()
                if not linhas_sem_essencial.empty:
                    linhas_sem_essencial["ARQUIVO_ORIGEM"] = doc.name
                    linhas_sem_essencial["ABA_ORIGEM"] = name
                    linhas_sem_essencial["MOTIVO_ERRO"] = "Faltou ESCRITÓRIO, PARTE AUTORA ou PARTE RÉ"
                    outros.append(linhas_sem_essencial)

                matched_mask = pd.Series(False, index=filtrados.index)

                if not columns_str.contains("DEPÓSITOS RECLAMANTE", case=False).any():

                    mask_ba = filtrados["PARTE AUTORA"].str.contains("BANCO BARI", case=False, na=False)
                    ba = filtrados[mask_ba]
                    if not ba.empty:
                        ba = ba.rename(columns=dict(zip(ba.columns, HEADERS["BH(ATIVAS)"])))
                        banco["ativas"].append(ba)

                    mask_bp = filtrados["PARTE RÉ"].str.contains("BANCO", case=False, na=False)
                    bp = filtrados[mask_bp]
                    if not bp.empty:
                        bp = bp.rename(columns=dict(zip(bp.columns, HEADERS["BH(PASSIVAS)"])))
                        banco["passivas"].append(bp)

                    mask_ha = filtrados["PARTE AUTORA"].str.contains("HIPOTECÁRIA", case=False, na=False)
                    ha = filtrados[mask_ha]
                    if not ha.empty:
                        ha = ha.rename(columns=dict(zip(ha.columns, HEADERS["BH(ATIVAS)"])))
                        hipo["ativas"].append(ha)

                    mask_hp = filtrados["PARTE RÉ"].str.contains("HIPOTECÁRIA", case=False, na=False)
                    hp = filtrados[mask_hp]
                    if not hp.empty:
                        hp = hp.rename(columns=dict(zip(hp.columns, HEADERS["BH(PASSIVAS)"])))
                        hipo["passivas"].append(hp)

                    mask_sa = filtrados["PARTE AUTORA"].str.contains("SECURITIZADORA", case=False, na=False)
                    sa = filtrados[mask_sa]
                    if not sa.empty:
                        sa = sa.rename(columns=dict(zip(sa.columns, HEADERS["SEC(ATIVAS)"])))
                        sec["ativas"].append(sa)

                    mask_sp = filtrados["PARTE RÉ"].str.contains("SECURITIZADORA", case=False, na=False)
                    sp = filtrados[mask_sp]
                    if not sp.empty:
                        sp = sp.rename(columns=dict(zip(sp.columns, HEADERS["SEC(PASSIVAS)"])))
                        sec["passivas"].append(sp)

                    matched_mask = mask_ba | mask_bp | mask_ha | mask_hp | mask_sa | mask_sp

                else:

                    mask_tb = filtrados["PARTE RÉ"].str.contains("BANCO", case=False, na=False)
                    tb = filtrados[mask_tb]
                    if not tb.empty:
                        trabalhistas["BANCO"].append(tb)

                    mask_ts = filtrados["PARTE RÉ"].str.contains("SERVICE", case=False, na=False)
                    ts = filtrados[mask_ts]
                    if not ts.empty:
                        trabalhistas["SERVICE"].append(ts)

                    mask_tp = filtrados["PARTE RÉ"].str.contains("PROMOTORA", case=False, na=False)
                    tp = filtrados[mask_tp]
                    if not tp.empty:
                        trabalhistas["PROMOTORA"].append(tp)

                    matched_mask = mask_tb | mask_ts | mask_tp

                df_outros = filtrados[~matched_mask].copy()
                if not df_outros.empty:
                    df_outros["ARQUIVO_ORIGEM"] = doc.name
                    df_outros["ABA_ORIGEM"] = name
                    outros.append(df_outros)

        except Exception as e:
            print(f"Erro na aba '{name}': {e}")

    print(f"{'='*35 + ' '} {((i+1)/numDocs*100):.1f}% {' ' + '='*35}")

with pd.ExcelWriter("FINAL - CONSOLIDADO - HIPOTECÁRIA_BANCO_SEC.xlsx", engine="openpyxl") as wr:
    salvar_aba(banco["ativas"], wr, "Banco(Ativas)")
    salvar_aba(banco["passivas"], wr, "Banco(Passivas)")
    salvar_aba(hipo["ativas"], wr, "HIPO(Ativas)")
    salvar_aba(hipo["passivas"], wr, "HIPO(Passivas)")
    salvar_aba(sec["ativas"], wr, "SEC(Ativas)")
    salvar_aba(sec["passivas"], wr, "SEC(Passivas)")

with pd.ExcelWriter("TRABALHISTA - SERVICE E PROMOTORA.xlsx", engine="openpyxl") as wr:
    salvar_aba(trabalhistas["SERVICE"], wr, "AÇÕES TRABALHISTAS - SERVICE")
    salvar_aba(trabalhistas["PROMOTORA"], wr, "AÇÕES TRABALHISTAS - PROMOTORA")

with pd.ExcelWriter("TRABALHISTA - BANCO E HIPO.xlsx", engine="openpyxl") as wr:
    salvar_aba(trabalhistas["BANCO"], wr, "AÇÕES TRABALHISTAS - BANCO")

if outros:
    with pd.ExcelWriter("VERIFICAR_OUTROS.xlsx", engine="openpyxl") as wr:
        salvar_aba(outros, wr, "Nao Classificados")
    print("\n⚠️ Alguns registros não foram classificados. Verifique o arquivo 'VERIFICAR_OUTROS.xlsx'")

print("\n✅ Relatórios exportados com sucesso!")
