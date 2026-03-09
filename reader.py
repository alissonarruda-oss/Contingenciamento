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
            if not columns_str.contains("ENCERRAMENTO", case=False).any() and columns_str.contains("ESCRITÓRIO", case=False).any() and not (page.columns == "OBS.").any():
                if all(columns_str.contains(header, case=False).any() for header in HEADERS["ESSENCIAIS"]):
                    pagina = page.dropna(how="all")

                    if not columns_str.contains("DEPÓSITOS RECLAMANTE", case=False).any():
                        for index, row in pagina.iterrows():
                            if "BANCO" in str(row["PARTE AUTORA"]):
                                validador(HEADERS["BH(ATIVAS)"], banco["ativas"], outros, row, doc, name)

                            elif "BANCO" in str(row["PARTE RÉ"]):
                                validador(HEADERS["BH(PASSIVAS)"], banco["passivas"], outros, row, doc, name)

                            elif "HIPOTECÁRIA" in str(row["PARTE AUTORA"]):
                                validador(HEADERS["BH(ATIVAS)"], hipo["ativas"], outros, row, doc, name)

                            elif "HIPOTECÁRIA" in str(row["PARTE RÉ"]):
                                validador(HEADERS["BH(PASSIVAS)"], hipo["passivas"], outros, row, doc, name)

                            elif "SECURITIZADORA" in str(row["PARTE AUTORA"]):
                                validador(HEADERS["SEC(ATIVAS)"], sec["ativas"], outros, row, doc, name)

                            elif "SECURITIZADORA" in str(row["PARTE RÉ"]):
                                validador(HEADERS["SEC(PASSIVAS)"], sec["passivas"], outros, row, doc, name)
                    else:
                        for index, row in pagina.iterrows():
                            if "BANCO" in str(row["PARTE RÉ"]):
                                validador([], trabalhistas["BANCO"], outros, row, doc, name)
                            elif "SERVICE" in str(row["PARTE RÉ"]):
                                validador([], trabalhistas["SERVICE"], outros, row, doc, name)
                            elif "PROMOTORA" in str(row["PARTE RÉ"]):
                                validador([], trabalhistas["PROMOTORA"], outros, row, doc, name)
                else:
                    pagina_invalida = page.dropna(how="all")

                    for index, row in pagina_invalida.iterrows():
                        row["ARQUIVO_ORIGEM"] = doc.name
                        row["ABA_ORIGEM"] = name
                        row["PROBLEMA"] = f"Colunas faltantes"
                        outros.append(row)

        except Exception as e:
            print(f"Erro na aba '{name}': {e}")

    print(f"{'='*35 + ' '} {((i+1)/numDocs*100):.1f}% {' ' + '='*35}")

try:
    with pd.ExcelWriter("CONSOLIDADO - HIPOTECÁRIA_BANCO_SEC.xlsx", engine="openpyxl") as wr:
        salvar_aba(banco["ativas"], wr, "BANCO - ATIVAS")
        salvar_aba(banco["passivas"], wr, "BANCO - PASSIVAS")
        salvar_aba(hipo["ativas"], wr, "HIPO - ATIVAS")
        salvar_aba(hipo["passivas"], wr, "HIPO - PASSIVAS")
        salvar_aba(sec["ativas"], wr, "SEC - ATIVAS")
        salvar_aba(sec["passivas"], wr, "SEC - PASSIVAS")
except Exception as e:
    print("ERRO AS CRIAR HIPOTECÁRIA_BANCO_SEC")

try:
    with pd.ExcelWriter("TRABALHISTA - SERVICE E PROMOTORA.xlsx", engine="openpyxl") as wr:
        salvar_aba(trabalhistas["SERVICE"], wr, "AÇÕES TRABALHISTAS - SERVICE")
        salvar_aba(trabalhistas["PROMOTORA"], wr, "AÇÕES TRABALHISTAS - PROMOTORA")
except Exception as e:
    print("ERRO AS CRIAR TRABALHISTAS - SERVICE E PROMOTORA")

try:
    with pd.ExcelWriter("TRABALHISTA - BANCO E HIPO.xlsx", engine="openpyxl") as wr:
        salvar_aba(trabalhistas["BANCO"], wr, "AÇÕES TRABALHISTAS - BANCO")
except Exception as e:
    print("ERRO AS CRIAR TRABALHISTAS - BANCO E HIPO")

if outros:
    with pd.ExcelWriter("VERIFICAR_OUTROS.xlsx", engine="openpyxl") as wr:
        problemas_unicos = set(row.get("PROBLEMA", "Erro Desconhecido") for row in outros)
        for problema in problemas_unicos:

            linhas_do_problema = [row for row in outros if row.get("PROBLEMA") == problema]

            nome_aba = str(problema).replace(":", "-").replace("/", "-")
            nome_aba = nome_aba[:31].strip()

            salvar_aba(linhas_do_problema, wr, nome_aba)

    print("\n⚠️ Alguns registros não foram classificados. Verifique o arquivo 'VERIFICAR_OUTROS.xlsx'")

print("\n✅ Relatórios exportados com sucesso!")
