import pandas as pd
from pathlib import Path

from services.constants import (
    HEADERS, COLS, MONETARIO, ABAS_IGNORADAS,
    banco, hipo, sec, trabalhistas, outros, encerradas
)

def rejeitar(row: pd.Series, outros: list[pd.Series], doc: Path, nome: str, problema: str) -> None:
    """Marca uma linha como inválida, anotando a origem e o motivo da rejeição, e a envia para a lista de registros com problemas."""
    row["ARQUIVO_ORIGEM"] = doc.name
    row["ABA_ORIGEM"] = nome
    row["PROBLEMA"] = problema
    outros.append(row)

def validador(page_headers: list[str], arr_final: list[pd.Series], outros: list[pd.Series], row: pd.Series, doc: Path, nome: str) -> None:
    """Valida uma linha verificando quantidade de colunas, campos obrigatórios e conversão de valores monetários; adiciona ao destino correto ou rejeita com o motivo."""
    if len(page_headers) > 0:
        if len(row) == len(page_headers):
            row.index = page_headers
            if all(pd.notna(row[col]) for col in COLS):
                for mon in MONETARIO:
                    try:
                        valor_str = str(row[mon]).replace("R$", "")
                        if valor_str and valor_str.lower() != "nan":
                            row[mon] = float(valor_str)
                    except ValueError:
                        rejeitar(row, outros, doc, nome, "Valor não monetário")
                        return
                arr_final.append(row)
            else:
                rejeitar(row, outros, doc, nome, "Informações não preenchidas")
        else:
            rejeitar(row, outros, doc, nome, "Quantidade de colunas incorretas")
    else:
        arr_final.append(row)

def classificar_civil(pagina: pd.DataFrame, doc: Path, name: str) -> None:
    """Classifica cada linha de uma aba cível identificando a entidade (Banco, Hipotecária ou Securitizadora) e a posição (ativa/passiva) com base nas partes do processo."""
    for _, row in pagina.iterrows():
        parte_autora = str(row["PARTE AUTORA"])
        parte_re = str(row["PARTE RÉ"])

        if "BANCO" in parte_autora:
            validador(HEADERS["BH(ATIVAS)"], banco["ativas"], outros, row, doc, name)
        elif "BANCO" in parte_re:
            validador(HEADERS["BH(PASSIVAS)"], banco["passivas"], outros, row, doc, name)
        elif "HIPOTECÁRIA" in parte_autora:
            validador(HEADERS["BH(ATIVAS)"], hipo["ativas"], outros, row, doc, name)
        elif "HIPOTECÁRIA" in parte_re:
            validador(HEADERS["BH(PASSIVAS)"], hipo["passivas"], outros, row, doc, name)
        elif "SECURITIZADORA" in parte_autora:
            validador(HEADERS["SEC(ATIVAS)"], sec["ativas"], outros, row, doc, name)
        elif "SECURITIZADORA" in parte_re:
            validador(HEADERS["SEC(PASSIVAS)"], sec["passivas"], outros, row, doc, name)


def classificar_trabalhista(pagina: pd.DataFrame, doc: Path, name: str) -> None:
    """Classifica cada linha de uma aba trabalhista roteando para a lista da entidade correspondente (Banco, Service, Promotora ou Hipotecária) com base na Parte Ré."""
    for _, row in pagina.iterrows():
        parte_re = str(row["PARTE RÉ"])

        if "BANCO" in parte_re:
            validador([], trabalhistas["BANCO"], outros, row, doc, name)
        elif "SERVICE" in parte_re:
            validador([], trabalhistas["SERVICE"], outros, row, doc, name)
        elif "PROMOTORA" in parte_re:
            validador([], trabalhistas["PROMOTORA"], outros, row, doc, name)
        elif "HIPOTECÁRIA" in parte_re:
            validador([], trabalhistas["HIPO"], outros, row, doc, name)


def processar_aba(name: str, page: pd.DataFrame, doc: Path) -> None:
    """Ponto de entrada por aba: ignora abas desnecessárias, detecta o tipo da planilha (encerrada, cível ou trabalhista) e delega para o classificador adequado."""
    if name.upper() in ABAS_IGNORADAS:
        return
    
    try:
        columns_str = page.columns.str
        if not columns_str.contains("ENCERRAMENTO", case=False).any():
            if not (page.columns == "OBS.").any():
                if all(columns_str.contains(header, case=False).any() for header in HEADERS["ESSENCIAIS"]):
                    pagina = page.dropna(how="all")

                    if not columns_str.contains("DEPÓSITOS RECLAMANTE", case=False).any():
                        classificar_civil(pagina, doc, name)
                    else:
                        classificar_trabalhista(pagina, doc, name)
                else:
                    pagina_invalida = page.dropna(how="all")
                    for _, row in pagina_invalida.iterrows():
                        rejeitar(row, outros, doc, name, "Colunas faltantes")
        else:
            pagina = page.dropna(how="all")
            for _, row in pagina.iterrows():
                encerradas.append(row)
            
            
    except Exception as e:
        print(f"Erro na aba '{name}': {e}")
