import pandas as pd
import random as rd

nomes = ["Ryan", "Junior", "Alisson", "Samuel", "Djeison"]
motivos = ["Manutenção", "Ajustes", "Materiais", "Defesa", "Auditoria", "Mensalidades"]
rows = 10


for i in range(8):
    with pd.ExcelWriter(f"planilhas/Planilha_{i+1}.xlsx") as wr:
        for j in range(4):
            df = pd.DataFrame(
                {
                    "Nomes": [rd.choice(nomes) for i in range(rows)],
                    "Valor gasto": [round(rd.random() * 20000 + 2000, 2) for i in range(rows)],
                    "Motivos": [rd.choice(motivos) for i in range(rows)],
                }
            ) 
        
            try:
                df.to_excel(wr, sheet_name=f"Pagina_{j+1}", index=False)
            except Exception as e:
                print(f"{"[ERRO]":<10} {e}")

