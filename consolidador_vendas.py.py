import os
import pandas as pd
import win32com.client as win32
from datetime import datetime
caminho = "bases/"
arquivos = os.listdir(caminho)
tabela_consolidada = pd.DataFrame()
for nome_arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho,nome_arquivo))
    tabela_vendas["Data de Venda"]= pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"],unit="d")
    tabela_consolidada = pd.concat([tabela_consolidada,tabela_vendas])
    #print(tabela_vendas)

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
tabela_consolidada = tabela_consolidada.reset_index(drop=True)
tabela_consolidada.to_excel("Vendas.xlsx",index=False)

data_de_hoje = datetime.today().strftime("%d/%m/%Y")

print(data_de_hoje)