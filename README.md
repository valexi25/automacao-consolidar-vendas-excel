# 🚀 Automação: Consolidador de Relatórios de Vendas

Este script em Python automatiza o processo de consolidar múltiplos arquivos de vendas em formato CSV, que estão em uma pasta, em um único relatório final em formato Excel.

## 🎯 O Problema Resolvido

Empresas frequentemente geram relatórios diários, semanais ou por loja em arquivos separados. Consolidar esses dados manualmente em uma única planilha é um trabalho repetitivo e sujeito a erros. Esta ferramenta automatiza completamente essa tarefa, economizando tempo e garantindo a precisão dos dados.

## ✨ Funcionalidades

-   **Leitura em Lote:** Lê todos os arquivos `.csv` localizados em uma subpasta chamada `bases/`.
-   **Consolidação de Dados:** Utiliza a biblioteca Pandas para unir (concatenar) todas as tabelas de vendas em um único DataFrame.
-   **Tratamento de Datas:** Converte as datas de um formato numérico específico do Excel para o formato padrão de data.
-   **Organização:** Ordena o relatório final por data e re-ajusta o índice para uma visualização limpa.
-   **Exportação para Excel:** Salva a tabela consolidada e organizada em um arquivo `.xlsx`.

## 🛠️ Tecnologias Utilizadas

-   Python 3
-   **Pandas**: Para toda a manipulação e consolidação dos dados.
-   **os**: Módulo nativo para interagir com o sistema de arquivos.
-   **pywin32**: (Opcional, para a parte de envio de e-mail) Para automação com aplicações Windows como o Outlook.

## ⚙️ Instalação

Para executar este projeto, você precisa instalar as bibliotecas necessárias. Abra seu terminal e rode o comando:
```bash
pip install pandas openpyxl pywin32

## 🚀 Como Usar

1.  Clone este repositório para o seu computador.
2.  Instale as dependências com o comando `pip` acima.
3.  Na pasta principal do projeto, crie uma subpasta chamada `bases`.
4.  Coloque todos os seus arquivos CSV de vendas que deseja consolidar dentro desta pasta `bases`.
5.  Abra o script `consolidador_vendas.py` e altere o e-mail do destinatário na linha `email.To = "destinatario@email.com"`.
6.  Execute o script principal no seu terminal:
    ```bash
    python consolidador_vendas.py
    ```
7.  Aguarde a execução. Ao final, um novo arquivo chamado `Vendas.xlsx` será criado na pasta principal e um e-mail com este arquivo em anexo será enviado.
