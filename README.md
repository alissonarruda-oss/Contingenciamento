## Contexto

Este é um sistema que percorre pastas contendo arquivos `.xlsx` enviados por escritórios de advocacia, identifica o tipo de cada ação (cível ou trabalhista), a entidade envolvida (Banco, Hipotecária, Securitizadora, Service ou Promotora) e a posição processual (ativa ou passiva). Ao final, gera arquivos Excel consolidados com formatação visual para facilitar a revisão.

---

## Estrutura do Projeto

```
Contingenciamento/
├── main.py                  # Ponto de entrada: orquestra leitura e exportação
├── dados.xlsx               # Aba utilizada como template de dados da planilha consolidada
├── services/
│   ├── constants.py         # Cabeçalhos esperados, colunas obrigatórias e estruturas de dados globais
│   ├── processor.py         # Lógica de validação e classificação de cada aba/linha
│   └── path_logic.py        # Busca de arquivos e geração dos relatórios Excel
└── arquivos/                  # Diretório onde devem ser colocados os arquivos de entrada
```

---

## Entrada

- **Localização:** subpastas dentro do diretório raiz do projeto (arquivos na raiz são ignorados).
- **Formato:** arquivos `.xlsx` com abas de processos judiciais.
- **Abas ignoradas automaticamente:** `GRÁFICOS`, `ANALYTICS`, `DADOS`, `TESTE`, `RELATÓRIO`.
- **Detecção do tipo de aba:**
  - Aba com coluna `ENCERRAMENTO` → ações **encerradas**
  - Aba com coluna `DEPÓSITOS RECLAMANTE` → ação **trabalhista**
  - Demais abas com colunas essenciais → ação **cível**

### Classificação cível (por palavra-chave na Parte Autora / Parte Ré)

| Palavra-chave | Entidade        | Posição   |
|---------------|-----------------|-----------|
| `BANCO`       | Banco           | Ativa/Passiva |
| `HIPOTECÁRIA` | Hipotecária     | Ativa/Passiva |
| `SECURITIZADORA` | Securitizadora | Ativa/Passiva |

### Classificação trabalhista (por palavra-chave na Parte Ré)

| Palavra-chave | Entidade    |
|---------------|-------------|
| `BANCO`       | Banco       |
| `SERVICE`     | Service     |
| `PROMOTORA`   | Promotora   |
| `HIPOTECÁRIA` | Hipotecária |

---

## Saída

Ao final da execução, os seguintes arquivos são gerados na raiz do projeto:

| Arquivo | Conteúdo |
|---|---|
| `CONSOLIDADO - HIPOTECÁRIA_BANCO_SEC.xlsx` | Ações cíveis de Banco, Hipotecária e Securitizadora (ativas e passivas), com a aba `DADOS` original |
| `TRABALHISTA_CONSOLIDADO - SERVICE e PROMOTORA.xlsx` | Ações trabalhistas de Service e Promotora |
| `TRABALHISTA_CONSOLIDADO - BANCO E HIPO.xlsx` | Ações trabalhistas de Banco e Hipotecária |
| `VERIFICAR_OUTROS.xlsx` | Registros que não puderam ser classificados, separados por escritório |
| `ENCERRADAS.xlsx` | Processos encerrados identificados nas planilhas |

---

## Validação de Registros

Cada linha passa por validação antes de ser aceita:

1. **Quantidade de colunas** deve corresponder ao cabeçalho esperado para a entidade
2. **Campos obrigatórios** não podem estar vazios: `ESCRITÓRIO`, `PARTE AUTORA`, `PARTE RÉ`, `NÚMERO DO PROCESSO`, `PRODUTO`, `VALOR DA CAUSA`, `VALOR DO RISCO ATUALIZADO`, `PROBABILIDADE DE PERDA`
3. **Valores monetários** (`VALOR DA CAUSA`, `VALOR DO RISCO ATUALIZADO`) devem ser conversíveis para número (prefixo `R$` é removido automaticamente)

Linhas rejeitadas vão para `VERIFICAR_OUTROS.xlsx` com o motivo anotado (`Colunas faltantes`, `Informações não preenchidas`, `Valor não monetário`, `Quantidade de colunas incorretas`).

---

## Como Executar

1. Coloque os arquivos `.xlsx` de entrada dentro de subpastas (ex.: `pastas/escritorio_x/arquivo.xlsx`)
2. Certifique-se de que o arquivo `dados.xlsx` existe na raiz (usado como template do consolidado)
3. Execute:

```bash
python main.py
```

O progresso é exibido no terminal com percentual e tempo decorrido. Ao final, uma mensagem confirma a exportação com o tempo total de execução.

---

## Dependências

| Pacote | Versão |
|--------|--------|
| pandas | 2.3.3 |
| openpyxl | 3.1.5 |
| arrow | 1.4.0 |

As dependências estão declaradas em [requirements.txt](requirements.txt). Instale com:

```bash
pip install -r requirements.txt
```
