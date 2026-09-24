# Control 📊

[English](#english) | [Português](#português)

---

## English

**Control** is a Python-based Command Line Interface (CLI) tool designed to automate the extraction, categorization, and consolidation of bank statements in PDF format into formatted Excel spreadsheets, featuring visual dashboards for budgeting and financial analysis.

### 🚀 Features

* **Automated PDF Parsing:** Extracts dates, descriptions, and amounts (credits and debits) directly from PDF bank statements.
* **Keyword-based Categorization:** Automatically maps transactions into predefined categories (e.g., Groceries, Transport, Personal Care, Utilities, Subscriptions, Investments).
* **Subsequent Month Rolling:** Generates the spreadsheet file titled with the subsequent month relative to the statement period (optimized for billing cycles and monthly planning).
* **Consolidated Excel Dashboard:**
  * Raw transaction register with color-coded categories.
  * Executive summary cards (Total Incomes, Total Expenses, Net Balance).
  * Side-by-side category-specific transaction breakdown.

### 🛠️ Prerequisites & Installation

1. Clone the repository:
```bash
git clone https://github.com/YOUR_USERNAME/control.git
cd control
```

2. Create and activate a virtual environment (recommended):
```bash
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. Install required libraries:
```bash
pip install -r requirements.txt
```

*(If you have not created `requirements.txt` yet, install via: `pip install pdfplumber pandas openpyxl`)*

### 💻 Usage

#### Automatic mode (parses the most recently added/updated PDF in the current folder):
```bash
python control.py
```

#### Specific file mode:
```bash
python control.py path/to/statement.pdf
```

#### Optional CLI alias (Linux/macOS):
To run the command directly from anywhere as `control`:
```bash
chmod +x control.py
echo 'alias control="python3 $(pwd)/control.py"' >> ~/.bashrc
source ~/.bashrc
```

### ⚙️ Customization

You can customize categories, search keywords, and color palettes by editing the `REGRAS_CATEGORIAS` and `CORES_CATEGORIAS` dictionaries directly inside `control.py`:

```python
REGRAS_CATEGORIAS = {
    'Alimentação': ['mercado', 'supermercado', 'restaurante', 'ifood', ...],
    'Transporte': ['uber', '99', 'posto', 'gasolina', ...]
}
```

### 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Português

O **Control** é uma ferramenta de linha de comando (CLI) desenvolvida em Python para automatizar a leitura, categorização e consolidação de extratos bancários em formato PDF, gerando planilhas Excel estruturadas com painéis visuais para controle financeiro e planejamento orçamentário.

### 🚀 Funcionalidades

* **Extração Automatizada de PDF:** Identifica lançamentos, datas, descrições e valores (créditos e débitos) diretamente de extratos bancários.
* **Categorização Inteligente:** Mapeia transações automaticamente por palavras-chave (ex: Alimentação, Transporte, Cuidados Pessoais, Contas, Lazer, Investimentos).
* **Projeção para Mês Posterior:** Nomeia o relatório consolidado com base no mês seguinte ao período do extrato (ideal para fechamento de faturas e planejamento de gastos).
* **Dashboard em Excel:**
  * Aba com o extrato completo formatado e colorido por categoria.
  * Aba com resumo executivo (Total Recebimentos, Total Gastos, Saldo Líquido).
  * Tabela consolidada lado a lado por setor de despesa.

### 🛠️ Pré-requisitos e Instalação

1. Clone o repositório:
```bash
git clone https://github.com/SEU_USUARIO/control.git
cd control
```

2. Crie e ative um ambiente virtual (recomendado):
```bash
python3 -m venv .venv
source .venv/bin/activate  # No Windows: .venv\Scripts\activate
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

*(Caso ainda não tenha o arquivo `requirements.txt`, instale via: `pip install pdfplumber pandas openpyxl`)*

### 💻 Como Usar

#### Modo Automático (busca o PDF mais recente na pasta atual):
```bash
python control.py
```

#### Modo Específico (informando o caminho do extrato):
```bash
python control.py caminho/para/seu_extrato.pdf
```

#### Criando atalho no terminal (Linux/macOS):
Para executar em qualquer lugar apenas digitando `control`:
```bash
chmod +x control.py
echo 'alias control="python3 $(pwd)/control.py"' >> ~/.bashrc
source ~/.bashrc
```

### ⚙️ Personalização de Regras

Você pode adaptar as categorias, termos de busca e cores editando diretamente os dicionários `REGRAS_CATEGORIAS` e `CORES_CATEGORIAS` dentro de `control.py`:

```python
REGRAS_CATEGORIAS = {
    'Alimentação': ['mercado', 'supermercado', 'restaurante', 'ifood', ...],
    'Transporte': ['uber', '99', 'posto', 'gasolina', ...]
}
```

### 📄 Licença

Este projeto está sob a licença MIT - consulte o arquivo [LICENSE](LICENSE) para mais detalhes.