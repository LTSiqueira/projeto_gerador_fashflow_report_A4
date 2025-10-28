# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based cash flow report generator that processes Excel financial data and generates daily cash flow reports. The system reads from a multi-sheet Excel workbook containing bank balances, accounts receivable, accounts payable, and general expenses, then outputs a consolidated daily report with projected balances.

## Running the Application

### Full Processing (Excel + PDF)
```bash
python gerador.py
```

### Test PDF Generation Only
```bash
python test_pdf_generation.py
```

### Install Dependencies
```bash
pip install -r requirements.txt
```

**Input**: `CashFlow Financeiro_new.xlsx`

**Outputs**:
- `relatorio_fluxo_caixa_completo.xlsx` - Excel with 3 sheets
- `relatorio_fluxo_caixa.pdf` - Professional PDF report
- `relatorio_fluxo_caixa_debug.html` - Intermediate HTML (useful for debugging)

**Note**: File paths are hardcoded in `main()` function at gerador.py:477-479. Update as needed for different environments.

## Architecture

### Core Component: FluxoCaixaProcessor

The main processing class (`FluxoCaixaProcessor`) orchestrates the entire cash flow analysis pipeline. It follows a sequential processing model:

1. **Extract Bank Balances** (`extrair_todos_saldos()`) - Reads the 'SALDO BANCÁRIO - R$' sheet and extracts all historical bank balance data across all dates
2. **Extract Receivables** (`extrair_contas_receber()`) - Processes 'CR - Produto' sheet for incoming payments
3. **Extract Product Payables** (`extrair_contas_pagar_produtos()`) - Processes 'CP - Produto' sheet for product-related expenses
4. **Extract General Expenses** (`extrair_saidas_gerais()`) - Processes 'CP - Saídas Gerais' sheet and aggregates by date
5. **Create Timeline** (`criar_timeline()`) - Combines all transaction types into a chronological timeline
6. **Generate Daily Report** (`gerar_relatorio_diario()`) - Creates day-by-day projections starting from the most recent bank balance

### Key Processing Logic

**Balance Projection Methodology** (gerador.py:244-330):
- The system uses the **most recent bank balance date** as the starting point for projections
- It filters transactions to include only those on or after this date
- Daily balances are calculated progressively: `Saldo[i] = Saldo[i-1] + Entradas[i] - Saídas[i]`
- The projection creates a continuous daily timeline from the most recent balance date to the last transaction date

**Excel Sheet Structure Requirements**:
- 'SALDO BANCÁRIO - R$': Row 0 contains dates in 'dd/mm/yyyy' format, columns 2+ contain balance data. Rows with "TOTAL" in columns 0 or 1 are excluded from summation
- 'CR - Produto': Header at row 6, data starts row 8. Columns: PED, CLIENTE, VENCIMENTO, VLR A RECEBER R$
- 'CP - Produto': Header at row 6, data starts row 8. Columns: PED, FORNECEDOR, VENCIMENTO, VLR R$
- 'CP - Saídas Gerais': Header at row 6, data starts row 8. Columns: DATA VENC., VALOR A PAGAR R$

### Output Structure

The Excel output contains three sheets:
1. **Relatório Diário**: Daily summary with bank balance, entries, exits, and projected final balance
2. **Timeline Detalhada**: Chronological list of all individual transactions with categories
3. **Histórico Saldos**: Historical bank balance data extracted from the source

## PDF Generation Architecture

### Module: `pdf_generator.py`

**Class: `CashFlowPDFGenerator`**

Implements a modular PDF generation system using **Jinja2 + WeasyPrint** stack:

**Why this stack?**
- **Jinja2**: Separates business logic (Python) from presentation (HTML/CSS)
- **WeasyPrint**: Native CSS Paged Media support (@page, page-break-inside, etc.), maintains design fidelity without browser overhead

**Key Methods**:
- `prepare_report_data()`: Transforms DataFrames into template-friendly dict structure with nested days/transactions
- `generate_html()`: Renders Jinja2 template with data, outputs HTML string
- `html_to_pdf()`: Converts HTML to PDF using WeasyPrint, handles stylesheets
- `generate_pdf_report()`: Main orchestrator method

**Custom Jinja2 Filters**:
- `format_currency`: Formats floats to Brazilian currency (R$ 1.234,56)
- `format_date`: Formats timestamps to dd/mm/yyyy

### Template: `templates/cashflow_report.html`

Production Jinja2 template with:
- **CSS Paged Media**: @page rules for A4 pagination, page counters, automatic page breaks
- **Timeline Design**: Green dots (::before) and vertical lines (::after) using CSS pseudo-elements
- **Two-Column Layout**: Grid-based Entradas/Saídas with dotted leaders between description and value
- **Pagination Control**: `page-break-inside: avoid` on `.day-section` to prevent splitting days across pages
- **Dynamic Content**: Loops through `{% for dia in dias %}` with conditional rendering for empty transaction lists

### Integration with FluxoCaixaProcessor

`gerar_relatorio_pdf()` method (gerador.py:384-410) bridges the processor with PDF generator:
1. Validates that timeline and daily report DataFrames exist
2. Instantiates `CashFlowPDFGenerator` with templates directory
3. Calls `generate_pdf_report()` passing both DataFrames
4. Outputs PDF and debug HTML

### Static Example Template

`example_initial_template.html` was the initial design prototype that informed the final Jinja2 template. It demonstrates the desired visual style but is not used in production - the actual template is `templates/cashflow_report.html`.

## Dependencies

Required Python packages (see requirements.txt):
- **Data Processing**: pandas, numpy, openpyxl
- **PDF Generation**: Jinja2, WeasyPrint
- **WeasyPrint deps**: Pillow, pycairo, cffi

**Important**: WeasyPrint on Windows requires GTK3 runtime. If installation fails, see README.md troubleshooting section.

## Data Flow

```
Excel Input (CashFlow Financeiro_new.xlsx)
  ├─ SALDO BANCÁRIO - R$ → extrair_todos_saldos()
  ├─ CR - Produto → extrair_contas_receber()
  ├─ CP - Produto → extrair_contas_pagar_produtos()
  └─ CP - Saídas Gerais → extrair_saidas_gerais()
                            ↓
                    criar_timeline()
                            ↓
                  gerar_relatorio_diario()
                            ↓
                    ┌───────┴───────┐
                    ↓               ↓
    exportar_relatorio_completo()  gerar_relatorio_pdf()
                    ↓               ↓
                    ↓          CashFlowPDFGenerator
                    ↓               ├─ prepare_report_data()
                    ↓               ├─ generate_html() (Jinja2)
                    ↓               └─ html_to_pdf() (WeasyPrint)
                    ↓               ↓
    Excel Output            PDF Output + HTML Debug
```
