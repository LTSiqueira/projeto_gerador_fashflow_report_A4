# Gerador de Relatório de Fluxo de Caixa

Sistema automatizado para processar planilhas Excel de fluxo de caixa e gerar relatórios profissionais em PDF e Excel.

## Características

- Processa múltiplas abas de fluxo de caixa (Saldo Bancário, Contas a Receber, Contas a Pagar)
- Gera relatório diário com projeção de saldos
- Exporta para Excel com múltiplas abas
- **Gera PDF profissional com design timeline** (A4, paginado)
- Arquitetura modular e extensível

## Instalação

### 1. Instalar Python 3.9+

Certifique-se de ter Python 3.9 ou superior instalado.

### 2. Instalar Dependências

```bash
pip install -r requirements.txt
```

### 3. Dependências do WeasyPrint (Windows)

O WeasyPrint requer GTK3. No Windows, você tem duas opções:

**Opção A: Usar versão standalone**
```bash
pip install WeasyPrint[all]
```

**Opção B: Instalar GTK3 manualmente**
1. Baixe GTK3: https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases
2. Execute o instalador
3. Adicione o path do GTK ao PATH do sistema

## Uso Básico

### Executar processamento completo

```bash
python gerador.py
```

Isso irá:
1. Processar o arquivo Excel de entrada
2. Gerar relatório diário
3. Exportar para Excel
4. **Gerar PDF profissional**

### Arquivos gerados

- `relatorio_fluxo_caixa_completo.xlsx` - Relatório Excel com 3 abas
- `relatorio_fluxo_caixa.pdf` - Relatório PDF com design profissional
- `relatorio_fluxo_caixa_debug.html` - HTML intermediário (útil para debug)

## Estrutura do Projeto

```
projeto_gerador_fashflow_report_A4/
├── gerador.py                      # Script principal com FluxoCaixaProcessor
├── pdf_generator.py                # Módulo de geração de PDF
├── templates/
│   └── cashflow_report.html        # Template Jinja2 do relatório
├── requirements.txt                # Dependências Python
├── CashFlow Financeiro_new.xlsx   # Arquivo de entrada (exemplo)
└── README.md                       # Este arquivo
```

## Arquitetura

### Módulo Principal: `gerador.py`

Classe `FluxoCaixaProcessor` que orquestra todo o processamento:

- `extrair_todos_saldos()` - Extrai saldos bancários históricos
- `extrair_contas_receber()` - Processa contas a receber
- `extrair_contas_pagar_produtos()` - Processa contas a pagar de produtos
- `extrair_saidas_gerais()` - Processa saídas gerais agregadas
- `criar_timeline()` - Cria linha do tempo de transações
- `gerar_relatorio_diario()` - Gera relatório diário com projeções
- `exportar_relatorio_completo()` - Exporta para Excel
- `gerar_relatorio_pdf()` - **Gera PDF profissional**

### Módulo de PDF: `pdf_generator.py`

Classe `CashFlowPDFGenerator` responsável pela geração de PDF:

- `prepare_report_data()` - Prepara dados para o template
- `generate_html()` - Gera HTML a partir do template Jinja2
- `html_to_pdf()` - Converte HTML para PDF usando WeasyPrint
- `generate_pdf_report()` - Método principal de geração

### Template: `templates/cashflow_report.html`

Template Jinja2 com:
- Design profissional em timeline
- Paginação automática (CSS Paged Media)
- Duas colunas (Entradas/Saídas)
- Bolinha verde e linha vertical estilo timeline
- Suporte a múltiplas páginas A4

## Customização

### Modificar design do PDF

Edite o arquivo `templates/cashflow_report.html`:
- CSS está incorporado no `<style>`
- Use `@page` para configurar páginas
- Use `page-break-inside: avoid` para controlar paginação

### Modificar lógica de processamento

Edite o arquivo `gerador.py`:
- Ajuste os métodos `extrair_*()` para suas necessidades
- Modifique `gerar_relatorio_diario()` para alterar cálculos

### Adicionar novos templates

1. Crie novo template em `templates/`
2. Use filtros Jinja2: `{{ valor|format_currency }}`, `{{ data|format_date }}`
3. Chame com `generator.generate_pdf_report(..., template_name='seu_template.html')`

## Solução de Problemas

### Erro ao instalar WeasyPrint no Windows

**Problema**: `OSError: cannot load library 'gobject-2.0-0'`

**Solução**: Instale GTK3 Runtime ou use:
```bash
pip install --force-reinstall WeasyPrint
```

### PDF não está sendo gerado

1. Verifique se o HTML intermediário foi criado (`*_debug.html`)
2. Abra o HTML no navegador para verificar o layout
3. Verifique logs de erro do WeasyPrint

### Layout quebrado no PDF

1. Verifique o HTML de debug
2. Ajuste CSS no template
3. Use `page-break-inside: avoid` para evitar quebras indesejadas

### Fontes não aparecem no PDF

WeasyPrint baixa fontes do Google Fonts automaticamente. Certifique-se de ter conexão com internet na primeira execução.

## Stack Tecnológica

- **Python 3.9+** - Linguagem base
- **Pandas** - Processamento de dados
- **Jinja2** - Template engine para HTML
- **WeasyPrint** - Conversão HTML → PDF
- **OpenPyXL** - Leitura/escrita de Excel

## Por que Jinja2 + WeasyPrint?

✅ **Separação de responsabilidades**: Lógica (Python) separada de design (HTML/CSS)
✅ **CSS Paged Media**: Suporte nativo a paginação, quebras de página, cabeçalhos
✅ **Fácil manutenção**: Designers podem editar templates sem tocar no código
✅ **Qualidade**: Mantém fidelidade do design HTML
✅ **Leve**: Não precisa de browser headless

## Licença

Código gerado por Claude (Anthropic) - 2025
