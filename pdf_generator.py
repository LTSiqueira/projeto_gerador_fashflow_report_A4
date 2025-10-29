"""
Módulo para geração de relatórios PDF de Fluxo de Caixa
Usa Jinja2 para templates HTML e WeasyPrint para conversão PDF

Autor: Claude
Data: 28/10/2025
"""

import os
import sys
import warnings

# Configurar ambiente para WeasyPrint no Windows (ANTES de importar weasyprint)
# Resolve o erro "Fontconfig error: Cannot load default config file"
if sys.platform == 'win32':
    # Suprimir warnings de Fontconfig que não afetam a funcionalidade
    warnings.filterwarnings('ignore', category=UserWarning, module='weasyprint')

    # Configurar variáveis de ambiente para Fontconfig
    # Criar um diretório temporário para cache do Fontconfig
    fontconfig_cache = os.path.join(os.path.dirname(__file__), '.fontconfig_cache')
    os.makedirs(fontconfig_cache, exist_ok=True)

    # Configurar variáveis de ambiente
    os.environ.setdefault('FONTCONFIG_PATH', fontconfig_cache)
    os.environ.setdefault('FONTCONFIG_FILE', os.path.join(fontconfig_cache, 'fonts.conf'))

    # Criar um arquivo fonts.conf básico se não existir
    fonts_conf_path = os.path.join(fontconfig_cache, 'fonts.conf')
    if not os.path.exists(fonts_conf_path):
        fonts_conf_content = '''<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <dir>C:/Windows/Fonts</dir>
  <cachedir>FONTCONFIG_CACHE</cachedir>
</fontconfig>
'''.replace('FONTCONFIG_CACHE', fontconfig_cache.replace('\\', '/'))

        with open(fonts_conf_path, 'w', encoding='utf-8') as f:
            f.write(fonts_conf_content)

import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML, CSS
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime


class CashFlowPDFGenerator:
    """
    Gerador modular de relatórios PDF de fluxo de caixa
    Separa lógica de negócio (Python) de apresentação (HTML/CSS)
    """

    def __init__(self, template_dir: str = 'templates'):
        """
        Inicializa o gerador com o diretório de templates

        Args:
            template_dir: Diretório onde estão os templates Jinja2
        """
        self.template_dir = Path(template_dir)
        self.env = Environment(
            loader=FileSystemLoader(str(self.template_dir)),
            autoescape=True
        )

        # Registrar filtros customizados
        self.env.filters['format_currency'] = self._format_currency
        self.env.filters['format_currency_accounting'] = self._format_currency_accounting
        self.env.filters['format_date'] = self._format_date

    @staticmethod
    def _format_currency(value: float) -> str:
        """Formata valor monetário para padrão brasileiro"""
        if pd.isna(value):
            return "R$ 0,00"
        return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    @staticmethod
    def _format_currency_accounting(value: float) -> Dict[str, str]:
        """
        Formata valor monetário em estilo contábil (sinal, moeda e valor separados)
        Retorna dicionário com 'sign', 'currency' e 'value' para alinhamento perfeito
        """
        if pd.isna(value):
            return {'sign': '', 'currency': 'R$', 'value': '0,00'}

        # Determinar sinal
        sign = '-' if value < 0 else ''

        # Formatar valor absoluto no padrão brasileiro
        abs_value = abs(value)
        formatted = f"{abs_value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        return {'sign': sign, 'currency': 'R$', 'value': formatted}

    @staticmethod
    def _format_date(date: pd.Timestamp) -> str:
        """Formata data para padrão brasileiro"""
        if pd.isna(date):
            return ""
        return date.strftime('%d/%m/%Y')

    @staticmethod
    def _get_saldo_status(saldo: float) -> str:
        """
        Determina o status do saldo baseado em faixas de valor

        Args:
            saldo: Valor do saldo final do dia

        Returns:
            String com a classe CSS correspondente ao status do saldo
        """
        if saldo >= 15_000_000:
            return 'saldo-excelente'  # Verde #91D9A6
        elif saldo >= 10_000_000:
            return 'saldo-bom'  # Amarelo claro #FFE295
        elif saldo > 5_000_000:
            return 'saldo-atencao'  # Laranja claro #FBBD89
        elif saldo > 0:
            return 'saldo-critico'  # Vermelho claro #F68C90
        else:
            return 'saldo-negativo'  # Vermelho forte #E73338

    @staticmethod
    def _gerar_alerta_pior_cenario(dias_data: list) -> dict:
        """
        Analisa todos os dias e retorna alerta baseado no pior cenário encontrado

        Args:
            dias_data: Lista com informações de todos os dias do relatório

        Returns:
            Dicionário com status e mensagem do alerta ou None se não houver alerta
        """
        if not dias_data:
            return None

        # Mapear prioridade dos status (quanto maior o número, pior o cenário)
        prioridade_status = {
            'saldo-excelente': 0,
            'saldo-bom': 1,
            'saldo-atencao': 2,
            'saldo-critico': 3,
            'saldo-negativo': 4
        }

        # Mensagens para cada status
        mensagens = {
            'saldo-excelente': 'Todos os saldos acima de R$ 15 Milhões',
            'saldo-bom': 'Há data(s) abaixo de R$ 15 Milhões!',
            'saldo-atencao': 'Há data(s) abaixo de R$ 10 Milhões!',
            'saldo-critico': 'Há data(s) abaixo de R$ 5 Milhões!',
            'saldo-negativo': 'Há data(s) com saldo negativo!'
        }

        # Encontrar o pior status entre todos os dias
        pior_status = 'saldo-excelente'
        pior_prioridade = 0

        for dia in dias_data:
            status_atual = dia['saldo_status']
            prioridade_atual = prioridade_status.get(status_atual, 0)

            if prioridade_atual > pior_prioridade:
                pior_prioridade = prioridade_atual
                pior_status = status_atual

        # Retornar alerta apenas se não for o melhor cenário
        if pior_status != 'saldo-excelente':
            return {
                'status': pior_status,
                'mensagem': mensagens[pior_status]
            }

        return None

    def prepare_report_data(
        self,
        df_relatorio_diario: pd.DataFrame,
        df_timeline: pd.DataFrame,
        arquivo_excel: str = None
    ) -> Dict:
        """
        Prepara dados do relatório em formato otimizado para o template

        Args:
            df_relatorio_diario: DataFrame com resumo diário
            df_timeline: DataFrame com todas as transações detalhadas
            arquivo_excel: Caminho do arquivo Excel fonte (opcional)

        Returns:
            Dicionário com dados estruturados para o template
        """
        print("\n🔄 Preparando dados para template HTML...")

        # Filtrar apenas dias com movimentação
        df_com_movimentacao = df_relatorio_diario[
            (df_relatorio_diario['ENTRADAS'] > 0) |
            (df_relatorio_diario['SAIDAS'] > 0)
        ].copy()

        dias_data = []

        for _, row in df_com_movimentacao.iterrows():
            data = row['DATA']
            data_formatada = data.strftime('%d/%m/%Y')

            # Filtrar transações deste dia
            transacoes_dia = df_timeline[df_timeline['DATA'] == data]

            # Separar entradas e saídas
            entradas = transacoes_dia[transacoes_dia['TIPO'] == 'ENTRADA'].copy()
            saidas = transacoes_dia[transacoes_dia['TIPO'] == 'SAÍDA'].copy()

            # Preparar lista de entradas
            entradas_list = []
            for _, t in entradas.iterrows():
                descricao = str(t['DESCRICAO'])
                if pd.notna(t['PED']) and str(t['PED']).strip():
                    # Adicionar prefixo "PV " para entradas (Pedido de Venda)
                    descricao = f"{descricao} | PV {t['PED']}"

                entradas_list.append({
                    'descricao': descricao,
                    'valor': float(t['VALOR'])
                })

            # Preparar lista de saídas
            saidas_list = []
            for _, t in saidas.iterrows():
                descricao = str(t['DESCRICAO'])
                if pd.notna(t['PED']) and str(t['PED']).strip():
                    # Adicionar prefixo "PC " para saídas (Pedido de Compra)
                    # Saídas gerais não têm PED, então não entram aqui
                    descricao = f"{descricao} | PC {t['PED']}"

                saidas_list.append({
                    'descricao': descricao,
                    'valor': float(t['VALOR'])
                })

            # Calcular largura máxima para entradas (incluindo total)
            max_width_entradas = 0
            if entradas_list:
                for e in entradas_list:
                    valor_fmt = f"{e['valor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    max_width_entradas = max(max_width_entradas, len(valor_fmt))
                # Verificar também o total
                total_fmt = f"{float(row['ENTRADAS']):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                max_width_entradas = max(max_width_entradas, len(total_fmt))

            # Calcular largura máxima para saídas (incluindo total)
            max_width_saidas = 0
            if saidas_list:
                for s in saidas_list:
                    valor_fmt = f"{s['valor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    max_width_saidas = max(max_width_saidas, len(valor_fmt))
                # Verificar também o total
                total_fmt = f"{float(row['SAIDAS']):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                max_width_saidas = max(max_width_saidas, len(total_fmt))

            # Dados do dia
            saldo_final_valor = float(row['SALDO_FINAL'])
            dia_info = {
                'data': data_formatada,
                'dia_semana': data.strftime('%A'),
                'saldo_final': saldo_final_valor,
                'saldo_status': self._get_saldo_status(saldo_final_valor),
                'entradas': entradas_list,
                'saidas': saidas_list,
                'total_entradas': float(row['ENTRADAS']),
                'total_saidas': float(row['SAIDAS']),
                'tem_movimentacao': len(entradas_list) > 0 or len(saidas_list) > 0,
                'max_width_entradas': max_width_entradas,
                'max_width_saidas': max_width_saidas
            }

            dias_data.append(dia_info)

        # Obter saldo base inicial (primeira linha do relatório diário)
        saldo_base = float(df_relatorio_diario.iloc[0]['SALDO_BANCARIO'])

        # Detectar o pior cenário de saldo no relatório
        alerta = self._gerar_alerta_pior_cenario(dias_data)

        # Obter data de atualização (modificação do arquivo Excel, se fornecido)
        if arquivo_excel and os.path.exists(arquivo_excel):
            # Pegar timestamp de modificação do arquivo
            timestamp_modificacao = os.path.getmtime(arquivo_excel)
            data_atualizacao = datetime.fromtimestamp(timestamp_modificacao)
            data_formatada = data_atualizacao.strftime('%d/%m/%Y %H:%Mh')
        else:
            # Fallback para data atual
            data_formatada = datetime.now().strftime('%d/%m/%Y %H:%Mh')

        # Dados gerais do relatório
        report_data = {
            'titulo': 'Report Cashflow detalhado (A4)',
            'subtitulo': 'analise dia a dia',
            'saldo_base': saldo_base,
            'data_geracao': data_formatada,
            'periodo_inicio': df_com_movimentacao['DATA'].min().strftime('%d/%m/%Y'),
            'periodo_fim': df_com_movimentacao['DATA'].max().strftime('%d/%m/%Y'),
            'total_dias': len(dias_data),
            'dias': dias_data,
            'alerta': alerta
        }

        print(f"   ✅ {len(dias_data)} dias com movimentação preparados")
        return report_data

    def generate_html(self, report_data: Dict, template_name: str = 'cashflow_report.html') -> str:
        """
        Gera HTML a partir dos dados e template Jinja2

        Args:
            report_data: Dados preparados do relatório
            template_name: Nome do arquivo template

        Returns:
            String com HTML completo renderizado
        """
        print(f"\n📝 Gerando HTML a partir do template '{template_name}'...")

        template = self.env.get_template(template_name)
        html_content = template.render(**report_data)

        print("   ✅ HTML gerado com sucesso")
        return html_content

    def html_to_pdf(
        self,
        html_content: str,
        output_path: str,
        custom_css: Optional[str] = None
    ) -> str:
        """
        Converte HTML para PDF usando WeasyPrint

        Args:
            html_content: String com HTML completo
            output_path: Caminho do arquivo PDF de saída
            custom_css: CSS adicional (opcional)

        Returns:
            Caminho do arquivo PDF gerado
        """
        print(f"\n📄 Convertendo HTML para PDF...")

        # Verificar se o arquivo existe e tentar deletá-lo
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
                print(f"   🗑️  Arquivo PDF existente removido")
            except PermissionError:
                print(f"\n❌ ERRO: O arquivo PDF está aberto em outro programa!")
                print(f"   📁 Arquivo: {output_path}")
                print(f"   💡 Solução: Feche o arquivo PDF e execute novamente")
                raise PermissionError(f"Não foi possível acessar o arquivo. Por favor, feche-o: {output_path}")

        # Criar objeto HTML
        html_obj = HTML(string=html_content, base_url=str(self.template_dir))

        # CSS adicional se fornecido
        stylesheets = []
        if custom_css:
            stylesheets.append(CSS(string=custom_css))

        # Gerar PDF com tratamento de erro
        try:
            html_obj.write_pdf(
                output_path,
                stylesheets=stylesheets
            )
        except PermissionError as e:
            print(f"\n❌ ERRO: Não foi possível salvar o arquivo PDF!")
            print(f"   📁 Arquivo: {output_path}")
            print(f"   💡 Solução: Feche o arquivo se estiver aberto e execute novamente")
            raise PermissionError(f"Não foi possível salvar o PDF. Por favor, feche-o se estiver aberto: {output_path}") from e

        print(f"   ✅ PDF gerado: {output_path}")
        return output_path

    def generate_pdf_report(
        self,
        df_relatorio_diario: pd.DataFrame,
        df_timeline: pd.DataFrame,
        output_path: str,
        template_name: str = 'cashflow_report.html',
        arquivo_excel: str = None
    ) -> str:
        """
        Método principal: Gera relatório PDF completo

        Args:
            df_relatorio_diario: DataFrame com resumo diário
            df_timeline: DataFrame com transações detalhadas
            output_path: Caminho do arquivo PDF de saída
            template_name: Nome do template a usar
            arquivo_excel: Caminho do arquivo Excel fonte (opcional)

        Returns:
            Caminho do arquivo PDF gerado
        """
        print("\n" + "="*100)
        print("🚀 INICIANDO GERAÇÃO DE RELATÓRIO PDF")
        print("="*100)

        # 1. Preparar dados
        report_data = self.prepare_report_data(df_relatorio_diario, df_timeline, arquivo_excel)

        # 2. Gerar HTML
        html_content = self.generate_html(report_data, template_name)

        # 3. Salvar HTML temporário (útil para debug)
        html_temp_path = output_path.replace('.pdf', '_debug.html')

        # Verificar se o arquivo HTML existe e tentar deletá-lo
        if os.path.exists(html_temp_path):
            try:
                os.remove(html_temp_path)
            except PermissionError:
                print(f"\n⚠️  AVISO: Não foi possível remover HTML de debug: {html_temp_path}")
                print(f"   💡 O arquivo pode estar aberto. Continuando...")

        try:
            with open(html_temp_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"   💾 HTML de debug salvo: {html_temp_path}")
        except PermissionError:
            print(f"\n⚠️  AVISO: Não foi possível salvar HTML de debug: {html_temp_path}")
            print(f"   💡 O arquivo pode estar aberto. Continuando com geração do PDF...")

        # 4. Converter para PDF
        pdf_path = self.html_to_pdf(html_content, output_path)

        print("\n" + "="*100)
        print("✅ RELATÓRIO PDF GERADO COM SUCESSO")
        print("="*100)

        return pdf_path


def example_usage():
    """Exemplo de uso do gerador"""
    import pandas as pd

    # Dados de exemplo (substituir por dados reais)
    df_relatorio = pd.DataFrame({
        'DATA': pd.date_range('2025-01-01', periods=10),
        'ENTRADAS': [1000, 0, 2000, 0, 1500, 0, 0, 3000, 0, 1000],
        'SAIDAS': [500, 0, 1000, 0, 800, 0, 0, 1500, 0, 600],
        'SALDO_FINAL': [10500, 10500, 11500, 11500, 12200, 12200, 12200, 13700, 13700, 14100]
    })

    df_timeline = pd.DataFrame({
        'DATA': pd.date_range('2025-01-01', periods=5),
        'DESCRICAO': ['Cliente A', 'Fornecedor B', 'Cliente C', 'Fornecedor D', 'Cliente E'],
        'PED': ['123', '456', '789', '101', '102'],
        'TIPO': ['ENTRADA', 'SAÍDA', 'ENTRADA', 'SAÍDA', 'ENTRADA'],
        'VALOR': [1000, 500, 2000, 1000, 1500]
    })

    # Gerar PDF
    generator = CashFlowPDFGenerator()
    generator.generate_pdf_report(
        df_relatorio,
        df_timeline,
        'relatorio_cashflow.pdf'
    )


if __name__ == "__main__":
    example_usage()
