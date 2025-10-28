"""
Script para processar Excel de Fluxo de Caixa e gerar linha do tempo detalhada
Autor: Claude
Data: 28/10/2025

FUNCIONALIDADES:
- Extrai saldos bancários de cada data
- Processa contas a receber (CR - Produto)
- Processa contas a pagar produtos (CP - Produto)
- Processa saídas gerais (CP - Saídas Gerais) agregadas por data
- Gera relatório detalhado por data com: entradas, saídas e saldo final
"""

# Configurar PATH para GTK3 (necessário para WeasyPrint no Windows)
import os
import sys

# Configurar encoding UTF-8 para o console do Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

gtk_path = r'C:\Program Files\GTK3-Runtime Win64\bin'
if sys.platform == 'win32' and os.path.exists(gtk_path):
    if gtk_path not in os.environ['PATH']:
        os.environ['PATH'] = gtk_path + os.pathsep + os.environ['PATH']

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Dict, List, Tuple
import warnings
warnings.filterwarnings('ignore')

# Importar gerador de PDF
from pdf_generator import CashFlowPDFGenerator


class FluxoCaixaProcessor:
    """Classe para processar dados de fluxo de caixa do Excel"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df_saldos_por_data = None
        self.df_timeline = None
        self.df_relatorio_diario = None
        
    def extrair_todos_saldos(self) -> pd.DataFrame:
        """
        Extrai TODOS os saldos bancários de TODAS as datas disponíveis
        Retorna um DataFrame com: DATA e SALDO_TOTAL
        """
        print("📊 Extraindo histórico completo de saldos bancários...")
        
        df_saldo = pd.read_excel(self.file_path, sheet_name='SALDO BANCÁRIO - R$')
        
        # Linha 0: tem as datas
        # Linha 1: tem os horários
        # Linha 2 em diante: tem os bancos e valores
        
        datas = []
        saldos = []
        
        # Iterar pelas colunas (cada coluna é uma data)
        for col_idx in range(2, len(df_saldo.columns)):  # Começar da coluna 2 (pula as 2 primeiras)
            # Verificar se é uma data válida
            valor_header = df_saldo.iloc[0, col_idx]

            if pd.notna(valor_header) and '/' in str(valor_header):
                # É uma data válida
                data = valor_header

                # Somar todos os saldos desta coluna (exceto linha com "TOTAL")
                saldo_total = 0
                for row_idx in range(2, len(df_saldo)):
                    # Verificar AMBAS as colunas (0 e 1) para detectar "TOTAL"
                    col_0 = str(df_saldo.iloc[row_idx, 0]) if pd.notna(df_saldo.iloc[row_idx, 0]) else ""
                    col_1 = str(df_saldo.iloc[row_idx, 1]) if pd.notna(df_saldo.iloc[row_idx, 1]) else ""

                    # IMPORTANTE: Ignorar linha com "TOTAL" em qualquer das colunas
                    if 'TOTAL' not in col_0.upper() and 'TOTAL' not in col_1.upper():
                        valor = df_saldo.iloc[row_idx, col_idx]
                        if pd.notna(valor):
                            try:
                                saldo_total += float(valor)
                            except:
                                pass

                datas.append(data)
                saldos.append(saldo_total)
        
        # Criar DataFrame
        df_saldos = pd.DataFrame({
            'DATA': datas,
            'SALDO_BANCARIO': saldos
        })
        
        # Converter data
        df_saldos['DATA'] = pd.to_datetime(df_saldos['DATA'], format='%d/%m/%Y', errors='coerce')
        
        # Remover datas inválidas
        df_saldos = df_saldos[df_saldos['DATA'].notna()].copy()
        
        # Ordenar por data
        df_saldos = df_saldos.sort_values('DATA').reset_index(drop=True)
        
        self.df_saldos_por_data = df_saldos
        
        print(f"   ✅ {len(df_saldos)} datas encontradas")
        print(f"   📅 Período: {df_saldos['DATA'].min().strftime('%d/%m/%Y')} até {df_saldos['DATA'].max().strftime('%d/%m/%Y')}")
        print(f"   💰 Primeiro saldo: R$ {df_saldos['SALDO_BANCARIO'].iloc[0]:,.2f}")
        print(f"   💰 Último saldo: R$ {df_saldos['SALDO_BANCARIO'].iloc[-1]:,.2f}")
        
        return df_saldos
    
    def extrair_contas_receber(self) -> pd.DataFrame:
        """
        Extrai dados da aba 'CR - Produto' (Contas a Receber)
        """
        print("\n📥 Processando Contas a Receber (CR - Produto)...")
        
        df = pd.read_excel(self.file_path, sheet_name='CR - Produto', header=6)
        
        # Primeira linha tem os headers reais
        df_headers = df.iloc[0]
        df_data = df.iloc[1:].copy()
        df_data.columns = df_headers
        
        # Selecionar colunas relevantes
        df_clean = df_data[['PED', 'CLIENTE', 'VENCIMENTO', 'VLR A RECEBER R$']].copy()
        df_clean.columns = ['PED', 'DESCRICAO', 'DATA', 'VALOR']
        
        # Filtrar registros válidos
        df_clean = df_clean[df_clean['PED'].notna()].copy()
        
        # Converter tipos
        df_clean['VALOR'] = pd.to_numeric(df_clean['VALOR'], errors='coerce')
        df_clean['DATA'] = pd.to_datetime(df_clean['DATA'], errors='coerce')
        
        # Adicionar tipo de transação
        df_clean['TIPO'] = 'ENTRADA'
        df_clean['CATEGORIA'] = 'CR - Produto'
        
        # Remover registros sem valor ou data
        df_clean = df_clean[df_clean['VALOR'].notna() & df_clean['DATA'].notna()].copy()
        
        print(f"   ✅ {len(df_clean)} registros | Total: R$ {df_clean['VALOR'].sum():,.2f}")
        return df_clean
    
    def extrair_contas_pagar_produtos(self) -> pd.DataFrame:
        """
        Extrai dados da aba 'CP - Produto' (Contas a Pagar - Produtos)
        """
        print("📤 Processando Contas a Pagar - Produtos...")
        
        df = pd.read_excel(self.file_path, sheet_name='CP - Produto', header=6)
        
        # Primeira linha tem os headers reais
        df_headers = df.iloc[0]
        df_data = df.iloc[1:].copy()
        df_data.columns = df_headers
        
        # Selecionar colunas relevantes
        df_clean = df_data[['PED', 'FORNECEDOR', 'VENCIMENTO', 'VLR R$']].copy()
        df_clean.columns = ['PED', 'DESCRICAO', 'DATA', 'VALOR']
        
        # Filtrar registros válidos
        df_clean = df_clean[df_clean['PED'].notna()].copy()
        
        # Converter tipos
        df_clean['VALOR'] = pd.to_numeric(df_clean['VALOR'], errors='coerce')
        df_clean['DATA'] = pd.to_datetime(df_clean['DATA'], errors='coerce')
        
        # Adicionar tipo de transação
        df_clean['TIPO'] = 'SAÍDA'
        df_clean['CATEGORIA'] = 'CP - Produto'
        
        # Remover registros sem valor ou data
        df_clean = df_clean[df_clean['VALOR'].notna() & df_clean['DATA'].notna()].copy()
        
        print(f"   ✅ {len(df_clean)} registros | Total: R$ {df_clean['VALOR'].sum():,.2f}")
        return df_clean
    
    def extrair_saidas_gerais(self) -> pd.DataFrame:
        """
        Extrai dados da aba 'CP - Saídas Gerais' e AGRUPA POR DATA
        """
        print("📤 Processando Saídas Gerais...")
        
        df = pd.read_excel(self.file_path, sheet_name='CP - Saídas Gerais', header=6)
        
        # Primeira linha tem os headers reais
        df_headers = df.iloc[0]
        df_data = df.iloc[1:].copy()
        df_data.columns = df_headers
        
        # Selecionar colunas relevantes
        df_clean = df_data[['DATA VENC.', 'VALOR A PAGAR R$']].copy()
        df_clean.columns = ['DATA', 'VALOR']
        
        # Filtrar registros válidos
        df_clean = df_clean[df_clean['DATA'].notna()].copy()
        
        # Converter tipos
        df_clean['VALOR'] = pd.to_numeric(df_clean['VALOR'], errors='coerce')
        df_clean['DATA'] = pd.to_datetime(df_clean['DATA'], errors='coerce')
        
        # Remover registros sem valor
        df_clean = df_clean[df_clean['VALOR'].notna()].copy()
        
        # AGRUPAR POR DATA (como solicitado pelo usuário)
        df_grouped = df_clean.groupby('DATA')['VALOR'].sum().reset_index()
        
        # Adicionar colunas padrão
        df_grouped['PED'] = ''
        df_grouped['DESCRICAO'] = 'SAÍDAS GERAIS'
        df_grouped['TIPO'] = 'SAÍDA'
        df_grouped['CATEGORIA'] = 'CP - Saídas Gerais'
        
        print(f"   ✅ {len(df_grouped)} datas únicas | Total: R$ {df_grouped['VALOR'].sum():,.2f}")
        return df_grouped
    
    def criar_timeline(self) -> pd.DataFrame:
        """
        Cria a linha do tempo completa com todas as transações
        """
        print("\n🔄 Criando timeline de transações...")
        
        # Extrair todos os dados
        df_cr = self.extrair_contas_receber()
        df_cp = self.extrair_contas_pagar_produtos()
        df_sg = self.extrair_saidas_gerais()
        
        # Combinar todos os DataFrames
        df_timeline = pd.concat([df_cr, df_cp, df_sg], ignore_index=True)
        
        # Ordenar por data
        df_timeline = df_timeline.sort_values('DATA').reset_index(drop=True)
        
        # Formatar data para exibição
        df_timeline['DATA_FORMATADA'] = df_timeline['DATA'].dt.strftime('%d/%m/%Y')
        
        self.df_timeline = df_timeline
        
        print(f"   ✅ Timeline criada com {len(df_timeline)} transações")
        print(f"   📊 Período: {df_timeline['DATA'].min().strftime('%d/%m/%Y')} até {df_timeline['DATA'].max().strftime('%d/%m/%Y')}")
        
        return df_timeline
    
    def gerar_relatorio_diario(self) -> pd.DataFrame:
        """
        Gera relatório detalhado POR DATA com:
        - Data
        - Saldo Bancário do dia (da planilha de saldos)
        - Total de Entradas do dia
        - Total de Saídas do dia
        - Saldo Final (Saldo Bancário + Entradas - Saídas acumuladas)

        IMPORTANTE: Usa a data MAIS RECENTE dos saldos como ponto de partida
        e projeta para frente com as transações.
        """
        print("\n📊 Gerando relatório diário consolidado...")

        if self.df_saldos_por_data is None:
            self.extrair_todos_saldos()

        if self.df_timeline is None:
            self.criar_timeline()

        # ===================================================================
        # MUDANÇA IMPORTANTE: Usar apenas a data MAIS RECENTE dos saldos
        # ===================================================================
        data_mais_recente_saldo = self.df_saldos_por_data['DATA'].max()
        saldo_inicial = self.df_saldos_por_data[
            self.df_saldos_por_data['DATA'] == data_mais_recente_saldo
        ]['SALDO_BANCARIO'].iloc[0]

        print(f"   📅 Data mais recente dos saldos: {data_mais_recente_saldo.strftime('%d/%m/%Y')}")
        print(f"   💰 Saldo inicial: R$ {saldo_inicial:,.2f}")

        # Filtrar transações para incluir apenas >= data mais recente
        df_timeline_filtrado = self.df_timeline[
            self.df_timeline['DATA'] >= data_mais_recente_saldo
        ].copy()

        print(f"   🔄 Transações filtradas: {len(df_timeline_filtrado)} (>= {data_mais_recente_saldo.strftime('%d/%m/%Y')})")

        # Agrupar transações por data
        df_transacoes_dia = df_timeline_filtrado.groupby('DATA').apply(
            lambda x: pd.Series({
                'ENTRADAS': x[x['TIPO'] == 'ENTRADA']['VALOR'].sum(),
                'SAIDAS': abs(x[x['TIPO'] == 'SAÍDA']['VALOR'].sum()),
                'QTD_ENTRADAS': len(x[x['TIPO'] == 'ENTRADA']),
                'QTD_SAIDAS': len(x[x['TIPO'] == 'SAÍDA'])
            })
        ).reset_index()

        # Determinar range de datas: da data mais recente até a última transação
        data_inicio = data_mais_recente_saldo
        data_fim_transacoes = df_timeline_filtrado['DATA'].max() if len(df_timeline_filtrado) > 0 else data_mais_recente_saldo
        data_fim = max(data_mais_recente_saldo, data_fim_transacoes)

        # Criar range de datas
        todas_datas = pd.date_range(start=data_inicio, end=data_fim, freq='D')
        df_relatorio = pd.DataFrame({'DATA': todas_datas})

        # Adicionar saldo inicial (apenas na primeira linha)
        df_relatorio['SALDO_BANCARIO'] = saldo_inicial

        # Merge com transações
        df_relatorio = df_relatorio.merge(df_transacoes_dia, on='DATA', how='left')

        # Preencher zeros onde não há transações
        df_relatorio['ENTRADAS'] = df_relatorio['ENTRADAS'].fillna(0)
        df_relatorio['SAIDAS'] = df_relatorio['SAIDAS'].fillna(0)
        df_relatorio['QTD_ENTRADAS'] = df_relatorio['QTD_ENTRADAS'].fillna(0).astype(int)
        df_relatorio['QTD_SAIDAS'] = df_relatorio['QTD_SAIDAS'].fillna(0).astype(int)

        # Calcular movimentação líquida diária
        df_relatorio['MOVIMENTACAO_DIA'] = df_relatorio['ENTRADAS'] - df_relatorio['SAIDAS']

        # Calcular saldo progressivo dia a dia
        # Primeiro dia: saldo inicial + movimentação do primeiro dia
        saldos = []
        saldo_atual = saldo_inicial

        for i in range(len(df_relatorio)):
            movimentacao = df_relatorio['MOVIMENTACAO_DIA'].iloc[i]
            saldo_atual = saldo_atual + movimentacao
            saldos.append(saldo_atual)

        df_relatorio['SALDO_FINAL'] = saldos

        # Formatar data
        df_relatorio['DATA_FORMATADA'] = df_relatorio['DATA'].dt.strftime('%d/%m/%Y')

        # Adicionar dia da semana
        df_relatorio['DIA_SEMANA'] = df_relatorio['DATA'].dt.day_name()

        self.df_relatorio_diario = df_relatorio

        print(f"   ✅ Relatório gerado com {len(df_relatorio)} dias")
        print(f"   📅 Período: {df_relatorio['DATA'].min().strftime('%d/%m/%Y')} até {df_relatorio['DATA'].max().strftime('%d/%m/%Y')}")
        print(f"   💰 Saldo inicial: R$ {saldo_inicial:,.2f}")
        print(f"   💰 Saldo final projetado: R$ {df_relatorio['SALDO_FINAL'].iloc[-1]:,.2f}")

        return df_relatorio
    
    def exportar_relatorio_completo(self, output_path: str = 'relatorio_fluxo_caixa_completo.xlsx'):
        """
        Exporta relatório completo para Excel com múltiplas abas
        """
        print(f"\n💾 Exportando relatório completo...")

        # Verificar se o arquivo existe e tentar deletá-lo
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
                print(f"   🗑️  Arquivo existente removido")
            except PermissionError:
                print(f"\n❌ ERRO: O arquivo está aberto em outro programa!")
                print(f"   📁 Arquivo: {output_path}")
                print(f"   💡 Solução: Feche o arquivo no Excel e execute novamente")
                raise PermissionError(f"Não foi possível acessar o arquivo. Por favor, feche-o no Excel: {output_path}")

        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

                # ABA 1: Relatório Diário
                if self.df_relatorio_diario is not None:
                    df_export_diario = self.df_relatorio_diario[[
                        'DATA_FORMATADA', 'DIA_SEMANA', 'SALDO_BANCARIO',
                        'QTD_ENTRADAS', 'ENTRADAS', 'QTD_SAIDAS', 'SAIDAS',
                        'MOVIMENTACAO_DIA', 'SALDO_FINAL'
                    ]].copy()

                    df_export_diario.columns = [
                        'Data', 'Dia da Semana', 'Saldo Bancário',
                        'Qtd Entradas', 'Total Entradas', 'Qtd Saídas', 'Total Saídas',
                        'Movimentação Líquida', 'Saldo Final'
                    ]

                    df_export_diario.to_excel(writer, sheet_name='Relatório Diário', index=False)

                # ABA 2: Timeline Detalhada
                if self.df_timeline is not None:
                    df_export_timeline = self.df_timeline[[
                        'DATA_FORMATADA', 'PED', 'DESCRICAO', 'CATEGORIA',
                        'TIPO', 'VALOR'
                    ]].copy()

                    df_export_timeline.columns = [
                        'Data', 'Pedido', 'Descrição', 'Categoria', 'Tipo', 'Valor'
                    ]

                    df_export_timeline.to_excel(writer, sheet_name='Timeline Detalhada', index=False)

                # ABA 3: Histórico de Saldos Bancários
                if self.df_saldos_por_data is not None:
                    df_export_saldos = self.df_saldos_por_data.copy()
                    df_export_saldos['DATA_FORMATADA'] = df_export_saldos['DATA'].dt.strftime('%d/%m/%Y')
                    df_export_saldos = df_export_saldos[['DATA_FORMATADA', 'SALDO_BANCARIO']]
                    df_export_saldos.columns = ['Data', 'Saldo Bancário']

                    df_export_saldos.to_excel(writer, sheet_name='Histórico Saldos', index=False)

            print(f"   ✅ Relatório exportado: {output_path}")
            return output_path

        except PermissionError as e:
            print(f"\n❌ ERRO: O arquivo está aberto em outro programa!")
            print(f"   📁 Arquivo: {output_path}")
            print(f"   💡 Solução: Feche o arquivo no Excel e execute novamente")
            raise PermissionError(f"Não foi possível acessar o arquivo. Por favor, feche-o no Excel: {output_path}") from e

    def gerar_relatorio_pdf(self, output_path: str = 'relatorio_fluxo_caixa.pdf') -> str:
        """
        Gera relatório em PDF com design profissional

        Args:
            output_path: Caminho do arquivo PDF de saída

        Returns:
            Caminho do arquivo PDF gerado
        """
        print(f"\n📄 Gerando relatório PDF...")

        if self.df_relatorio_diario is None or self.df_timeline is None:
            print("⚠️  Execute gerar_relatorio_diario() e criar_timeline() primeiro!")
            return None

        # Inicializar gerador de PDF
        pdf_generator = CashFlowPDFGenerator(template_dir='templates')

        # Gerar PDF
        pdf_path = pdf_generator.generate_pdf_report(
            df_relatorio_diario=self.df_relatorio_diario,
            df_timeline=self.df_timeline,
            output_path=output_path
        )

        return pdf_path

    def imprimir_resumo(self):
        """
        Imprime resumo detalhado do processamento
        """
        if self.df_relatorio_diario is None:
            print("⚠️  Execute gerar_relatorio_diario() primeiro!")
            return
        
        df = self.df_relatorio_diario
        
        print("\n" + "="*100)
        print("📊 RESUMO DO FLUXO DE CAIXA")
        print("="*100)
        
        print(f"\n📅 Período Analisado: {df['DATA'].min().strftime('%d/%m/%Y')} até {df['DATA'].max().strftime('%d/%m/%Y')}")
        print(f"🔢 Total de Dias: {len(df)}")
        
        print(f"\n💰 SALDOS:")
        print(f"   • Saldo Inicial: R$ {df['SALDO_BANCARIO'].iloc[0]:,.2f}")
        print(f"   • Saldo Final Projetado: R$ {df['SALDO_FINAL'].iloc[-1]:,.2f}")
        print(f"   • Variação: R$ {df['SALDO_FINAL'].iloc[-1] - df['SALDO_BANCARIO'].iloc[0]:,.2f}")
        
        print(f"\n📊 MOVIMENTAÇÕES TOTAIS:")
        total_entradas = df['ENTRADAS'].sum()
        total_saidas = df['SAIDAS'].sum()
        
        print(f"   • Total de Entradas: R$ {total_entradas:,.2f}")
        print(f"   • Total de Saídas: R$ {total_saidas:,.2f}")
        print(f"   • Movimentação Líquida: R$ {total_entradas - total_saidas:,.2f}")
        
        print(f"\n📈 ESTATÍSTICAS:")
        dias_com_movimentacao = len(df[df['MOVIMENTACAO_DIA'] != 0])
        dias_com_entradas = len(df[df['ENTRADAS'] > 0])
        dias_com_saidas = len(df[df['SAIDAS'] > 0])
        
        print(f"   • Dias com Movimentação: {dias_com_movimentacao}")
        print(f"   • Dias com Entradas: {dias_com_entradas}")
        print(f"   • Dias com Saídas: {dias_com_saidas}")
        
        # Maior entrada e saída
        if dias_com_entradas > 0:
            maior_entrada_idx = df['ENTRADAS'].idxmax()
            print(f"   • Maior Entrada: R$ {df.loc[maior_entrada_idx, 'ENTRADAS']:,.2f} em {df.loc[maior_entrada_idx, 'DATA_FORMATADA']}")
        
        if dias_com_saidas > 0:
            maior_saida_idx = df['SAIDAS'].idxmax()
            print(f"   • Maior Saída: R$ {df.loc[maior_saida_idx, 'SAIDAS']:,.2f} em {df.loc[maior_saida_idx, 'DATA_FORMATADA']}")
        
        # Menor e maior saldo
        menor_saldo_idx = df['SALDO_FINAL'].idxmin()
        maior_saldo_idx = df['SALDO_FINAL'].idxmax()
        
        print(f"\n📉 EXTREMOS DE SALDO:")
        print(f"   • Menor Saldo: R$ {df.loc[menor_saldo_idx, 'SALDO_FINAL']:,.2f} em {df.loc[menor_saldo_idx, 'DATA_FORMATADA']}")
        print(f"   • Maior Saldo: R$ {df.loc[maior_saldo_idx, 'SALDO_FINAL']:,.2f} em {df.loc[maior_saldo_idx, 'DATA_FORMATADA']}")
        
        print("="*100)


def main():
    """Função principal - Exemplo de uso"""
    
    # ===========================
    # CONFIGURAÇÃO
    # ===========================
    file_path = r'G:\Meu Drive\projeto_gerador_fashflow_report_A4\CashFlow Financeiro_new.xlsx'
    output_excel_path = r'G:\Meu Drive\projeto_gerador_fashflow_report_A4\relatorio_fluxo_caixa_completo.xlsx'
    output_pdf_path = r'G:\Meu Drive\projeto_gerador_fashflow_report_A4\relatorio_fluxo_caixa.pdf'
    
    print("\n" + "="*100)
    print("🚀 PROCESSAMENTO DE FLUXO DE CAIXA")
    print("="*100 + "\n")
    
    # ===========================
    # PROCESSAMENTO
    # ===========================
    processor = FluxoCaixaProcessor(file_path)
    
    # 1. Extrair todos os saldos bancários
    df_saldos = processor.extrair_todos_saldos()
    
    # 2. Criar timeline de transações
    df_timeline = processor.criar_timeline()
    
    # 3. Gerar relatório diário consolidado
    df_relatorio = processor.gerar_relatorio_diario()
    
    # 4. Imprimir resumo
    processor.imprimir_resumo()

    # 5. Exportar para Excel
    processor.exportar_relatorio_completo(output_excel_path)

    # 6. Gerar relatório PDF
    processor.gerar_relatorio_pdf(output_pdf_path)
    
    # ===========================
    # VISUALIZAÇÃO
    # ===========================
    print("\n\n📋 AMOSTRA DO RELATÓRIO DIÁRIO (Primeiras 20 linhas):")
    print("-" * 100)
    print(df_relatorio[[
        'DATA_FORMATADA', 'SALDO_BANCARIO', 'ENTRADAS', 'SAIDAS', 
        'MOVIMENTACAO_DIA', 'SALDO_FINAL'
    ]].head(20).to_string(index=False))
    
    print("\n\n📋 ÚLTIMAS 20 LINHAS (com transações):")
    print("-" * 100)
    # Filtrar apenas linhas com movimentação
    df_com_mov = df_relatorio[df_relatorio['MOVIMENTACAO_DIA'] != 0]
    if len(df_com_mov) > 0:
        print(df_com_mov[[
            'DATA_FORMATADA', 'ENTRADAS', 'SAIDAS', 
            'MOVIMENTACAO_DIA', 'SALDO_FINAL'
        ]].tail(20).to_string(index=False))
    else:
        print("Nenhuma movimentação encontrada no período.")
    
    print("\n\n✅ Processamento concluído com sucesso!")
    print(f"📁 Arquivos gerados:")
    print(f"   • Excel: {output_excel_path}")
    print(f"   • PDF: {output_pdf_path}")

    return processor, df_relatorio


if __name__ == "__main__":
    processor, df_relatorio = main()