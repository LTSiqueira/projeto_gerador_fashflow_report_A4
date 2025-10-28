"""
Script de teste para geração de PDF
Testa o módulo pdf_generator com dados mockados
"""

import pandas as pd
from pdf_generator import CashFlowPDFGenerator
from datetime import datetime, timedelta


def create_mock_data():
    """Cria dados mockados para teste"""
    print("📊 Criando dados de teste...")

    # Criar 30 dias de dados
    start_date = datetime(2025, 1, 1)
    dates = [start_date + timedelta(days=i) for i in range(30)]

    # Relatório diário
    df_relatorio = pd.DataFrame({
        'DATA': dates,
        'ENTRADAS': [1000 + i*100 for i in range(30)],
        'SAIDAS': [500 + i*50 for i in range(30)],
        'SALDO_FINAL': [10000 + i*50 for i in range(30)],
    })

    # Timeline de transações (duplicar alguns dias para ter múltiplas transações)
    transacoes = []
    for i, date in enumerate(dates[:15]):  # Apenas primeiros 15 dias para ter variação
        # Entrada
        transacoes.append({
            'DATA': date,
            'DESCRICAO': f'Cliente {chr(65+i)}',
            'PED': f'PC{1000+i}',
            'TIPO': 'ENTRADA',
            'VALOR': 500 + i*50,
            'CATEGORIA': 'CR - Produto'
        })

        # Mais uma entrada
        transacoes.append({
            'DATA': date,
            'DESCRICAO': f'Cliente {chr(65+i+1)}',
            'PED': f'PC{2000+i}',
            'TIPO': 'ENTRADA',
            'VALOR': 500 + i*50,
            'CATEGORIA': 'CR - Produto'
        })

        # Saída
        transacoes.append({
            'DATA': date,
            'DESCRICAO': f'Fornecedor {chr(90-i)}',
            'PED': f'PG{3000+i}',
            'TIPO': 'SAÍDA',
            'VALOR': 300 + i*30,
            'CATEGORIA': 'CP - Produto'
        })

        # Saída geral
        transacoes.append({
            'DATA': date,
            'DESCRICAO': 'SAÍDAS GERAIS',
            'PED': '',
            'TIPO': 'SAÍDA',
            'VALOR': 200 + i*20,
            'CATEGORIA': 'CP - Saídas Gerais'
        })

    df_timeline = pd.DataFrame(transacoes)

    print(f"   ✅ Criados {len(df_relatorio)} dias de relatório")
    print(f"   ✅ Criadas {len(df_timeline)} transações")

    return df_relatorio, df_timeline


def test_pdf_generation():
    """Testa a geração de PDF"""
    print("\n" + "="*100)
    print("🧪 TESTE DE GERAÇÃO DE PDF")
    print("="*100 + "\n")

    try:
        # 1. Criar dados mockados
        df_relatorio, df_timeline = create_mock_data()

        # 2. Inicializar gerador
        print("\n📦 Inicializando gerador de PDF...")
        generator = CashFlowPDFGenerator(template_dir='templates')
        print("   ✅ Gerador inicializado")

        # 3. Gerar PDF
        output_path = r'G:\Meu Drive\projeto_gerador_fashflow_report_A4\test_relatorio.pdf'

        pdf_path = generator.generate_pdf_report(
            df_relatorio_diario=df_relatorio,
            df_timeline=df_timeline,
            output_path=output_path
        )

        print("\n" + "="*100)
        print("✅ TESTE CONCLUÍDO COM SUCESSO!")
        print("="*100)
        print(f"\n📁 Arquivos gerados:")
        print(f"   • PDF: {pdf_path}")
        print(f"   • HTML (debug): {output_path.replace('.pdf', '_debug.html')}")
        print("\n💡 Dica: Abra o HTML no navegador para verificar o layout antes do PDF")

        return True

    except Exception as e:
        print("\n" + "="*100)
        print("❌ ERRO NO TESTE")
        print("="*100)
        print(f"\n{type(e).__name__}: {str(e)}")
        print("\n💡 Verifique se:")
        print("   1. WeasyPrint está instalado corretamente")
        print("   2. A pasta 'templates' existe")
        print("   3. O template 'cashflow_report.html' está na pasta templates")

        import traceback
        print("\n📋 Stack trace completo:")
        traceback.print_exc()

        return False


if __name__ == "__main__":
    success = test_pdf_generation()

    if success:
        print("\n🎉 Teste bem-sucedido! O sistema está pronto para uso.")
    else:
        print("\n⚠️  Corrija os erros acima e tente novamente.")
