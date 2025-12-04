#!/usr/bin/env python3
"""
Exemplo prático de uso do sistema de auditoria de estoque aprimorado.
Este script demonstra as principais funcionalidades do sistema.
"""

import os
import tempfile
import sys
sys.path.insert(0, '/workspace')

from app_melhorado import Estoque

def exemplo_pratico():
    print("=" * 60)
    print("EXEMPLO PRÁTICO - Sistema de Auditoria de Estoque Aprimorado")
    print("=" * 60)
    
    # Criar uma instância do estoque
    estoque = Estoque("exemplo_estoque.json")
    
    print("\n1. ADICIONANDO ITENS AO ESTOQUE")
    print("-" * 40)
    estoque.adicionar_item("Notebook Dell", 15)
    estoque.adicionar_item("Mouse Logitech", 50)
    estoque.adicionar_item("Teclado Mecânico", 30)
    estoque.adicionar_item("Monitor 24\"", 20)
    estoque.adicionar_item("Cadeira Gamer", 8)
    
    print("\n2. LISTANDO ITENS EM ORDEM ALFABÉTICA")
    print("-" * 40)
    estoque.listar_itens("nome")
    
    print("\n3. LISTANDO ITENS POR QUANTIDADE (DECRESCE NTE)")
    print("-" * 40)
    estoque.listar_itens("quantidade_desc")
    
    print("\n4. REMOVENDO ALGUNS ITENS")
    print("-" * 40)
    estoque.remover_item("Cadeira Gamer", 3)  # Remover 3 cadeiras
    print("Removendo 3 cadeiras gamer...")
    estoque.remover_item("Mouse Logitech")  # Remover todos os mouses
    print("Removendo todos os mouses logitech...")
    
    print("\n5. LISTANDO ITENS APÓS REMOÇÕES")
    print("-" * 40)
    estoque.listar_itens("quantidade_desc")
    
    print("\n6. REALIZANDO UM BALANÇO DE ESTOQUE (SIMULADO)")
    print("-" * 40)
    print("Simulando balanço com quantidades físicas diferentes:")
    
    # Adicionando itens para o balanço
    estoque.adicionar_item("HD Externo", 25)
    estoque.adicionar_item("Webcam", 12)
    
    # Criando um balanço simulado (isso normalmente seria feito interativamente)
    from app_melhorado import BalancoItem
    import datetime
    
    # Simular balanço para alguns itens
    balanco_atual = []
    
    # Para cada item no estoque, criar um item de balanço com quantidade física diferente
    for nome, item in estoque.itens.items():
        # Simular contagem física com alguma diferença
        if nome == "notebook dell":
            quantidade_fisica = item.quantidade - 2  # 2 notebooks a menos
        elif nome == "teclado mecânico":
            quantidade_fisica = item.quantidade + 1  # 1 teclado a mais
        elif nome == "monitor 24\"":
            quantidade_fisica = item.quantidade  # quantidade correta
        elif nome == "hd externo":
            quantidade_fisica = item.quantidade - 5  # 5 HDs a menos
        elif nome == "webcam":
            quantidade_fisica = 0  # nenhuma webcam encontrada
        else:
            quantidade_fisica = item.quantidade  # quantidade correta para outros itens
        
        balanco_item = BalancoItem(nome, item.quantidade, quantidade_fisica)
        balanco_atual.append(balanco_item)
    
    # Salvar no histórico
    estoque.historico_balanco.append({
        "data": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "balanco": [item.to_dict() for item in balanco_atual]
    })
    
    print("Balanço simulado realizado!")
    estoque.mostrar_balanco(balanco_atual)
    
    print("\n7. GERANDO RELATÓRIO DETALHADO")
    print("-" * 40)
    print(estoque.gerar_relatorio_estoque())
    
    print("\n8. EXPORTANDO DADOS DO ESTOQUE PARA CSV")
    print("-" * 40)
    with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, prefix='estoque_export_') as tmp:
        caminho_csv = tmp.name
    
    if estoque.exportar_estoque(caminho_csv):
        print(f"Dados exportados para: {caminho_csv}")
        # Mostrar conteúdo do arquivo
        import pandas as pd
        df = pd.read_csv(caminho_csv)
        print(f"Arquivo contém {len(df)} registros")
        print(df.head())  # Mostrar primeiras linhas
        
        # Remover arquivo temporário
        os.remove(caminho_csv)
    
    print("\n9. HISTÓRICO DE BALANÇOS")
    print("-" * 40)
    estoque.mostrar_historico_balanco()
    
    print("\n10. PERSISTÊNCIA DE DADOS")
    print("-" * 40)
    print(f"Os dados do estoque estão sendo automaticamente salvos em: exemplo_estoque.json")
    print(f"Arquivo existe: {os.path.exists('exemplo_estoque.json')}")
    
    # Limpar o arquivo de exemplo
    if os.path.exists("exemplo_estoque.json"):
        os.remove("exemplo_estoque.json")
    
    print("\n" + "=" * 60)
    print("EXEMPLO CONCLUÍDO COM SUCESSO!")
    print("O sistema aprimorado oferece todas essas funcionalidades e mais:")
    print("• Gerenciamento completo de estoque")
    print("• Histórico de balanços")
    print("• Exportação para CSV e Excel")
    print("• Relatórios detalhados")
    print("• Persistência automática de dados")
    print("• Validações e tratamento de erros robustos")
    print("• Interface intuitiva e amigável")
    print("=" * 60)

if __name__ == "__main__":
    exemplo_pratico()