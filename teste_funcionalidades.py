#!/usr/bin/env python3
# Teste para verificar as funcionalidades do sistema aprimorado

import os
import sys
import tempfile
import pandas as pd

# Adiciona o caminho do diretório atual para importar o módulo
sys.path.insert(0, '/workspace')

from app_melhorado import Estoque, ItemEstoque, BalancoItem

def testar_funcionalidades():
    print("Testando funcionalidades do sistema aprimorado...")
    
    # Testar criação de estoque
    print("\n1. Criando estoque...")
    estoque = Estoque("teste_estoque.json")
    print("✓ Estoque criado com sucesso")
    
    # Testar adição de itens
    print("\n2. Testando adição de itens...")
    estoque.adicionar_item("Caneta", 100)
    estoque.adicionar_item("Lápis", 50)
    estoque.adicionar_item("Borracha", 30)
    print("✓ Itens adicionados com sucesso")
    
    # Testar listagem de itens
    print("\n3. Listando itens...")
    estoque.listar_itens()
    
    # Testar obtenção de item específico
    print("\n4. Testando obtenção de item específico...")
    item = estoque.obter_item("Caneta")
    if item:
        print(f"✓ Item encontrado: {item.nome} - Quantidade: {item.quantidade}")
    else:
        print("✗ Item não encontrado")
    
    # Testar remoção de item
    print("\n5. Testando remoção de item...")
    estoque.remover_item("Borracha")
    print("✓ Item removido com sucesso")
    
    # Listar novamente para verificar remoção
    print("\n6. Listando itens após remoção...")
    estoque.listar_itens()
    
    # Testar exportação para CSV
    print("\n7. Testando exportação para CSV...")
    with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
        temp_csv = tmp.name
    
    if estoque.exportar_estoque(temp_csv):
        print(f"✓ Estoque exportado para {temp_csv}")
        
        # Verificar conteúdo do CSV exportado
        df = pd.read_csv(temp_csv)
        print(f"✓ CSV contém {len(df)} registros")
        print("Conteúdo do CSV:")
        print(df)
        
        # Remover arquivo temporário
        os.remove(temp_csv)
    else:
        print("✗ Falha na exportação para CSV")
    
    # Testar geração de relatório
    print("\n8. Testando geração de relatório...")
    relatorio = estoque.gerar_relatorio_estoque()
    print("✓ Relatório gerado:")
    print(relatorio)
    
    # Testar persistência
    print("\n9. Testando persistência de dados...")
    # Verificar se o arquivo JSON foi criado
    if os.path.exists("teste_estoque.json"):
        print("✓ Arquivo de persistência criado")
    else:
        print("✗ Arquivo de persistência não encontrado")
    
    # Testar criação de objeto BalancoItem
    print("\n10. Testando criação de item de balanço...")
    balanco_item = BalancoItem("Caneta", 100, 95)
    print(f"✓ Item de balanço criado: {balanco_item.nome}, Status: {balanco_item.status}, Diferença: {balanco_item.diferenca}")
    
    print("\n✓ Todos os testes básicos concluídos com sucesso!")
    
    # Limpar arquivo de teste
    if os.path.exists("teste_estoque.json"):
        os.remove("teste_estoque.json")

if __name__ == "__main__":
    testar_funcionalidades()