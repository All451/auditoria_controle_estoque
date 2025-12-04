#!/usr/bin/env python3
"""
Script de teste para o sistema de caixa de mercado
"""

from caixa_mercado import SistemaCaixa

def teste_basico():
    print("Testando sistema de caixa...")
    
    # Criar instância do sistema
    sistema = SistemaCaixa("teste_dados.json")
    
    # Adicionar alguns produtos
    print("\n1. Adicionando produtos...")
    sistema.adicionar_produto("Arroz", 10, 18.50, "001")
    sistema.adicionar_produto("Feijão", 5, 7.90, "002")
    sistema.adicionar_produto("Macarrão", 8, 3.45, "003")
    print("Produtos adicionados com sucesso!")
    
    # Listar produtos
    print("\n2. Listando produtos...")
    produtos = sistema.listar_produtos()
    for produto in produtos:
        print(f"- {produto.nome}: R$ {produto.preco:.2f}, Estoque: {produto.quantidade}, Código: {produto.codigo}")
    
    # Registrar uma venda
    print("\n3. Registrando venda...")
    itens_venda = [
        {"nome": "Arroz", "quantidade": 2},
        {"nome": "Feijão", "quantidade": 1}
    ]
    
    sucesso = sistema.registrar_venda(itens_venda, "Dinheiro", 2.00)  # Com R$ 2,00 de desconto
    if sucesso:
        print("Venda registrada com sucesso!")
    else:
        print("Erro ao registrar venda!")
    
    # Verificar estoque após venda
    print("\n4. Estoque após venda...")
    produtos = sistema.listar_produtos()
    for produto in produtos:
        print(f"- {produto.nome}: Estoque: {produto.quantidade}")
    
    # Gerar relatório
    print("\n5. Relatório de vendas...")
    relatorio = sistema.get_relatorio_vendas()
    print(f"Total de vendas: {relatorio['total_vendas']}")
    print(f"Total recebido: R$ {relatorio['total_recebido']:.2f}")
    print(f"Total no caixa: R$ {relatorio['total_caixa']:.2f}")
    
    print("\nTeste concluído com sucesso!")

if __name__ == "__main__":
    teste_basico()