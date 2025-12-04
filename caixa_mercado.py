#!/usr/bin/env python3
"""
Sistema de Caixa para Mercado
Transformação do sistema de estoque em um caixa eletrônico de mercado
"""

import json
import os
from datetime import datetime
from typing import Dict, List, Optional
from enum import Enum


class TipoMovimentacao(Enum):
    ENTRADA = "entrada"
    SAIDA = "saida"


class ItemEstoque:
    def __init__(self, nome: str, quantidade: int, preco: float = 0.0, codigo: str = ""):
        self.nome = nome
        self.quantidade = quantidade
        self.preco = preco
        self.codigo = codigo
        self.data_cadastro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def to_dict(self):
        return {
            "nome": self.nome,
            "quantidade": self.quantidade,
            "preco": self.preco,
            "codigo": self.codigo,
            "data_cadastro": self.data_cadastro
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        item = cls(data["nome"], data["quantidade"], data["preco"], data["codigo"])
        item.data_cadastro = data["data_cadastro"]
        return item


class Venda:
    def __init__(self, itens: List[Dict], total: float, forma_pagamento: str, desconto: float = 0.0):
        self.itens = itens  # Lista de dicionários com nome, quantidade e preço
        self.total = total
        self.forma_pagamento = forma_pagamento
        self.desconto = desconto
        self.data_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def to_dict(self):
        return {
            "itens": self.itens,
            "total": self.total,
            "forma_pagamento": self.forma_pagamento,
            "desconto": self.desconto,
            "data_registro": self.data_registro
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        venda = cls(data["itens"], data["total"], data["forma_pagamento"], data["desconto"])
        venda.data_registro = data["data_registro"]
        return venda


class SistemaCaixa:
    def __init__(self, arquivo_dados: str = "dados_caixa.json"):
        self.arquivo_dados = arquivo_dados
        self.itens: Dict[str, ItemEstoque] = {}
        self.vendas: List[Venda] = []
        self.total_caixa = 0.0
        self.carregar_dados()
    
    def carregar_dados(self):
        if os.path.exists(self.arquivo_dados):
            with open(self.arquivo_dados, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if content:  # Verificar se o arquivo não está vazio
                    dados = json.loads(content)
                    
                    # Carregar itens
                    for nome, item_data in dados.get("itens", {}).items():
                        self.itens[nome] = ItemEstoque.from_dict(item_data)
                    
                    # Carregar vendas
                    for venda_data in dados.get("vendas", []):
                        self.vendas.append(Venda.from_dict(venda_data))
                    
                    # Carregar total do caixa
                    self.total_caixa = dados.get("total_caixa", 0.0)
    
    def salvar_dados(self):
        dados = {
            "itens": {nome: item.to_dict() for nome, item in self.itens.items()},
            "vendas": [venda.to_dict() for venda in self.vendas],
            "total_caixa": self.total_caixa
        }
        
        with open(self.arquivo_dados, 'w', encoding='utf-8') as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    
    def adicionar_produto(self, nome: str, quantidade: int, preco: float, codigo: str = "") -> bool:
        """Adiciona ou atualiza um produto no estoque"""
        if not nome or quantidade < 0:
            return False
        
        nome = nome.strip().lower()
        if nome in self.itens:
            self.itens[nome].quantidade += quantidade
            if preco > 0:
                self.itens[nome].preco = preco
            if codigo:
                self.itens[nome].codigo = codigo
        else:
            self.itens[nome] = ItemEstoque(nome, quantidade, preco, codigo)
        
        self.salvar_dados()
        return True
    
    def remover_produto(self, nome: str, quantidade: int = None) -> bool:
        """Remove quantidade específica ou todo o produto do estoque"""
        nome = nome.strip().lower()
        if nome not in self.itens:
            return False
        
        if quantidade is None:
            # Remover item completamente
            del self.itens[nome]
        else:
            # Reduzir quantidade
            if quantidade > self.itens[nome].quantidade:
                return False  # Não pode remover mais do que tem
            
            self.itens[nome].quantidade -= quantidade
        
        self.salvar_dados()
        return True
    
    def pesquisar_produto(self, termo: str) -> List[ItemEstoque]:
        """Pesquisa produtos pelo nome"""
        termo = termo.strip().lower()
        resultados = []
        
        for item in self.itens.values():
            if termo in item.nome or termo == item.codigo:
                resultados.append(item)
        
        return resultados
    
    def verificar_estoque(self, nome: str, quantidade: int) -> bool:
        """Verifica se há estoque suficiente para a venda"""
        nome = nome.strip().lower()
        if nome not in self.itens:
            return False
        return self.itens[nome].quantidade >= quantidade
    
    def registrar_venda(self, itens_venda: List[Dict], forma_pagamento: str, desconto: float = 0.0) -> bool:
        """Registra uma venda e atualiza o estoque"""
        total = 0.0
        itens_atualizados = []
        
        # Verificar estoque e calcular total
        for item_info in itens_venda:
            nome = item_info["nome"].strip().lower()
            quantidade = item_info["quantidade"]
            
            if not self.verificar_estoque(nome, quantidade):
                print(f"Estoque insuficiente para {nome}")
                return False
            
            preco = self.itens[nome].preco
            subtotal = preco * quantidade
            total += subtotal
            
            # Adicionar item atualizado para salvar na venda
            itens_atualizados.append({
                "nome": self.itens[nome].nome,
                "quantidade": quantidade,
                "preco_unitario": preco,
                "subtotal": subtotal
            })
        
        # Aplicar desconto
        total_com_desconto = total - desconto
        
        # Atualizar estoque
        for item_info in itens_venda:
            nome = item_info["nome"].strip().lower()
            quantidade = item_info["quantidade"]
            self.itens[nome].quantidade -= quantidade
        
        # Registrar venda
        venda = Venda(itens_atualizados, total_com_desconto, forma_pagamento, desconto)
        self.vendas.append(venda)
        
        # Atualizar total do caixa
        self.total_caixa += total_com_desconto
        
        self.salvar_dados()
        return True
    
    def listar_produtos(self) -> List[ItemEstoque]:
        """Lista todos os produtos ordenados por nome"""
        return sorted(self.itens.values(), key=lambda x: x.nome)
    
    def get_total_vendas_hoje(self) -> float:
        """Calcula o total de vendas do dia"""
        hoje = datetime.now().strftime("%Y-%m-%d")
        total = 0.0
        
        for venda in self.vendas:
            if venda.data_registro.startswith(hoje):
                total += venda.total
        
        return total
    
    def get_relatorio_vendas(self) -> Dict:
        """Gera relatório de vendas"""
        total_vendas = len(self.vendas)
        total_recebido = sum(venda.total for venda in self.vendas)
        vendas_hoje = self.get_total_vendas_hoje()
        
        return {
            "total_vendas": total_vendas,
            "total_recebido": total_recebido,
            "vendas_hoje": vendas_hoje,
            "total_caixa": self.total_caixa,
            "data_geracao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }


def validar_numero(valor: str) -> Optional[float]:
    try:
        return float(valor)
    except ValueError:
        return None


def validar_inteiro(valor: str) -> Optional[int]:
    try:
        return int(valor)
    except ValueError:
        return None


def main():
    sistema = SistemaCaixa()
    
    print("┌" + "─" * 50 + "┐")
    print("│        SISTEMA DE CAIXA - MERCADO          │")
    print("└" + "─" * 50 + "┘")
    
    while True:
        print("\n" + "="*50)
        print("MENU PRINCIPAL")
        print("="*50)
        print("1. Registrar venda")
        print("2. Adicionar/Atualizar produto")
        print("3. Listar produtos")
        print("4. Pesquisar produto")
        print("5. Remover produto")
        print("6. Relatório de vendas")
        print("7. Verificar estoque")
        print("0. Sair")
        print("-"*50)
        
        opcao = input("Escolha uma opção: ").strip()
        
        if opcao == "1":
            print("\nREGISTRO DE VENDA")
            print("-" * 30)
            
            itens_venda = []
            total_parcial = 0.0
            
            while True:
                nome_produto = input("Nome ou código do produto (Enter para finalizar): ").strip()
                if not nome_produto:
                    break
                
                # Procurar produto
                produtos = sistema.pesquisar_produto(nome_produto)
                if not produtos:
                    print("Produto não encontrado!")
                    continue
                
                # Mostrar opções se houver múltiplos resultados
                if len(produtos) > 1:
                    print("Múltiplos produtos encontrados:")
                    for i, produto in enumerate(produtos):
                        print(f"{i+1}. {produto.nome} - R$ {produto.preco:.2f} - Estoque: {produto.quantidade}")
                    
                    escolha = input("Escolha o número do produto (ou Enter para cancelar): ").strip()
                    if not escolha:
                        continue
                    
                    try:
                        escolha_idx = int(escolha) - 1
                        if 0 <= escolha_idx < len(produtos):
                            produto = produtos[escolha_idx]
                        else:
                            print("Escolha inválida!")
                            continue
                    except ValueError:
                        print("Escolha inválida!")
                        continue
                else:
                    produto = produtos[0]
                
                print(f"Produto: {produto.nome}")
                print(f"Preço: R$ {produto.preco:.2f}")
                print(f"Estoque disponível: {produto.quantidade}")
                
                quantidade_str = input("Quantidade: ").strip()
                quantidade = validar_inteiro(quantidade_str)
                
                if quantidade is None or quantidade <= 0:
                    print("Quantidade inválida!")
                    continue
                
                if quantidade > produto.quantidade:
                    print(f"Quantidade solicitada ({quantidade}) maior que estoque disponível ({produto.quantidade})!")
                    continue
                
                subtotal = produto.preco * quantidade
                print(f"Subtotal: R$ {subtotal:.2f}")
                
                itens_venda.append({
                    "nome": produto.nome,
                    "quantidade": quantidade
                })
                
                total_parcial += subtotal
            
            if not itens_venda:
                print("Nenhum item adicionado à venda!")
                continue
            
            print(f"\nTotal parcial: R$ {total_parcial:.2f}")
            
            # Aplicar desconto
            desconto_str = input("Desconto (Enter para nenhum): ").strip()
            desconto = 0.0
            if desconto_str:
                desconto = validar_numero(desconto_str)
                if desconto is None:
                    print("Desconto inválido!")
                    continue
                if desconto > total_parcial:
                    print("Desconto maior que o total da venda!")
                    continue
            
            total_final = total_parcial - desconto
            print(f"Total com desconto: R$ {total_final:.2f}")
            
            # Forma de pagamento
            print("\nFormas de pagamento:")
            print("1. Dinheiro")
            print("2. Cartão de débito")
            print("3. Cartão de crédito")
            print("4. Pix")
            
            forma_pagamento = input("Escolha a forma de pagamento (1-4): ").strip()
            formas = {"1": "Dinheiro", "2": "Cartão de débito", "3": "Cartão de crédito", "4": "Pix"}
            
            if forma_pagamento in formas:
                forma_pagamento = formas[forma_pagamento]
            else:
                forma_pagamento = "Dinheiro"  # Padrão
            
            # Confirmar venda
            confirmar = input(f"\nConfirmar venda de R$ {total_final:.2f}? (S/N): ").strip().lower()
            if confirmar == 's':
                if sistema.registrar_venda(itens_venda, forma_pagamento, desconto):
                    print("Venda registrada com sucesso!")
                    print(f"Forma de pagamento: {forma_pagamento}")
                    print(f"Total: R$ {total_final:.2f}")
                else:
                    print("Erro ao registrar venda!")
            else:
                print("Venda cancelada!")
        
        elif opcao == "2":
            print("\nADICIONAR/ATUALIZAR PRODUTO")
            print("-" * 30)
            
            nome = input("Nome do produto: ").strip()
            if not nome:
                print("Nome do produto não pode ser vazio!")
                continue
            
            codigo = input("Código do produto (opcional): ").strip()
            
            quantidade_str = input("Quantidade inicial: ").strip()
            quantidade = validar_inteiro(quantidade_str)
            if quantidade is None or quantidade < 0:
                print("Quantidade inválida!")
                continue
            
            preco_str = input("Preço unitário: ").strip()
            preco = validar_numero(preco_str)
            if preco is None or preco < 0:
                print("Preço inválido!")
                continue
            
            if sistema.adicionar_produto(nome, quantidade, preco, codigo):
                print(f"Produto '{nome}' adicionado/atualizado com sucesso!")
            else:
                print("Erro ao adicionar produto!")
        
        elif opcao == "3":
            print("\nLISTA DE PRODUTOS")
            print("-" * 30)
            
            produtos = sistema.listar_produtos()
            if not produtos:
                print("Nenhum produto cadastrado!")
            else:
                print(f"{'Produto':<30} {'Código':<10} {'Preço':<10} {'Estoque':<8}")
                print("-" * 60)
                for produto in produtos:
                    print(f"{produto.nome:<30} {produto.codigo:<10} R$ {produto.preco:<8.2f} {produto.quantidade:<8}")
        
        elif opcao == "4":
            print("\nPESQUISAR PRODUTO")
            print("-" * 30)
            
            termo = input("Termo de pesquisa (nome ou código): ").strip()
            if not termo:
                print("Termo de pesquisa não pode ser vazio!")
                continue
            
            resultados = sistema.pesquisar_produto(termo)
            if not resultados:
                print("Nenhum produto encontrado!")
            else:
                print(f"{'Produto':<30} {'Código':<10} {'Preço':<10} {'Estoque':<8}")
                print("-" * 60)
                for produto in resultados:
                    print(f"{produto.nome:<30} {produto.codigo:<10} R$ {produto.preco:<8.2f} {produto.quantidade:<8}")
        
        elif opcao == "5":
            print("\nREMOVER PRODUTO")
            print("-" * 30)
            
            nome = input("Nome do produto para remover: ").strip()
            if not nome:
                print("Nome do produto não pode ser vazio!")
                continue
            
            produtos = sistema.pesquisar_produto(nome)
            if not produtos:
                print("Produto não encontrado!")
                continue
            
            if len(produtos) > 1:
                print("Múltiplos produtos encontrados:")
                for i, produto in enumerate(produtos):
                    print(f"{i+1}. {produto.nome} - Estoque: {produto.quantidade}")
                
                escolha = input("Escolha o número do produto para remover (ou Enter para cancelar): ").strip()
                if not escolha:
                    continue
                
                try:
                    escolha_idx = int(escolha) - 1
                    if 0 <= escolha_idx < len(produtos):
                        produto = produtos[escolha_idx]
                        nome = produto.nome
                    else:
                        print("Escolha inválida!")
                        continue
                except ValueError:
                    print("Escolha inválida!")
                    continue
            
            remover_tudo = input("Remover produto completamente? (s/N): ").strip().lower()
            if remover_tudo == 's':
                if sistema.remover_produto(nome):
                    print(f"Produto '{nome}' removido completamente!")
                else:
                    print(f"Produto '{nome}' não encontrado!")
            else:
                quantidade_str = input("Quantidade a remover: ").strip()
                quantidade = validar_inteiro(quantidade_str)
                if quantidade is None or quantidade <= 0:
                    print("Quantidade inválida!")
                    continue
                
                if sistema.remover_produto(nome, quantidade):
                    print(f"{quantidade} unidades de '{nome}' removidas!")
                else:
                    print(f"Não foi possível remover {quantidade} unidades de '{nome}'!")
        
        elif opcao == "6":
            print("\nRELATÓRIO DE VENDAS")
            print("-" * 30)
            
            relatorio = sistema.get_relatorio_vendas()
            print(f"Total de vendas registradas: {relatorio['total_vendas']}")
            print(f"Total recebido: R$ {relatorio['total_recebido']:.2f}")
            print(f"Vendas de hoje: R$ {relatorio['vendas_hoje']:.2f}")
            print(f"Total no caixa: R$ {relatorio['total_caixa']:.2f}")
            print(f"Data: {relatorio['data_geracao']}")
        
        elif opcao == "7":
            print("\nVERIFICAÇÃO DE ESTOQUE")
            print("-" * 30)
            
            produtos = sistema.listar_produtos()
            if not produtos:
                print("Nenhum produto cadastrado!")
            else:
                baixo_estoque = [p for p in produtos if p.quantidade < 5]
                if baixo_estoque:
                    print("Produtos com baixo estoque (menos de 5 unidades):")
                    for produto in baixo_estoque:
                        print(f"- {produto.nome}: {produto.quantidade} unidades")
                else:
                    print("Nenhum produto com baixo estoque!")
        
        elif opcao == "0":
            print("Saindo do sistema de caixa...")
            break
        
        else:
            print("Opção inválida!")


if __name__ == "__main__":
    main()