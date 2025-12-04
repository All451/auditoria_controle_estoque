import json
import os
from datetime import datetime
from typing import Dict, List, Optional
from enum import Enum

class TipoMovimentacao(Enum):
    ENTRADA = "entrada"
    SAIDA = "saida"

class ItemEstoque:
    def __init__(self, nome: str, quantidade: int, preco: float = 0.0):
        self.nome = nome
        self.quantidade = quantidade
        self.preco = preco
        self.data_cadastro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def to_dict(self):
        return {
            "nome": self.nome,
            "quantidade": self.quantidade,
            "preco": self.preco,
            "data_cadastro": self.data_cadastro
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        item = cls(data["nome"], data["quantidade"], data["preco"])
        item.data_cadastro = data["data_cadastro"]
        return item

class Movimentacao:
    def __init__(self, tipo: TipoMovimentacao, item_nome: str, quantidade: int, observacao: str = ""):
        self.tipo = tipo
        self.item_nome = item_nome
        self.quantidade = quantidade
        self.observacao = observacao
        self.data_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def to_dict(self):
        return {
            "tipo": self.tipo.value,
            "item_nome": self.item_nome,
            "quantidade": self.quantidade,
            "observacao": self.observacao,
            "data_registro": self.data_registro
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        tipo = TipoMovimentacao(data["tipo"])
        mov = cls(tipo, data["item_nome"], data["quantidade"], data["observacao"])
        mov.data_registro = data["data_registro"]
        return mov

class Balanco:
    def __init__(self, itens_fisicos: Dict[str, int], observacao: str = ""):
        self.itens_fisicos = itens_fisicos
        self.observacao = observacao
        self.data_registro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def to_dict(self):
        return {
            "itens_fisicos": self.itens_fisicos,
            "observacao": self.observacao,
            "data_registro": self.data_registro
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        balanco = cls(data["itens_fisicos"], data["observacao"])
        balanco.data_registro = data["data_registro"]
        return balanco

class SistemaEstoque:
    def __init__(self, arquivo_dados: str = "dados_estoque.json"):
        self.arquivo_dados = arquivo_dados
        self.itens: Dict[str, ItemEstoque] = {}
        self.movimentacoes: List[Movimentacao] = []
        self.balancos: List[Balanco] = []
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
                    
                    # Carregar movimentações
                    for mov_data in dados.get("movimentacoes", []):
                        self.movimentacoes.append(Movimentacao.from_dict(mov_data))
                    
                    # Carregar balanços
                    for balanco_data in dados.get("balancos", []):
                        self.balancos.append(Balanco.from_dict(balanco_data))
    
    def salvar_dados(self):
        dados = {
            "itens": {nome: item.to_dict() for nome, item in self.itens.items()},
            "movimentacoes": [mov.to_dict() for mov in self.movimentacoes],
            "balancos": [balanco.to_dict() for balanco in self.balancos]
        }
        
        with open(self.arquivo_dados, 'w', encoding='utf-8') as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    
    def adicionar_item(self, nome: str, quantidade: int, preco: float = 0.0) -> bool:
        if not nome or quantidade < 0:
            return False
        
        nome = nome.strip().lower()
        if nome in self.itens:
            self.itens[nome].quantidade += quantidade
        else:
            self.itens[nome] = ItemEstoque(nome, quantidade, preco)
        
        # Registrar movimentação de entrada
        movimentacao = Movimentacao(TipoMovimentacao.ENTRADA, nome, quantidade, "Adição inicial ou reabastecimento")
        self.movimentacoes.append(movimentacao)
        self.salvar_dados()
        return True
    
    def remover_item(self, nome: str, quantidade: int = None) -> bool:
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
            
            # Registrar movimentação de saída
            movimentacao = Movimentacao(TipoMovimentacao.SAIDA, nome, quantidade, "Remoção de estoque")
            self.movimentacoes.append(movimentacao)
        
        self.salvar_dados()
        return True
    
    def registrar_entrada(self, nome: str, quantidade: int, observacao: str = "") -> bool:
        if quantidade <= 0:
            return False
        
        nome = nome.strip().lower()
        if nome not in self.itens:
            self.itens[nome] = ItemEstoque(nome, 0)
        
        self.itens[nome].quantidade += quantidade
        
        movimentacao = Movimentacao(TipoMovimentacao.ENTRADA, nome, quantidade, observacao)
        self.movimentacoes.append(movimentacao)
        self.salvar_dados()
        return True
    
    def registrar_saida(self, nome: str, quantidade: int, observacao: str = "") -> bool:
        if quantidade <= 0:
            return False
        
        nome = nome.strip().lower()
        if nome not in self.itens or self.itens[nome].quantidade < quantidade:
            return False
        
        self.itens[nome].quantidade -= quantidade
        
        movimentacao = Movimentacao(TipoMovimentacao.SAIDA, nome, quantidade, observacao)
        self.movimentacoes.append(movimentacao)
        self.salvar_dados()
        return True
    
    def listar_itens(self, ordenar_por: str = "nome") -> List[ItemEstoque]:
        itens_lista = list(self.itens.values())
        
        if ordenar_por == "quantidade":
            itens_lista.sort(key=lambda x: x.quantidade, reverse=True)
        else:  # ordenar por nome
            itens_lista.sort(key=lambda x: x.nome)
        
        return itens_lista
    
    def pesquisar_item(self, termo: str) -> List[ItemEstoque]:
        termo = termo.strip().lower()
        resultados = []
        
        for item in self.itens.values():
            if termo in item.nome:
                resultados.append(item)
        
        return resultados
    
    def get_historico_movimentacoes(self, item_nome: str = None) -> List[Movimentacao]:
        if item_nome:
            item_nome = item_nome.strip().lower()
            return [mov for mov in self.movimentacoes if mov.item_nome == item_nome]
        return self.movimentacoes
    
    def realizar_balanco(self, itens_fisicos: Dict[str, int], observacao: str = "") -> Dict[str, int]:
        itens_fisicos_lower = {k.strip().lower(): v for k, v in itens_fisicos.items()}
        
        # Registrar o balanço
        balanco = Balanco(itens_fisicos_lower, observacao)
        self.balancos.append(balanco)
        
        # Identificar discrepâncias
        discrepancias = {}
        for nome, qtd_fisica in itens_fisicos_lower.items():
            qtd_sistema = self.itens.get(nome, ItemEstoque(nome, 0)).quantidade
            diferenca = qtd_fisica - qtd_sistema
            if diferenca != 0:
                discrepancias[nome] = diferenca
        
        # Atualizar quantidades no sistema com base no balanço físico
        for nome, qtd_fisica in itens_fisicos_lower.items():
            if nome in self.itens:
                diferenca = qtd_fisica - self.itens[nome].quantidade
                if diferenca != 0:
                    tipo = TipoMovimentacao.ENTRADA if diferenca > 0 else TipoMovimentacao.SAIDA
                    observacao_ajuste = f"Ajuste de balanço físico: diferença de {diferenca}"
                    movimentacao = Movimentacao(tipo, nome, abs(diferenca), observacao_ajuste)
                    self.movimentacoes.append(movimentacao)
                    self.itens[nome].quantidade = qtd_fisica
            else:
                # Item encontrado no balanço físico mas não no sistema
                self.itens[nome] = ItemEstoque(nome, qtd_fisica)
                movimentacao = Movimentacao(TipoMovimentacao.ENTRADA, nome, qtd_fisica, "Registro de balanço físico")
                self.movimentacoes.append(movimentacao)
        
        # Verificar itens que existem no sistema mas não no balanço físico
        for nome in self.itens.keys():
            if nome not in itens_fisicos_lower and self.itens[nome].quantidade > 0:
                discrepancias[nome] = -self.itens[nome].quantidade
                # Registrar saída para zerar o item
                movimentacao = Movimentacao(TipoMovimentacao.SAIDA, nome, self.itens[nome].quantidade, "Item não encontrado no balanço físico")
                self.movimentacoes.append(movimentacao)
                self.itens[nome].quantidade = 0
        
        self.salvar_dados()
        return discrepancias
    
    def gerar_relatorio_estoque(self) -> Dict:
        total_itens = len(self.itens)
        total_quantidade = sum(item.quantidade for item in self.itens.values())
        valor_total = sum(item.quantidade * item.preco for item in self.itens.values())
        
        itens_baixa_quantidade = [item for item in self.itens.values() if item.quantidade < 5]
        
        return {
            "total_itens": total_itens,
            "total_quantidade": total_quantidade,
            "valor_total": valor_total,
            "itens_baixa_quantidade": len(itens_baixa_quantidade),
            "data_geracao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    
    def get_ultimo_balanco(self) -> Optional[Balanco]:
        if not self.balancos:
            return None
        return self.balancos[-1]

def validar_numero(valor: str) -> Optional[int]:
    try:
        return int(valor)
    except ValueError:
        return None

def main():
    sistema = SistemaEstoque()
    
    while True:
        print("\n" + "="*50)
        print("SISTEMA DE GERENCIAMENTO DE ESTOQUE E AUDITORIA")
        print("="*50)
        print("1. Adicionar item")
        print("2. Remover item")
        print("3. Registrar entrada")
        print("4. Registrar saída")
        print("5. Listar itens")
        print("6. Pesquisar item")
        print("7. Histórico de movimentações")
        print("8. Realizar balanço")
        print("9. Relatório de estoque")
        print("10. Histórico de balanços")
        print("0. Sair")
        print("-"*50)
        
        opcao = input("Escolha uma opção: ").strip()
        
        if opcao == "1":
            nome = input("Nome do item: ").strip()
            if not nome:
                print("Nome do item não pode ser vazio!")
                continue
            
            quantidade_str = input("Quantidade inicial: ").strip()
            quantidade = validar_numero(quantidade_str)
            if quantidade is None or quantidade < 0:
                print("Quantidade inválida!")
                continue
            
            preco_str = input("Preço unitário (opcional, pressione Enter para pular): ").strip()
            preco = 0.0
            if preco_str:
                try:
                    preco = float(preco_str)
                except ValueError:
                    print("Preço inválido!")
                    continue
            
            if sistema.adicionar_item(nome, quantidade, preco):
                print(f"Item '{nome}' adicionado com sucesso!")
            else:
                print("Erro ao adicionar item!")
        
        elif opcao == "2":
            nome = input("Nome do item para remover: ").strip()
            if not nome:
                print("Nome do item não pode ser vazio!")
                continue
            
            remover_tudo = input("Remover item completamente? (s/N): ").strip().lower()
            if remover_tudo == 's':
                if sistema.remover_item(nome):
                    print(f"Item '{nome}' removido completamente!")
                else:
                    print(f"Item '{nome}' não encontrado!")
            else:
                quantidade_str = input("Quantidade a remover: ").strip()
                quantidade = validar_numero(quantidade_str)
                if quantidade is None or quantidade <= 0:
                    print("Quantidade inválida!")
                    continue
                
                if sistema.remover_item(nome, quantidade):
                    print(f"{quantidade} unidades de '{nome}' removidas!")
                else:
                    print(f"Não foi possível remover {quantidade} unidades de '{nome}'!")
        
        elif opcao == "3":
            nome = input("Nome do item: ").strip()
            if not nome:
                print("Nome do item não pode ser vazio!")
                continue
            
            quantidade_str = input("Quantidade a adicionar: ").strip()
            quantidade = validar_numero(quantidade_str)
            if quantidade is None or quantidade <= 0:
                print("Quantidade inválida!")
                continue
            
            observacao = input("Observação (opcional): ").strip()
            
            if sistema.registrar_entrada(nome, quantidade, observacao):
                print(f"Entrada de {quantidade} unidades de '{nome}' registrada!")
            else:
                print(f"Erro ao registrar entrada de '{nome}'!")
        
        elif opcao == "4":
            nome = input("Nome do item: ").strip()
            if not nome:
                print("Nome do item não pode ser vazio!")
                continue
            
            quantidade_str = input("Quantidade a remover: ").strip()
            quantidade = validar_numero(quantidade_str)
            if quantidade is None or quantidade <= 0:
                print("Quantidade inválida!")
                continue
            
            observacao = input("Observação (opcional): ").strip()
            
            if sistema.registrar_saida(nome, quantidade, observacao):
                print(f"Saída de {quantidade} unidades de '{nome}' registrada!")
            else:
                print(f"Não foi possível registrar saída de '{nome}'!")
        
        elif opcao == "5":
            print("\nOrdenar por:")
            print("1. Nome")
            print("2. Quantidade")
            ordem = input("Escolha a ordem (1-2, Enter para nome): ").strip()
            
            ordenar_por = "quantidade" if ordem == "2" else "nome"
            itens = sistema.listar_itens(ordenar_por)
            
            if not itens:
                print("Nenhum item no estoque!")
            else:
                print(f"\nItens no estoque (ordenado por {ordenar_por}):")
                print("-" * 60)
                for item in itens:
                    print(f"{item.nome:<20} | Qtd: {item.quantidade:<5} | Preço: R$ {item.preco:.2f}")
        
        elif opcao == "6":
            termo = input("Termo de pesquisa: ").strip()
            if not termo:
                print("Termo de pesquisa não pode ser vazio!")
                continue
            
            resultados = sistema.pesquisar_item(termo)
            
            if not resultados:
                print("Nenhum item encontrado!")
            else:
                print(f"\nResultados da pesquisa por '{termo}':")
                print("-" * 60)
                for item in resultados:
                    print(f"{item.nome:<20} | Qtd: {item.quantidade:<5} | Preço: R$ {item.preco:.2f}")
        
        elif opcao == "7":
            nome = input("Nome do item para histórico (Enter para todos): ").strip()
            movimentacoes = sistema.get_historico_movimentacoes(nome if nome else None)
            
            if not movimentacoes:
                print("Nenhuma movimentação encontrada!")
            else:
                print(f"\nHistórico de movimentações{' para ' + nome if nome else ''}:")
                print("-" * 80)
                for mov in movimentacoes[-20:]:  # Mostrar as últimas 20 movimentações
                    tipo_str = "ENTRADA" if mov.tipo == TipoMovimentacao.ENTRADA else "SAÍDA"
                    print(f"{mov.data_registro} | {tipo_str:<7} | {mov.item_nome:<15} | Qtd: {mov.quantidade:<5} | {mov.observacao}")
        
        elif opcao == "8":
            print("\nRealizando balanço físico...")
            itens_fisicos = {}
            
            while True:
                nome = input("Nome do item físico (Enter para finalizar): ").strip()
                if not nome:
                    break
                
                quantidade_str = input("Quantidade física: ").strip()
                quantidade = validar_numero(quantidade_str)
                if quantidade is None or quantidade < 0:
                    print("Quantidade inválida!")
                    continue
                
                itens_fisicos[nome] = quantidade
            
            if not itens_fisicos:
                print("Nenhum item informado para o balanço!")
                continue
            
            observacao = input("Observação sobre o balanço (opcional): ").strip()
            discrepancias = sistema.realizar_balanco(itens_fisicos, observacao)
            
            print("\nBalanço realizado com sucesso!")
            if discrepancias:
                print("Discrepâncias encontradas:")
                for nome, diferenca in discrepancias.items():
                    status = "SOBRANDO" if diferenca > 0 else "FALTANDO"
                    print(f"  {nome}: {abs(diferenca)} unidades {status}")
            else:
                print("Nenhuma discrepância encontrada!")
        
        elif opcao == "9":
            relatorio = sistema.gerar_relatorio_estoque()
            print(f"\nRelatório de Estoque - {relatorio['data_geracao']}")
            print("-" * 40)
            print(f"Total de itens: {relatorio['total_itens']}")
            print(f"Quantidade total: {relatorio['total_quantidade']}")
            print(f"Valor total do estoque: R$ {relatorio['valor_total']:.2f}")
            print(f"Itens com baixa quantidade (menos de 5): {relatorio['itens_baixa_quantidade']}")
        
        elif opcao == "10":
            if not sistema.balancos:
                print("Nenhum balanço realizado!")
            else:
                print(f"\nHistórico de {len(sistema.balancos)} balanço(s):")
                print("-" * 60)
                for i, balanco in enumerate(sistema.balancos[-5:], 1):  # Mostrar os últimos 5
                    print(f"{i}. {balanco.data_registro} - {len(balanco.itens_fisicos)} itens verificados")
                    if balanco.observacao:
                        print(f"   Observação: {balanco.observacao}")
        
        elif opcao == "0":
            print("Saindo do sistema...")
            break
        
        else:
            print("Opção inválida!")

if __name__ == "__main__":
    main()