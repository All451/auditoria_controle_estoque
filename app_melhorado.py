import pandas as pd
from tabulate import tabulate
import os
import json
import csv
from openpyxl import Workbook
from typing import Dict, Any, Optional, List
from datetime import datetime
import re


class ItemEstoque:
    """Classe que representa um item no estoque."""
    
    def __init__(self, nome: str, quantidade: float = 0):
        self.nome = nome.strip()
        self.quantidade = max(0, float(quantidade))
    
    def __str__(self):
        return f"{self.nome}: {self.quantidade}"
    
    def to_dict(self):
        return {
            "nome": self.nome,
            "quantidade": self.quantidade
        }


class BalancoItem:
    """Classe que representa um item no balanço de estoque."""
    
    def __init__(self, nome: str, quantidade_sistema: float, quantidade_fisica: float):
        self.nome = nome
        self.quantidade_sistema = float(quantidade_sistema)
        self.quantidade_fisica = float(quantidade_fisica)
        self.diferenca = self.quantidade_fisica - self.quantidade_sistema
        self.status = self._calcular_status()
    
    def _calcular_status(self) -> str:
        if self.quantidade_fisica == 0 and self.quantidade_sistema == 0:
            return "OK"
        elif self.quantidade_fisica < self.quantidade_sistema:
            return "baixa"
        elif self.quantidade_fisica > self.quantidade_sistema:
            return "retorno"
        else:
            return "OK"
    
    def to_dict(self):
        return {
            "item": self.nome,
            "quantidade_sistema": self.quantidade_sistema,
            "quantidade_fisica": self.quantidade_fisica,
            "diferenca": self.diferenca,
            "status": self.status
        }


class Estoque:
    """Classe principal para gerenciar o estoque."""
    
    def __init__(self, arquivo_dados: str = "estoque.json") -> None:
        """
        Inicializa os dicionários para itens e balanço do estoque.
        
        Args:
            arquivo_dados (str): Caminho para o arquivo de persistência de dados
        """
        self.itens: Dict[str, ItemEstoque] = {}
        self.historico_balanco: List[Dict[str, Any]] = []
        self.arquivo_dados = arquivo_dados
        self.carregar_dados()
    
    def carregar_dados(self) -> None:
        """Carrega os dados do estoque a partir de um arquivo JSON."""
        try:
            if os.path.exists(self.arquivo_dados):
                with open(self.arquivo_dados, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                    for item_data in dados.get('itens', []):
                        item = ItemEstoque(item_data['nome'], item_data['quantidade'])
                        self.itens[item.nome] = item
                    self.historico_balanco = dados.get('historico_balanco', [])
                print("Dados carregados com sucesso.")
        except Exception as e:
            print(f"Erro ao carregar dados: {e}")
    
    def salvar_dados(self) -> None:
        """Salva os dados do estoque em um arquivo JSON."""
        try:
            dados = {
                'itens': [item.to_dict() for item in self.itens.values()],
                'historico_balanco': self.historico_balanco,
                'data_ultima_atualizacao': datetime.now().isoformat()
            }
            with open(self.arquivo_dados, 'w', encoding='utf-8') as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Erro ao salvar dados: {e}")
    
    def validar_nome_item(self, nome: str) -> bool:
        """Valida o nome do item."""
        if not nome or not nome.strip():
            print("Erro: O nome do item não pode ser vazio.")
            return False
        if len(nome.strip()) < 2:
            print("Erro: O nome do item deve ter pelo menos 2 caracteres.")
            return False
        return True
    
    def adicionar_item(self, nome: str, quantidade: float) -> bool:
        """
        Adiciona um item ao estoque com a quantidade especificada.
        
        Args:
            nome (str): Nome do item.
            quantidade (float): Quantidade do item.
            
        Returns:
            bool: True se a operação foi bem-sucedida, False caso contrário.
        """
        if not self.validar_nome_item(nome):
            return False
            
        if quantidade <= 0:
            print("Erro: A quantidade deve ser um número positivo.")
            return False
        
        nome = nome.strip().lower()
        
        if nome in self.itens:
            self.itens[nome].quantidade += quantidade
        else:
            self.itens[nome] = ItemEstoque(nome, quantidade)
        
        print(f"Adicionado {quantidade} de {nome} ao estoque.")
        self.salvar_dados()
        return True
    
    def remover_item(self, nome: str, quantidade: float = None) -> bool:
        """
        Remove um item do estoque ou uma quantidade específica.
        
        Args:
            nome (str): Nome do item.
            quantidade (float, optional): Quantidade a remover. Se None, remove o item completamente.
            
        Returns:
            bool: True se a operação foi bem-sucedida, False caso contrário.
        """
        nome = nome.strip().lower()
        
        if nome not in self.itens:
            print(f"Erro: Item '{nome}' não encontrado no estoque.")
            return False
        
        if quantidade is None:
            # Remover item completamente
            del self.itens[nome]
            print(f"Item '{nome}' removido do estoque.")
        else:
            if quantidade <= 0:
                print("Erro: A quantidade deve ser um número positivo.")
                return False
            
            if quantidade >= self.itens[nome].quantidade:
                # Remover item completamente se quantidade a remover for maior ou igual
                del self.itens[nome]
                print(f"Item '{nome}' removido completamente do estoque.")
            else:
                # Reduzir a quantidade
                self.itens[nome].quantidade -= quantidade
                print(f"Removido {quantidade} de {nome} do estoque.")
        
        self.salvar_dados()
        return True
    
    def listar_itens(self, ordenar_por: str = "nome") -> None:
        """
        Lista todos os itens presentes no estoque e suas quantidades.
        
        Args:
            ordenar_por (str): Critério de ordenação ("nome", "quantidade", "quantidade_desc")
        """
        if not self.itens:
            print("O estoque está vazio.")
            return
        
        itens_lista = list(self.itens.values())
        
        # Ordenar conforme critério
        if ordenar_por == "quantidade":
            itens_lista.sort(key=lambda x: x.quantidade)
        elif ordenar_por == "quantidade_desc":
            itens_lista.sort(key=lambda x: x.quantidade, reverse=True)
        else:  # ordenar_por == "nome"
            itens_lista.sort(key=lambda x: x.nome)
        
        tabela = [[item.nome, item.quantidade] for item in itens_lista]
        print(tabulate(tabela, headers=["Item", "Quantidade"], tablefmt="grid"))
    
    def obter_item(self, nome: str) -> Optional[ItemEstoque]:
        """
        Obtém um item do estoque pelo nome.
        
        Args:
            nome (str): Nome do item
            
        Returns:
            Optional[ItemEstoque]: O item se encontrado, None caso contrário
        """
        nome = nome.strip().lower()
        return self.itens.get(nome)
    
    def importar_estoque(self, caminho_arquivo: str, delimitador: str = ',') -> bool:
        """
        Importa dados de estoque de um arquivo CSV.
        
        Args:
            caminho_arquivo (str): Caminho do arquivo CSV a ser importado.
            delimitador (str): Delimitador do arquivo CSV (padrão: ',')
            
        Returns:
            bool: True se a operação foi bem-sucedida, False caso contrário.
        """
        try:
            if not os.path.isfile(caminho_arquivo):
                print("Erro: Arquivo não encontrado.")
                return False

            # Lê o arquivo CSV usando pandas
            df = pd.read_csv(caminho_arquivo, sep=delimitador, encoding='utf-8')
            
            # Normaliza os nomes das colunas
            df.columns = df.columns.str.lower().str.strip()
            
            # Verifica se as colunas necessárias estão presentes
            colunas_necessarias = ['item', 'quantidade']
            if not all(col in df.columns for col in colunas_necessarias):
                print(f"Erro: O arquivo não contém as colunas necessárias: {colunas_necessarias}")
                return False
            
            itens_importados = 0
            for index, row in df.iterrows():
                nome = str(row['item']).strip()
                try:
                    quantidade = float(row['quantidade'])
                    if self.adicionar_item(nome, quantidade):
                        itens_importados += 1
                except ValueError:
                    print(f"Erro na linha {index+1}: Quantidade '{row['quantidade']}' não é um número válido.")
            
            print(f"Importação concluída. {itens_importados} itens importados com sucesso.")
            return True
            
        except pd.errors.EmptyDataError:
            print("Erro: O arquivo CSV está vazio.")
            return False
        except Exception as e:
            print(f"Erro ao importar estoque: {e}")
            return False
    
    def exportar_estoque(self, caminho_arquivo: str) -> bool:
        """
        Exporta os dados do estoque para um arquivo CSV.
        
        Args:
            caminho_arquivo (str): Caminho do arquivo CSV para exportação.
            
        Returns:
            bool: True se a operação foi bem-sucedida, False caso contrário.
        """
        try:
            if not self.itens:
                print("O estoque está vazio. Nada para exportar.")
                return False
            
            # Criar DataFrame com os dados do estoque
            dados = {
                'item': [item.nome for item in self.itens.values()],
                'quantidade': [item.quantidade for item in self.itens.values()]
            }
            df = pd.DataFrame(dados)
            
            # Exportar para CSV
            df.to_csv(caminho_arquivo, index=False, encoding='utf-8')
            print(f"Estoque exportado para {caminho_arquivo}")
            return True
            
        except Exception as e:
            print(f"Erro ao exportar estoque: {e}")
            return False
    
    def balanca_estoque(self) -> None:
        """Realiza o balanço do estoque comparando quantidades físicas e no sistema."""
        if not self.itens:
            print("O estoque está vazio.")
            return

        balanco_atual = []
        print("\nIniciando balanço de estoque...")
        print("Digite as quantidades físicas para cada item:")

        for nome, item in self.itens.items():
            while True:
                try:
                    # Solicita a quantidade física do item
                    quantidade_fisica_input = input(f"Quantidade física de '{nome}' (atual: {item.quantidade}): ").strip()
                    
                    # Permitir sair do balanço digitando 'sair'
                    if quantidade_fisica_input.lower() in ['sair', 's']:
                        print("Balanço interrompido pelo usuário.")
                        return
                    
                    quantidade_fisica = float(quantidade_fisica_input)
                    break
                except ValueError:
                    print("Erro: A quantidade deve ser um número (inteiro ou decimal). Digite 'sair' para cancelar.")
            
            balanco_item = BalancoItem(nome, item.quantidade, quantidade_fisica)
            balanco_atual.append(balanco_item)
            
            # Exibir resultado parcial para o item
            status_simbolo = {"OK": "✓", "baixa": "↓", "retorno": "↑"}[balanco_item.status]
            print(f"  → {status_simbolo} {balanco_item.nome}: Sistema={balanco_item.quantidade_sistema}, Físico={balanco_item.quantidade_fisica}, Diferença={balanco_item.diferenca}")

        # Salvar balanço no histórico
        self.historico_balanco.append({
            "data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "balanco": [item.to_dict() for item in balanco_atual]
        })
        
        print("\nBalanço de estoque concluído!")
        self.mostrar_balanco(balanco_atual)
    
    def mostrar_balanco(self, balanco: List[BalancoItem] = None) -> None:
        """
        Mostra o balanço atual do estoque.
        
        Args:
            balanco (List[BalancoItem], optional): Lista específica de itens para exibir. Se None, exibe o último balanço.
        """
        if balanco is None:
            if not self.historico_balanco:
                print("Nenhum balanço foi realizado.")
                return
            
            # Usar o último balanço realizado
            balanco = [BalancoItem(item['item'], item['quantidade_sistema'], item['quantidade_fisica']) 
                      for item in self.historico_balanco[-1]['balanco']]
        
        tabela = []
        for item in balanco:
            status_simbolo = {"OK": "✓", "baixa": "↓", "retorno": "↑"}[item.status]
            tabela.append([
                item.nome, 
                item.quantidade_sistema, 
                item.quantidade_fisica, 
                item.diferenca, 
                f"{status_simbolo} {item.status}"
            ])

        print(tabulate(tabela, 
                      headers=["Item", "Quantidade no Sistema", "Quantidade Física", "Diferença", "Status"], 
                      tablefmt="grid"))
    
    def mostrar_historico_balanco(self) -> None:
        """Mostra o histórico de balanços realizados."""
        if not self.historico_balanco:
            print("Nenhum balanço foi realizado até o momento.")
            return
        
        print(f"\nHistórico de {len(self.historico_balanco)} balanço(s) realizado(s):")
        for i, balanco in enumerate(self.historico_balanco, 1):
            print(f"{i}. Data: {balanco['data']} - Itens: {len(balanco['balanco'])}")
    
    def exportar_balanco(self, caminho_arquivo: str) -> bool:
        """
        Exporta o último balanço do estoque para um arquivo Excel.
        
        Args:
            caminho_arquivo (str): Caminho do arquivo Excel para exportação.
            
        Returns:
            bool: True se a operação foi bem-sucedida, False caso contrário.
        """
        if not self.historico_balanco:
            print("Nenhum balanço foi realizado.")
            return False

        try:
            # Pegar o último balanço
            ultimo_balanco = self.historico_balanco[-1]['balanco']
            
            wb = Workbook()
            ws = wb.active
            ws.title = f"Balanço {datetime.now().strftime('%Y-%m-%d')}"
            ws.append(["Item", "Quantidade no Sistema", "Quantidade Física", "Diferença", "Status", "Data do Balanço"])
            
            data_balanco = self.historico_balanco[-1]['data']
            
            for item in ultimo_balanco:
                ws.append([
                    item['item'], 
                    item['quantidade_sistema'], 
                    item['quantidade_fisica'], 
                    item['diferenca'], 
                    item['status'],
                    data_balanco
                ])

            wb.save(caminho_arquivo)
            print(f"Balanço exportado para {caminho_arquivo}")
            return True

        except Exception as e:
            print(f"Erro ao exportar balanço: {e}")
            return False
    
    def gerar_relatorio_estoque(self) -> str:
        """Gera um relatório detalhado do estoque."""
        if not self.itens:
            return "O estoque está vazio."
        
        total_itens = len(self.itens)
        total_quantidade = sum(item.quantidade for item in self.itens.values())
        itens_sem_estoque = sum(1 for item in self.itens.values() if item.quantidade == 0)
        item_maior_quantidade = max(self.itens.values(), key=lambda x: x.quantidade)
        item_menor_quantidade = min(self.itens.values(), key=lambda x: x.quantidade)
        
        relatorio = f"""
RELATÓRIO DE ESTOQUE
====================
Data: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Total de itens: {total_itens}
Quantidade total: {total_quantidade}
Itens sem estoque: {itens_sem_estoque}
Item com maior quantidade: {item_maior_quantidade.nome} ({item_maior_quantidade.quantidade})
Item com menor quantidade: {item_menor_quantidade.nome} ({item_menor_quantidade.quantidade})

Itens em estoque:
"""
        
        # Adiciona a lista de itens
        itens_lista = sorted(self.itens.values(), key=lambda x: x.quantidade, reverse=True)
        for item in itens_lista:
            relatorio += f"  - {item.nome}: {item.quantidade}\n"
        
        return relatorio


def validar_entrada_numero(mensagem: str, tipo: type = float, positivo: bool = True) -> Optional[float]:
    """
    Valida a entrada de um número do usuário.
    
    Args:
        mensagem (str): Mensagem para solicitar a entrada
        tipo (type): Tipo do número (int ou float)
        positivo (bool): Se o número deve ser positivo
        
    Returns:
        Optional[float]: Número validado ou None se inválido
    """
    while True:
        try:
            entrada = input(mensagem).strip()
            if not entrada:
                print("Entrada vazia. Por favor, informe um valor.")
                continue
                
            numero = tipo(entrada)
            
            if positivo and numero <= 0:
                print(f"Erro: O valor deve ser um número positivo.")
                continue
                
            return numero
        except ValueError:
            print(f"Erro: A entrada deve ser um número {'positivo' if positivo else ''} válido.")


def main() -> None:
    """Função principal que gerencia o controle de estoque."""
    estoque = Estoque()
    
    print("Sistema de Auditoria de Estoque")
    print("=" * 40)
    
    while True:
        print("\nControle de Estoque:")
        print("1. Adicionar Item")
        print("2. Remover Item")
        print("3. Listar Itens")
        print("4. Obter Informações de Item")
        print("5. Importar Estoque de Arquivo CSV")
        print("6. Exportar Estoque para Arquivo CSV")
        print("7. Realizar Balanço de Estoque")
        print("8. Visualizar Histórico de Balanço")
        print("9. Exportar Balanço para Arquivo Excel")
        print("10. Gerar Relatório de Estoque")
        print("0. Sair")
        
        escolha = input("\nEscolha uma opção: ").strip()
        
        if escolha == '1':
            nome = input("Nome do item: ").strip()
            if nome:
                quantidade = validar_entrada_numero("Quantidade: ", float, True)
                if quantidade is not None:
                    estoque.adicionar_item(nome, quantidade)
            else:
                print("Erro: O nome do item não pode ser vazio.")
        
        elif escolha == '2':
            nome = input("Nome do item a remover: ").strip()
            if nome:
                remover_quantidade = input("Deseja remover quantidade específica? (s/n): ").strip().lower()
                if remover_quantidade == 's':
                    quantidade = validar_entrada_numero("Quantidade a remover: ", float, True)
                    if quantidade is not None:
                        estoque.remover_item(nome, quantidade)
                else:
                    confirmacao = input(f"Tem certeza que deseja remover completamente o item '{nome}'? (s/n): ").strip().lower()
                    if confirmacao == 's':
                        estoque.remover_item(nome)
                    else:
                        print("Remoção cancelada.")
            else:
                print("Erro: O nome do item não pode ser vazio.")
        
        elif escolha == '3':
            print("\nOpções de ordenação:")
            print("1. Por nome")
            print("2. Por quantidade (crescente)")
            print("3. Por quantidade (decrescente)")
            ordem = input("Escolha a ordenação (1-3): ").strip()
            
            ordem_map = {'1': 'nome', '2': 'quantidade', '3': 'quantidade_desc'}
            ordenar_por = ordem_map.get(ordem, 'nome')
            
            estoque.listar_itens(ordenar_por)
        
        elif escolha == '4':
            nome = input("Nome do item: ").strip()
            if nome:
                item = estoque.obter_item(nome)
                if item:
                    print(f"Informações do item '{item.nome}':")
                    print(f"  Quantidade: {item.quantidade}")
                else:
                    print(f"Item '{nome}' não encontrado no estoque.")
            else:
                print("Erro: O nome do item não pode ser vazio.")
        
        elif escolha == '5':
            caminho_arquivo = input("Digite o caminho do arquivo CSV: ").strip()
            if caminho_arquivo:
                delimitador = input("Delimitador do CSV (padrão: ','): ").strip() or ','
                estoque.importar_estoque(caminho_arquivo, delimitador)
            else:
                print("Erro: O caminho do arquivo não pode ser vazio.")
        
        elif escolha == '6':
            caminho_arquivo = input("Digite o caminho do arquivo CSV para exportação: ").strip()
            if caminho_arquivo:
                estoque.exportar_estoque(caminho_arquivo)
            else:
                print("Erro: O caminho do arquivo não pode ser vazio.")
        
        elif escolha == '7':
            estoque.balanca_estoque()
        
        elif escolha == '8':
            estoque.mostrar_historico_balanco()
        
        elif escolha == '9':
            caminho_arquivo = input("Digite o caminho do arquivo Excel para exportação: ").strip()
            if caminho_arquivo:
                estoque.exportar_balanco(caminho_arquivo)
            else:
                print("Erro: O caminho do arquivo não pode ser vazio.")
        
        elif escolha == '10':
            print(estoque.gerar_relatorio_estoque())
        
        elif escolha == '0':
            print("Saindo do sistema...")
            break
        
        else:
            print("Opção inválida. Tente novamente.")


if __name__ == "__main__":
    main()