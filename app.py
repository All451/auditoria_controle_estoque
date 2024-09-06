import pandas as pd
from tabulate import tabulate
import os
from openpyxl import Workbook
from typing import Dict, Any

class Estoque:
    def __init__(self) -> None:
        """
        Inicializa os dicionários para itens e balanço do estoque.
        """
        self.itens: Dict[str, float] = {}  # Dicionário para armazenar itens e suas quantidades
        self.balanco: Dict[str, Dict[str, Any]] = {}  # Dicionário para armazenar o balanço do estoque

    def adicionar_item(self, nome: str, quantidade: float) -> None:
        """
        Adiciona um item ao estoque com a quantidade especificada.

        Args:
            nome (str): Nome do item.
            quantidade (float): Quantidade do item.
        """
        if quantidade <= 0:
            print("Erro: A quantidade deve ser um número positivo.")
            return
        if nome in self.itens:
            self.itens[nome] += quantidade
        else:
            self.itens[nome] = quantidade
        print(f"Adicionado {quantidade} de {nome} ao estoque.")

    def listar_itens(self) -> None:
        """
        Lista todos os itens presentes no estoque e suas quantidades.
        """
        if not self.itens:
            print("O estoque está vazio.")
        else:
            tabela = [[nome, quantidade] for nome, quantidade in self.itens.items()]
            print(tabulate(tabela, headers=["Item", "Quantidade"], tablefmt="grid"))

    def importar_estoque(self, caminho_arquivo: str) -> None:
        """
        Importa dados de estoque de um arquivo CSV.

        Args:
            caminho_arquivo (str): Caminho do arquivo CSV a ser importado.
        """
        try:
            if not os.path.isfile(caminho_arquivo):
                raise FileNotFoundError("Arquivo não encontrado.")

            # Lê o arquivo CSV usando pandas
            df = pd.read_csv(caminho_arquivo, sep=',', encoding='utf-8')
            df.columns = df.columns.str.lower().str.strip()

            # Verifica se as colunas necessárias estão presentes
            if 'item' in df.columns and 'quantidade' in df.columns:
                for index, row in df.iterrows():
                    nome = row['item']
                    quantidade = row['quantidade']
                    if isinstance(quantidade, (int, float)):
                        self.adicionar_item(nome, float(quantidade))
                    else:
                        print(f"Erro na linha {index+1}: Quantidade não é um número.")
                print("Estoque importado com sucesso.")
            else:
                print("Erro: O arquivo não contém as colunas 'item' e 'quantidade'.")
        except FileNotFoundError:
            print("Erro: Arquivo não encontrado.")
        except ValueError as ve:
            print(f"Erro de valor: {ve}")
        except Exception as e:
            print(f"Erro ao importar estoque: {e}")

    def balanca_estoque(self) -> None:
        """
        Realiza o balanço do estoque comparando quantidades físicas e no sistema.
        """
        if not self.itens:
            print("O estoque está vazio.")
            return

        for nome in self.itens.keys():
            while True:
                try:
                    # Solicita a quantidade física do item
                    quantidade_fisica = float(input(f"Quantidade física de {nome}: "))
                    break
                except ValueError:
                    print("Erro: A quantidade deve ser um número (inteiro ou decimal).")

            quantidade_sistema = self.itens[nome]
            diferenca = quantidade_fisica - quantidade_sistema
            self.balanco[nome] = {
                "quantidade_sistema": quantidade_sistema,
                "quantidade_fisica": quantidade_fisica,
                "diferenca": diferenca
            }

        print("Balanço de estoque atualizado.")
        self.mostrar_balanco()

    def mostrar_balanco(self) -> None:
        """
        Mostra o balanço atual do estoque com o status conforme especificado.
        """
        if not self.balanco:
            print("Nenhum balanço foi realizado.")
        else:
            tabela = []
            for nome, info in self.balanco.items():
                status = ""
                if info["quantidade_fisica"] == 0:
                    status = "OK"
                elif info["quantidade_fisica"] < info["quantidade_sistema"]:
                    status = "baixa"
                elif info["quantidade_fisica"] > info["quantidade_sistema"]:
                    status = "retorno"

                tabela.append([nome, info["quantidade_sistema"], info["quantidade_fisica"], info["diferenca"], status])

            print(tabulate(tabela, headers=["Item", "Quantidade no Sistema", "Quantidade Física", "Diferença", "Status"], tablefmt="grid"))

    def exportar_balanco(self, caminho_arquivo: str) -> None:
        """
        Exporta o balanço do estoque para um arquivo Excel.

        Args:
            caminho_arquivo (str): Caminho do arquivo Excel para exportação.
        """
        if not self.balanco:
            print("Nenhum balanço foi realizado.")
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.append(["Item", "Quantidade no Sistema", "Quantidade Física", "Diferença", "Status"])

            for nome, info in self.balanco.items():
                status = ""
                if info["quantidade_fisica"] == 0:
                    status = "OK"
                elif info["quantidade_fisica"] < info["quantidade_sistema"]:
                    status = "baixa"
                elif info["quantidade_fisica"] > info["quantidade_sistema"]:
                    status = "retorno"

                ws.append([nome, info["quantidade_sistema"], info["quantidade_fisica"], info["diferenca"], status])

            wb.save(caminho_arquivo)
            print(f"Balanço exportado para {caminho_arquivo}")

        except Exception as e:
            print(f"Erro ao exportar balanço: {e}")

def main() -> None:
    """
    Função principal que gerencia o controle de estoque.
    """
    estoque = Estoque()
    while True:
        print("\nControle de Estoque:")
        print("1. Adicionar Item")
        print("2. Listar Itens")
        print("3. Importar Estoque de Arquivo CSV")
        print("4. Realizar Balanço de Estoque")
        print("5. Exportar Balanço para Arquivo Excel")
        print("6. Sair")

        escolha = input("Escolha uma opção: ")

        if escolha == '1':
            try:
                nome = input("Nome do item: ").strip()
                quantidade = float(input("Quantidade: "))
                if quantidade <= 0:
                    raise ValueError("A quantidade deve ser um número positivo.")
                estoque.adicionar_item(nome, quantidade)
            except ValueError as ve:
                print(f"Erro: {ve}")

        elif escolha == '2':
            estoque.listar_itens()

        elif escolha == '3':
            caminho_arquivo = input("Digite o caminho do arquivo CSV: ").strip()
            estoque.importar_estoque(caminho_arquivo)

        elif escolha == '4':
            estoque.balanca_estoque()

        elif escolha == '5':
            caminho_arquivo = input("Digite o caminho do arquivo Excel para exportação: ").strip()
            estoque.exportar_balanco(caminho_arquivo)

        elif escolha == '6':
            print("Saindo...")
            break

        else:
            print("Opção inválida. Tente novamente.")

if __name__ == "__main__":
    main()
