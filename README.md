# Sistema de auditoria de Estoque

Este projeto é um sistema de controle de estoque desenvolvido em Python. O sistema permite gerenciar e controlar o estoque de forma eficiente, oferecendo funcionalidades para adicionar itens, listar itens, importar e exportar dados, e realizar balanços de estoque.

## Funcionalidades

- **Adicionar Itens ao Estoque:** Permite adicionar novos itens ao estoque ou atualizar a quantidade de itens existentes.
- **Listar Itens no Estoque:** Exibe uma lista de todos os itens no estoque com suas quantidades atuais.
- **Importar Dados de Estoque:** Importa dados de um arquivo CSV para atualizar o estoque.
- **Realizar Balanço de Estoque:** Compara as quantidades físicas dos itens com as quantidades registradas no sistema e apura as diferenças.
- **Exportar Balanço para Arquivo Excel:** Exporta o balanço de estoque para um arquivo Excel para análise e relatórios.

## Como Usar

1. **Adicionar Itens ao Estoque:**
   - Escolha a opção "1" no menu principal.
   - Informe o nome do item e a quantidade desejada.

2. **Listar Itens no Estoque:**
   - Escolha a opção "2" no menu principal para visualizar todos os itens e suas quantidades.

3. **Importar Dados de Estoque:**
   - Escolha a opção "3" no menu principal.
   - Forneça o caminho para o arquivo CSV contendo os dados do estoque.

4. **Realizar Balanço de Estoque:**
   - Escolha a opção "4" no menu principal.
   - Informe a quantidade física dos itens quando solicitado.

5. **Exportar Balanço para Arquivo Excel:**
   - Escolha a opção "5" no menu principal.
   - Forneça o caminho para o arquivo Excel onde o balanço será salvo.

## Requisitos

Para executar o projeto, você precisa do Python 3.x e das seguintes bibliotecas Python:

- `pandas`
- `tabulate`
- `openpyxl`

Para instalar as dependências, execute o seguinte comando:

```bash
pip install pandas tabulate openpyxl

