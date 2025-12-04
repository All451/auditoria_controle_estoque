# Sistema de Auditoria de Estoque - Versão Aprimorada

Este projeto é uma versão aprimorada do sistema de controle de estoque desenvolvido em Python. O sistema permite gerenciar e controlar o estoque de forma eficiente, oferecendo funcionalidades avançadas para adicionar itens, remover itens, listar itens, importar e exportar dados, realizar balanços de estoque e gerar relatórios detalhados.

## Melhorias Implementadas

### 1. Estrutura de Código Aprimorada
- **Classes Especializadas**: Criadas classes `ItemEstoque` e `BalancoItem` para representar entidades do domínio
- **Separação de Responsabilidades**: Melhor organização do código com métodos mais específicos
- **Tipagem de Dados**: Uso de type hints para melhor legibilidade e manutenção

### 2. Persistência de Dados
- **Armazenamento em JSON**: Dados do estoque são salvos automaticamente em arquivo JSON
- **Recuperação Automática**: Sistema carrega dados salvos ao iniciar

### 3. Validações e Tratamento de Erros
- **Validação de Entrada**: Verificação mais robusta de nomes de itens e quantidades
- **Tratamento de Exceções**: Melhor tratamento de erros com mensagens claras
- **Entradas Numéricas Validadas**: Função auxiliar para validação de números

### 4. Novas Funcionalidades
- **Remoção de Itens**: Possibilidade de remover itens ou quantidades específicas
- **Ordenação de Listagem**: Opções para ordenar itens por nome ou quantidade
- **Histórico de Balanços**: Armazenamento de múltiplos balanços com datas
- **Relatórios Detalhados**: Geração de relatórios com estatísticas do estoque
- **Exportação de Estoque**: Exportar dados do estoque para CSV

### 5. Melhorias na Experiência do Usuário
- **Interface Mais Intuitiva**: Menu com mais opções e navegação melhorada
- **Feedback em Tempo Real**: Informações detalhadas durante o balanço
- **Símbolos de Status**: Indicadores visuais para status do balanço (✓, ↓, ↑)
- **Cancelamento de Operações**: Possibilidade de interromper operações

## Funcionalidades

- **Adicionar Itens ao Estoque**: Permite adicionar novos itens ao estoque ou atualizar a quantidade de itens existentes.
- **Remover Itens do Estoque**: Remove itens ou quantidades específicas do estoque.
- **Listar Itens no Estoque**: Exibe uma lista de todos os itens no estoque com suas quantidades atuais, com opções de ordenação.
- **Obter Informações de Item**: Visualiza informações detalhadas de um item específico.
- **Importar Dados de Estoque**: Importa dados de um arquivo CSV para atualizar o estoque.
- **Exportar Dados de Estoque**: Exporta dados do estoque para um arquivo CSV.
- **Realizar Balanço de Estoque**: Compara as quantidades físicas dos itens com as quantidades registradas no sistema e apura as diferenças.
- **Visualizar Histórico de Balanços**: Consulta todos os balanços realizados com datas e quantidades.
- **Exportar Balanço para Arquivo Excel**: Exporta o último balanço de estoque para um arquivo Excel para análise e relatórios.
- **Gerar Relatório de Estoque**: Cria um relatório detalhado com estatísticas e informações do estoque.

## Como Usar

1. **Adicionar Itens ao Estoque:**
   - Escolha a opção "1" no menu principal.
   - Informe o nome do item e a quantidade desejada.

2. **Remover Itens do Estoque:**
   - Escolha a opção "2" no menu principal.
   - Informe o nome do item e escolha entre remover quantidade específica ou o item completamente.

3. **Listar Itens no Estoque:**
   - Escolha a opção "3" no menu principal para visualizar todos os itens e suas quantidades.
   - Escolha o critério de ordenação (por nome, por quantidade crescente ou decrescente).

4. **Obter Informações de Item:**
   - Escolha a opção "4" no menu principal.
   - Informe o nome do item para ver detalhes específicos.

5. **Importar Dados de Estoque:**
   - Escolha a opção "5" no menu principal.
   - Forneça o caminho para o arquivo CSV contendo os dados do estoque.
   - Especifique o delimitador do CSV se for diferente de vírgula.

6. **Exportar Dados de Estoque:**
   - Escolha a opção "6" no menu principal.
   - Forneça o caminho para o arquivo CSV onde os dados serão salvos.

7. **Realizar Balanço de Estoque:**
   - Escolha a opção "7" no menu principal.
   - Informe a quantidade física dos itens quando solicitado.
   - O sistema mostrará o status de cada item (OK, baixa ou retorno).

8. **Visualizar Histórico de Balanços:**
   - Escolha a opção "8" no menu principal.
   - Consulte todos os balanços realizados com datas e quantidades.

9. **Exportar Balanço para Arquivo Excel:**
   - Escolha a opção "9" no menu principal.
   - Forneça o caminho para o arquivo Excel onde o último balanço será salvo.

10. **Gerar Relatório de Estoque:**
    - Escolha a opção "10" no menu principal.
    - Visualize um relatório detalhado com estatísticas do estoque.

## Requisitos

Para executar o projeto, você precisa do Python 3.x e das seguintes bibliotecas Python:

- `pandas`
- `tabulate`
- `openpyxl`

Para instalar as dependências, execute o seguinte comando:

```bash
pip install pandas tabulate openpyxl
```

## Execução

Para executar o sistema, utilize o seguinte comando:

```bash
python app_melhorado.py
```

## Estrutura do Projeto

- `app_melhorado.py`: Arquivo principal com a implementação do sistema aprimorado
- `estoque.json`: Arquivo de persistência automática dos dados do estoque
- `README.md`: Documentação das melhorias implementadas

## Considerações Finais

O sistema aprimorado oferece uma experiência mais completa e robusta para o controle e auditoria de estoque, com foco na usabilidade, confiabilidade e manutenibilidade do código.