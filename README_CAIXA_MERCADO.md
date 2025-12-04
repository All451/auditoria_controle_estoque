# Sistema de Caixa para Mercado

Este é um sistema de caixa eletrônico desenvolvido a partir da transformação do sistema de controle de estoque original, adaptado para funcionar como um caixa de mercado com todas as funcionalidades necessárias para vendas e controle de estoque.

## Funcionalidades

### 1. Registro de Vendas
- Adiciona produtos à venda com nome ou código
- Calcula subtotal e total automaticamente
- Aplica descontos
- Escolha de forma de pagamento (dinheiro, cartão de débito, cartão de crédito, Pix)
- Verificação de estoque disponível antes da venda
- Confirmação da venda antes de processar

### 2. Gerenciamento de Produtos
- Adicionar novos produtos com nome, código, preço e quantidade
- Atualizar informações de produtos existentes
- Remover produtos completamente ou reduzir estoque
- Pesquisa por nome ou código do produto

### 3. Controle de Estoque
- Verificação automática de estoque durante vendas
- Alerta de produtos com baixo estoque (menos de 5 unidades)
- Atualização automática do estoque após cada venda

### 4. Relatórios
- Total de vendas registradas
- Valor total recebido
- Vendas do dia
- Total no caixa
- Histórico de vendas

## Como Usar

1. Execute o programa:
```bash
python3 caixa_mercado.py
```

2. Utilize o menu principal para navegar entre as opções:
   - 1: Registrar venda
   - 2: Adicionar/Atualizar produto
   - 3: Listar produtos
   - 4: Pesquisar produto
   - 5: Remover produto
   - 6: Relatório de vendas
   - 7: Verificar estoque
   - 0: Sair

## Persistência de Dados

O sistema salva automaticamente todas as informações em um arquivo JSON chamado `dados_caixa.json`, que mantém:
- Informações dos produtos
- Histórico de vendas
- Total acumulado no caixa

## Estrutura do Sistema

- `SistemaCaixa`: Classe principal que gerencia produtos, vendas e dados do caixa
- `ItemEstoque`: Representa cada produto com nome, quantidade, preço e código
- `Venda`: Representa cada transação de venda com itens, total e forma de pagamento
- Interface amigável com validação de dados

## Benefícios

- Interface intuitiva e fácil de usar
- Controle completo de estoque
- Gestão de vendas e pagamentos
- Relatórios financeiros
- Segurança contra vendas com estoque insuficiente
- Persistência de dados entre sessões

O sistema está pronto para ser utilizado em um ambiente de mercado real, com todas as funcionalidades necessárias para operação de caixa e controle de estoque.