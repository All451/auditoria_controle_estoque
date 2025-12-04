import os
import tempfile
from sistema_estoque import SistemaEstoque, TipoMovimentacao

def testar_funcionalidades():
    # Criar um sistema temporário para testes
    with tempfile.NamedTemporaryFile(delete=False, suffix='.json') as temp_file:
        temp_filename = temp_file.name
    
    try:
        sistema = SistemaEstoque(temp_filename)
        
        # Testar adição de itens
        print("Testando adição de itens...")
        sistema.adicionar_item("Camiseta", 50, 29.90)
        sistema.adicionar_item("Calça Jeans", 30, 79.90)
        
        assert len(sistema.itens) == 2
        assert sistema.itens["camiseta"].quantidade == 50
        assert sistema.itens["calça jeans"].quantidade == 30
        print("✓ Adição de itens funcionando")
        
        # Testar movimentações
        print("Testando movimentações...")
        sistema.registrar_saida("Camiseta", 5, "Venda")
        sistema.registrar_entrada("Camiseta", 10, "Reposição")
        
        assert sistema.itens["camiseta"].quantidade == 55  # 50 - 5 + 10
        assert len(sistema.movimentacoes) == 4  # 2 adições iniciais + 2 movimentações
        print("✓ Movimentações funcionando")
        
        # Testar balanço
        print("Testando balanço...")
        itens_fisicos = {
            "Camiseta": 52,  # Esperado: 55, encontrado: 52 (3 a menos)
            "Calça Jeans": 30  # Mesma quantidade
        }
        
        discrepancias = sistema.realizar_balanco(itens_fisicos, "Teste de balanço")
        
        assert sistema.itens["camiseta"].quantidade == 52  # Agora com valor físico
        assert "camiseta" in discrepancias
        assert discrepancias["camiseta"] == -3  # 3 a menos
        print("✓ Balanço funcionando")
        
        # Testar relatórios
        print("Testando relatórios...")
        relatorio = sistema.gerar_relatorio_estoque()
        assert relatorio["total_itens"] == 2
        assert relatorio["total_quantidade"] == 82  # 52 (camiseta) + 30 (calça jeans)
        print("✓ Relatórios funcionando")
        
        # Testar persistência
        print("Testando persistência...")
        sistema2 = SistemaEstoque(temp_filename)
        assert len(sistema2.itens) == 2
        assert sistema2.itens["camiseta"].quantidade == 52
        print("✓ Persistência funcionando")
        
        print("\n✓ Todos os testes passaram!")
        
    finally:
        # Limpar arquivo temporário
        if os.path.exists(temp_filename):
            os.remove(temp_filename)

if __name__ == "__main__":
    testar_funcionalidades()