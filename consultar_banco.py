"""
Script auxiliar para consultar e exibir os registros persistidos no banco SQLite.
Mostra a última execução armazenada em banco/rpa_logiexpress.db.
"""

from pathlib import Path
import sqlite3


BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_BANCO = BASE_DIR / "banco" / "rpa_logiexpress.db"


def consultar_ultima_execucao() -> None:
    if not ARQUIVO_BANCO.exists():
        print(f"[AVISO] Banco de dados não encontrado em: {ARQUIVO_BANCO}")
        return

    conn = sqlite3.connect(ARQUIVO_BANCO)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT id_execucao FROM historico_processamento ORDER BY id DESC LIMIT 1"
    )
    row = cursor.fetchone()
    if not row:
        print("[AVISO] Nenhum registro encontrado na tabela historico_processamento.")
        conn.close()
        return

    ultima_execucao = row[0]
    cursor.execute(
        """
        SELECT id_cliente, nome, cpf_formatado, cep_formatado, uf,
               valor_frete, valor_total, status, mensagem_detalhe
        FROM historico_processamento
        WHERE id_execucao = ?
        ORDER BY id ASC
        """,
        (ultima_execucao,),
    )
    registros = cursor.fetchall()
    conn.close()

    sucessos = sum(1 for r in registros if r[7] == "SUCESSO")
    insucessos = len(registros) - sucessos

    print("=" * 110)
    print(f"BANCO SQLITE: {ARQUIVO_BANCO.name} | Última Execução: {ultima_execucao}")
    print(f"Total: {len(registros)} | SUCESSO: {sucessos} | INSUCESSO: {insucessos}")
    print("=" * 110)
    print(f"{'ID':<6} {'Nome':<18} {'CPF':<16} {'CEP':<11} {'UF':<4} {'Frete':<9} {'Total':<10} {'Status':<10} {'Detalhe'}")
    print("-" * 110)

    for r in registros:
        id_cli, nome, cpf, cep, uf, frete, total, status, detalhe = r
        frete_str = f"R${frete:>6.2f}" if frete is not None else "    -    "
        total_str = f"R${total:>7.2f}" if total is not None else "     -    "
        uf_str = uf or "-"
        detalhe_curto = (detalhe[:45] + "...") if len(detalhe) > 48 else detalhe
        print(
            f"{id_cli:<6} {nome:<18} {str(cpf or '-'):<16} {str(cep or '-'):<11} "
            f"{uf_str:<4} {frete_str:<9} {total_str:<10} {status:<10} {detalhe_curto}"
        )
    print("=" * 110)


if __name__ == "__main__":
    consultar_ultima_execucao()
