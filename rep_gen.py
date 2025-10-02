from rep_tools import (
    load_data,
    show_menu,
    MONTHS,
    relatorio_mensal_geral,
    relatorio_anual_geral,
    relatorio_paciente_mes,
    relatorio_paciente_anual,
    relatorio_totais_por_paciente,
    relatorio_personalizado_por_periodo
)
from datetime import datetime

def main():
    dados = load_data()

    if not dados:
        print("❌ Nenhum dado carregado. Verifique o arquivo Excel.")
        return

    while True:
        show_menu()
        opcao = input("\nEscolha uma opção: ")

        if opcao == "1":
            # Relatório do mês atual
            hoje = datetime.today()
            mes = hoje.month
            ano = hoje.year
            print(f"\n📅 Gerando relatório geral de {MONTHS[mes]} de {ano}...")
            relatorio_mensal_geral(dados, mes, ano)

        elif opcao == "2":
            # Relatório de um paciente (mensal ou anual)
            paciente = input("Nome do paciente: ")
            esc = input("Deseja relatório [1] Mensal ou [2] Anual? ")

            if esc == "1":
                mes = int(input("Mês (1-12): "))
                ano = int(input("Ano: "))
                relatorio_paciente_mes(dados, paciente, mes, ano)
            elif esc == "2":
                ano = int(input("Ano: "))
                relatorio_paciente_anual(dados, paciente, ano)
            else:
                print("⚠️ Opção inválida.")

        elif opcao == "3":
            # Relatório anual geral
            ano = int(input("Ano: "))
            relatorio_anual_geral(dados, ano)

        elif opcao == "4":
            # Totais por paciente no ano
            ano = int(input("Ano: "))
            relatorio_totais_por_paciente(dados, ano)

        elif opcao == "5":
            # Relatório personalizado por intervalo
            data_inicio = input("Data inicial (YYYY-MM-DD): ")
            data_fim = input("Data final (YYYY-MM-DD): ")
            filtrar = input("Deseja filtrar por paciente específico? (s/n): ").lower()

            if filtrar == "s":
                paciente = input("Nome do paciente: ")
                relatorio_personalizado_por_periodo(dados, data_inicio, data_fim, paciente)
            else:
                relatorio_personalizado_por_periodo(dados, data_inicio, data_fim)

        elif opcao == "6":
            print("Encerrando o programa.")
            break

        else:
            print("❌ Opção inválida. Tente novamente.")

if __name__ == "__main__":
    main()
