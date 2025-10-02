import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from inv_tools import date_converter, comma_check
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill


arquivo_excel = "notas_fiscais.xlsx"

def file_create():

    if os.path.exists(arquivo_excel):
        wb = load_workbook(arquivo_excel)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        headers = [
            "Número da nota fiscal", "Data da Consulta/Teste", "Data de Pagamento",
            "Nome do Paciente (dependente)", "CPF do Tomador", "CPF do Dependente",
            "Quem pagou (e relação)", "Valor (R$)", "Forma de pagamento",
            "Data de Registro"
        ]
        ws.append(headers)
    
    return wb, ws

def get_inputs():
    print("📋 Preencha os dados da nota fiscal:")
    nr_nota = input("Número da Nota Fiscal: ")
    d_consulta = date_converter("Data da Consulta/Teste")
    d_pagamento = date_converter("Data de Pagamento")
    paciente = input("Nome do paciente (Dependente): ")
    cpf_tomador = input("CPF do Tomador: ")
    cpf_dependente = input("CPF do Dependente: ")
    q_pagou = input("Quem pagou (Nome e relação com o paciente): ")
    valor = comma_check()
    pagamento = input("Forma de pagamento (PIX, Dinheiro, etc.): ")
    d_registro = datetime.today().strftime("%d/%m/%Y")
    return [
    nr_nota,
    d_consulta,
    d_pagamento,
    paciente,
    cpf_tomador,
    cpf_dependente,
    q_pagou,   # 7. Quem pagou
    valor,     # 8. Valor (R$)
    pagamento,
    d_registro
]

def formatar_planilha(ws):
    # Congela o cabeçalho
    ws.freeze_panes = "A2"

    # Fonte em negrito para o cabeçalho
    bold_font = Font(bold=True)

    # Fundo cinza para o cabeçalho
    fundo_cabecalho = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

    # Alinhamento centralizado
    alinhamento_centralizado = Alignment(horizontal="center", vertical="center")

    # Borda padrão fina
    borda = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )

    # Ajusta largura das colunas
    for col in ws.columns:
        max_length = 0
        coluna = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[coluna].width = max_length + 2

    # Aplica estilos às células preenchidas
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None:
                cell.alignment = alinhamento_centralizado
                cell.border = borda

                if cell.row == 1:
                    cell.font = bold_font
                    cell.fill = fundo_cabecalho

    # Formata a coluna "Valor (R$)" como moeda (4ª coluna)
    col_valor_index = 8
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[col_valor_index - 1]
        if isinstance(cell.value, float):
            cell.number_format = u'R$ #,##0.00'

    # Altura maior para o cabeçalho
    ws.row_dimensions[1].height = 25


def run():
    while True:
        wb, ws = file_create()
        dados = get_inputs()
        ws.append(dados)
        formatar_planilha(ws)
        wb.save(arquivo_excel)
        print("✅ Nota fiscal registrada com sucesso!")

        while True:
            resposta = input("Deseja registrar outra nota? (s/n): ").strip().lower()
            if resposta in ["s", "n"]:
                break
            else:
                print("⚠️ Resposta inválida. Digite apenas 's' para sim ou 'n' para não.")
        if resposta == "n":
            print("Encerrando o programa. Até mais! 👋")
            break


if __name__ == "__main__":
    run()

