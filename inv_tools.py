from datetime import datetime

def date_converter(label):
    while True:
        date = input(f"{label} (formato DD/MM/AAAA): ")
        try:
            d_form = datetime.strptime(date, "%d/%m/%Y")
            return d_form.strftime("%d/%m/%Y")
        except ValueError:
            print("⚠️ Data inválida. Tente novamente.")


def comma_check():
    while True:
        inp = input("Digite o valor em reais (R$): ")
        try:
            return float(inp.replace(",", "."))
        except ValueError:
            print("⚠️ Valor inválido. Digite apenas números.")
