from datetime import datetime

def date_converter(label):
    while True:
        date = input(f"{label} (DD/MM/AAAA): ")
        try:
            d_form = datetime.strptime(date, "%d/%m/%Y")
            return d_form.strftime("%d/%m/%Y")
        except ValueError:
            print("⚠️ Invalid date. Try again.")


def comma_check():
    while True:
        inp = input("Write the  value in dollar ($): ")
        try:
            return float(inp.replace(",", "."))
        except ValueError:
            print("⚠️ Invalid value. Write only numbers.")
