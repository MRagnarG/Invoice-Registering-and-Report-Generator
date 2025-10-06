from rep_tools import (
    load_data,
    show_menu,
    MONTHS,
    monthly_general_report,
    yearly_general_report,
    patient_monthly_report,
    patient_yearly_report,
    totals_per_patient_report,
    custom_period_report,
)
from datetime import datetime


def main():
    data = load_data()

    if not data:
        print("❌ No data loaded. Check the Excel file.")
        return

    while True:
        show_menu()
        option = input("\nChoose an option: ")

        if option == "1":
            # Current month report
            today = datetime.today()
            month = today.month
            year = today.year
            print(
                f"\n📅 Generating general report for {MONTHS[month]} {year}..."
            )
            monthly_general_report(data, month, year)

        elif option == "2":
            # Patient report (monthly or yearly)
            patient = input("Patient name: ")
            choice = input("Do you want a [1] Monthly or [2] Yearly report? ")

            if choice == "1":
                month = int(input("Month (1-12): "))
                year = int(input("Year: "))
                patient_monthly_report(data, patient, month, year)
            elif choice == "2":
                year = int(input("Year: "))
                patient_yearly_report(data, patient, year)
            else:
                print("⚠️ Invalid option.")

        elif option == "3":
            # General yearly report
            year = int(input("Year: "))
            yearly_general_report(data, year)

        elif option == "4":
            # Totals per patient in the year
            year = int(input("Year: "))
            totals_per_patient_report(data, year)

        elif option == "5":
            # Custom report by date range
            start_date = input("Start date (YYYY-MM-DD): ")
            end_date = input("End date (YYYY-MM-DD): ")
            filter_patient = input(
                "Do you want to filter by a specific patient? (y/n): "
            ).lower()

            if filter_patient == "y":
                patient = input("Patient name: ")
                custom_period_report(data, start_date, end_date, patient)
            else:
                custom_period_report(data, start_date, end_date)

        elif option == "6":
            print("Closing the program.")
            break

        else:
            print("❌ Invalid option. Try again.")


if __name__ == "__main__":
    main()
