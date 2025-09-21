import numpy as np
import pandas as pd

def insert_col_formula(names, formulas, index):
    for i, (name, formula) in enumerate(zip(names, formulas), start=index):
        salaries.insert(gift_tickets + i, name, formula)
        index += 1
    return index

salaries_raw = pd.read_excel("../input/Salaries.xls")
accounting_emp_codes = pd.read_excel("../input/Codes.xls")
benefits = pd.read_excel("../input/Benefits.xlsx")
salaries_old = pd.read_excel("../input/Salaries_old.xls")

salaries = salaries_raw.merge(accounting_emp_codes[["EmployeeCode", "NationalID"]], on="EmployeeCode", how="left", validate="m:1")
cod_index = salaries.columns.get_loc("EmployeeCode")
salaries.insert(cod_index + 1, "NationalID", salaries.pop("NationalID"))
del cod_index, accounting_emp_codes

salaries = salaries.merge(benefits[["EmployeeCode", "EmployeeID"]], on="EmployeeCode", how="left", validate="m:1")
national_id_index = salaries.columns.get_loc("NationalID")
salaries.insert(national_id_index + 1, "Employee ID", salaries.pop("EmployeeID"))
del national_id_index

salaries = salaries.merge(benefits[["EmployeeCode", "CostCenter"]], on="EmployeeCode", how="left", validate="m:1")
prenume_index = salaries.columns.get_loc("FirstName")
salaries.insert(prenume_index + 1, "Cost Center", salaries.pop("CostCenter"))

salaries = salaries.merge(salaries_old[["EmployeeCode", "Department", "Sub Department", "Region"]], on="EmployeeCode", how="left", validate="m:1")

new_cols = ["Department", "Sub Department", "Region"]

for i, x in enumerate(new_cols, start=1):
    salaries.insert(prenume_index + i, x, salaries.pop(x))

del new_cols

rest_plata_index = salaries.columns.get_loc("NetRemainder") 
salaries.insert(rest_plata_index + 1, "NET", salaries["AdvancePay"] + salaries["Deductions"] + salaries["NetRemainder"])

new_cols = ["InsuranceBenefit", "BenefitCard"]

for i, x in enumerate(new_cols, start=2):
    salaries = salaries.merge(benefits[["EmployeeCode", x]], on="EmployeeCode", how="left", validate="m:1")
    salaries.insert(rest_plata_index + i, x, salaries.pop(x))

del new_cols

salaries.insert(rest_plata_index + 4, "Benefits", salaries["InsuranceBenefit"] + salaries["BenefitCard"])
del benefits, rest_plata_index

gift_tickets = salaries.columns.get_loc("GiftTickets")

new_cols = ["Staff Benefits", "Employee Taxes"]
new_values = [salaries["Benefits"] + salaries["MealTickets"] + salaries["GiftTickets"], salaries["Pension"] + salaries["Health"] + salaries["IncomeTax"]]
index = 1
index = insert_col_formula(new_cols, new_values, index)

new_cols = ["Employee Gross", "Employer Tax"]
new_values = [salaries["NET"] + salaries["Staff Benefits"] + salaries["Employee Taxes"], salaries["GrossBase"] * 0.0225]

index = insert_col_formula(new_cols, new_values, index)

new_cols = ["Total Gross", " ", "Total Benefits", "Overtime"]
new_values = [salaries["Employee Gross"] + salaries["Employer Tax"], " ", salaries["Staff Benefits"], salaries["OvertimePay"] + salaries["OtherEarnings"]]

index = insert_col_formula(new_cols, new_values, index)

new_cols = ["Salaries"]
new_values = [salaries["Total Gross"] - salaries["Total Benefits"] - salaries["Overtime"]]

index = insert_col_formula(new_cols, new_values, index)

del new_cols, new_values, index

salaries.to_excel("../output/Salaries_Final.xlsx", index=False)
