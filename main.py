from openpyxl import load_workbook
from tabulate import tabulate

workbook = load_workbook("тестовые_оценки.xlsx")
sheet = workbook.active

table_data = []

for row in sheet.iter_rows(2, values_only=True):
    stats_0 = []
    stats_2 = []
    stats_3 = []
    stats_4 = []
    stats_5 = []
    
    a = row[2:63]
    fio = str(row[1])
    for i in range(60):
        if a[i] == 0:
            stats_0.append(a[i])
        elif a[i] == 2:
            stats_2.append(a[i])
        elif a[i] == 3:
            stats_3.append(a[i])
        elif a[i] == 4:
            stats_4.append(a[i])
        elif a[i] == 5:
            stats_5.append(a[i])

    table_data.append([fio, f"Пропуски: {len(stats_0)}", f"Оценка 2: {len(stats_2)}",
                       f"Оценка 3: {len(stats_3)}", f"Оценка 4: {len(stats_4)}", f"Оценка 5: {len(stats_5)}"])

    stats_0.clear()
    stats_2.clear()
    stats_3.clear()
    stats_4.clear()
    stats_5.clear()

headers = ["ФИО", "Пропуски", "Оценка 2", "Оценка 3", "Оценка 4", "Оценка 5"]
print(tabulate(table_data, headers=headers, tablefmt="pipe"))
