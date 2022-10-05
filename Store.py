import openpyxl as op
data = ("name", "gender", "date", "date_registration",
        "situation", "email", "cpf")

arquivo_excel = op.Workbook()
planilha_clientes = arquivo_excel.active
planilha_clientes.title = "gastos "
posicaoLivre = int(planilha_clientes._get_cell(1,9).value)
add = 1

planilha_clientes["A1"] = data[0]
planilha_clientes["B1"] = data[1]
planilha_clientes["C1"] = data[2]
planilha_clientes["D1"] = data[3]
planilha_clientes["E1"] = data[4]
planilha_clientes["F1"] = data[5]
planilha_clientes["G1"] = data[6]


def adcionar(*dados):
    coluna = "ABCDEFG"
    for x in range(7):
        planilha_clientes[f"{coluna[x]}{posicaoLivre}"] = dados[x]

while True:
    print("posicaoLivre")
    if (add == 1):
        adcionar("Gabriel", "Mascu", "07/04/2002", "07/04/2003",
                    "OK", "gbrielti096@gmail.com", "12331232466")
    else:
        break
    add = int(input("deseja continuar ?"))
    posicaoLivre += 1


arquivo_excel.save("planilha_clientes.xlsx")
