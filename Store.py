import openpyxl as op

data = ("name", "gender", "date", "date_registration",
        "situation", "email", "cpf")

arquivo_excel = op.load_workbook("planilha_clientes.xlsx")
planilha_clientes = arquivo_excel.active
planilha_clientes.title = "gastos "
add = 1

planilha_clientes["A1"] = data[0]
planilha_clientes["B1"] = data[1]
planilha_clientes["C1"] = data[2]
planilha_clientes["D1"] = data[3]
planilha_clientes["E1"] = data[4]
planilha_clientes["F1"] = data[5]
planilha_clientes["G1"] = data[6]

def cadastrar(*dados):
    coluna = "ABCDEFG"
    posicao = ultimoElememto()
    for x in range(7):
        planilha_clientes[f"{coluna[x]}{posicao}"] = dados[x]

def ultimoElememto():
    cont = 1
    for celula in planilha_clientes["A"]:
        cont += 1
    return cont;


while True:
    if (add == 1):
        cadastrar("Rafel", "Mascu", "07/04/2002", "07/04/2003",
                    "OK", "gbrielti096@gmail.com", "12331232466")
    else:
        break
    add = int(input("deseja continuar ?"))


arquivo_excel.save("planilha_clientes.xlsx")
