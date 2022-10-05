import openpyxl as op

dados = ("name", "gender", "date", "date_registration","situação", "email", "cpf")

arquivo_excel = op.load_workbook("planilha_clientes.xlsx")
planilha_clientes = arquivo_excel.active
planilha_clientes.title = "gastos"
add = 1


data = ("test","teste","2312","0342","OK","asdfasdf","123")

planilha_clientes["A1"] = dados[0]
planilha_clientes["B1"] = dados[1]
planilha_clientes["C1"] = dados[2]
planilha_clientes["D1"] = dados[3]
planilha_clientes["E1"] = dados[4]
planilha_clientes["F1"] = dados[5]
planilha_clientes["G1"] = dados[6]



def cadastrar(*dados):
    coluna = "ABCDEFG"
    posicao = ultimoElememto()
    for x in range(7):
        planilha_clientes[f"{coluna[x]}{posicao}"].value = data[x]



def ultimoElememto():
    cont = 1
    for celula in planilha_clientes["A"]:
        cont += 1
    return cont


def removerCliente(CPF):
    cont = 2
    colunaCPF = "G"
    colunaSituacao = "E"
    for celula in planilha_clientes["E"]:
        if CPF == planilha_clientes[f"{colunaCPF}{cont}"].value:
            planilha_clientes[f"{colunaSituacao}{cont}"].value = "FALSE"
        cont += 1





while True:
    if (add == 1):
        removerCliente("123")
    else:
        break
    add = int(input("deseja continuar ?"))


arquivo_excel.save("planilha_clientes.xlsx")
