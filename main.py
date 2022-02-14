from openpyxl import workbook, load_workbook



class Pessoa:
    def __init__(self, nome, profissao, salariobruto,desconto):
        self.nome = nome
        self.professao = profissao
        self.salariobruto = salariobruto
        self.desconto=desconto

    def calcula_salario(self):
        return self.salariobruto - self.desconto


dados = Pessoa('Clara Maria','Veterinária',1500,420)

salario_liquido = dados.calcula_salario()

#abrindo arquivo excel

template_excel = 'preenche.xlsx'
wb = load_workbook(template_excel)

ws = wb.active

#inserindo dados na planilha
ws['A3'] = dados.nome
ws['B3'] = dados.professao
ws['C3'] = dados.salariobruto
ws['D3'] = dados.desconto

#salvando os dados na planilha
wb.save(template_excel)

#pegando valores

sb = ws['C3'].value
desc = ws['D3'].value

ws['E3'] = sb - desc

#salvando os dados na planilha
wb.save(template_excel)




