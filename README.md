# Alimentar-template-excel-com-Python-
Esse repositório contém um código que usa a biblioteca openpyxl importanto workbook e load_workbook.

Foi criada uma classe chamada Pessoa no main.py e uma função que calcula o salário liquido.

Os dados foram passados na variável "dados".

Com isso conseguimos partir para a abertura do template excel utilizando a função "load_workbook", iserimos os valores nas colunas e linhas desejadas e salvamos utilizando a função "save".

Também, na mesma estrutura, pegamos um valor que já estava na planilha, para isso fizemos uma variável receber a planilha na linha e coluna utilizando a função "value".

OBs: sempre que for usar a função "save" é necessário que a planilha a ser preeenchida esteja fechada, caso contrário apresentará erro.

