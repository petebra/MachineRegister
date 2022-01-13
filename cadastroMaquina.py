import tkinter as tk
import datetime as dt
from tkinter import ttk
import pandas as pd

maquinas = pd.read_excel('Cardomaq.xlsx', engine='openpyxl')

categories = [
    'Afiadoras',
    'Curvadoras de Tudo',
    'Furadeiras',
    'Fresadora',
    'Laminadora',
    'Guilhotina',
    'Plaina',
    'Prensa',
    'Retifica'
]

codes = []


# Métodos
def addCategoria():
    newCategory = comboBox_category.get()
    categories.append(newCategory)
    print(categories)


def addMachine():
    description = entry_description.get()
    category = comboBox_category.get()
    paidValue = entry_paid_value.get()
    futurePrice = entry_price.get()
    createDate = dt.datetime.now()
    createDate = createDate.strftime('%d/%m/%Y %H:%M')
    code = maquinas.shape[0] + len(codes) + 1
    code_str = f'COD{code}'
    codes.append((code_str, description, category, paidValue, futurePrice, createDate))

    label_success.config(text='Máquina cadastrada com sucesso')


def clearInfo():
    entry_description.config(text='')


janela = tk.Tk()

# Titulo da janela
janela.title('Cadastro de Novas Máquinas')

# Description Label
label_description = tk.Label(text='Descrição da Máquina')
label_description.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=4)

# Description Data Entry
entry_description = tk.Entry()
entry_description.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=4)

# Machine Category
label_category = tk.Label(text='Categoria:')
label_category.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=1)

# Machine Data Entry (ComboBox)
comboBox_category = ttk.Combobox(values=categories)
comboBox_category.grid(row=4, column=1, padx=10, pady=10, sticky='nswe', columnspan=2)

# Botão para adicionar Categoria
btn_add_category = tk.Button(text='Adicionar Categoria', command=addCategoria)
btn_add_category.grid(row=4, column=3, padx=10, pady=10, sticky='nswe', columnspan=1)

label_paid_value = tk.Label(text='Valor pago:')
label_paid_value.grid(row=5, column=0, padx=10, pady=10, sticky='nswe', columnspan=1)

entry_paid_value = tk.Entry()
entry_paid_value.grid(row=5, column=1, padx=10, pady=10, sticky='nswe', columnspan=1)

label_price = tk.Label(text='Novo preço:')
label_price.grid(row=5, column=2, padx=10, pady=10, sticky='nswe', columnspan=1)

entry_price = tk.Entry()
entry_price.grid(row=5, column=3, padx=10, pady=10, sticky='nswe', columnspan=1)

# Botão de registrar
btn_machine_registry = tk.Button(text='Registrar Máquina', command=addMachine)
btn_machine_registry.grid(row=6, column=0, padx=10, pady=10, sticky='nswe', columnspan=4)

label_success = tk.Label(text='')
label_success.grid(row=7, column=0, padx=10, pady=10, sticky='nswe', columnspan=4)

janela.mainloop()

novoDF = pd.DataFrame(
    codes, columns=['Código', 'Categoria', 'Valor Pago', 'Preço Futuro', 'Data de Criação', 'Descrição']
)
maquinas = maquinas.append(novoDF, ignore_index=True)

maquinas.to_excel('Cardomaq.xlsx', index=False)
