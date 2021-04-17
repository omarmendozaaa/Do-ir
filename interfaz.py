import tkinter as tk
from tkinter.ttk import Combobox, Progressbar
from tkinter.filedialog import askopenfile 
from tkinter import filedialog
from algorithm import algorith

file2 = 'abd/registro.xlsx'
registro_oficial = 'registro_oficial.xlsx'

ventana = tk.Tk()
logova = tk.PhotoImage(file='img/va.png')

ventana.title("Tomar Asistencia")
ventana.geometry("600x410")
ventana.configure(bg='light steel blue')

tk.Label(ventana,image=logova).place(x = 5, y = 5)

days = []

combo = Combobox(ventana,
                    values = [ 
                        1,2,3,4
                    ]
                , width = 5, state = 'readonly')

bar = Progressbar(ventana, length = 500, style='grey.Horizontal.TProgressbar')

def open_file(days):
    file_path = filedialog.askopenfilename(initialdir = "C:/Users/Hysteria/Documents/",title = "Select file", filetypes=[('Archivos Excel', '*xlsx')])
    if file_path is not None:
        pass
    days.append(file_path) 

fil1 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days))
fil2 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days))
fil3 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days))
fil4 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days))
fil5 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days))

fil1.place(x = 430, y = 100)
fil2.place(x = 430, y = 150)
fil3.place(x = 430, y = 200)
fil4.place(x = 430, y = 250)
fil5.place(x = 430, y = 300)
combo.place(x = 200, y = 352)
bar.place(x = 50, y = 380)



def tomar_lista(days, file2, semana):
    bar['value'] = 2
    bar.place(x = 50, y = 380)
    print(days)
    algorith(days[0],file2,semana,1)
    bar['value'] += 20
    bar.place(x = 50, y = 380)
    algorith(days[1],file2,semana,2)
    bar['value'] += 20
    bar.place(x = 50, y = 380)
    algorith(days[2],file2,semana,3)
    bar['value'] += 20
    bar.place(x = 50, y = 380)
    algorith(days[3],file2,semana,4)
    bar['value'] += 20
    bar.place(x = 50, y = 380)
    algorith(days[4],file2,semana,5)
    bar['value'] += 20
    bar.place(x = 50, y = 380)

start = tk.Button(ventana, text = 'Llenar Asistencia', width = 20, height = 1, command = lambda:tomar_lista(days,file2,combo.get()))
start.place(x = 260, y = 350)
ventana.mainloop()




