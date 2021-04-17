import tkinter as tk
from tkinter.ttk import Combobox, Progressbar
from tkinter.filedialog import askopenfile 
from tkinter import filedialog
from algorithm import algorith, crear_registro

file2 = 'abd/registro.xlsx'
registro_oficial = 'registro_oficial.xlsx'

ventana = tk.Tk()
logova = tk.PhotoImage(file='img/va.png')

ventana.title("Tomar Asistencia")
ventana.geometry("600x410")
ventana.configure(bg='light steel blue')

tk.Label(ventana,image=logova).place(x = 5, y = 5)

days = ['', '', '', '', '']
oficial = []

combo = Combobox(ventana,
                    values = [ 
                        1,2,3,4
                    ]
                , width = 5, state = 'readonly')

bar = Progressbar(ventana, length = 500, style='grey.Horizontal.TProgressbar')

def open_file(days, num):
    file_path = filedialog.askopenfilename(initialdir = "C:/Users/Hysteria/Documents/MANUELA/ASISTENCIAS - WILLIAM",title = "Select file", filetypes=[('Archivos Excel', '*xlsx')])
    if file_path is not None:
        days[num - 1] = file_path
        pass
    else:
        days[num - 1] = ''
    

def seleccion_oficial(oficial):
    file_path = filedialog.askopenfilename(initialdir = "C:/Users/Hysteria/Documents/MANUELA/REGISTROS OFICIALES",title = "Select file", filetypes=[('Archivos Excel', '*xlsx')])
    if file_path is not None:
        pass
    oficial.append(file_path) 

ofi = tk.Button(ventana, text = 'Seleccionar Registro Oficial', width = 20, height = 1, command = lambda:seleccion_oficial(oficial))
fil1 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days,1))
fil2 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days,2))
fil3 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days,3))
fil4 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days,4))
fil5 = tk.Button(ventana, text = 'Seleccionar Archivo', width = 20, height = 1, command = lambda:open_file(days,5))


ofi.place(x = 430, y = 10)
fil1.place(x = 430, y = 100)
fil2.place(x = 430, y = 150)
fil3.place(x = 430, y = 200)
fil4.place(x = 430, y = 250)
fil5.place(x = 430, y = 300)

combo.place(x = 200, y = 352)




def tomar_lista(days, file2, semana):
    registro = crear_registro(oficial)
    i = 1
    for day in days:
        if day == '':
            i += 1
            continue
        else:
            algorith(day,file2,semana,i, registro,oficial[0])
            i += 1
        
    print('-------------------COMPLETADO-----------------------')
    ventana.destroy()

start = tk.Button(ventana, text = 'Llenar Asistencia', width = 20, height = 1, command = lambda:tomar_lista(days,file2,combo.get()))

start.place(x = 260, y = 350)
ventana.mainloop()




