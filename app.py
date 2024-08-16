import tkinter as tk
from tkinter import messagebox

def show_message():
    messagebox.showinfo("Mensagem", "Olá, este é um aplicativo local simples!")

app = tk.Tk()
app.title("Aplicativo Local Simples")

# Configurar a geometria da janela
app.geometry("1080x1920")

# Criar um botão e associar a função show_message
button = tk.Button(app, text="Clique aqui", command=show_message)
button.pack(pady=20)

app.mainloop()