import tkinter as tk
from tkinter import ttk
import subprocess

def consiglog_button_clicked():
    subprocess.run(["python", "botconsiglog.py"])

def rf_button_clicked():
    subprocess.run(["python", "botrf1.py"])

def install_drivers():
    # Coloque aqui o código para instalar os drivers
    print("Instalando os drivers...")

def ver_localizacao_button_clicked():
    subprocess.run(["python", "printtela.py"])

def main():
    root = tk.Tk()
    root.title("Tela Inicial")
    root.geometry("400x300")

    # Estilo para os botões
    style = ttk.Style()
    style.configure('TButton', font=('Arial', 14), padding=10)

    rf_button = ttk.Button(root, text="RF1", command=rf_button_clicked, style='TButton')
    rf_button.pack(side=tk.LEFT, padx=20, pady=20)

    consiglog_button = ttk.Button(root, text="CONSIGLOG", command=consiglog_button_clicked, style='TButton')
    consiglog_button.pack(side=tk.LEFT, padx=20, pady=20)

    ver_localizacao_button = ttk.Button(root, text="Ver Localização", command=ver_localizacao_button_clicked, style='TButton')
    ver_localizacao_button.pack(side=tk.LEFT, padx=20, pady=20)

    install_label = ttk.Label(root, text="Não possui os drivers necessários? Clique aqui para instalar.", cursor="hand2", anchor="w")
    install_label.pack(side=tk.LEFT, padx=20, pady=0)

    # Liga a função de instalação de drivers ao evento de clique no label
    install_label.bind("<Button-1>", lambda e: install_drivers())

    root.mainloop()

    
if __name__ == "__main__":
    main()
