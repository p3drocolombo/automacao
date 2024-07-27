import tkinter as tk
from tkinter import filedialog, messagebox
import json
from rf1 import capturar_e_preencher_captcha

# Função para salvar os dados inseridos nos campos
def salvar_dados():
    dados = {
        "login": entry_login.get(),
        "senha": entry_senha.get(),
        "tipo": var_tipo.get(),
        "arquivo": entry_arquivo.get()
    }
    with open("ultimos_dados.json", "w") as file:
        json.dump(dados, file)
    messagebox.showinfo("Salvo", "Últimos dados salvos com sucesso!")

# Função para carregar os últimos dados salvos
def carregar_dados():
    try:
        with open("ultimos_dados.json", "r") as file:
            dados = json.load(file)
            entry_login.delete(0, tk.END)
            entry_login.insert(0, dados["login"])
            entry_senha.delete(0, tk.END)
            entry_senha.insert(0, dados["senha"])
            var_tipo.set(dados["tipo"])
            entry_arquivo.delete(0, tk.END)
            entry_arquivo.insert(0, dados["arquivo"])
            messagebox.showinfo("Carregado", "Últimos dados carregados com sucesso!")
    except FileNotFoundError:
        messagebox.showerror("Erro", "Nenhum dado salvo encontrado!")

# Função para abrir o explorador de arquivos e obter o caminho do arquivo selecionado
def escolher_arquivo():
    arquivo = filedialog.askopenfilename()
    entry_arquivo.delete(0, tk.END)
    entry_arquivo.insert(0, arquivo)

# Função para automatizar o processo de captura e preenchimento do captcha
def automatizar():
    # Capturar dados dos campos
    tipo = var_tipo.get()
    login = entry_login.get()
    senha = entry_senha.get()
    arquivo_excel = entry_arquivo.get()  # Obter o caminho do arquivo Excel

    # Verificar se algum campo está vazio
    if not tipo or not login or not senha or not arquivo_excel:
        messagebox.showwarning("Atenção", "Por favor, preencha todos os campos antes de automatizar o processo.")
        return False  # Retorna False para indicar que o processo não foi iniciado devido a campos vazios

    # Chamar a função para capturar e preencher o captcha
    resultado = capturar_e_preencher_captcha(tipo, login, senha, arquivo_excel)  # Passar o caminho do arquivo Excel

    if resultado:
        messagebox.showinfo("Concluído", "Processo do arquivo Excel concluído com sucesso!")
        pass
    else:
        messagebox.showerror("Erro", "Ocorreu um erro durante o processamento do arquivo Excel.")

    return resultado  # Retornar True se o processo for concluído com sucesso, False caso contrário

# Criar a janela principal
root = tk.Tk()
root.title("Tela")

# Campos e widgets
tk.Label(root, text="Login:").grid(row=0, column=0, sticky="w")
entry_login = tk.Entry(root)
entry_login.grid(row=0, column=1)

tk.Label(root, text="Senha:").grid(row=1, column=0, sticky="w")
entry_senha = tk.Entry(root, show="*")
entry_senha.grid(row=1, column=1)

tk.Label(root, text="Tipo:").grid(row=2, column=0, sticky="w")
var_tipo = tk.StringVar(root)
var_tipo.set("Cajamar")  # Valor padrão
option_menu = tk.OptionMenu(root, var_tipo, "Cajamar", "Boavista", "Juruti", "Portel", "Caxias")
option_menu.grid(row=2, column=1)

tk.Label(root, text="Arquivo:").grid(row=3, column=0, sticky="w")
entry_arquivo = tk.Entry(root)
entry_arquivo.grid(row=3, column=1)
btn_escolher_arquivo = tk.Button(root, text="Escolher arquivo", command=escolher_arquivo)
btn_escolher_arquivo.grid(row=3, column=2)

btn_salvar = tk.Button(root, text="Salvar", command=salvar_dados)
btn_salvar.grid(row=4, column=0, pady=10)

btn_carregar = tk.Button(root, text="Carregar", command=carregar_dados)
btn_carregar.grid(row=4, column=1, pady=10)

btn_automatizar = tk.Button(root, text="Automatizar", command=automatizar)
btn_automatizar.grid(row=5, column=0, columnspan=2, pady=10)

root.mainloop()
