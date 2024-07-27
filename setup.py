from cx_Freeze import setup, Executable

setup(
    name="MinhaAutomacao",
    version="1.0",
    description="Uma descrição da minha aplicação",
    executables=[Executable("tela_inicial.py")],
)
