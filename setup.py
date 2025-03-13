from cx_Freeze import setup, Executable

setup(
    name="MeuPrograma",
    version="1.0",
    description="Descrição do meu programa",
    executables=[Executable("main.py")],
)
