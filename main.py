import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime

# ===============================
#  BACK-END
# ===============================

def quebrar_celulas_mescladas(ws):
    """
    Desmescla todas as células de uma planilha, copiando o valor da
    célula superior esquerda para todas as células que faziam parte
    do intervalo mesclado.
    """
    merged_ranges = list(ws.merged_cells.ranges)
    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        valor = ws.cell(row=min_row, column=min_col).value

        # Desmescla o intervalo
        ws.unmerge_cells(range_string=str(merged_range))

        # Preenche as células do intervalo com o valor original
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                ws.cell(row=row, column=col, value=valor)
    return ws

def encontrar_cabecalho_personalizado(df, colunas_busca, max_linhas=15):
    """
    Tenta encontrar, nas primeiras 'max_linhas' do DataFrame,
    as colunas especificadas em 'colunas_busca', mesmo que
    estejam em linhas diferentes.

    Retorna:
    - Uma lista de nomes de colunas unificados, se encontrado
    - O índice (linha) até onde foi usado para compor o cabeçalho
    """
    num_cols = df.shape[1]
    header_acumulado = [""] * num_cols
    linha_final_cabecalho = 0

    limite = min(max_linhas, len(df))
    for i in range(limite):
        row_values = df.iloc[i].fillna("").astype(str).tolist()
        for c in range(num_cols):
            val = row_values[c].strip()
            if val:
                if header_acumulado[c] == "":
                    header_acumulado[c] = val
                else:
                    header_acumulado[c] += f" {val}"
        colunas_encontradas = [h for h in header_acumulado if any(b in h for b in colunas_busca)]
        if len(colunas_encontradas) >= len(colunas_busca):
            linha_final_cabecalho = i
            break

    return header_acumulado, linha_final_cabecalho

def carregar_planilha_e_filtrar(caminho_arquivo):
    """
    Carrega a planilha, desmescla as células (apenas se não for .xlsb)
    e filtra os dados com base nas colunas necessárias e nos valores
    'SIM' ou 'X' nas colunas 'Optante de transporte' e 'Usará transporte na HE'.
    Agora a comparação é por 'Reg.' (renomeado para 'Registro').
    """
    # Se for .xlsb, carregamos diretamente com pandas (pyxlsb)
    # e PULAMOS o trecho de openpyxl (pois não há suporte para .xlsb).
    if caminho_arquivo.lower().endswith(".xlsb"):
        df = pd.read_excel(
            caminho_arquivo,
            engine='pyxlsb',
            header=None
        )
    else:
        df = pd.read_excel(
            caminho_arquivo,
            engine=None,  # openpyxl por padrão
            header=None
        )
        wb = load_workbook(caminho_arquivo)
        ws = wb.active
        ws = quebrar_celulas_mescladas(ws)
        wb.save(caminho_arquivo)

        # Recarrega após desmesclar
        df = pd.read_excel(
            caminho_arquivo,
            engine=None,
            header=None
        )

    # Colunas que queremos identificar no cabeçalho
    colunas_desejadas = [
        "Reg.", "Nome empregado", "Unidade de Negócio", "Turno",
        "Optante de transporte", "Usará transporte na HE", "LANCHE",
        "HORARIO DE SAÍDA", "OBSERVAÇÃO"
    ]

    # Monta o cabeçalho personalizadamente
    header_unificado, idx_cabecalho = encontrar_cabecalho_personalizado(df, colunas_desejadas, max_linhas=15)
    df.columns = header_unificado
    df = df.iloc[idx_cabecalho + 1:].copy()

    # Mantém só as colunas de interesse
    colunas_existentes = [col for col in df.columns if any(c.lower() in col.lower() for c in colunas_desejadas)]
    if not colunas_existentes:
        raise ValueError("Nenhuma das colunas necessárias foi encontrada na planilha.")

    df = df[colunas_existentes].copy()

    # Renomeia colunas para o padrão que usaremos
    for c in df.columns:
        if "reg." in c.lower():
            df.rename(columns={c: "Registro"}, inplace=True)
        if "optante de transporte" in c.lower():
            df.rename(columns={c: "Optante de transporte"}, inplace=True)
        if "usará transporte na he" in c.lower():
            df.rename(columns={c: "Usará transporte na HE"}, inplace=True)
        if "nome empregado" in c.lower():
            df.rename(columns={c: "Nome empregado"}, inplace=True)

    # Remove linhas onde "Nome empregado" é NaN (opcional)
    if "Nome empregado" in df.columns:
        df = df[df["Nome empregado"].notna()]

    # Normaliza valores de transporte (SIM ou X)
    if "Optante de transporte" in df.columns:
        df["Optante de transporte"] = (
            df["Optante de transporte"].astype(str).str.upper().str.strip()
        )
    if "Usará transporte na HE" in df.columns:
        df["Usará transporte na HE"] = (
            df["Usará transporte na HE"].astype(str).str.upper().str.strip()
        )

    # Filtra somente quem tiver SIM ou X em ambas as colunas
    if "Optante de transporte" in df.columns and "Usará transporte na HE" in df.columns:
        df = df[
            (df["Optante de transporte"].isin(["SIM", "X"])) &
            (df["Usará transporte na HE"].isin(["SIM", "X"]))
        ]

    return df

def comparar_planilhas(caminho_mestre, caminhos_comparacao):
    """
    Compara a planilha mestre com diversas planilhas de comparação
    pela coluna 'Registro' (removendo zeros à esquerda para ambas).
    """
    if not caminho_mestre or not caminhos_comparacao:
        raise ValueError("Selecione a planilha mestre e as planilhas para comparação.")

    # Carrega a planilha mestre
    try:
        planilha_mestre = pd.read_excel(caminho_mestre, header=None)
        planilha_mestre.columns = [
            "Linha", "Turno", "Itinerário", "Registro",
            "Nome dos Passageiros", "Endereço", "Bairro", "Telefone"
        ]
        # Remove zeros à esquerda na planilha mestre
        planilha_mestre["Registro"] = (
            planilha_mestre["Registro"]
            .astype(str)
            .str.strip()
            .str.lstrip("0")  # <-- REMOVE zeros à esquerda
        )
    except Exception as e:
        raise ValueError(f"Erro ao processar a planilha mestre: {e}")

    # Carrega e filtra as planilhas de comparação
    dfs = []
    for caminho in caminhos_comparacao:
        try:
            df_filtrado = carregar_planilha_e_filtrar(caminho)
            if df_filtrado.empty:
                print(f"A planilha {os.path.basename(caminho)} não contém dados válidos após filtragem.")
                continue
            if "Registro" not in df_filtrado.columns:
                print(f"A planilha {os.path.basename(caminho)} não possui coluna 'Reg.' / 'Registro'.")
                continue

            # Remove zeros à esquerda nas planilhas de comparação
            df_filtrado["Registro"] = (
                df_filtrado["Registro"]
                .astype(str)
                .str.strip()
                .str.lstrip("0")  # <-- REMOVE zeros à esquerda
            )

            dfs.append(df_filtrado)
        except ValueError as e:
            print(f"Erro ao filtrar a planilha {os.path.basename(caminho)}: {e}")
            continue
        except Exception as e:
            print(f"Erro inesperado ao carregar a planilha {os.path.basename(caminho)}: {e}")
            continue

    if not dfs:
        raise ValueError("Nenhuma planilha válida foi encontrada para comparação.")

    # Concatena todas as planilhas de comparação
    try:
        todas_planilhas = pd.concat(dfs, ignore_index=True, axis=0)
    except ValueError as e:
        raise ValueError("Erro ao concatenar as planilhas: " + str(e))

    # Remove linhas completamente vazias
    todas_planilhas = todas_planilhas.dropna(how='all')

    # Compara somente por 'Registro' (já sem zeros à esquerda)
    comparacao = planilha_mestre[
        planilha_mestre["Registro"].isin(todas_planilhas["Registro"])
    ]

    # Ordena por Turno e Itinerário
    comparacao["Itinerário"] = comparacao["Itinerário"].astype(str)
    comparacao = comparacao.sort_values(by=["Turno", "Itinerário"])

    # Cria o layout final, separando por turno e itinerário
    separada_por_turnos = []
    for turno, grupo_turno in comparacao.groupby("Turno"):
        # Linha separadora do turno
        separator_turno = pd.DataFrame([[""] * len(comparacao.columns)], columns=comparacao.columns)
        separator_turno.at[0, "Turno"] = f"Turno: {turno}"
        separada_por_turnos.append(separator_turno)

        # Linha vazia (opcional)
        linha_vazia = pd.DataFrame([[""] * len(comparacao.columns)], columns=comparacao.columns)
        separada_por_turnos.append(linha_vazia)

        # Para cada itinerário dentro do turno
        for itinerario, grupo_itinerario in grupo_turno.groupby("Itinerário"):
            # Cabeçalho
            cabecalho_df = pd.DataFrame([planilha_mestre.columns], columns=planilha_mestre.columns)
            separada_por_turnos.append(cabecalho_df)

            # Dados do itinerário
            separada_por_turnos.append(grupo_itinerario)

            # Duas linhas vazias
            linhas_vazias = pd.DataFrame([[""] * len(grupo_itinerario.columns)] * 2, columns=grupo_itinerario.columns)
            separada_por_turnos.append(linhas_vazias)

    planilha_final = pd.concat(separada_por_turnos, ignore_index=True)
    return planilha_final

def gerar_nome_sheet_com_data():
    """
    Gera o nome da aba (sheet) no formato dd.mm,
    por exemplo: '10.03'.
    """
    data_atual = datetime.now()
    return data_atual.strftime("%d.%m")

def salvar_planilha_com_estilo(planilha, caminho_saida):
    """
    Salva o DataFrame 'planilha' em um arquivo Excel,
    aplicando estilos e formatação com openpyxl.
    """
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active

    nome_sheet = gerar_nome_sheet_com_data()
    ws.title = nome_sheet

    estilo_amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    estilo_azul_claro = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    estilo_laranja = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    estilo_negrito = Font(bold=True)
    centralizado = Alignment(horizontal="center", vertical="center")

    border_style = Border(
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000"),
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000")
    )

    # Cabeçalho (primeira linha)
    for col_num, col_name in enumerate(planilha.columns, 1):
        cell = ws.cell(row=1, column=col_num, value=col_name)
        cell.font = estilo_negrito
        cell.alignment = centralizado
        if col_name in ["Linha", "Turno", "Itinerário", "Registro"]:
            cell.fill = estilo_amarelo
        else:
            cell.fill = estilo_azul_claro
        cell.border = border_style

    # Demais linhas
    for row_num, row in enumerate(planilha.itertuples(index=False, name=None), start=2):
        linha_vazia = all(value == "" for value in row)
        linha_turno = any(isinstance(value, str) and value.startswith("Turno:") for value in row)

        for col_num, value in enumerate(row, start=1):
            cell = ws.cell(row=row_num, column=col_num, value=value)

            # Se for linha de turno, aplica o estilo laranja
            if linha_turno:
                cell.fill = estilo_laranja
                cell.font = estilo_negrito
                cell.alignment = centralizado
            else:
                # Se for linha de cabeçalho repetida
                if str(value).upper().strip() in [str(col).upper().strip() for col in planilha.columns]:
                    cell.font = estilo_negrito
                    # Preenchimento de acordo com a coluna
                    if planilha.columns[col_num - 1] in ["Linha", "Turno", "Itinerário", "Registro"]:
                        cell.fill = estilo_amarelo
                    else:
                        cell.fill = estilo_azul_claro
                    cell.alignment = centralizado
                else:
                    cell.alignment = centralizado

            # Aplica borda somente se a linha não for completamente vazia
            if not linha_vazia:
                cell.border = border_style

    wb.save(caminho_saida)

# ===============================
#  FRONT-END (Tkinter)
# ===============================

class ComparadorPlanilhas:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Planilhas Excel")
        self.root.geometry("400x400")
        self.root.config(bg="#282c34")

        self.caminho_mestre = None
        self.caminhos_comparacao = []

        titulo = tk.Label(root, text="Comparador de Planilhas", font=("Helvetica", 20, "bold"), bg="#282c34", fg="#61afef")
        titulo.pack(pady=(10, 5))

        linha = tk.Canvas(root, height=2, bg="white", bd=0, relief="sunken")
        linha.pack(fill="x", padx=10, pady=5)

        self.container = tk.Frame(root, bg="#282c34")
        self.container.pack(pady=10)

        self.botao_mestre = tk.Button(
            self.container,
            text="Planilha Mestre",
            command=self.selecionar_mestre,
            bg="#61afef",
            fg="white",
            font=("Arial", 14),
            borderwidth=0,
            activebackground="#528bb7"
        )
        self.botao_mestre.grid(row=0, column=0, padx=20, pady=5, sticky="ew")

        self.label_mestre = tk.Label(
            self.container,
            text="Nenhuma planilha mestre selecionada.",
            bg="#282c34",
            fg="white",
            font=("Arial", 10)
        )
        self.label_mestre.grid(row=1, column=0, padx=20, pady=5)

        self.botao_comparacao = tk.Button(
            self.container,
            text="Planilhas para Comparação",
            command=self.selecionar_comparacao,
            bg="#61afef",
            fg="white",
            font=("Arial", 14),
            borderwidth=0,
            activebackground="#528bb7"
        )
        self.botao_comparacao.grid(row=2, column=0, padx=20, pady=5, sticky="ew")

        self.label_comparacao = tk.Label(
            self.container,
            text="Nenhuma planilha selecionada.",
            bg="#282c34",
            fg="white",
            font=("Arial", 10)
        )
        self.label_comparacao.grid(row=3, column=0, padx=20, pady=5)

        self.botao_comparar = tk.Button(
            self.container,
            text="Comparar Planilhas",
            command=self.comparar_planilhas,
            bg="#98c379",
            fg="white",
            font=("Arial", 14),
            borderwidth=0,
            activebackground="#8fbc4f"
        )
        self.botao_comparar.grid(row=4, column=0, padx=20, pady=10, sticky="ew")

        self.botao_sair = tk.Button(
            self.container,
            text="Sair",
            command=self.root.quit,
            bg="#e06c75",
            fg="white",
            font=("Arial", 14),
            borderwidth=0,
            activebackground="#d56c6c"
        )
        self.botao_sair.grid(row=5, column=0, padx=20, pady=10, sticky="ew")

    def selecionar_mestre(self):
        """
        Abre um diálogo para o usuário selecionar a planilha mestre
        e armazena o caminho no atributo self.caminho_mestre.
        """
        self.caminho_mestre = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls;*.xlsb")])
        if self.caminho_mestre:
            self.label_mestre.config(text=f"Planilha Mestre: {os.path.basename(self.caminho_mestre)}")
        else:
            self.label_mestre.config(text="Nenhuma planilha mestre selecionada.")

    def selecionar_comparacao(self):
        """
        Abre um diálogo para o usuário selecionar as planilhas
        que serão comparadas com a planilha mestre.
        """
        caminhos_comparacao = filedialog.askopenfilenames(filetypes=[("Arquivos Excel", "*.xlsx;*.xls;*.xlsb")])
        if caminhos_comparacao:
            self.caminhos_comparacao = list(caminhos_comparacao)
            self.label_comparacao.config(text=f"{len(caminhos_comparacao)} planilhas selecionadas.")
        else:
            self.label_comparacao.config(text="Nenhuma planilha selecionada.")

    def comparar_planilhas(self):
        """
        Chama a função de comparação e salva o resultado
        em um arquivo Excel escolhido pelo usuário.
        """
        try:
            if self.caminho_mestre and self.caminhos_comparacao:
                planilha_comparada = comparar_planilhas(self.caminho_mestre, self.caminhos_comparacao)
                caminho_saida = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Arquivos Excel", "*.xlsx")]
                )
                if caminho_saida:
                    salvar_planilha_com_estilo(planilha_comparada, caminho_saida)
                    messagebox.showinfo("Sucesso", f"Planilha comparada salva em: {caminho_saida}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

# ===============================
# MAIN
# ===============================
if __name__ == "__main__":
    root = tk.Tk()
    app = ComparadorPlanilhas(root)
    root.mainloop()
