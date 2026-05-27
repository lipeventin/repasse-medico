import tkinter as tk
from tkinter import filedialog, scrolledtext
import pandas as pd
import re
import os

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ── Adicione novos nomes de coluna aqui ──────────────────────────
COLUNAS_MEDICO = [
    "Médico Exec.",
    "Médico Exec",
    "Medico Exec.",
    "Medico Exec",
    "Medico",
    "Médico",
]

COLUNAS_VALOR = [
    "Repasse",
    "Valor Fixo",
    "Honorário",
    "Honorarios",
    "Pagamento",
    "Valor",
    "Total",
]
# ─────────────────────────────────────────────────────────────────

def parse_valor(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    if not re.fullmatch(r'[\d.,]+', s):
        return None
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return None

def processar(filepath):
    filepath = filepath.strip().strip('{}')
    if not os.path.isfile(filepath):
        mostrar_texto(f"ERRO: Arquivo não encontrado:\n{filepath}")
        return

    try:
        df_raw = pd.read_excel(filepath)

        # Detecta coluna de médico
        col_medico = next((c for c in COLUNAS_MEDICO if c in df_raw.columns), None)
        if col_medico is None:
            mostrar_texto(
                f"ERRO: Coluna de médico não encontrada.\n"
                f"Colunas no arquivo: {list(df_raw.columns)}\n\n"
                f"Adicione o nome correto em COLUNAS_MEDICO no script."
            )
            return

        colunas_encontradas = [c for c in COLUNAS_VALOR if c in df_raw.columns]
        if not colunas_encontradas:
            mostrar_texto(
                f"ERRO: Nenhuma coluna de valor reconhecida.\n"
                f"Colunas no arquivo: {list(df_raw.columns)}\n\n"
                f"Adicione o nome correto em COLUNAS_VALOR no script."
            )
            return

        linhas_saida = []
        dfs_para_csv = []

        for coluna in colunas_encontradas:
            df = df_raw[[col_medico, coluna]].copy()
            df[coluna] = df[coluna].apply(parse_valor)
            df = df.dropna(subset=[coluna])
            resultado = df.groupby(col_medico)[coluna].sum().reset_index()
            resultado.columns = ["Médico Exec.", coluna]
            resultado = resultado.sort_values(coluna, ascending=False)

            linhas_saida.append(f"── {coluna} {'─' * (40 - len(coluna))}")
            linhas_saida.append(resultado.to_string(index=False))
            linhas_saida.append("")

            dfs_para_csv.append(resultado.set_index("Médico Exec."))

        output_file = filepath.replace(".xlsx", "_repasse.csv")
        df_final = pd.concat(dfs_para_csv, axis=1).reset_index()
        df_final.to_csv(output_file, index=False, encoding="utf-8-sig")

        texto = "\n".join(linhas_saida) + f"\nSalvo: {output_file}"
        mostrar_texto(texto)
        label_drop.config(text=f"✔  {os.path.basename(filepath)}", fg="#7ee787")

    except Exception as e:
        mostrar_texto(f"ERRO: {e}")

def mostrar_texto(texto):
    output_area.config(state=tk.NORMAL)
    output_area.delete("1.0", tk.END)
    output_area.insert(tk.END, texto)

def abrir_arquivo():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        processar(path)

def on_drop(event):
    processar(event.data)

if HAS_DND:
    root = TkinterDnD.Tk()
else:
    root = tk.Tk()

root.title("Repasse por Médico")
root.geometry("600x480")
root.configure(bg="#0d1117")
root.resizable(True, True)

frame_drop = tk.Frame(root, bg="#161b22", highlightbackground="#30363d", highlightthickness=1)
frame_drop.pack(fill="x", padx=24, pady=(24, 10))

label_drop = tk.Label(
    frame_drop,
    text="Arraste o .xlsx aqui  /  clique para selecionar",
    bg="#161b22", fg="#8b949e",
    font=("Courier New", 11),
    pady=20, cursor="hand2"
)
label_drop.pack(fill="x")

label_drop.bind("<Button-1>", lambda e: abrir_arquivo())
frame_drop.bind("<Button-1>", lambda e: abrir_arquivo())

if HAS_DND:
    for w in (frame_drop, label_drop):
        w.drop_target_register(DND_FILES)
        w.dnd_bind('<<Drop>>', on_drop)

output_area = scrolledtext.ScrolledText(
    root,
    font=("Courier New", 11),
    bg="#0d1117", fg="#c9d1d9",
    insertbackground="#c9d1d9",
    relief="flat", bd=0,
    padx=14, pady=14,
    wrap=tk.NONE,
    selectbackground="#264f78"
)
output_area.pack(fill="both", expand=True, padx=24, pady=(0, 24))
output_area.insert(tk.END, "Aguardando arquivo...")

root.mainloop()
