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

SOCIOS = ["THIAGO", "LUCAS", "RAQUEL", "FELIPE", "BRUNO", "STHEFANIE"]

COLUNAS_CSV = [
    "info", "Nota numero", "data", "Pagamento N", "Valor",
    "Retenções fonte", "IRRF", "CSLL fonte", "Valor Liq", "Reter trim",
    "Celia", "De quem", "pg?", "qnd?", "Transferir", "Entrou na CC", "Obs",
    "THIAGO", "LUCAS", "RAQUEL", "FELIPE", "BRUNO", "STHEFANIE"
]
# ─────────────────────────────────────────────────────────────────

resultados_globais = {}  # { numero: (nome_coluna, resultado_df, soma_total) }

def normalizar_medico(nome):
    if pd.isna(nome):
        return None
    s = str(nome).strip().upper()
    primeiro = s.split()[0] if s.split() else s
    return primeiro

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
    global resultados_globais
    resultados_globais = {}

    filepath = filepath.strip().strip('{}')
    if not os.path.isfile(filepath):
        mostrar_texto(f"ERRO: Arquivo não encontrado:\n{filepath}")
        return

    try:
        df_raw = pd.read_excel(filepath)

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

        for i, coluna in enumerate(colunas_encontradas, start=1):
            df = df_raw[[col_medico, coluna]].copy()
            df[col_medico] = df[col_medico].apply(normalizar_medico)
            df = df.dropna(subset=[col_medico])
            df[coluna] = df[coluna].apply(parse_valor)
            df = df.dropna(subset=[coluna])

            resultado = df.groupby(col_medico)[coluna].sum().reset_index()
            resultado.columns = ["Médico", coluna]
            resultado = resultado.sort_values(coluna, ascending=False)
            soma = resultado[coluna].sum()

            resultados_globais[i] = (coluna, resultado, soma)

            linhas_saida.append(f"[{i}] ── {coluna} {'─' * (36 - len(coluna))}")
            linhas_saida.append(resultado.to_string(index=False))
            linhas_saida.append(f"    TOTAL: {soma:_.2f}".replace('.', ',').replace('_', '.'))
            linhas_saida.append("")

        mostrar_texto("\n".join(linhas_saida))
        label_drop.config(text=f"✔  {os.path.basename(filepath)}", fg="#7ee787")
        frame_export.pack(fill="x", padx=24, pady=(0, 16))

    except Exception as e:
        mostrar_texto(f"ERRO: {e}")

def exportar():
    try:
        numero = int(entry_modelo.get().strip())
    except ValueError:
        mostrar_texto(mostrar_texto_atual() + "\n\nERRO: Digite um número válido.")
        return

    if numero not in resultados_globais:
        mostrar_texto(mostrar_texto_atual() + f"\n\nERRO: Modelo [{numero}] não existe.")
        return

    coluna, resultado, soma = resultados_globais[numero]

    # Monta linha única
    linha = {c: "" for c in COLUNAS_CSV}
    linha["Valor"] = f"{soma:_.2f}".replace('.', ',').replace('_', '.')

    for socio in SOCIOS:
        valor_socio = resultado.loc[resultado["Médico"] == socio, coluna]
        if not valor_socio.empty and soma > 0:
            pct = (valor_socio.values[0] / soma) * 100
            linha[socio] = f"{pct:.2f}".replace('.', ',') + "%"
        else:
            linha[socio] = "0,00%"

    df_export = pd.DataFrame([linha], columns=COLUNAS_CSV)

    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv")],
        initialfile=f"export_modelo_{numero}.csv"
    )
    if path:
        df_export.to_csv(path, index=False, encoding="utf-8-sig")
        mostrar_texto(mostrar_texto_atual() + f"\n\nExportado: {path}")

def mostrar_texto_atual():
    return output_area.get("1.0", tk.END).strip()

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

# ── UI ───────────────────────────────────────────────────────────
if HAS_DND:
    root = TkinterDnD.Tk()
else:
    root = tk.Tk()

root.title("Repasse por Médico")
root.geometry("620x560")
root.configure(bg="#0d1117")
root.resizable(True, True)

# Drop zone
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

# Output
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
output_area.pack(fill="both", expand=True, padx=24, pady=(0, 0))
output_area.insert(tk.END, "Aguardando arquivo...")

# Painel de exportação (oculto até processar)
frame_export = tk.Frame(root, bg="#161b22", highlightbackground="#30363d", highlightthickness=1)

tk.Label(
    frame_export,
    text="Escolha um modelo a ser exportado:",
    bg="#161b22", fg="#8b949e",
    font=("Courier New", 10),
    pady=6
).pack(side="left", padx=(12, 6))

entry_modelo = tk.Entry(frame_export, width=4, font=("Courier New", 11),
                        bg="#0d1117", fg="#c9d1d9", insertbackground="#c9d1d9",
                        relief="flat", bd=4)
entry_modelo.pack(side="left", padx=(0, 10))

tk.Button(
    frame_export,
    text="Exportar CSV",
    command=exportar,
    bg="#238636", fg="white",
    font=("Courier New", 10, "bold"),
    relief="flat", padx=10, pady=4,
    cursor="hand2"
).pack(side="left", padx=(0, 12))

root.mainloop()
