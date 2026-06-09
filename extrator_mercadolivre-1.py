"""
Extrator de Dados - Mercado Livre / Mercado Shops
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
import os

# ─────────────────────────────────────────────
# COLUNAS DO MERCADO LIVRE → nome na planilha nova
# ─────────────────────────────────────────────
COLUNAS_DESEJADAS = {
    "Título do anúncio":                "Nome do produto",
    "Receita por produtos (BRL)":       "Receita",
    "Tarifa de venda e impostos (BRL)": "Tarifa 12% ou 17%",
    "Receita por envio (BRL)":          "Receita por envio",
    "Tarifas de envio (BRL)":           "Tarifa de envio",
    "Descontos e bônus":                "Desconto bônus",
    "Total (BRL)":                      "Total sobre",
}

COLUNA_GUBEE = "Taxa Gubee"
VALOR_GUBEE  = 1.0


# ─────────────────────────────────────────────
# HELPERS DE ESTILO
# ─────────────────────────────────────────────

def cor(hex6):
    """Cria um objeto Color com opacidade total (FF) a partir de 6 chars hex."""
    return Color(rgb="FF" + hex6.upper())

def fill(hex6):
    return PatternFill(fill_type="solid", fgColor=cor(hex6))

def borda_cel():
    c = "AAAAAA"
    return Border(
        left=Side(style="thin", color=c), right=Side(style="thin", color=c),
        top=Side(style="thin", color=c),  bottom=Side(style="thin", color=c),
    )


# ─────────────────────────────────────────────
# PROCESSAMENTO
# ─────────────────────────────────────────────

def encontrar_linha_cabecalho(caminho):
    df_raw = pd.read_excel(caminho, header=None, nrows=20)
    for i, row in df_raw.iterrows():
        vals = [str(v) for v in row.values]
        if any("N.º de venda" in v or "Receita por produtos" in v for v in vals):
            return i
    return 4


def processar_planilha(caminho_entrada, caminho_saida, status_label):
    status_label.config(text="⏳ Lendo arquivo...")
    status_label.update()

    linha_cab = encontrar_linha_cabecalho(caminho_entrada)
    df = pd.read_excel(caminho_entrada, sheet_name=0, header=linha_cab, dtype=str)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

    encontradas = {}
    faltando    = []
    for orig, novo in COLUNAS_DESEJADAS.items():
        if orig in df.columns:
            encontradas[orig] = novo
        else:
            faltando.append(orig)

    if not encontradas:
        raise ValueError("Nenhuma coluna conhecida foi encontrada.\nVerifique se é um relatório de Vendas do Mercado Livre.")

    df_saida = df[list(encontradas.keys())].copy()
    df_saida.rename(columns=encontradas, inplace=True)
    df_saida = df_saida.dropna(how="all").reset_index(drop=True)
    df_saida[COLUNA_GUBEE] = VALOR_GUBEE

    status_label.config(text="💾 Salvando planilha...")
    status_label.update()

    wb = Workbook()
    wb.loaded_theme = None   # remove tema padrão que conflita com Excel
    ws = wb.active
    ws.title = "Vendas Extraídas"

    colunas = list(df_saida.columns)

    # ── Cabeçalho ──
    ws.row_dimensions[1].height = 30
    for c_idx, nome in enumerate(colunas, start=1):
        cel = ws.cell(row=1, column=c_idx, value=nome)
        cel.fill      = fill("1B4F72") if nome != COLUNA_GUBEE else fill("1E8449")
        cel.font      = Font(name="Arial", size=11, bold=True, color=cor("FFFFFF"))
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border    = borda_cel()

    # ── Dados ──
    for r_idx, (_, row) in enumerate(df_saida.iterrows(), start=2):
        ws.row_dimensions[r_idx].height = 16
        cor_linha = "D6EAF8" if r_idx % 2 == 0 else "FFFFFF"

        for c_idx, nome_col in enumerate(colunas, start=1):
            v = row[nome_col]

            if nome_col == COLUNA_GUBEE:
                val = VALOR_GUBEE
            else:
                try:
                    val = float(str(v).replace(",", ".")) if str(v) not in ("", "nan", "None") else ""
                except (ValueError, AttributeError):
                    val = str(v) if str(v) not in ("nan", "None") else ""

            cel = ws.cell(row=r_idx, column=c_idx, value=val)
            cel.font   = Font(name="Arial", size=10)
            cel.border = borda_cel()

            if nome_col == COLUNA_GUBEE:
                cel.fill      = fill("D5F5E3")
                cel.alignment = Alignment(horizontal="center", vertical="center")
                cel.number_format = 'R$ #,##0.00'
            else:
                cel.fill      = fill(cor_linha)
                cel.alignment = Alignment(vertical="center")
                if isinstance(val, float):
                    cel.number_format = 'R$ #,##0.00'

    # ── Larguras ──
    larguras = {"Nome do produto": 50, COLUNA_GUBEE: 14}
    for c_idx, nome_col in enumerate(colunas, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = larguras.get(nome_col, 20)

    ws.freeze_panes = "A2"

    wb.save(caminho_saida)

    aviso = ""
    if faltando:
        aviso = "\n\n⚠️ Colunas não encontradas:\n" + "\n".join(f"  • {c}" for c in faltando)

    return len(df_saida), aviso



# ─────────────────────────────────────────────
# INTERFACE GRÁFICA
# ─────────────────────────────────────────────

UI = {
    "bg_escuro":    "#0F2942",
    "bg_claro":     "#F5F7FA",
    "bg_card":      "#FFFFFF",
    "azul":         "#1565C0",
    "azul_hover":   "#0D47A1",
    "verde":        "#2E7D32",
    "verde_hover":  "#1B5E20",
    "borda":        "#DCE3ED",
    "texto":        "#1A2A3A",
    "subtexto":     "#607D8B",
    "tag_bg":       "#E3F2FD",
    "tag_fg":       "#1565C0",
    "tag_gubee":    "#E8F5E9",
    "tag_gubee_fg": "#2E7D32",
}


def _hover(btn, normal, hover):
    btn.bind("<Enter>", lambda e: btn.config(bg=hover))
    btn.bind("<Leave>", lambda e: btn.config(bg=normal))


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator ML – Mercado Livre")
        self.root.geometry("640x570")
        self.root.resizable(False, False)
        self.root.configure(bg=UI["bg_claro"])
        self.arquivo_entrada = tk.StringVar()
        self.arquivo_saida   = tk.StringVar()
        self._construir_interface()

    def _card(self, parent, pady=(6, 0)):
        outer = tk.Frame(parent, bg=UI["borda"])
        outer.pack(fill="x", padx=28, pady=pady)
        inner = tk.Frame(outer, bg=UI["bg_card"], padx=16, pady=12)
        inner.pack(fill="x", padx=1, pady=1)
        return inner

    def _num_label(self, parent, numero, titulo):
        row = tk.Frame(parent, bg=UI["bg_card"])
        row.pack(fill="x", pady=(0, 8))
        tk.Label(row, text=f" {numero} ", font=("Arial", 9, "bold"),
            bg=UI["azul"], fg="white"
        ).pack(side="left", padx=(0, 8))
        tk.Label(row, text=titulo, font=("Arial", 10, "bold"),
            bg=UI["bg_card"], fg=UI["texto"]
        ).pack(side="left")

    def _entry_row(self, parent, var, btn_texto, cmd):
        row = tk.Frame(parent, bg=UI["bg_card"])
        row.pack(fill="x")
        e = tk.Entry(row, textvariable=var, font=("Arial", 9),
            state="readonly", bg="#F0F4F8", fg=UI["texto"],
            relief="flat", highlightthickness=1,
            highlightbackground=UI["borda"], readonlybackground="#F0F4F8"
        )
        e.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        b = tk.Button(row, text=btn_texto, command=cmd,
            bg=UI["azul"], fg="white", font=("Arial", 9, "bold"),
            relief="flat", cursor="hand2", padx=12, pady=5
        )
        b.pack(side="left")
        _hover(b, UI["azul"], UI["azul_hover"])

    def _construir_interface(self):
        # ── Header ──
        hdr = tk.Frame(self.root, bg=UI["bg_escuro"], height=82)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📊  Extrator de Dados",
            font=("Arial", 17, "bold"), bg=UI["bg_escuro"], fg="white"
        ).place(relx=0.5, rely=0.36, anchor="center")
        tk.Label(hdr, text="Mercado Livre  /  Mercado Shops",
            font=("Arial", 9), bg=UI["bg_escuro"], fg="#90CAF9"
        ).place(relx=0.5, rely=0.70, anchor="center")

        tk.Frame(self.root, bg=UI["bg_claro"], height=14).pack(fill="x")

        # ── Card 1: entrada ──
        c1 = self._card(self.root)
        self._num_label(c1, "1", "Planilha de origem  (Mercado Livre)")
        self._entry_row(c1, self.arquivo_entrada, "  Procurar  ", self._selecionar_entrada)

        # ── Card 2: saída ──
        c2 = self._card(self.root)
        self._num_label(c2, "2", "Onde salvar a planilha nova")
        self._entry_row(c2, self.arquivo_saida, "  Salvar em  ", self._selecionar_saida)

        # ── Card 3: tags das colunas ──
        c3 = self._card(self.root, pady=(6, 6))
        tk.Label(c3, text="Colunas que serão geradas na planilha:",
            font=("Arial", 9, "bold"), bg=UI["bg_card"], fg=UI["subtexto"]
        ).pack(anchor="w", pady=(0, 7))

        tags = tk.Frame(c3, bg=UI["bg_card"])
        tags.pack(fill="x")

        for nome in COLUNAS_DESEJADAS.values():
            tk.Label(tags, text=nome, font=("Arial", 8, "bold"),
                bg=UI["tag_bg"], fg=UI["tag_fg"],
                padx=8, pady=4
            ).pack(side="left", padx=3, pady=2)

        tk.Label(tags, text=f"  {COLUNA_GUBEE}  (R$ 1,00 fixo)  ",
            font=("Arial", 8, "bold"),
            bg=UI["tag_gubee"], fg=UI["tag_gubee_fg"],
            padx=8, pady=4
        ).pack(side="left", padx=3, pady=2)

        # ── Botão extrair ──
        btn_frame = tk.Frame(self.root, bg=UI["bg_claro"])
        btn_frame.pack(pady=18)

        self.btn_extrair = tk.Button(btn_frame, text="▶   Extrair Dados",
            command=self._processar,
            bg=UI["verde"], fg="white",
            font=("Arial", 13, "bold"),
            relief="flat", padx=32, pady=11, cursor="hand2"
        )
        self.btn_extrair.pack()
        _hover(self.btn_extrair, UI["verde"], UI["verde_hover"])

        # ── Barra de status ──
        status_bar = tk.Frame(self.root, bg="#E8EDF2", height=32)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)

        self.label_status = tk.Label(status_bar,
            text="  ●  Aguardando arquivo...",
            font=("Arial", 9), bg="#E8EDF2", fg=UI["subtexto"], anchor="w"
        )
        self.label_status.pack(side="left", fill="y", padx=8)

    def _selecionar_entrada(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha do Mercado Livre",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.arquivo_entrada.set(caminho)
            pasta     = os.path.dirname(caminho)
            nome_base = os.path.splitext(os.path.basename(caminho))[0]
            self.arquivo_saida.set(os.path.join(pasta, f"{nome_base}_EXTRAIDO.xlsx"))
            self.label_status.config(text="  ●  Arquivo selecionado.", fg="#2E7D32")

    def _selecionar_saida(self):
        caminho = filedialog.asksaveasfilename(
            title="Salvar planilha nova como...",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")]
        )
        if caminho:
            self.arquivo_saida.set(caminho)

    def _processar(self):
        entrada = self.arquivo_entrada.get()
        saida   = self.arquivo_saida.get()

        if not entrada:
            messagebox.showwarning("Atenção", "Selecione a planilha de origem primeiro.")
            return
        if not saida:
            messagebox.showwarning("Atenção", "Escolha onde salvar a planilha nova.")
            return

        self.btn_extrair.config(state="disabled", text="⏳  Processando...")
        self.root.update()

        try:
            qtd, aviso = processar_planilha(entrada, saida, self.label_status)
            self.label_status.config(
                text=f"  ●  Concluído! {qtd} linhas extraídas.", fg="#2E7D32"
            )
            msg = f"Planilha criada com sucesso!\n\n{qtd} vendas extraídas.\n\nSalvo em:\n{saida}"
            if aviso:
                msg += aviso
            messagebox.showinfo("Concluído!", msg)

            if messagebox.askyesno("Abrir pasta?", "Deseja abrir a pasta onde foi salvo?"):
                pasta = os.path.dirname(saida)
                if os.name == "nt":
                    os.startfile(pasta)
                else:
                    os.system(f'open "{pasta}"')

        except Exception as e:
            self.label_status.config(text="  ●  Erro ao processar.", fg="#C0392B")
            messagebox.showerror("Erro", f"Ocorreu um erro:\n\n{str(e)}")

        finally:
            self.btn_extrair.config(state="normal", text="▶   Extrair Dados")


# ─────────────────────────────────────────────
# INÍCIO
# ─────────────────────────────────────────────
if __name__ == "__main__":
    janela = tk.Tk()
    App(janela)
    janela.mainloop()
