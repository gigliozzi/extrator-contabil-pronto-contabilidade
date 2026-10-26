#!/usr/bin/env python3
"""
===============================================
DocFlow — Interface Gráfica
Automação Contábil Inteligente
Versão: 1.0
===============================================
"""

import os
import sys
import json
import threading
import subprocess
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont

# ── Chave de API (embutida na distribuição por cliente) ────────
API_KEY = "SUA_CHAVE_AQUI"

# ── Cores e estilos ────────────────────────────────────────────
NAVY       = "#0B1628"
NAVY2      = "#112240"
CARD       = "#13213E"
BLUE       = "#1A6CF6"
BLUE_HOV   = "#1558D6"
CYAN       = "#00D4FF"
GREEN      = "#00E5A0"
RED        = "#FF5C5C"
YELLOW     = "#FFD166"
WHITE      = "#F0F4FF"
MUTED      = "#8899BB"
BORDER     = "#1E3560"


# ══════════════════════════════════════════════════════════════
# IMPORTA O EXTRATOR (mesmo diretório)
# ══════════════════════════════════════════════════════════════

def importar_extrator():
    """Importa a função de processamento do extrator_visao_claude.py"""
    script_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    sys.path.insert(0, str(script_dir))
    try:
        import extrator_visao_claude as ext
        ext.API_KEY = API_KEY
        return ext.processar_pdf
    except ImportError as e:
        raise ImportError(f"Arquivo extrator_visao_claude.py não encontrado em {script_dir}\n{e}")


# ══════════════════════════════════════════════════════════════
# JANELA PRINCIPAL
# ══════════════════════════════════════════════════════════════

class DocFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DocFlow")
        self.root.geometry("780x640")
        self.root.minsize(700, 580)
        self.root.configure(bg=NAVY)
        self.root.resizable(True, True)

        # Ícone (se existir)
        icon_path = Path(__file__).parent / "docflow_icon.ico"
        if icon_path.exists():
            try: self.root.iconbitmap(str(icon_path))
            except: pass

        self._ultimo_excel = None
        self._processando = False

        self._build_ui()
        self._pasta_padrao = str(Path.home() / "Downloads")

    # ── BUILD UI ──────────────────────────────────────────────

    def _build_ui(self):
        # Fontes
        self.f_title  = tkfont.Font(family="Segoe UI", size=22, weight="bold")
        self.f_sub    = tkfont.Font(family="Segoe UI", size=10)
        self.f_label  = tkfont.Font(family="Segoe UI", size=9, weight="bold")
        self.f_btn    = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        self.f_log    = tkfont.Font(family="Consolas", size=8)
        self.f_mono   = tkfont.Font(family="Consolas", size=9)

        # ── HEADER ──
        header = tk.Frame(self.root, bg=NAVY2, height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        hinner = tk.Frame(header, bg=NAVY2)
        hinner.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(hinner, text="Doc", font=self.f_title,
                 bg=NAVY2, fg=WHITE).pack(side="left")
        tk.Label(hinner, text="Flow", font=self.f_title,
                 bg=NAVY2, fg=CYAN).pack(side="left")
        tk.Label(hinner, text="  •  Automação Contábil Inteligente",
                 font=self.f_sub, bg=NAVY2, fg=MUTED).pack(side="left", pady=(6,0))

        # ── BODY ──
        body = tk.Frame(self.root, bg=NAVY, padx=32, pady=24)
        body.pack(fill="both", expand=True)

        # Seção de seleção
        sel_frame = tk.Frame(body, bg=CARD, bd=0, relief="flat",
                             highlightthickness=1, highlightbackground=BORDER)
        sel_frame.pack(fill="x", pady=(0, 18))

        sel_inner = tk.Frame(sel_frame, bg=CARD, padx=24, pady=20)
        sel_inner.pack(fill="x")

        tk.Label(sel_inner, text="SELECIONAR ORIGEM",
                 font=self.f_label, bg=CARD, fg=CYAN).pack(anchor="w")

        # Campo de caminho
        path_row = tk.Frame(sel_inner, bg=CARD)
        path_row.pack(fill="x", pady=(10, 0))

        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(
            path_row, textvariable=self.path_var,
            font=self.f_mono, bg=NAVY, fg=WHITE,
            insertbackground=WHITE, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=BLUE, bd=0
        )
        self.path_entry.pack(side="left", fill="x", expand=True,
                             ipady=8, ipadx=8)

        # Botões de seleção
        btn_row = tk.Frame(sel_inner, bg=CARD)
        btn_row.pack(fill="x", pady=(12, 0))

        self._btn(btn_row, "📄  Selecionar PDF",
                  self._selecionar_pdf, BLUE, WHITE).pack(side="left", padx=(0,10))
        self._btn(btn_row, "📁  Selecionar Pasta",
                  self._selecionar_pasta, NAVY2, WHITE,
                  border=BORDER).pack(side="left")

        tk.Label(btn_row,
                 text="💡 Pasta Downloads pré-selecionada por padrão",
                 font=tkfont.Font(family="Segoe UI", size=8),
                 bg=CARD, fg=MUTED).pack(side="right", padx=4)

        # Botão processar
        self.btn_processar = self._btn(
            body, "▶   PROCESSAR DOCUMENTOS",
            self._iniciar_processamento,
            BLUE, WHITE,
            height=2, fullwidth=True
        )
        self.btn_processar.pack(fill="x", pady=(0, 18))

        # Progresso
        prog_frame = tk.Frame(body, bg=CARD,
                              highlightthickness=1, highlightbackground=BORDER)
        prog_frame.pack(fill="x", pady=(0, 18))

        prog_inner = tk.Frame(prog_frame, bg=CARD, padx=24, pady=16)
        prog_inner.pack(fill="x")

        prog_top = tk.Frame(prog_inner, bg=CARD)
        prog_top.pack(fill="x")

        tk.Label(prog_top, text="PROGRESSO",
                 font=self.f_label, bg=CARD, fg=CYAN).pack(side="left")
        self.lbl_paginas = tk.Label(prog_top, text="",
                                    font=self.f_mono, bg=CARD, fg=MUTED)
        self.lbl_paginas.pack(side="right")

        # Barra de progresso customizada
        self.prog_bg = tk.Frame(prog_inner, bg=BORDER, height=6)
        self.prog_bg.pack(fill="x", pady=(10, 6))
        self.prog_fill = tk.Frame(self.prog_bg, bg=BLUE, height=6)
        self.prog_fill.place(relx=0, rely=0, relwidth=0, relheight=1)

        self.lbl_status = tk.Label(prog_inner, text="Aguardando seleção...",
                                   font=self.f_sub, bg=CARD, fg=MUTED,
                                   anchor="w")
        self.lbl_status.pack(fill="x")

        # Log
        log_frame = tk.Frame(body, bg=CARD,
                             highlightthickness=1, highlightbackground=BORDER)
        log_frame.pack(fill="both", expand=True)

        log_inner = tk.Frame(log_frame, bg=CARD, padx=24, pady=16)
        log_inner.pack(fill="both", expand=True)

        log_top = tk.Frame(log_inner, bg=CARD)
        log_top.pack(fill="x", pady=(0,8))

        tk.Label(log_top, text="LOG DE PROCESSAMENTO",
                 font=self.f_label, bg=CARD, fg=CYAN).pack(side="left")

        self.btn_limpar = tk.Button(
            log_top, text="Limpar", font=self.f_sub,
            bg=CARD, fg=MUTED, bd=0, cursor="hand2",
            activebackground=CARD, activeforeground=WHITE,
            command=self._limpar_log
        )
        self.btn_limpar.pack(side="right")

        log_scroll = tk.Scrollbar(log_inner)
        log_scroll.pack(side="right", fill="y")

        self.log_text = tk.Text(
            log_inner, font=self.f_log,
            bg=NAVY, fg=WHITE,
            insertbackground=WHITE, relief="flat",
            bd=0, wrap="word",
            yscrollcommand=log_scroll.set,
            state="disabled", cursor="arrow"
        )
        self.log_text.pack(fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)

        # Tags de cor no log
        self.log_text.tag_config("info",    foreground=MUTED)
        self.log_text.tag_config("ok",      foreground=GREEN)
        self.log_text.tag_config("erro",    foreground=RED)
        self.log_text.tag_config("aviso",   foreground=YELLOW)
        self.log_text.tag_config("destaque",foreground=CYAN)
        self.log_text.tag_config("bold",    foreground=WHITE,
                                 font=tkfont.Font(family="Consolas", size=8, weight="bold"))

        # Botão abrir Excel (inicialmente oculto)
        self.btn_abrir = self._btn(
            body, "📊  Abrir Excel gerado",
            self._abrir_excel, GREEN, NAVY,
            fullwidth=True
        )

        # Footer
        tk.Label(body, text="DocFlow v1.0  •  Dados processados localmente",
                 font=tkfont.Font(family="Segoe UI", size=7),
                 bg=NAVY, fg=BORDER).pack(pady=(12, 0))

    # ── HELPER: criar botão estilizado ────────────────────────

    def _btn(self, parent, text, cmd, bg, fg,
             height=1, fullwidth=False, border=None):
        btn = tk.Button(
            parent, text=text, command=cmd,
            font=self.f_btn, bg=bg, fg=fg,
            activebackground=BLUE_HOV, activeforeground=WHITE,
            relief="flat", bd=0, cursor="hand2",
            padx=20, pady=8, height=height,
        )
        if border:
            btn.config(highlightthickness=1, highlightbackground=border)
        # Hover
        btn.bind("<Enter>", lambda e: btn.config(bg=BLUE_HOV if bg==BLUE else NAVY2))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    # ── SELEÇÃO DE ARQUIVO / PASTA ────────────────────────────

    def _selecionar_pdf(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar PDF",
            initialdir=self._pasta_padrao,
            filetypes=[("Arquivos PDF", "*.pdf *.PDF"), ("Todos", "*.*")]
        )
        if caminho:
            self.path_var.set(caminho)
            self._log(f"PDF selecionado: {Path(caminho).name}", "destaque")

    def _selecionar_pasta(self):
        caminho = filedialog.askdirectory(
            title="Selecionar pasta com PDFs",
            initialdir=self._pasta_padrao,
        )
        if caminho:
            self.path_var.set(caminho)
            pdfs = list(Path(caminho).glob("*.pdf")) + list(Path(caminho).glob("*.PDF"))
            self._log(f"Pasta selecionada: {caminho}", "destaque")
            self._log(f"{len(pdfs)} PDF(s) encontrado(s)", "info")

    # ── LOG ───────────────────────────────────────────────────

    def _log(self, msg, tag="info"):
        self.log_text.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _limpar_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    # ── PROGRESSO ─────────────────────────────────────────────

    def _set_progresso(self, atual, total, status_text, tag="info"):
        pct = atual / total if total > 0 else 0
        self.prog_fill.place(relwidth=pct)
        self.lbl_paginas.config(text=f"{atual}/{total} páginas")
        self.lbl_status.config(text=status_text,
                               fg=GREEN if tag=="ok" else
                                  RED   if tag=="erro" else MUTED)

    # ── PROCESSAMENTO ─────────────────────────────────────────

    def _iniciar_processamento(self):
        caminho = self.path_var.get().strip()
        if not caminho:
            messagebox.showwarning("DocFlow", "Selecione um PDF ou pasta antes de processar.")
            return

        if self._processando:
            return

        self._processando = True
        self.btn_processar.config(state="disabled", text="⏳  Processando...", bg=NAVY2)
        self.btn_abrir.pack_forget()
        self._ultimo_excel = None

        thread = threading.Thread(target=self._processar_thread,
                                  args=(caminho,), daemon=True)
        thread.start()

    def _processar_thread(self, caminho):
        try:
            processar_pdf = importar_extrator()
        except ImportError as e:
            self.root.after(0, self._log, str(e), "erro")
            self.root.after(0, self._finalizar, False)
            return

        entrada = Path(caminho)

        if entrada.is_file():
            pdfs = [entrada]
        elif entrada.is_dir():
            pdfs = sorted(entrada.glob("*.pdf")) + sorted(entrada.glob("*.PDF"))
        else:
            self.root.after(0, self._log, f"Caminho não encontrado: {caminho}", "erro")
            self.root.after(0, self._finalizar, False)
            return

        if not pdfs:
            self.root.after(0, self._log, "Nenhum PDF encontrado.", "aviso")
            self.root.after(0, self._finalizar, False)
            return

        ultimo_excel = None

        for pdf in pdfs:
            self.root.after(0, self._log, f"Iniciando: {pdf.name}", "bold")

            # Conta páginas para a barra
            try:
                import fitz
                doc = fitz.open(str(pdf))
                total_pgs = len(doc)
                doc.close()
            except:
                total_pgs = 1

            self.root.after(0, self._set_progresso, 0, total_pgs,
                            f"Processando {pdf.name}...", "info")

            # Monkey-patch para capturar progresso em tempo real
            import extrator_visao_claude as ext
            ext.API_KEY = API_KEY
            original_print = ext.print if hasattr(ext, '_print_orig') else None

            pg_atual = [0]

            def print_interceptado(*args, **kwargs):
                msg = " ".join(str(a) for a in args)
                # Detecta linha de página
                if "→" in msg and "/" in msg:
                    try:
                        parte = msg.strip().split("/")
                        pg_atual[0] = int(parte[0].split()[-1])
                    except:
                        pass
                    self.root.after(0, self._set_progresso,
                                    pg_atual[0], total_pgs, msg.strip(), "info")
                    self.root.after(0, self._log, msg.strip(), "info")
                elif "✅" in msg:
                    self.root.after(0, self._log, msg.strip(), "ok")
                elif "ERRO" in msg.upper():
                    self.root.after(0, self._log, msg.strip(), "erro")
                else:
                    self.root.after(0, self._log, msg.strip(), "info")

            import builtins
            builtins.print = print_interceptado

            try:
                saida = pdf.parent / f"{pdf.stem}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
                ext.processar_pdf(str(pdf), str(saida))
                ultimo_excel = str(saida)
                self.root.after(0, self._set_progresso, total_pgs, total_pgs,
                                f"Concluído: {saida.name}", "ok")
            except Exception as e:
                self.root.after(0, self._log, f"Erro ao processar {pdf.name}: {e}", "erro")
            finally:
                builtins.print = print

        self._ultimo_excel = ultimo_excel
        self.root.after(0, self._finalizar, ultimo_excel is not None)

    def _finalizar(self, sucesso):
        self._processando = False
        self.btn_processar.config(
            state="normal",
            text="▶   PROCESSAR DOCUMENTOS",
            bg=BLUE
        )
        if sucesso and self._ultimo_excel:
            self.btn_abrir.pack(fill="x", pady=(12, 0))
            self._log("Processamento concluído! Clique em 'Abrir Excel' para ver o resultado.", "ok")
        elif not sucesso:
            self._log("Processamento finalizado com erros.", "erro")

    # ── ABRIR EXCEL ───────────────────────────────────────────

    def _abrir_excel(self):
        if self._ultimo_excel and Path(self._ultimo_excel).exists():
            os.startfile(self._ultimo_excel)
        else:
            messagebox.showwarning("DocFlow", "Arquivo Excel não encontrado.")


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    app = DocFlowApp(root)

    # Centraliza na tela
    root.update_idletasks()
    w, h = 780, 640
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    root.mainloop()


if __name__ == "__main__":
    main()
