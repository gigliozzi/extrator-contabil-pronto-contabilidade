#!/usr/bin/env python3
"""
===============================================
Script : interface.py
Descrição : Interface gráfica para o sistema
            de extração de PDFs contábeis.
            Chama extrator_visao_claude.py e
            exibe o progresso em tempo real.
Versão : v1.0 - 2026-10-12
Autor: William Gigliozzi
Suporte: gigliozzi@gmail.com
===============================================
"""

import os
import sys
import threading
import subprocess
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ── Pastas do projeto ─────────────────────────────────────────
BASE_DIR = Path(__file__).parent
PASTA_SAIDA = BASE_DIR / "saida"
PASTA_ENTRADA = BASE_DIR.parent / "entrada"
SCRIPT_EXTRATOR = BASE_DIR / "extrator_visao_claude.py"

PASTA_ENTRADA.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════════════
# JANELA PRINCIPAL
# ══════════════════════════════════════════════════════════════


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Pronto Contabilidade :::::: Extrator de Documentos Contábeis")
        self.geometry("780x600")
        self.resizable(False, False)
        self.configure(bg="#F0F2F5")

        # Estado
        self.pdf_selecionado = tk.StringVar(value="")
        self.processando = False
        self.excel_gerado = None

        self._construir_ui()
        self._centralizar()

    # ── Layout ────────────────────────────────────────────────

    def _construir_ui(self):
        # ── Cabeçalho ─────────────────────────────────────────
        cabecalho = tk.Frame(self, bg="#1F4E79", height=70)
        cabecalho.pack(fill="x")
        cabecalho.pack_propagate(False)

        tk.Label(
            cabecalho,
            text="📄  Pronto Contabilidade",
            font=("Segoe UI", 16, "bold"),
            bg="#1F4E79",
            fg="#FFFFFF",
        ).pack(side="left", padx=24, pady=18)

        tk.Label(
            cabecalho,
            text="v1.0",
            font=("Segoe UI", 9),
            bg="#1F4E79",
            fg="#90B8D8",
        ).pack(side="right", padx=20, pady=22)

        # ── Área de seleção de PDF ────────────────────────────
        frame_sel = tk.LabelFrame(
            self,
            text="  Arquivo PDF  ",
            font=("Segoe UI", 10),
            bg="#F0F2F5",
            fg="#1F4E79",
            bd=1,
            relief="groove",
            padx=16,
            pady=12,
        )
        frame_sel.pack(fill="x", padx=20, pady=(18, 0))

        frame_linha = tk.Frame(frame_sel, bg="#F0F2F5")
        frame_linha.pack(fill="x")

        self.entry_pdf = tk.Entry(
            frame_linha,
            textvariable=self.pdf_selecionado,
            font=("Segoe UI", 10),
            state="readonly",
            relief="solid",
            bd=1,
            bg="#FFFFFF",
            fg="#333333",
            readonlybackground="#FFFFFF",
        )
        self.entry_pdf.pack(side="left", fill="x", expand=True, ipady=6)

        tk.Button(
            frame_linha,
            text="  Carregar PDF ",
            font=("Segoe UI", 10),
            bg="#1F4E79",
            fg="#FFFFFF",
            relief="flat",
            activebackground="#2E6FAA",
            activeforeground="#FFFFFF",
            cursor="hand2",
            command=self._selecionar_pdf,
        ).pack(side="left", padx=(10, 0), ipady=6, ipadx=4)

        # ── Pasta de saída ────────────────────────────────────
        frame_saida = tk.Frame(self, bg="#F0F2F5")
        frame_saida.pack(fill="x", padx=20, pady=(8, 0))

        tk.Label(
            frame_saida,
            text=f"📁  Saída: {BASE_DIR.parent / 'excel_gerado'}",
            font=("Segoe UI", 9),
            bg="#F0F2F5",
            fg="#666666",
        ).pack(side="left")

        # ── Botões de ação ────────────────────────────────────
        frame_botoes = tk.Frame(self, bg="#F0F2F5")
        frame_botoes.pack(pady=14)

        self.btn_processar = tk.Button(
            frame_botoes,
            text="  ▶  Processar PDF  ",
            font=("Segoe UI", 11, "bold"),
            bg="#217346",
            fg="#FFFFFF",
            relief="flat",
            activebackground="#2A9B5C",
            activeforeground="#FFFFFF",
            cursor="hand2",
            command=self._iniciar_processamento,
            state="disabled",
        )
        self.btn_processar.pack(side="left", ipadx=8, ipady=6)

        self.btn_abrir = tk.Button(
            frame_botoes,
            text="  📊  Abrir Excel  ",
            font=("Segoe UI", 11),
            bg="#C0C0C0",
            fg="#FFFFFF",
            relief="flat",
            activebackground="#A0A0A0",
            cursor="hand2",
            command=self._abrir_excel,
            state="disabled",
        )
        self.btn_abrir.pack(side="left", padx=(12, 0), ipadx=8, ipady=6)

        self.btn_limpar = tk.Button(
            frame_botoes,
            text="  🗑  Limpar  ",
            font=("Segoe UI", 11),
            bg="#888888",
            fg="#FFFFFF",
            relief="flat",
            activebackground="#666666",
            cursor="hand2",
            command=self._limpar,
        )
        self.btn_limpar.pack(side="left", padx=(12, 0), ipadx=8, ipady=6)

        # ── Barra de progresso ────────────────────────────────
        frame_prog = tk.Frame(self, bg="#F0F2F5")
        frame_prog.pack(fill="x", padx=20, pady=(0, 6))

        self.label_status = tk.Label(
            frame_prog,
            text="Aguardando seleção de arquivo…",
            font=("Segoe UI", 9),
            bg="#F0F2F5",
            fg="#888888",
        )
        self.label_status.pack(anchor="w")

        self.barra = ttk.Progressbar(
            frame_prog,
            mode="indeterminate",
            length=740,
        )
        self.barra.pack(fill="x", pady=(4, 0))

        # ── Log de processamento ──────────────────────────────
        frame_log = tk.LabelFrame(
            self,
            text="  Log de processamento  ",
            font=("Segoe UI", 10),
            bg="#F0F2F5",
            fg="#1F4E79",
            bd=1,
            relief="groove",
            padx=0,
            pady=0,
        )
        frame_log.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        self.texto_log = tk.Text(
            frame_log,
            font=("Consolas", 9),
            bg="#1E1E1E",
            fg="#D4D4D4",
            relief="flat",
            state="disabled",
            wrap="word",
            padx=10,
            pady=8,
        )
        self.texto_log.pack(side="left", fill="both", expand=True)

        scroll = tk.Scrollbar(frame_log, command=self.texto_log.yview)
        scroll.pack(side="right", fill="y")
        self.texto_log.configure(yscrollcommand=scroll.set)

        # Tags de cor no log
        self.texto_log.tag_configure("ok", foreground="#4EC9B0")
        self.texto_log.tag_configure("erro", foreground="#F44747")
        self.texto_log.tag_configure("info", foreground="#569CD6")
        self.texto_log.tag_configure("aviso", foreground="#CE9178")
        self.texto_log.tag_configure("dim", foreground="#858585")

    # ── Ações ─────────────────────────────────────────────────

    def _selecionar_pdf(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar PDF",
            filetypes=[("Arquivos PDF", "*.pdf *.PDF"), ("Todos", "*.*")],
            initialdir=str(PASTA_ENTRADA),
        )
        if caminho:
            self.pdf_selecionado.set(caminho)
            self.btn_processar.configure(state="normal")
            self._log(f"Arquivo selecionado: {Path(caminho).name}", "info")
            self.label_status.configure(
                text="Arquivo selecionado. Clique em Processar PDF."
            )

    def _iniciar_processamento(self):
        if self.processando:
            return

        pdf = self.pdf_selecionado.get()
        if not pdf:
            messagebox.showwarning("Atenção", "Selecione um arquivo PDF primeiro.")
            return

        if not SCRIPT_EXTRATOR.exists():
            messagebox.showerror(
                "Erro",
                f"Script extrator não encontrado:\n{SCRIPT_EXTRATOR}\n\n"
                "Verifique se extrator_visao_claude.py está na mesma pasta.",
            )
            return

        self.processando = True
        self.excel_gerado = None
        self.btn_processar.configure(state="disabled", text="  ⏳  Processando…  ")
        self.btn_abrir.configure(state="disabled", bg="#C0C0C0")
        self.barra.start(12)

        thread = threading.Thread(target=self._processar, args=(pdf,), daemon=True)
        thread.start()

    def _processar(self, caminho_pdf):
        try:
            self._log("=" * 55, "dim")
            self._log(f"Iniciando processamento", "info")
            self._log(f"PDF   : {Path(caminho_pdf).name}", "dim")
            self._log("=" * 55, "dim")

            self._atualizar_status("Processando — aguarde…")

            processo = subprocess.Popen(
                [sys.executable, str(SCRIPT_EXTRATOR), caminho_pdf],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
            )

            saida = None
            for linha in processo.stdout:
                linha = linha.rstrip()
                if not linha:
                    continue

                # Capturar caminho real do Excel gerado pelo extrator
                if "✅ Excel salvo em:" in linha:
                    caminho_str = linha.split("✅ Excel salvo em:")[-1].strip()
                    saida = Path(caminho_str)

                # Colorir linhas por conteúdo
                if "ERRO" in linha or "❌" in linha:
                    tag = "erro"
                elif "✅" in linha or "Excel salvo" in linha:
                    tag = "ok"
                elif "===" in linha or "---" in linha:
                    tag = "dim"
                elif "→" in linha:
                    tag = "aviso"
                else:
                    tag = None

                self._log(linha, tag)

            processo.wait()

            if processo.returncode == 0 and saida and saida.exists():
                self.excel_gerado = saida
                self._log("=" * 55, "dim")
                self._log(f"✅  Concluído com sucesso!", "ok")
                self._log(f"📊  Excel: {saida.name}", "ok")
                self._log("=" * 55, "dim")
                self._atualizar_status(f"✅  Concluído — {saida.name}")
                self.after(0, self._habilitar_abrir)
            else:
                self._log("❌  Processamento encerrou com erro.", "erro")
                self._atualizar_status("❌  Erro durante o processamento.")

        except Exception as e:
            self._log(f"❌  Exceção: {e}", "erro")
            self._atualizar_status("❌  Erro inesperado.")

        finally:
            self.processando = False
            self.after(0, self._finalizar_ui)

    def _habilitar_abrir(self):
        self.btn_abrir.configure(state="normal", bg="#1F4E79")

    def _finalizar_ui(self):
        self.barra.stop()
        self.btn_processar.configure(state="normal", text="  ▶  Processar PDF  ")

    def _abrir_excel(self):
        if self.excel_gerado and self.excel_gerado.exists():
            os.startfile(str(self.excel_gerado))
        else:
            messagebox.showinfo("Info", "Nenhum Excel gerado ainda.")

    def _limpar(self):
        self.pdf_selecionado.set("")
        self.excel_gerado = None
        self.btn_processar.configure(state="disabled")
        self.btn_abrir.configure(state="disabled", bg="#C0C0C0")
        self.label_status.configure(text="Aguardando seleção de arquivo…")
        self.texto_log.configure(state="normal")
        self.texto_log.delete("1.0", "end")
        self.texto_log.configure(state="disabled")

    # ── Helpers ───────────────────────────────────────────────

    def _log(self, mensagem, tag=None):
        def _escrever():
            self.texto_log.configure(state="normal")
            if tag:
                self.texto_log.insert("end", mensagem + "\n", tag)
            else:
                self.texto_log.insert("end", mensagem + "\n")
            self.texto_log.see("end")
            self.texto_log.configure(state="disabled")

        self.after(0, _escrever)

    def _atualizar_status(self, texto):
        self.after(0, lambda: self.label_status.configure(text=texto))

    def _centralizar(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")


# ══════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
