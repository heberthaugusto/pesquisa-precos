# ============================================================
#  PESQUISA DE PREÇOS
#  Lê os hyperlinks diretamente da planilha: qualquer produto
#  adicionado com hyperlink nas colunas das lojas é buscado
#  automaticamente. As lojas são definidas em lojas.json.
# ============================================================
import json
import os
import re
import threading
import tkinter as tk
from datetime import date, datetime
from tkinter import filedialog, messagebox, ttk

PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_LOJAS = os.path.join(PASTA_SCRIPT, "lojas.json")


# ── Configuração das lojas ───────────────────────────────────────────────────
def carregar_lojas():
    """Lê lojas.json. Cada loja tem:
    id                  identificador interno
    cabecalhos          fragmentos de texto que identificam a coluna na linha 1 da planilha
    regra_preco         "primeiro" (primeiro R$ da página) ou "maior_dos_dois_primeiros"
                        (para páginas que mostram preço "de/por" invertido)
    indisponivel_sem    trecho de HTML que, quando ausente, indica produto indisponível (opcional)
    requer_login        true para lojas que exigem sessão autenticada (opcional)
    url_login           página de login da loja, aberta no navegador do usuário (opcional)
    """
    if not os.path.isfile(ARQUIVO_LOJAS):
        raise RuntimeError(
            f"Arquivo de configuração não encontrado: {ARQUIVO_LOJAS}\n"
            "Copie lojas.exemplo.json para lojas.json e ajuste as lojas.")
    with open(ARQUIVO_LOJAS, encoding="utf-8") as f:
        lojas = json.load(f)
    for l in lojas:
        l.setdefault("regra_preco", "primeiro")
        l.setdefault("requer_login", False)
        l["cabecalhos"] = [c.strip().upper() for c in l["cabecalhos"]]
    return {l["id"]: l for l in lojas}


def detectar_aba_mais_recente(wb):
    """Retorna o nome da aba com data mais recente no formato DDMMAAAA.
    Se nenhuma aba tiver esse formato, retorna a última aba da planilha."""
    melhor_data = melhor_nome = None
    for nome in wb.sheetnames:
        for fmt in ("%d%m%Y", "%d%m%y"):
            try:
                dt = datetime.strptime(nome.split("_")[0], fmt)
                if melhor_data is None or dt > melhor_data:
                    melhor_data, melhor_nome = dt, nome
                break
            except ValueError:
                continue
    return melhor_nome or wb.sheetnames[-1]


# ── Leitura da planilha ──────────────────────────────────────────────────────
def ler_celulas(xlsx_path, lojas):
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    aba = detectar_aba_mais_recente(wb)
    ws = wb[aba]

    # Mapeia colunas da linha 1 (cabeçalho) para o id da loja
    col_loja = {}
    for cell in ws[1]:
        if not cell.value:
            continue
        nome = str(cell.value).strip().upper()
        for loja_id, loja in lojas.items():
            if any(chave in nome for chave in loja["cabecalhos"]):
                col_loja[cell.column] = loja_id
                break

    if not col_loja:
        esperados = ", ".join(l["cabecalhos"][0] for l in lojas.values())
        raise RuntimeError(
            "Nenhuma coluna de loja reconhecida na linha 1.\n"
            f"Cabeçalhos esperados: {esperados}")

    sku_col, desc_col = 1, 2  # coluna A = SKU, coluna B = descritivo

    cells = []
    for row in ws.iter_rows(min_row=2):
        row_idx = row[0].row
        for cell in row:
            if cell.column not in col_loja or not cell.hyperlink:
                continue
            url = cell.hyperlink.target or ""
            if not url.startswith("http"):
                continue

            value = 0.0
            try:
                if cell.value is not None:
                    value = float(cell.value)
            except (ValueError, TypeError):
                pass

            cells.append({
                "coord": cell.coordinate,
                "row": row_idx,
                "col": cell.column,
                "loja": col_loja[cell.column],
                "sku": str(ws.cell(row_idx, sku_col).value or "").strip(),
                "product": str(ws.cell(row_idx, desc_col).value or "").strip() or f"Linha {row_idx}",
                "value": value,
                "url": url,
            })

    return cells, aba


# ── Extração de preço ────────────────────────────────────────────────────────
PRECO_RE = re.compile(r'R\$\s*([\d]{1,3}(?:\.[\d]{3})*,[\d]{2})')


def parse_brl(txt):
    return float(txt.replace('.', '').replace(',', '.'))


def texto_visivel(page):
    """Concatena o texto visível da página, ignorando scripts e estilos."""
    return page.evaluate("""() => {
        const skip = new Set(['SCRIPT','STYLE','NOSCRIPT','META','HEAD']);
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
        let t = '', node;
        while (node = walker.nextNode()) {
            const tag = node.parentElement && node.parentElement.tagName;
            if (!skip.has(tag)) t += node.nodeValue + ' ';
        }
        return t;
    }""")


def extrair_preco(page, loja):
    texto = texto_visivel(page)
    valores = [parse_brl(m) for m in PRECO_RE.findall(texto) if 5 < parse_brl(m) < 10000]
    if not valores:
        return None, texto
    if loja["regra_preco"] == "maior_dos_dois_primeiros" and len(valores) >= 2:
        return max(valores[0], valores[1]), texto
    return valores[0], texto


def esta_esgotado(page, texto, loja):
    """Produto indisponível quando falta o trecho configurado no HTML
    (ex.: o botão de compra) ou quando o texto traz avisos de esgotamento."""
    trecho = loja.get("indisponivel_sem")
    if trecho:
        try:
            return trecho not in page.content()
        except Exception:
            return False
    t = texto.lower()
    return any(p in t for p in ("fora de estoque", "produto esgotado", "sem estoque", "indisponível"))


# ── Interface ────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pesquisa de Preços")
        self.configure(bg="#f0f0f0")
        self.resizable(False, False)
        self.results = {}
        self.cells = []
        self.running = False
        self.aba_atual = ""
        self._login_confirmado = False
        try:
            self.lojas = carregar_lojas()
        except Exception as e:
            messagebox.showerror("Configuração", str(e))
            raise SystemExit(1)
        self._build()

    def _build(self):
        frm_xlsx = tk.LabelFrame(self, text="  📄  Planilha  ", bg="#f0f0f0",
                                 font=("Segoe UI", 10, "bold"), padx=10, pady=6)
        frm_xlsx.pack(fill="x", padx=10, pady=(10, 4))

        self.xlsx_path = tk.StringVar()
        tk.Entry(frm_xlsx, textvariable=self.xlsx_path, width=52, font=("Segoe UI", 9)).pack(side="left")
        tk.Button(frm_xlsx, text="📂 Procurar", font=("Segoe UI", 9), command=self._browse).pack(side="left", padx=(6, 0))
        self.lbl_cells = tk.Label(frm_xlsx, text="", bg="#f0f0f0", fg="#27ae60", font=("Segoe UI", 9))
        self.lbl_cells.pack(side="left", padx=(8, 0))

        lojas_login = [l for l in self.lojas.values() if l["requer_login"]]
        if lojas_login:
            nomes = ", ".join(l["id"] for l in lojas_login)
            frm_login = tk.LabelFrame(self, text="  🔐  Lojas com login  ", bg="#f0f0f0",
                                      font=("Segoe UI", 10, "bold"), padx=10, pady=8)
            frm_login.pack(fill="x", padx=10, pady=(0, 4))
            tk.Label(frm_login,
                     text=f"ℹ️  Ao iniciar a busca, o navegador abrirá na página de login de: {nomes}.\n"
                          "    Faça o login normalmente e clique no botão laranja para continuar.",
                     bg="#f0f0f0", fg="#555", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=(2, 0))

        frm_prog = tk.Frame(self, bg="#f0f0f0", padx=10)
        frm_prog.pack(fill="x")
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(frm_prog, variable=self.progress_var, maximum=100, length=780)
        self.progress.pack(fill="x", pady=4)
        self.status_lbl = tk.Label(frm_prog, text="Selecione a planilha para iniciar.",
                                   bg="#f0f0f0", font=("Segoe UI", 9), anchor="w")
        self.status_lbl.pack(fill="x")

        frm_btn = tk.Frame(self, bg="#f0f0f0", padx=10)
        frm_btn.pack(fill="x", pady=(2, 6))
        self.btn_start = tk.Button(frm_btn, text="▶  Iniciar Busca", font=("Segoe UI", 10, "bold"),
                                   bg="#27ae60", fg="white", width=18, state="disabled", command=self._start)
        self.btn_start.pack(side="left")
        self.btn_stop = tk.Button(frm_btn, text="⏹  Parar", font=("Segoe UI", 10), bg="#e74c3c", fg="white",
                                  width=10, state="disabled", command=self._stop)
        self.btn_stop.pack(side="left", padx=(6, 0))
        self.btn_save = tk.Button(frm_btn, text="💾  Salvar Planilha", font=("Segoe UI", 10), bg="#2980b9",
                                  fg="white", width=18, state="disabled", command=self._save)
        self.btn_save.pack(side="left", padx=(6, 0))
        self.btn_confirmar_login = tk.Button(frm_btn, text="✅  Estou logado, CONTINUAR",
                                             font=("Segoe UI", 10, "bold"), bg="#e67e22", fg="white",
                                             command=self._confirmar_login)
        # só aparece quando necessário

        cols = ("Célula", "SKU", "Descritivo", "Loja", "Antes", "Depois", "Status")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        for c, w in zip(cols, [55, 70, 260, 120, 80, 80, 110]):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="center" if w < 200 else "w")
        self.tree.tag_configure("ok", background="#d5f5e3")
        self.tree.tag_configure("same", background="#fef9e7")
        self.tree.tag_configure("unavail", background="#ecf0f1")
        self.tree.tag_configure("error", background="#fadbd8")
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(0, 10))
        sb.pack(side="left", fill="y", pady=(0, 10), padx=(0, 10))

    # ── Ações ─────────────────────────────────────────────────────────────────
    def _browse(self):
        p = filedialog.askopenfilename(title="Selecione a planilha",
                                       filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if p:
            self.xlsx_path.set(p)
            self._ler_planilha()

    def _ler_planilha(self):
        path = self.xlsx_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Erro", "Selecione um arquivo .xlsx válido.")
            return
        try:
            self.cells, self.aba_atual = ler_celulas(path, self.lojas)
            n = len(self.cells)
            lojas = sorted(set(c["loja"] for c in self.cells))
            self.lbl_cells.config(text=f"✅  {n} células  |  Aba: {self.aba_atual}  |  {', '.join(lojas)}")
            self.progress.config(maximum=max(n, 1))
            self.btn_start.config(state="normal")
            self._log(f"Aba '{self.aba_atual}' detectada, {n} produtos. Clique em Iniciar Busca.")
        except Exception as e:
            messagebox.showerror("Erro ao ler planilha", str(e))
            self.lbl_cells.config(text="❌ Erro ao ler")
            self.btn_start.config(state="disabled")

    def _confirmar_login(self):
        self._login_confirmado = True
        self.btn_confirmar_login.pack_forget()
        self._log("Login confirmado! Continuando busca...")

    def _start(self):
        if not self.cells:
            messagebox.showerror("Erro", "Leia a planilha primeiro.")
            return
        self.running = True
        self.results.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.progress_var.set(0)
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_save.config(state="disabled")
        threading.Thread(target=self._run, daemon=True).start()

    def _stop(self):
        self.running = False
        self._log("⏹ Busca interrompida.")

    def _log(self, msg):
        if msg == "AGUARDANDO_LOGIN":
            self.status_lbl.config(text="⏳  Faça o login no navegador e clique no botão laranja quando estiver logado.")
            self.btn_confirmar_login.pack(side="left", padx=(12, 0))
        else:
            self.status_lbl.config(text=msg)
        self.update_idletasks()

    def _add_row(self, cell, novo_valor, tag):
        old_val = cell["value"]
        old_str = f"R$ {old_val:.2f}".replace('.', ',') if old_val else "—"
        if isinstance(novo_valor, float) and tag == "unavail":
            new_str = "R$ 0,00 (indisponível)"
        elif isinstance(novo_valor, float):
            new_str = f"R$ {novo_valor:.2f}".replace('.', ',')
        else:
            new_str = str(novo_valor)
        status = {"ok": "✔ Atualizado", "same": "≈ Igual", "unavail": "⊘ Indisponível", "error": "✖ Erro"}[tag]
        self.tree.insert("", "end", values=(cell["coord"], cell["sku"], cell["product"][:45],
                                            cell["loja"], old_str, new_str, status), tags=(tag,))
        self.tree.yview_moveto(1)

    # ── Navegador com a sessão do usuário (lojas com login) ───────────────────
    def _abrir_navegador_logado(self, pw, loja):
        """Abre o navegador do próprio usuário (perfil persistente) na página de
        login da loja e aguarda a confirmação manual. Nenhuma credencial é
        armazenada pelo programa."""
        import subprocess
        import time

        self._log("Fechando o navegador se estiver aberto...")
        try:
            subprocess.run(["taskkill", "/f", "/im", "msedge.exe"],
                           creationflags=0x08000000, capture_output=True)  # CREATE_NO_WINDOW
            time.sleep(2)
        except Exception:
            pass

        localappdata = os.environ.get("LOCALAPPDATA", "")
        exe_candidatos = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            os.path.join(localappdata, "Microsoft", "Edge", "Application", "msedge.exe"),
        ]
        user_data = os.path.join(localappdata, "Microsoft", "Edge", "User Data")
        exe = next((p for p in exe_candidatos if os.path.isfile(p)), None)
        if not exe or not os.path.isdir(user_data):
            self._log("Navegador (Edge) não encontrado, lojas com login serão ignoradas.")
            return None

        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=user_data, executable_path=exe, headless=False,
            args=["--profile-directory=Default", "--no-first-run", "--no-default-browser-check"],
            no_viewport=True, locale="pt-BR")
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.set_default_timeout(30000)
        page.goto(loja["url_login"], wait_until="domcontentloaded")

        self._login_confirmado = False
        self._log("AGUARDANDO_LOGIN")
        for _ in range(300):  # até 5 minutos
            if not self.running or self._login_confirmado:
                break
            page.wait_for_timeout(1000)

        if not self._login_confirmado:
            self._log("Login não confirmado, lojas com login serão ignoradas.")
            return None
        self._log("Login confirmado! Iniciando busca...")
        return page

    # ── Thread de busca ───────────────────────────────────────────────────────
    def _run(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            self._log("❌ Playwright não instalado.")
            return

        cells, total = self.cells, len(self.cells)

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                           "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
                locale="pt-BR")
            page = ctx.new_page()
            page.set_default_timeout(20000)

            # Lojas que exigem login usam o navegador do usuário
            page_logada = None
            lojas_login = [self.lojas[c["loja"]] for c in cells if self.lojas[c["loja"]]["requer_login"]]
            if lojas_login:
                try:
                    page_logada = self._abrir_navegador_logado(pw, lojas_login[0])
                except Exception as e:
                    self._log("Erro ao abrir o navegador: " + str(e))

            for i, cell in enumerate(cells):
                if not self.running:
                    break

                loja = self.lojas[cell["loja"]]
                coord = cell["coord"]
                self._log(f"[{i+1}/{total}]  {coord}  {loja['id']}  |  {cell['product'][:45]}...")
                self.progress_var.set(i + 1)

                if loja["requer_login"] and not page_logada:
                    self._add_row(cell, "sem login", "error")
                    continue

                p = page_logada if loja["requer_login"] else page

                try:
                    p.goto(cell["url"], wait_until="domcontentloaded", timeout=25000)
                    p.wait_for_timeout(4000 if loja["requer_login"] else 2500)

                    if loja["requer_login"]:
                        if p.evaluate("() => !!document.querySelector(\"input[type='password']\")"):
                            self._add_row(cell, "sessão expirada", "error")
                            continue

                    novo, texto = extrair_preco(p, loja)

                    if esta_esgotado(p, texto, loja):
                        self.results[coord] = 0.0
                        self._add_row(cell, 0.0, "unavail")
                    elif novo is None:
                        self.results[coord] = 0.0
                        self._add_row(cell, 0.0, "error")
                    elif abs(novo - cell["value"]) < 0.01:
                        self.results[coord] = novo
                        self._add_row(cell, novo, "same")
                    else:
                        self.results[coord] = novo
                        self._add_row(cell, novo, "ok")

                except Exception as e:
                    self._add_row(cell, str(e)[:40], "error")

            for fechar in (lambda: browser.close(), lambda: page_logada.context.close()):
                try:
                    fechar()
                except Exception:
                    pass

        total_ok = len(self.tree.tag_has("ok"))
        self._log(f"✅ Concluído, {total_ok} atualizados, {len(self.results)} para salvar.")
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.btn_save.config(state="normal")
        self.running = False

    # ── Salvar ────────────────────────────────────────────────────────────────
    def _save(self):
        try:
            import copy
            import openpyxl
        except ImportError:
            messagebox.showerror("Erro", "openpyxl não instalado.")
            return

        if not self.results:
            messagebox.showwarning("Aviso", "Nenhum preço para salvar.")
            return

        src = self.xlsx_path.get()
        try:
            wb = openpyxl.load_workbook(src)

            # Nova aba com a data de hoje (ex.: 11032026); sufixo se já existir
            base = nome_nova_aba = date.today().strftime("%d%m%Y")
            sufixo = 1
            while nome_nova_aba in wb.sheetnames:
                nome_nova_aba = f"{base}_{sufixo}"
                sufixo += 1

            # Copia a aba de origem (preserva o histórico) e sua formatação condicional
            ws_orig = wb[self.aba_atual or detectar_aba_mais_recente(wb)]
            ws_nova = wb.copy_worksheet(ws_orig)
            ws_nova.title = nome_nova_aba
            for cf, rules in ws_orig.conditional_formatting._cf_rules.items():
                for cr in cf.sqref.ranges:
                    for rule in rules:
                        ws_nova.conditional_formatting.add(str(cr), copy.deepcopy(rule))

            gravados = 0
            for coord, preco in self.results.items():
                if isinstance(preco, float):
                    ws_nova[coord].value = preco
                    gravados += 1

            wb.save(src)
            messagebox.showinfo("Salvo!", f"{gravados} preços gravados na aba '{nome_nova_aba}'!\n\n{src}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))


if __name__ == "__main__":
    App().mainloop()
