import json
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import requests
import smtplib
from email.message import EmailMessage
from email.utils import formatdate

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


# PRODUÇÃO (você validou)
DANFSE_PROD = "https://adn.nfse.gov.br/danfse"

APP_NAME = "NFS-e | DANFSE em Lote (Colar ou Excel)"
CONFIG_FILE = "config.json"


def normalize_key(s: str) -> str:
    return "".join(ch for ch in (s or "").strip() if ch.isdigit())


def is_pdf_response(resp: requests.Response) -> bool:
    ct = (resp.headers.get("content-type") or "").lower()
    return ("application/pdf" in ct) or (resp.content[:4] == b"%PDF")


def get_with_retry(url: str, cert_tuple, timeout: int, retries: int = 4):
    """
    GET com mTLS + retentativas para instabilidade 5xx/Cloudflare.
    """
    last = None
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(
                url,
                cert=cert_tuple,
                timeout=timeout,
                headers={
                    "Accept": "application/pdf, application/json;q=0.9, */*;q=0.8",
                    "User-Agent": "NfseClient/1.0 (Python requests)"
                }
            )
            last = r
            # gateways instáveis
            if r.status_code in (502, 503, 504, 520, 521, 522):
                time.sleep(1.5 * attempt)
                continue
            return r
        except requests.RequestException as e:
            last = e
            time.sleep(1.5 * attempt)
    return last


def send_email_smtp(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_pass: str,
    from_email: str,
    to_email: str,
    subject: str,
    body: str,
    attachment_path: Path
):
    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.set_content(body)

    data = attachment_path.read_bytes()
    msg.add_attachment(data, maintype="application", subtype="pdf", filename=attachment_path.name)

    with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as s:
        s.ehlo()
        s.starttls()
        s.ehlo()
        s.login(smtp_user, smtp_pass)
        s.send_message(msg)


def create_excel_template(path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "DANFSE"

    headers = ["CHAVE", "EMAIL", "ASSUNTO", "CORPO"]
    ws.append(headers)
    ws.append(["", "", "", ""])

    widths = [48, 30, 40, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    wb.save(path)


def read_excel_rows(path: Path):
    """
    Retorna lista de dicts:
      {"chave":..., "email":..., "assunto":..., "corpo":...}
    """
    wb = load_workbook(path)
    ws = wb.active

    header = []
    for cell in ws[1]:
        header.append((cell.value or "").strip().upper())
    idx = {name: header.index(name) for name in header if name}

    if "CHAVE" not in idx:
        raise RuntimeError("Planilha inválida: coluna CHAVE é obrigatória.")

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        chave_raw = r[idx["CHAVE"]] if idx["CHAVE"] < len(r) else ""
        chave = normalize_key(str(chave_raw or ""))
        if not chave:
            continue

        email = ""
        assunto = ""
        corpo = ""
        if "EMAIL" in idx and idx["EMAIL"] < len(r):
            email = str(r[idx["EMAIL"]] or "").strip()
        if "ASSUNTO" in idx and idx["ASSUNTO"] < len(r):
            assunto = str(r[idx["ASSUNTO"]] or "").strip()
        if "CORPO" in idx and idx["CORPO"] < len(r):
            corpo = str(r[idx["CORPO"]] or "").strip()

        rows.append({"chave": chave, "email": email, "assunto": assunto, "corpo": corpo})

    # remove duplicadas preservando ordem
    seen = set()
    uniq = []
    for item in rows:
        if item["chave"] not in seen:
            uniq.append(item)
            seen.add(item["chave"])
    return uniq


def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def salvar_falhas_txt(falhas: list[str], pasta: Path) -> Path:
    p = pasta / f"falhas_{now_stamp()}.txt"
    p.write_text("\n".join(falhas), encoding="utf-8")
    return p


def salvar_resultado_xlsx(linhas: list[dict], pasta: Path) -> Path:
    """
    linhas: lista de dict:
      chave, status, pdf_path, email_to, erro, origem
    """
    p = pasta / f"resultado_{now_stamp()}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "RESULTADO"

    headers = ["CHAVE", "STATUS", "PDF_PATH", "EMAIL_TO", "ERRO", "ORIGEM"]
    ws.append(headers)

    for row in linhas:
        ws.append([
            row.get("chave", ""),
            row.get("status", ""),
            row.get("pdf_path", ""),
            row.get("email_to", ""),
            row.get("erro", ""),
            row.get("origem", ""),
        ])

    widths = [48, 12, 60, 30, 60, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    wb.save(p)
    return p


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1060x740")
        self.resizable(False, False)

        self.cfg = self.load_config()

        # cert e saída
        self.cert_pem = tk.StringVar(value=self.cfg.get("cert_pem", str(Path.cwd() / "cert_out" / "client_cert.pem")))
        self.key_pem = tk.StringVar(value=self.cfg.get("key_pem", str(Path.cwd() / "cert_out" / "client_key.pem")))
        self.save_dir = tk.StringVar(value=self.cfg.get("save_dir", str(Path.cwd() / "downloads")))

        # SMTP (opcional)
        self.smtp_host = tk.StringVar(value=self.cfg.get("smtp_host", ""))
        self.smtp_port = tk.StringVar(value=str(self.cfg.get("smtp_port", "587")))
        self.smtp_user = tk.StringVar(value=self.cfg.get("smtp_user", ""))
        self.smtp_pass = tk.StringVar(value=self.cfg.get("smtp_pass", ""))
        self.from_email = tk.StringVar(value=self.cfg.get("from_email", ""))
        self.default_to = tk.StringVar(value=self.cfg.get("default_to", ""))
        self.default_subject = tk.StringVar(value=self.cfg.get("default_subject", "NFSe - DANFSE (PDF)"))
        self.default_body = tk.StringVar(value=self.cfg.get("default_body", "Segue em anexo o DANFSE (PDF)."))

        # execução
        self.send_email_each = tk.BooleanVar(value=bool(self.cfg.get("send_email_each", False)))
        self.progress = tk.DoubleVar(value=0.0)

        # robustez lote
        self.pause_between_items = tk.DoubleVar(value=float(self.cfg.get("pause_between_items", 0.5)))  # segundos
        self.reprocess_rounds = tk.IntVar(value=int(self.cfg.get("reprocess_rounds", 3)))
        self.cooldown_between_rounds = tk.IntVar(value=int(self.cfg.get("cooldown_between_rounds", 10)))  # segundos
        self.cooldown_every_n = tk.IntVar(value=int(self.cfg.get("cooldown_every_n", 5)))
        self.cooldown_seconds = tk.IntVar(value=int(self.cfg.get("cooldown_seconds", 3)))

        # dados carregados do Excel
        self.excel_rows = []

        self._build_ui()

    # ---------- Config ----------
    def load_config(self):
        p = Path(CONFIG_FILE)
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def save_config(self):
        cfg = {
            "cert_pem": self.cert_pem.get().strip(),
            "key_pem": self.key_pem.get().strip(),
            "save_dir": self.save_dir.get().strip(),

            "smtp_host": self.smtp_host.get().strip(),
            "smtp_port": int(self.smtp_port.get().strip() or "587"),
            "smtp_user": self.smtp_user.get().strip(),
            "smtp_pass": self.smtp_pass.get(),
            "from_email": self.from_email.get().strip(),
            "default_to": self.default_to.get().strip(),
            "default_subject": self.default_subject.get().strip(),
            "default_body": self.body_txt.get("1.0", "end").strip(),
            "send_email_each": bool(self.send_email_each.get()),

            "pause_between_items": float(self.pause_between_items.get()),
            "reprocess_rounds": int(self.reprocess_rounds.get()),
            "cooldown_between_rounds": int(self.cooldown_between_rounds.get()),
            "cooldown_every_n": int(self.cooldown_every_n.get()),
            "cooldown_seconds": int(self.cooldown_seconds.get()),
        }
        Path(CONFIG_FILE).write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        self.cfg = cfg

    # ---------- Helpers ----------
    def log(self, msg: str):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.update_idletasks()

    def _cert_tuple(self):
        cert = Path(self.cert_pem.get().strip().strip('"'))
        key = Path(self.key_pem.get().strip().strip('"'))
        if not cert.exists():
            raise FileNotFoundError(f"client_cert.pem não encontrado: {cert}")
        if not key.exists():
            raise FileNotFoundError(f"client_key.pem não encontrado: {key}")
        return (str(cert), str(key))

    def _out_dir(self):
        d = Path(self.save_dir.get().strip().strip('"'))
        d.mkdir(parents=True, exist_ok=True)
        return d

    def smtp_ready(self):
        host = self.smtp_host.get().strip()
        user = self.smtp_user.get().strip()
        pw = self.smtp_pass.get()
        to_e = self.default_to.get().strip()
        from_e = self.from_email.get().strip() or user
        return bool(host and user and pw and to_e and from_e)

    def baixar_pdf(self, chave: str) -> tuple[Path | None, str]:
        """
        Retorna (pdf_path_ou_None, erro_string)
        """
        url = f"{DANFSE_PROD}/{chave}"
        resp = get_with_retry(url, self._cert_tuple(), timeout=70, retries=4)

        if isinstance(resp, Exception):
            return None, f"rede: {resp}"

        if resp.status_code == 200 and is_pdf_response(resp):
            out = self._out_dir() / f"DANFSE_{chave}.pdf"
            out.write_bytes(resp.content)
            return out, ""

        # Captura detalhes
        ct = resp.headers.get("content-type", "")
        err_txt = f"status={resp.status_code} ct={ct}"
        try:
            snippet = resp.text[:200].replace("\n", " ").strip()
            if snippet:
                err_txt += f" resp={snippet}"
        except Exception:
            pass
        return None, err_txt

    def send_email_for_pdf(self, chave: str, pdf: Path, email: str = "", assunto: str = "", corpo: str = ""):
        host = self.smtp_host.get().strip()
        port = int(self.smtp_port.get().strip() or "587")
        user = self.smtp_user.get().strip()
        pw = self.smtp_pass.get()
        from_e = self.from_email.get().strip() or user

        to_e = (email or "").strip() or self.default_to.get().strip()
        subject = (assunto or "").strip() or self.default_subject.get().strip()
        body = (corpo or "").strip() or self.body_txt.get("1.0", "end").strip()

        subject = f"{subject} | {chave}"
        send_email_smtp(host, port, user, pw, from_e, to_e, subject, body, pdf)
        return to_e

    # ---------- UI ----------
    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        tab_lote = ttk.Frame(nb, padding=12)
        tab_cfg = ttk.Frame(nb, padding=12)

        nb.add(tab_lote, text="Lote (Colar ou Excel)")
        nb.add(tab_cfg, text="Configuração")

        # ----- TAB LOTE -----
        ttk.Label(tab_lote, text="Opção 1 — Colar chaves (1 por linha):").grid(row=0, column=0, sticky="w")
        self.txt_keys = tk.Text(tab_lote, width=120, height=10)
        self.txt_keys.grid(row=1, column=0, columnspan=5, sticky="w", pady=(6, 8))

        rowbtn = ttk.Frame(tab_lote)
        rowbtn.grid(row=2, column=0, columnspan=5, sticky="w")
        ttk.Button(rowbtn, text="Limpar", command=lambda: self.txt_keys.delete("1.0", "end")).pack(side="left")
        ttk.Button(rowbtn, text="Executar lote (do texto)", command=self.run_from_text).pack(side="left", padx=(8, 0))

        ttk.Separator(tab_lote).grid(row=3, column=0, columnspan=5, sticky="we", pady=(12, 10))

        ttk.Label(tab_lote, text="Opção 2 — Excel (.xlsx):").grid(row=4, column=0, sticky="w")

        rowbtn2 = ttk.Frame(tab_lote)
        rowbtn2.grid(row=5, column=0, columnspan=5, sticky="w", pady=(6, 6))
        ttk.Button(rowbtn2, text="Baixar planilha modelo", command=self.download_template).pack(side="left")
        ttk.Button(rowbtn2, text="Importar Excel", command=self.import_excel).pack(side="left", padx=(8, 0))
        ttk.Button(rowbtn2, text="Executar lote (do Excel importado)", command=self.run_from_excel).pack(side="left", padx=(8, 0))

        self.lbl_excel = ttk.Label(tab_lote, text="Nenhum Excel importado.")
        self.lbl_excel.grid(row=6, column=0, columnspan=5, sticky="w")

        ttk.Checkbutton(
            tab_lote, text="Enviar e-mail para cada PDF baixado",
            variable=self.send_email_each
        ).grid(row=7, column=0, sticky="w", pady=(10, 0))

        ttk.Label(tab_lote, text="Progresso:").grid(row=8, column=0, sticky="w", pady=(12, 0))
        self.pb = ttk.Progressbar(tab_lote, variable=self.progress, maximum=100.0, length=900)
        self.pb.grid(row=9, column=0, columnspan=5, sticky="w", pady=(6, 0))

        ttk.Separator(tab_lote).grid(row=10, column=0, columnspan=5, sticky="we", pady=(12, 10))

        ttk.Label(tab_lote, text="Log:").grid(row=11, column=0, sticky="nw")
        self.txt_log = tk.Text(tab_lote, width=120, height=18)
        self.txt_log.grid(row=11, column=1, columnspan=4, sticky="w")

        # ----- TAB CONFIG -----
        ttk.Label(tab_cfg, text="Certificado (mTLS):").grid(row=0, column=0, sticky="w")

        ttk.Label(tab_cfg, text="client_cert.pem:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        c1 = ttk.Frame(tab_cfg)
        c1.grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Entry(c1, textvariable=self.cert_pem, width=78).pack(side="left", padx=(0, 8))
        ttk.Button(c1, text="Selecionar...", command=self.pick_cert).pack(side="left")

        ttk.Label(tab_cfg, text="client_key.pem:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        c2 = ttk.Frame(tab_cfg)
        c2.grid(row=2, column=1, sticky="w", pady=(8, 0))
        ttk.Entry(c2, textvariable=self.key_pem, width=78).pack(side="left", padx=(0, 8))
        ttk.Button(c2, text="Selecionar...", command=self.pick_key).pack(side="left")

        ttk.Label(tab_cfg, text="Pasta para salvar PDFs:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        c3 = ttk.Frame(tab_cfg)
        c3.grid(row=3, column=1, sticky="w", pady=(8, 0))
        ttk.Entry(c3, textvariable=self.save_dir, width=78).pack(side="left", padx=(0, 8))
        ttk.Button(c3, text="Escolher...", command=self.pick_dir).pack(side="left")

        ttk.Separator(tab_cfg).grid(row=4, column=0, columnspan=2, sticky="we", pady=(14, 10))

        ttk.Label(tab_cfg, text="Robustez do lote (recomendado manter assim):").grid(row=5, column=0, sticky="w")

        rb = ttk.Frame(tab_cfg)
        rb.grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 0))

        def add_row(r, label, var, width=12):
            ttk.Label(rb, text=label).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(rb, textvariable=var, width=width).grid(row=r, column=1, sticky="w", padx=(8, 0), pady=3)

        add_row(0, "Pausa entre itens (s):", self.pause_between_items)
        add_row(1, "Rodadas de reprocessamento:", self.reprocess_rounds)
        add_row(2, "Cooldown entre rodadas (s):", self.cooldown_between_rounds)
        add_row(3, "Cooldown a cada N itens:", self.cooldown_every_n)
        add_row(4, "Cooldown (s):", self.cooldown_seconds)

        ttk.Separator(tab_cfg).grid(row=7, column=0, columnspan=2, sticky="we", pady=(14, 10))

        ttk.Label(tab_cfg, text="SMTP (opcional — necessário se marcar envio automático):").grid(row=8, column=0, sticky="w")

        grid = ttk.Frame(tab_cfg)
        grid.grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))

        def add_smtp(r, label, var, width=36, show=None):
            ttk.Label(grid, text=label).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(grid, textvariable=var, width=width, show=show).grid(row=r, column=1, sticky="w", padx=(8, 0), pady=3)

        add_smtp(0, "SMTP Host:", self.smtp_host)
        add_smtp(1, "SMTP Port:", self.smtp_port, width=10)
        add_smtp(2, "SMTP User:", self.smtp_user)
        add_smtp(3, "SMTP Pass:", self.smtp_pass, show="*")
        add_smtp(4, "From:", self.from_email)
        add_smtp(5, "To padrão:", self.default_to)
        add_smtp(6, "Assunto padrão:", self.default_subject, width=60)

        ttk.Label(grid, text="Corpo padrão:").grid(row=7, column=0, sticky="nw", pady=3)
        self.body_txt = tk.Text(grid, width=62, height=5)
        self.body_txt.grid(row=7, column=1, sticky="w", padx=(8, 0), pady=3)
        self.body_txt.insert("1.0", self.default_body.get())

        ttk.Button(tab_cfg, text="Salvar configuração", command=self.on_save_config).grid(row=10, column=0, sticky="w", pady=(12, 0))

    # ---------- Buttons ----------
    def on_save_config(self):
        self.save_config()
        messagebox.showinfo("OK", f"Configuração salva em {Path(CONFIG_FILE).resolve()}")

    def pick_cert(self):
        p = filedialog.askopenfilename(title="Selecione client_cert.pem", filetypes=[("PEM", "*.pem"), ("Todos", "*.*")])
        if p:
            self.cert_pem.set(p)

    def pick_key(self):
        p = filedialog.askopenfilename(title="Selecione client_key.pem", filetypes=[("PEM", "*.pem"), ("Todos", "*.*")])
        if p:
            self.key_pem.set(p)

    def pick_dir(self):
        p = filedialog.askdirectory(title="Selecione a pasta de saída")
        if p:
            self.save_dir.set(p)

    def download_template(self):
        path = filedialog.asksaveasfilename(
            title="Salvar planilha modelo",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return
        p = Path(path)
        create_excel_template(p)
        messagebox.showinfo("OK", f"Planilha modelo salva em:\n{p}")

    def import_excel(self):
        p = filedialog.askopenfilename(
            title="Importar planilha Excel",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not p:
            return
        rows = read_excel_rows(Path(p))
        self.excel_rows = rows
        self.lbl_excel.config(text=f"Excel importado: {Path(p).name} | Linhas válidas: {len(rows)}")
        messagebox.showinfo("OK", f"Importado com sucesso.\nLinhas válidas: {len(rows)}")

    # ---------- Batch runners ----------
    def run_from_text(self):
        keys = []
        for line in self.txt_keys.get("1.0", "end").splitlines():
            k = normalize_key(line)
            if k:
                keys.append(k)

        # dedup
        seen = set()
        uniq = []
        for k in keys:
            if k not in seen:
                uniq.append(k)
                seen.add(k)

        if not uniq:
            messagebox.showerror("Erro", "Nenhuma chave válida no texto.")
            return

        self.save_config()
        items = [{"chave": k, "email": "", "assunto": "", "corpo": "", "origem": "texto"} for k in uniq]
        self.run_batch_items(items)

    def run_from_excel(self):
        if not self.excel_rows:
            messagebox.showerror("Erro", "Nenhum Excel importado.")
            return

        self.save_config()
        items = []
        for r in self.excel_rows:
            items.append({
                "chave": r.get("chave", ""),
                "email": r.get("email", ""),
                "assunto": r.get("assunto", ""),
                "corpo": r.get("corpo", ""),
                "origem": "excel"
            })
        self.run_batch_items(items)

    def run_batch_items(self, items: list[dict]):
        do_email = self.send_email_each.get()
        if do_email and not self.smtp_ready():
            messagebox.showerror("Erro", "Envio por e-mail marcado, mas SMTP/To padrão não está completo na Configuração.")
            return

        total = len(items)
        ok = 0
        fail = 0
        falhas = []

        resultados = []  # para resultado.xlsx

        self.progress.set(0.0)
        self.log("=" * 110)
        self.log(f"Iniciando lote: {total} item(ns) | Pausa={self.pause_between_items.get()}s | Reprocess={self.reprocess_rounds.get()} rodada(s)")

        # --- 1ª passada ---
        for i, item in enumerate(items, start=1):
            chave = item["chave"]
            origem = item.get("origem", "")

            self.log(f"[{i}/{total}] Processando chave: {chave}")
            pdf, err = self.baixar_pdf(chave)

            email_to = ""
            if pdf:
                ok += 1
                if do_email:
                    try:
                        email_to = self.send_email_for_pdf(
                            chave=chave,
                            pdf=pdf,
                            email=item.get("email", ""),
                            assunto=item.get("assunto", ""),
                            corpo=item.get("corpo", "")
                        )
                    except Exception as e:
                        # se email falhar, não consideramos "falha de pdf"
                        self.log(f"  [EMAIL] ERRO: {e}")

                resultados.append({
                    "chave": chave,
                    "status": "OK",
                    "pdf_path": str(pdf),
                    "email_to": email_to,
                    "erro": "",
                    "origem": origem
                })
            else:
                fail += 1
                falhas.append(chave)
                resultados.append({
                    "chave": chave,
                    "status": "FALHA",
                    "pdf_path": "",
                    "email_to": "",
                    "erro": err,
                    "origem": origem
                })
                self.log(f"  FALHA → {err}")

            # pausa curta (evita rajada/502)
            time.sleep(float(self.pause_between_items.get()))
            self.progress.set((i / total) * 100.0)

            # cooldown a cada N itens
            n = int(self.cooldown_every_n.get())
            if n > 0 and i % n == 0:
                time.sleep(int(self.cooldown_seconds.get()))

        # --- Reprocessamento das falhas (fila) ---
        rounds = int(self.reprocess_rounds.get())
        if falhas and rounds > 0:
            self.log("-" * 110)
            self.log(f"Entrando em reprocessamento: {len(falhas)} falha(s)")

            restantes = falhas[:]
            for rodada in range(1, rounds + 1):
                if not restantes:
                    break

                self.log(f"=== Rodada {rodada}/{rounds} | Restantes: {len(restantes)} ===")
                time.sleep(int(self.cooldown_between_rounds.get()))

                novas_restantes = []
                for j, chave in enumerate(restantes, start=1):
                    self.log(f"[RETRY {rodada}] ({j}/{len(restantes)}) {chave}")
                    pdf, err = self.baixar_pdf(chave)

                    if pdf:
                        ok += 1
                        fail -= 1  # recuperou uma falha anterior

                        # Atualiza linha correspondente em resultados (marca como RECUPERADO)
                        for row in resultados:
                            if row["chave"] == chave and row["status"] == "FALHA":
                                row["status"] = "RECUPERADO"
                                row["pdf_path"] = str(pdf)
                                row["erro"] = ""
                                break

                        if do_email:
                            try:
                                # tenta achar item original para email/assunto/corpo
                                original = next((x for x in items if x["chave"] == chave), {})
                                email_to = self.send_email_for_pdf(
                                    chave=chave,
                                    pdf=pdf,
                                    email=original.get("email", ""),
                                    assunto=original.get("assunto", ""),
                                    corpo=original.get("corpo", "")
                                )
                                # salva email_to no resultado
                                for row in resultados:
                                    if row["chave"] == chave and row["status"] == "RECUPERADO":
                                        row["email_to"] = email_to
                                        break
                            except Exception as e:
                                self.log(f"  [EMAIL] ERRO: {e}")

                        self.log("  RECUPERADO ✅")
                    else:
                        novas_restantes.append(chave)
                        # atualiza erro mais recente
                        for row in resultados:
                            if row["chave"] == chave and row["status"] == "FALHA":
                                row["erro"] = err
                                break
                        self.log(f"  AINDA FALHA → {err}")

                    # cooldown leve para não “estressar” o serviço
                    if int(self.cooldown_every_n.get()) > 0 and j % int(self.cooldown_every_n.get()) == 0:
                        time.sleep(int(self.cooldown_seconds.get()))
                    else:
                        time.sleep(float(self.pause_between_items.get()))

                restantes = novas_restantes

            # salva falhas finais (se houver)
            falhas_finais = restantes[:]
        else:
            falhas_finais = falhas[:]

        # --- Salvar relatórios ---
        out_dir = self._out_dir()
        resultado_path = salvar_resultado_xlsx(resultados, out_dir)
        self.log("-" * 110)
        self.log(f"Relatório salvo: {resultado_path}")

        falhas_path = None
        if falhas_finais:
            falhas_path = salvar_falhas_txt(falhas_finais, out_dir)
            self.log(f"Falhas finais salvas: {falhas_path}")

        self.log("-" * 110)
        self.log(f"Fim do lote. Sucesso total: {ok} | Falhas finais: {len(falhas_finais)}")
        if falhas_path:
            messagebox.showwarning(
                "Concluído (com falhas)",
                f"Lote finalizado.\n\nSucesso: {ok}\nFalhas finais: {len(falhas_finais)}\n\n"
                f"Relatório: {resultado_path}\nFalhas: {falhas_path}"
            )
        else:
            messagebox.showinfo(
                "Concluído",
                f"Lote finalizado com sucesso!\n\nSucesso: {ok}\n\nRelatório: {resultado_path}"
            )


if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    App().mainloop()
