"""
renomear_certificados.py
IA extrai data e palestra. Emissor é inserido manualmente com autocomplete.
Formato: AAAAMMDD - Nome da Palestra - Emissor.pdf

Dependências:
    pip install pypdf anthropic pymupdf Pillow
"""

import os
import re
import json
import base64
import threading
import winreg
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader
import anthropic
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF

# ── Configuração ───────────────────────────────────────────────────────────────

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = "claude-opus-4-5"

BG      = "#0f0f11"
SURFACE = "#1a1a1f"
BORDER  = "#2a2a32"
ACCENT  = "#7c6af7"
GREEN   = "#4ade80"
RED     = "#f87171"
TEXT    = "#e8e8f0"
MUTED   = "#6b6b7e"

# ── Configurações persistentes ────────────────────────────────────────────────

class ConfigManager:
    ARQUIVO = Path.home() / ".certificados_config.json"
    DEFAULTS = {
        "ultima_pasta": "",
        "modo": "lote",
        "geometry": "",
    }

    def __init__(self):
        self._dados: dict = self._carregar()

    def _carregar(self) -> dict:
        try:
            dados = json.loads(self.ARQUIVO.read_text(encoding="utf-8"))
            return {**self.DEFAULTS, **dados}
        except Exception:
            return dict(self.DEFAULTS)

    def salvar(self):
        self.ARQUIVO.write_text(
            json.dumps(self._dados, ensure_ascii=False, indent=2), encoding="utf-8")

    def get(self, chave: str):
        return self._dados.get(chave, self.DEFAULTS.get(chave, ""))

    def set(self, chave: str, valor):
        self._dados[chave] = valor
        self.salvar()


# ── Gerenciador de emissores (persistente) ─────────────────────────────────────

class EmissorManager:
    ARQUIVO = Path.home() / ".certificados_emissores.json"

    def __init__(self):
        self._lista: list[str] = self._carregar()

    def _carregar(self) -> list[str]:
        try:
            dados = json.loads(self.ARQUIVO.read_text(encoding="utf-8"))
            return sorted([str(e) for e in dados if e], key=str.lower)
        except Exception:
            return []

    def adicionar(self, emissor: str) -> bool:
        emissor = emissor.strip()
        if not emissor or emissor in self._lista:
            return False
        self._lista.append(emissor)
        self._lista.sort(key=str.lower)
        self.ARQUIVO.write_text(
            json.dumps(self._lista, ensure_ascii=False, indent=2), encoding="utf-8")
        return True

    def filtrar(self, texto: str) -> list[str]:
        t = texto.lower()
        return [e for e in self._lista if t in e.lower()]

    @property
    def lista(self) -> list[str]:
        return list(self._lista)

# ── OneDrive ───────────────────────────────────────────────────────────────────

def detectar_onedrives() -> list[dict]:
    drives = []
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"SOFTWARE\Microsoft\OneDrive\Accounts")
        i = 0
        while True:
            try:
                sub = winreg.OpenKey(key, winreg.EnumKey(key, i))
                try:
                    email  = winreg.QueryValueEx(sub, "UserEmail")[0]
                    folder = winreg.QueryValueEx(sub, "UserFolder")[0]
                    drives.append({"label": email, "caminho": Path(folder)})
                except Exception:
                    pass
                winreg.CloseKey(sub)
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
    except Exception:
        pass
    if not drives:
        for p in Path.home().iterdir():
            if (p.is_dir()
                    and "onedrive" in p.name.lower()
                    and "temp" not in p.name.lower()):
                drives.append({"label": p.name, "caminho": p})
    return drives

# ── PDF ────────────────────────────────────────────────────────────────────────

# WinError 4392 = ERROR_CLOUD_FILE_NOT_IN_SYNC
# WinError 362  = ERROR_CLOUD_FILE_PROVIDER_NOT_RUNNING (0x8007016A)
_WINERRORS_CLOUD = {362, 363, 364, 4390, 4391, 4392, 4393, 4394}

def verificar_arquivo(caminho: Path) -> str | None:
    """Retorna mensagem de erro se o arquivo não estiver acessível localmente."""
    try:
        with open(caminho, "rb") as f:
            f.read(16)
        return None
    except OSError as e:
        winerr = getattr(e, "winerror", 0) or 0
        if winerr in _WINERRORS_CLOUD or "cloud" in str(e).lower():
            return "☁ arquivo disponível apenas online — faça o download primeiro no OneDrive"
        return f"arquivo inacessível: {e}"
    except Exception as e:
        return f"erro ao acessar: {e}"


def extrair_texto(caminho: Path, max_chars: int = 3000) -> str:
    try:
        reader = PdfReader(str(caminho))
        texto = ""
        for page in reader.pages[:4]:
            texto += page.extract_text() or ""
            if len(texto) >= max_chars:
                break
        return texto[:max_chars]
    except Exception as e:
        return f"[ERRO: {e}]"


def texto_tem_qualidade(texto: str, min_letras: int = 40, min_ratio: float = 0.35) -> bool:
    """Retorna False se o texto tiver poucos caracteres alfabéticos (provável garbled/encoding ruim)."""
    letras = sum(1 for c in texto if c.isalpha())
    nao_espaco = sum(1 for c in texto if not c.isspace())
    if letras < min_letras:
        return False
    if nao_espaco == 0:
        return False
    return (letras / nao_espaco) >= min_ratio


def extrair_campos(texto: str, client: anthropic.Anthropic) -> dict:
    prompt = f"""Você é um extrator de informações de certificados de eventos/palestras.

Analise o texto abaixo e extraia EXATAMENTE dois campos:
1. data: a data do evento/palestra no formato AAAAMMDD (ex: 20230615). Se só tiver mês e ano, use 01 como dia. Se só tiver ano, use 0101.
2. palestra: o nome/título da palestra, curso ou evento (seja conciso, máx 60 chars)

Responda APENAS com JSON válido, sem explicações, sem markdown:
{{"data": "...", "palestra": "..."}}

Se não conseguir encontrar um campo com certeza, use "" para aquele campo.

TEXTO DO CERTIFICADO:
{texto}
"""
    try:
        resp = client.messages.create(
            model=MODEL, max_tokens=200,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = re.sub(r"```(?:json)?|```", "", resp.content[0].text.strip()).strip()
        return json.loads(raw)
    except Exception as e:
        return {"data": "", "palestra": "", "_erro": str(e)}


def extrair_campos_via_visao(caminho: Path, client: anthropic.Anthropic) -> dict:
    """Extrai campos de PDF image-only usando Claude Vision (sem OCR externo)."""
    prompt = """Você é um extrator de informações de certificados de eventos/palestras.

Analise a imagem do certificado e extraia EXATAMENTE dois campos:
1. data: a data do evento/palestra no formato AAAAMMDD (ex: 20230615). Se só tiver mês e ano, use 01 como dia. Se só tiver ano, use 0101.
2. palestra: o nome/título da palestra, curso ou evento (seja conciso, máx 60 chars)

Responda APENAS com JSON válido, sem explicações, sem markdown:
{"data": "...", "palestra": "..."}

Se não conseguir encontrar um campo com certeza, use "" para aquele campo."""
    try:
        doc  = fitz.open(str(caminho))
        page = doc[0]
        pix  = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
        img  = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        buf  = BytesIO()
        img.save(buf, format="JPEG", quality=85)
        img_b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")

        resp = client.messages.create(
            model=MODEL, max_tokens=200,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": "image/jpeg", "data": img_b64}},
                {"type": "text", "text": prompt},
            ]}]
        )
        raw = re.sub(r"```(?:json)?|```", "", resp.content[0].text.strip()).strip()
        return json.loads(raw)
    except Exception as e:
        return {"data": "", "palestra": "", "_erro": str(e)}


_RE_JA_FORMATADO = re.compile(r"^\d{8}( - .+)?$")

def ja_formatado(nome_stem: str) -> bool:
    """Retorna True se o nome já segue o padrão AAAAMMDD[ - ...][.pdf]."""
    stem = Path(nome_stem).stem if nome_stem.lower().endswith(".pdf") else nome_stem
    return bool(_RE_JA_FORMATADO.match(stem))


def montar_nome(data: str, palestra: str, emissor: str = "") -> str:
    partes = [data.strip(), palestra.strip(), emissor.strip()]
    partes = [p for p in partes if p]
    nome = " - ".join(partes)
    nome = re.sub(r'[\\/*?:"<>|]', "", nome)
    return nome.strip() or "sem_nome"


def renderizar_pdf(caminho: Path, largura: int = 360) -> ImageTk.PhotoImage | None:
    try:
        doc = fitz.open(str(caminho))
        page = doc[0]
        scale = largura / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None

# ── Aplicação ──────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Renomear Certificados")
        self.geometry("1200x800")
        self.minsize(900, 640)
        self.configure(bg=BG)

        self.config        = ConfigManager()
        self.emissores     = EmissorManager()
        self.client        = None
        self.resultados:   list[dict] = []
        self._foto_ref     = None
        self._idx:         int | None = None
        self._preview_lock = threading.Lock()
        self._preview_token = 0
        self._drives:      list[dict] = []

        self._estilos()
        self._build_ui()
        self._detectar_onedrives()
        self._restaurar_config()
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    # ── Estilos ────────────────────────────────────────────────────────────────

    def _estilos(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",    background=BG)
        s.configure("TLabel",    background=BG, foreground=TEXT, font=("Segoe UI", 10))
        s.configure("H.TLabel",  background=BG, foreground=TEXT, font=("Segoe UI", 13, "bold"))
        s.configure("Muted.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))
        s.configure("Cap.TLabel",   background=BG, foreground=MUTED,
                    font=("Courier New", 8, "bold"))

        s.configure("TButton", background=ACCENT, foreground="white",
                    font=("Segoe UI", 9, "bold"), relief="flat",
                    borderwidth=0, padding=(10, 5))
        s.map("TButton", background=[("active", "#6a59e0"), ("disabled", "#3a3a4a")])

        s.configure("Ghost.TButton", background=SURFACE, foreground=TEXT,
                    font=("Segoe UI", 9), relief="flat", borderwidth=1, padding=(8, 5))
        s.map("Ghost.TButton", background=[("active", BORDER)])

        s.configure("Danger.TButton", background=SURFACE, foreground=RED,
                    font=("Segoe UI", 9), relief="flat", borderwidth=1, padding=(8, 5))
        s.map("Danger.TButton", background=[("active", "#2a1010")])

        s.configure("TCombobox", fieldbackground=SURFACE, background=SURFACE,
                    foreground=TEXT, bordercolor=BORDER, font=("Segoe UI", 9))
        s.map("TCombobox", fieldbackground=[("readonly", SURFACE)],
              selectbackground=[("readonly", ACCENT)])

        s.configure("TEntry", fieldbackground=SURFACE, foreground=TEXT,
                    insertcolor=TEXT, bordercolor=BORDER, relief="flat")

        s.configure("TRadiobutton", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        s.map("TRadiobutton", background=[("active", BG)])

        s.configure("TCheckbutton", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        s.map("TCheckbutton", background=[("active", BG)])

        s.configure("Treeview", background=SURFACE, foreground=TEXT,
                    fieldbackground=SURFACE, rowheight=28, font=("Segoe UI", 9),
                    bordercolor=BORDER, relief="flat")
        s.configure("Treeview.Heading", background=BG, foreground=MUTED,
                    font=("Courier New", 8, "bold"), relief="flat", padding=(8, 6))
        s.map("Treeview", background=[("selected", "#2a2a42")],
              foreground=[("selected", TEXT)])

        s.configure("Horizontal.TProgressbar",
                    background=ACCENT, troughcolor=BORDER, bordercolor=BG)
        s.configure("TSeparator", background=BORDER)

    # ── Interface ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Cabeçalho
        hdr = ttk.Frame(self, padding=(16, 14, 16, 10))
        hdr.pack(fill="x")
        ttk.Label(hdr, text="Renomear Certificados", style="H.TLabel").pack(side="left")
        ttk.Label(hdr, text="  ·  powered by Claude AI",
                  style="Muted.TLabel").pack(side="left", pady=(3, 0))
        ttk.Separator(self).pack(fill="x")

        # Controles
        ctrl = ttk.Frame(self, padding=(16, 10))
        ctrl.pack(fill="x")

        ttk.Label(ctrl, text="Conta OneDrive:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.combo_drive = ttk.Combobox(ctrl, state="readonly", width=34)
        self.combo_drive.grid(row=0, column=1, sticky="w", padx=(0, 12))
        self.combo_drive.bind("<<ComboboxSelected>>", self._on_drive_selected)

        ttk.Label(ctrl, text="Pasta:").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.var_pasta = tk.StringVar()
        ttk.Entry(ctrl, textvariable=self.var_pasta, width=30,
                  font=("Segoe UI", 9)).grid(row=0, column=3, sticky="w", padx=(0, 4))
        ttk.Button(ctrl, text="Navegar", style="Ghost.TButton",
                   command=self._escolher_pasta).grid(row=0, column=4, padx=(0, 16))

        ttk.Label(ctrl, text="Modo:").grid(row=0, column=5, sticky="w", padx=(0, 6))
        self.var_modo = tk.StringVar(value="lote")
        ttk.Radiobutton(ctrl, text="Em lote",
                        variable=self.var_modo, value="lote").grid(row=0, column=6, padx=(0, 4))
        ttk.Radiobutton(ctrl, text="Individual",
                        variable=self.var_modo, value="individual").grid(row=0, column=7, padx=(0, 16))

        self.btn_analisar = ttk.Button(ctrl, text="▶  Analisar com IA", command=self._analisar)
        self.btn_analisar.grid(row=0, column=8)

        ttk.Separator(self).pack(fill="x")

        # Progresso (oculto até ser necessário)
        self.frame_prog = ttk.Frame(self, padding=(16, 8, 16, 4))
        self.progress = ttk.Progressbar(self.frame_prog, style="Horizontal.TProgressbar")
        self.progress.pack(fill="x")
        self.lbl_prog = ttk.Label(self.frame_prog, text="", style="Muted.TLabel")
        self.lbl_prog.pack(anchor="w")

        # Painel principal
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=16, pady=10)

        # ── Esquerdo: lista ────────────────────────────────────────────────────
        fl = ttk.Frame(paned)
        paned.add(fl, weight=3)

        tb = ttk.Frame(fl)
        tb.pack(fill="x", pady=(0, 6))
        self.var_todos = tk.BooleanVar(value=True)
        ttk.Checkbutton(tb, text="Todos", variable=self.var_todos,
                        command=self._toggle_todos).pack(side="left")
        self.lbl_count = ttk.Label(tb, text="", style="Muted.TLabel")
        self.lbl_count.pack(side="left", padx=(10, 0))
        ttk.Button(tb, text="Renomear selecionados",
                   command=self._renomear_selecionados).pack(side="right")
        ttk.Button(tb, text="Resetar nomes", style="Ghost.TButton",
                   command=self._resetar_todos).pack(side="right", padx=(0, 6))
        ttk.Button(tb, text="Abrir PDF", style="Ghost.TButton",
                   command=self._abrir_pdf).pack(side="right", padx=(0, 6))

        cols = ("sel", "original", "novo_nome", "status")
        self.tree = ttk.Treeview(fl, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("sel",       text="")
        self.tree.heading("original",  text="ARQUIVO ORIGINAL")
        self.tree.heading("novo_nome", text="NOVO NOME")
        self.tree.heading("status",    text="STATUS")
        self.tree.column("sel",      width=30, stretch=False, anchor="center")
        self.tree.column("original", width=180)
        self.tree.column("novo_nome",width=240)
        self.tree.column("status",   width=70, stretch=False, anchor="center")
        for tag, color in [("ok", GREEN), ("erro", RED), ("ignorado", MUTED), ("pendente", TEXT)]:
            self.tree.tag_configure(tag, foreground=color)

        sbv = ttk.Scrollbar(fl, orient="vertical",   command=self.tree.yview)
        sbh = ttk.Scrollbar(fl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sbv.set, xscrollcommand=sbh.set)
        sbh.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        sbv.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Button-1>",         self._on_click)
        self.tree.bind("<space>",            self._on_space)
        self.tree.bind("<Double-1>",         self._on_double_click)

        # ── Direito: preview + campos ──────────────────────────────────────────
        fd = ttk.Frame(paned, padding=(12, 0, 0, 0))
        paned.add(fd, weight=2)

        ttk.Label(fd, text="PREVIEW", style="Cap.TLabel").pack(anchor="w")
        self.canvas = tk.Canvas(fd, bg=SURFACE, bd=0, highlightthickness=0,
                                width=350, height=340)
        self.canvas.pack(fill="both", expand=True, pady=(4, 10))
        self._sem_preview()

        ttk.Separator(fd).pack(fill="x", pady=(0, 8))

        # DATA + PALESTRA na mesma linha
        row1 = ttk.Frame(fd)
        row1.pack(fill="x", pady=(0, 6))
        ttk.Label(row1, text="DATA", style="Cap.TLabel").pack(side="left")
        self.var_data = tk.StringVar()
        ttk.Entry(row1, textvariable=self.var_data, width=11,
                  font=("Courier New", 9)).pack(side="left", padx=(4, 12))
        self.var_data.trace_add("write", lambda *_: self._atualizar_nome_final())

        ttk.Label(row1, text="PALESTRA", style="Cap.TLabel").pack(side="left")
        self.var_palestra = tk.StringVar()
        ttk.Entry(row1, textvariable=self.var_palestra,
                  font=("Courier New", 9)).pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.var_palestra.trace_add("write", lambda *_: self._atualizar_nome_final())

        # EMISSOR com autocomplete + botão "Aplicar a todos"
        ttk.Label(fd, text="EMISSOR", style="Cap.TLabel").pack(anchor="w", pady=(0, 3))
        row2 = ttk.Frame(fd)
        row2.pack(fill="x", pady=(0, 6))
        self.var_emissor = tk.StringVar()
        self.combo_emissor = ttk.Combobox(
            row2, textvariable=self.var_emissor,
            values=self.emissores.lista, font=("Segoe UI", 9), width=30
        )
        self.combo_emissor.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.combo_emissor.bind("<KeyRelease>",         self._on_emissor_key)
        self.combo_emissor.bind("<<ComboboxSelected>>", self._on_emissor_selected)
        ttk.Button(row2, text="Aplicar a todos", style="Ghost.TButton",
                   command=self._aplicar_emissor_todos).pack(side="left")

        # NOME FINAL (preview em tempo real)
        ttk.Label(fd, text="NOME FINAL", style="Cap.TLabel").pack(anchor="w", pady=(2, 3))
        frame_nome = tk.Frame(fd, bg=SURFACE, padx=8, pady=6)
        frame_nome.pack(fill="x", pady=(0, 8))
        self.var_nome_final = tk.StringVar(value="—")
        tk.Label(frame_nome, textvariable=self.var_nome_final,
                 bg=SURFACE, fg=ACCENT, font=("Courier New", 9),
                 anchor="w", wraplength=340, justify="left").pack(fill="x")

        # Botões de ação
        bf = ttk.Frame(fd)
        bf.pack(fill="x")
        ttk.Button(bf, text="Renomear este",
                   command=self._renomear_um).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="Resetar", style="Ghost.TButton",
                   command=self._resetar_um).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="Ignorar", style="Danger.TButton",
                   command=self._ignorar_um).pack(side="left")

        self._toast_job = None

        # Status bar
        ttk.Separator(self).pack(fill="x", side="bottom")
        self.lbl_status = ttk.Label(self, text="Pronto.", style="Muted.TLabel", padding=(16, 5))
        self.lbl_status.pack(fill="x", side="bottom")

        # Toast — overlay fixo no rodapé, não afeta o layout
        self.frame_toast = tk.Frame(self, bg="#162616", pady=7)
        self.lbl_toast = tk.Label(self.frame_toast, text="", bg="#162616",
                                  fg=GREEN, font=("Segoe UI", 10, "bold"))
        self.lbl_toast.pack(expand=True)

    # ── OneDrive ───────────────────────────────────────────────────────────────

    # ── Configurações ──────────────────────────────────────────────────────────

    def _restaurar_config(self):
        geo = self.config.get("geometry")
        if geo:
            try:
                self.geometry(geo)
            except Exception:
                pass

        modo = self.config.get("modo")
        if modo in ("lote", "individual"):
            self.var_modo.set(modo)

        pasta_salva = self.config.get("ultima_pasta")
        if pasta_salva and Path(pasta_salva).is_dir():
            self.var_pasta.set(pasta_salva)
            # Sincroniza o combo de OneDrive se a pasta salva for raiz de um drive
            for d in self._drives:
                if Path(pasta_salva) == d["caminho"]:
                    self.combo_drive.set(d["label"])
                    break

    def _ao_fechar(self):
        self.config.set("geometry", self.geometry())
        self.config.set("modo", self.var_modo.get())
        pasta = self.var_pasta.get().strip()
        if pasta:
            self.config.set("ultima_pasta", pasta)
        self.destroy()

    # ── OneDrive ───────────────────────────────────────────────────────────────

    def _detectar_onedrives(self):
        self._drives = detectar_onedrives()
        self.combo_drive["values"] = [d["label"] for d in self._drives]
        if self._drives:
            self.combo_drive.set(self._drives[0]["label"])
            self.var_pasta.set(str(self._drives[0]["caminho"]))

    def _on_drive_selected(self, _=None):
        idx = self.combo_drive.current()
        if 0 <= idx < len(self._drives):
            self.var_pasta.set(str(self._drives[idx]["caminho"]))

    def _escolher_pasta(self):
        p = filedialog.askdirectory(initialdir=self.var_pasta.get() or str(Path.home()))
        if p:
            self.var_pasta.set(p)

    # ── Análise ────────────────────────────────────────────────────────────────

    def _analisar(self):
        api_key = ANTHROPIC_API_KEY or os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            messagebox.showerror("Erro", "Defina a variável ANTHROPIC_API_KEY.")
            return
        pasta = Path(self.var_pasta.get())
        if not pasta.is_dir():
            messagebox.showerror("Erro", f"Pasta não encontrada:\n{pasta}")
            return

        if self.var_modo.get() == "individual":
            arqs = filedialog.askopenfilenames(
                initialdir=str(pasta),
                filetypes=[("PDF", "*.pdf *.PDF")],
                title="Selecione os PDFs para analisar"
            )
            pdfs = [Path(a) for a in arqs]
        else:
            seen: set[str] = set()
            pdfs = []
            for p in sorted(list(pasta.glob("*.pdf")) + list(pasta.glob("*.PDF"))):
                if p.name.lower() not in seen:
                    seen.add(p.name.lower())
                    pdfs.append(p)

        if not pdfs:
            messagebox.showinfo("Aviso", "Nenhum PDF encontrado.")
            return

        self.resultados.clear()
        self.tree.delete(*self.tree.get_children())
        self._idx = None
        self._sem_preview()
        self.var_data.set(""); self.var_palestra.set("")
        self.var_emissor.set(""); self.var_nome_final.set("—")
        self.client = anthropic.Anthropic(api_key=api_key)
        self.btn_analisar.configure(state="disabled")
        self.frame_prog.pack(fill="x", padx=16)
        self.progress.configure(maximum=len(pdfs), value=0)

        def worker():
            erros = []
            for i, pdf in enumerate(pdfs):
                self.after(0, self.lbl_prog.configure,
                           {"text": f"Analisando {i+1}/{len(pdfs)}: {pdf.name}"})

                # Arquivo já segue o padrão — não precisa renomear
                if ja_formatado(pdf.stem):
                    item = {
                        "original": pdf.name, "caminho": str(pdf),
                        "data": pdf.stem[:8], "palestra": pdf.stem[11:] if len(pdf.stem) > 11 else "",
                        "emissor": "", "nome_sugerido": pdf.stem,
                        "novo_nome": pdf.stem, "status": "ok", "selecionado": False,
                    }
                    self.resultados.append(item)
                    self.after(0, self._add_row, len(self.resultados) - 1)
                    self.after(0, self.progress.configure, {"value": i + 1})
                    continue

                erro_acesso = verificar_arquivo(pdf)
                if erro_acesso:
                    item = {
                        "original": pdf.name, "caminho": str(pdf),
                        "data": "", "palestra": erro_acesso,
                        "emissor": "", "nome_sugerido": pdf.stem,
                        "novo_nome": pdf.stem, "status": "erro", "selecionado": False,
                    }
                    self.resultados.append(item)
                    self.after(0, self._add_row, len(self.resultados) - 1)
                    self.after(0, self.progress.configure, {"value": i + 1})
                    continue

                texto = extrair_texto(pdf)
                usou_visao = False
                if len(texto.strip()) < 80 or not texto_tem_qualidade(texto):
                    # PDF image-only ou texto corrompido — usa Claude Vision
                    campos = extrair_campos_via_visao(pdf, self.client)
                    usou_visao = True
                else:
                    campos = extrair_campos(texto, self.client)

                # Fallback: se o modo texto não extraiu nada útil, tenta Vision
                if not usou_visao and not campos.get("data") and not campos.get("palestra"):
                    campos = extrair_campos_via_visao(pdf, self.client)
                if campos.get("_erro"):
                    erros.append(f"{pdf.name}:\n{campos['_erro']}")
                data     = campos.get("data", "")
                palestra = campos.get("palestra", "").strip().title()
                nao_extraiu = not data and not palestra
                item = {
                    "original":      pdf.name,
                    "caminho":       str(pdf),
                    "data":          data,
                    "palestra":      palestra if not nao_extraiu else "⚠ preencha manualmente",
                    "emissor":       "",
                    "nome_sugerido": montar_nome(data, palestra),
                    "novo_nome":     montar_nome(data, palestra),
                    "status":        "pendente",
                    "selecionado":   not nao_extraiu,
                }
                self.resultados.append(item)
                self.after(0, self._add_row, len(self.resultados) - 1)
                self.after(0, self.progress.configure, {"value": i + 1})
            self.after(0, self._analise_ok, len(pdfs), erros)

        threading.Thread(target=worker, daemon=True).start()

    def _add_row(self, idx: int):
        r = self.resultados[idx]
        self.tree.insert("", "end", iid=str(idx), values=(
            "☑" if r["selecionado"] else "☐",
            r["original"], r["novo_nome"],
            self._icone(r["status"])
        ), tags=(r["status"],))
        self._atualizar_count()

    def _analise_ok(self, total: int, erros: list[str]):
        self.frame_prog.pack_forget()
        self.btn_analisar.configure(state="normal")
        self.lbl_status.configure(text=f"Análise concluída — {total} arquivo(s). Selecione um arquivo e preencha o emissor.")
        if erros:
            messagebox.showerror("Erro na análise",
                                 f"{len(erros)} arquivo(s) com erro:\n\n" + "\n\n".join(erros[:3]))

    # ── Seleção e preview ──────────────────────────────────────────────────────

    def _on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        self._idx = int(sel[0])
        r = self.resultados[self._idx]
        self.var_data.set(r["data"])
        self.var_palestra.set(r["palestra"])
        self.var_emissor.set(r["emissor"])
        self._resetar_combo_emissor()
        self._atualizar_nome_final()
        token = self._proximo_token()
        threading.Thread(target=self._carregar_preview,
                         args=(Path(r["caminho"]), token), daemon=True).start()

    def _proximo_token(self) -> int:
        with self._preview_lock:
            self._preview_token += 1
            return self._preview_token

    def _carregar_preview(self, caminho: Path, token: int):
        erro = verificar_arquivo(caminho)
        if erro:
            with self._preview_lock:
                if token != self._preview_token:
                    return
            self.after(0, self._sem_preview, "☁ Arquivo offline\nFaça o download no OneDrive")
            return
        w = max(self.canvas.winfo_width(), 350)
        foto = renderizar_pdf(caminho, largura=w)
        with self._preview_lock:
            if token != self._preview_token:
                return
        self.after(0, self._mostrar_preview, foto)

    def _mostrar_preview(self, foto):
        self.canvas.delete("all")
        if foto:
            self._foto_ref = foto
            cx = max(self.canvas.winfo_width(), 350) // 2
            cy = max(self.canvas.winfo_height(), 340) // 2
            self.canvas.create_image(cx, cy, image=foto, anchor="center")
        else:
            self._sem_preview("Não foi possível renderizar")

    def _sem_preview(self, msg: str = "Selecione um arquivo\npara visualizar"):
        self.canvas.delete("all")
        self._foto_ref = None
        cx = max(self.canvas.winfo_width(), 350) // 2
        cy = max(self.canvas.winfo_height(), 340) // 2
        self.canvas.create_text(cx, cy, text=msg, fill=MUTED,
                                font=("Segoe UI", 10), justify="center")

    def _on_click(self, event):
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if col == "#1" and iid:
            r = self.resultados[int(iid)]
            if r["status"] != "ok":
                r["selecionado"] = not r["selecionado"]
                self._refresh_row(int(iid))

    def _on_space(self, _):
        if self._idx is not None:
            r = self.resultados[self._idx]
            if r["status"] != "ok":
                r["selecionado"] = not r["selecionado"]
                self._refresh_row(self._idx)

    def _on_double_click(self, event):
        col = self.tree.identify_column(event.x)
        if col == "#1":  # coluna de checkbox — não abre
            return
        iid = self.tree.identify_row(event.y)
        if iid:
            self._abrir_pdf(int(iid))

    def _abrir_pdf(self, idx: int | None = None):
        i = idx if idx is not None else self._idx
        if i is None:
            return
        caminho = Path(self.resultados[i]["caminho"])
        if not caminho.exists():
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{caminho}")
            return
        os.startfile(str(caminho))

    # ── Emissor ────────────────────────────────────────────────────────────────

    def _on_emissor_key(self, event):
        if event.keysym in ("Tab", "Return", "KP_Enter", "Escape", "Up", "Down"):
            self._atualizar_nome_final()
            return
        typed = self.var_emissor.get()
        self.combo_emissor["values"] = (
            self.emissores.filtrar(typed) if typed else self.emissores.lista
        )
        self._atualizar_nome_final()
        if self._idx is not None:
            self.resultados[self._idx]["emissor"] = typed

    def _on_emissor_selected(self, _=None):
        emissor = self.var_emissor.get()
        self._atualizar_nome_final()
        if self._idx is not None:
            self.resultados[self._idx]["emissor"] = emissor

    def _resetar_combo_emissor(self):
        self.combo_emissor["values"] = self.emissores.lista

    def _aplicar_emissor_todos(self):
        emissor = self.var_emissor.get().strip()
        if not emissor:
            messagebox.showwarning("Aviso", "Digite um emissor antes de aplicar a todos.")
            return
        for i, r in enumerate(self.resultados):
            if r["status"] == "pendente":
                r["emissor"] = emissor
                r["novo_nome"] = montar_nome(r["data"], r["palestra"], emissor)
                self._refresh_row(i)
        self.lbl_status.configure(
            text=f"Emissor '{emissor}' aplicado a todos os arquivos pendentes.")

    def _atualizar_nome_final(self):
        data     = self.var_data.get().strip()
        palestra = self.var_palestra.get().strip()
        emissor  = self.var_emissor.get().strip()
        nome = montar_nome(data, palestra, emissor)
        self.var_nome_final.set(nome + ".pdf")
        if self._idx is not None:
            self.resultados[self._idx]["novo_nome"] = nome
            self._refresh_row(self._idx)

    # ── Ações ──────────────────────────────────────────────────────────────────

    def _renomear_um(self):
        if self._idx is None:
            return
        data     = self.var_data.get().strip()
        palestra = self.var_palestra.get().strip()
        emissor  = self.var_emissor.get().strip()
        nome = montar_nome(data, palestra, emissor)
        if not nome or nome == "sem_nome":
            messagebox.showwarning("Aviso", "Preencha ao menos a data ou o nome da palestra.")
            return
        r = self.resultados[self._idx]
        r.update({"data": data, "palestra": palestra, "emissor": emissor, "novo_nome": nome})
        self._executar_rename(self._idx)
        if emissor and self.emissores.adicionar(emissor):
            self._resetar_combo_emissor()

    def _renomear_selecionados(self):
        pendentes = [i for i, r in enumerate(self.resultados)
                     if r["selecionado"] and r["status"] == "pendente"]
        if not pendentes:
            messagebox.showinfo("Aviso", "Nenhum arquivo pendente selecionado.")
            return
        for i in pendentes:
            self._executar_rename(i)
        ok = sum(1 for r in self.resultados if r["status"] == "ok")
        self.lbl_status.configure(text=f"{ok} arquivo(s) renomeado(s) com sucesso.")

    def _executar_rename(self, idx: int):
        r    = self.resultados[idx]
        orig = Path(r["caminho"])
        nome = r["novo_nome"].strip()
        if not nome.lower().endswith(".pdf"):
            nome += ".pdf"
        dest = orig.parent / nome
        try:
            orig.rename(dest)
            r.update({"caminho": str(dest), "original": nome,
                       "novo_nome": nome, "status": "ok"})
            self._mostrar_toast(f"✓  {nome}")
        except Exception as e:
            r["status"] = "erro"
            self._mostrar_toast(f"✗  Erro: {e}", cor=RED)
        self._refresh_row(idx)
        self._atualizar_count()

    def _mostrar_toast(self, msg: str, cor: str = GREEN):
        if self._toast_job:
            self.after_cancel(self._toast_job)
        self.lbl_toast.configure(text=msg, fg=cor,
                                 bg="#162616" if cor == GREEN else "#261616")
        self.frame_toast.configure(bg="#162616" if cor == GREEN else "#261616")
        # Posiciona como overlay no rodapé, sem alterar o layout do pack
        self.frame_toast.place(relx=0, rely=1.0, relwidth=1.0, anchor="sw", y=-34)
        self._toast_job = self.after(3000, self._ocultar_toast)

    def _ocultar_toast(self):
        self.frame_toast.place_forget()
        self._toast_job = None

    def _ignorar_um(self):
        if self._idx is None:
            return
        r = self.resultados[self._idx]
        if r["status"] == "ok":
            return
        r["status"] = "ignorado" if r["status"] != "ignorado" else "pendente"
        r["selecionado"] = r["status"] == "pendente"
        self._refresh_row(self._idx)

    def _resetar_um(self):
        if self._idx is None:
            return
        r = self.resultados[self._idx]
        r.update({"emissor": "", "novo_nome": r["nome_sugerido"]})
        self.var_data.set(r["data"])
        self.var_palestra.set(r["palestra"])
        self.var_emissor.set("")
        self._atualizar_nome_final()

    def _toggle_todos(self):
        sel = self.var_todos.get()
        for i, r in enumerate(self.resultados):
            if r["status"] != "ok":
                r["selecionado"] = sel
                self._refresh_row(i)

    def _resetar_todos(self):
        for i, r in enumerate(self.resultados):
            r.update({"emissor": "", "novo_nome": r["nome_sugerido"]})
            self._refresh_row(i)
        if self._idx is not None:
            self.var_emissor.set("")
            self._atualizar_nome_final()

    # ── Utilitários ────────────────────────────────────────────────────────────

    @staticmethod
    def _icone(status: str) -> str:
        return {"pendente": "⏳", "ok": "✓", "erro": "✗", "ignorado": "—"}.get(status, status)

    def _refresh_row(self, idx: int):
        r = self.resultados[idx]
        self.tree.item(str(idx), values=(
            "☑" if r["selecionado"] else "☐",
            r["original"], r["novo_nome"],
            self._icone(r["status"])
        ), tags=(r["status"],))

    def _atualizar_count(self):
        total = len(self.resultados)
        ok    = sum(1 for r in self.resultados if r["status"] == "ok")
        pend  = sum(1 for r in self.resultados if r["status"] == "pendente")
        self.lbl_count.configure(
            text=f"{total} arquivo(s)  ·  {ok} ✓  ·  {pend} pendente(s)")
        self.lbl_status.configure(
            text=f"{total} arquivo(s) · {ok} renomeado(s) · {pend} pendente(s)")


# ── Entrada ────────────────────────────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
