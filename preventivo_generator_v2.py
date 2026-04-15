#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
preventivo_generator_v2.py
Generatore Preventivi PDF - Cessione del Quinto / Delega
QuintoQuote v1 — Generatore Preventivi Open Source

Modalità:
- CLI guidata (default): python quintoquote.py
- CLI con argomenti:      python quintoquote.py --cliente "Mario Rossi" ...
- Web localhost:          python quintoquote.py --web
"""

from __future__ import annotations

import argparse
from calendar import monthrange
import html
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import uuid
import webbrowser
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

try:
    import fitz
except Exception:
    fitz = None

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    Flowable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# =========================
#  UTILITA' FORMAT / DATE
# =========================

def parse_date_it(s: str) -> date:
    """Parsa 'GG/MM/AAAA' -> date; solleva ValueError se non valida."""
    s = s.strip()
    m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", s)
    if not m:
        raise ValueError("Formato data non valido. Usa GG/MM/AAAA (es: 15/05/1980).")
    gg, mm, aaaa = map(int, m.groups())
    return date(aaaa, mm, gg)

def calc_eta(born: date, today: Optional[date] = None) -> int:
    today = today or date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))

def euro(amount: float) -> str:
    """EUR con formato italiano (punti migliaia, virgola decimali)."""
    return f"EUR {amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def sanitize_filename(s: str) -> str:
    s = s.strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r'[\\/:*?"<>|]+', "", s)
    s = re.sub(r"_{2,}", "_", s)
    return s or "Cliente"

def ask(prompt: str, default: Optional[str] = None) -> str:
    if default is not None and default != "":
        full = f"{prompt} [{default}]: "
    else:
        full = f"{prompt}: "
    v = input(full).strip()
    return v if v else (default or "")

def ask_float(prompt: str, default: Optional[float] = None, min_value: Optional[float] = 0.0) -> float:
    while True:
        d = f"{default:.2f}" if default is not None else None
        raw = ask(prompt, d)
        raw = raw.replace("EUR", "").replace("€", "").strip()
        raw = raw.replace(".", "").replace(",", ".")  # consenti "1.234,56"
        try:
            v = float(raw)
            if min_value is not None and v < min_value:
                print(f"Valore troppo basso (min {min_value}). Riprova.")
                continue
            return v
        except ValueError:
            print("Numero non valido. Esempio: 350,00 oppure 350.00")

def ask_int(prompt: str, default: Optional[int] = None, min_value: Optional[int] = 1) -> int:
    while True:
        raw = ask(prompt, str(default) if default is not None else None)
        try:
            v = int(raw)
            if min_value is not None and v < min_value:
                print(f"Valore troppo basso (min {min_value}). Riprova.")
                continue
            return v
        except ValueError:
            print("Intero non valido. Esempio: 120")

def ask_choice(prompt: str, choices: list[str], default: Optional[str] = None) -> str:
    choices_lower = [c.lower() for c in choices]
    default = default or choices[0]
    while True:
        raw = ask(f"{prompt} ({'/'.join(choices)})", default).strip()
        if raw.lower() in choices_lower:
            return choices[choices_lower.index(raw.lower())]
        print(f"Scelta non valida. Opzioni: {', '.join(choices)}")

# =========================
#  MODELLO DATI
# =========================

@dataclass
class Preventivo:
    cliente: str
    data_nascita: str  # GG/MM/AAAA
    tipo_lavoro: str
    provincia: str
    note: str
    tipo_finanziamento: str  # "Cessione del Quinto" | "Delega di Pagamento"
    importo_rata: float
    durata_mesi: int
    tan: float
    taeg: float
    importo_erogato: float

    # Derivati (calcolati)
    eta: int = 0
    montante: float = 0.0
    interessi: float = 0.0
    data_preventivo: str = ""

    def compute(self) -> None:
        born = parse_date_it(self.data_nascita)
        self.eta = calc_eta(born)
        self.montante = round(self.importo_rata * self.durata_mesi, 2)
        self.interessi = round(self.montante - self.importo_erogato, 2)
        self.data_preventivo = datetime.now().strftime("%d/%m/%Y")

    def validate(self) -> None:
        if not self.cliente.strip():
            raise ValueError("Cliente vuoto.")
        if self.durata_mesi <= 0:
            raise ValueError("Durata mesi non valida.")
        if self.importo_rata <= 0:
            raise ValueError("Rata non valida.")
        if self.importo_erogato <= 0:
            raise ValueError("Importo erogato non valido.")
        if not (0 <= self.tan < 50) or not (0 <= self.taeg < 50):
            raise ValueError("TAN/TAEG fuori range plausibile.")
        # Attenzione: interessi può anche essere basso, ma negativo di solito è errore
        if self.interessi < 0:
            raise ValueError("Interessi negativi: controlla rata/durata/netto erogato.")

# =========================
#  BRANDING PROFILE
# =========================

@dataclass
class BrandingProfile:
    nome_agente: str = ""
    rete_mandante: str = ""
    codice_oam: str = ""
    telefono: str = ""
    logo_path: str = ""
    bollino_path: str = ""
    colore_primario: str = "#0a1628"
    colore_accento: str = "#c9a227"

    @property
    def is_configured(self) -> bool:
        return bool(self.nome_agente.strip())


def _is_frozen_runtime() -> bool:
    return bool(getattr(sys, "frozen", False))


def _get_user_data_root() -> Path:
    base = Path(os.getenv("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
    return base / "QuintoQuote"


def _default_config_path() -> Path:
    if _is_frozen_runtime():
        return _get_user_data_root() / "config.json"
    return Path(__file__).parent / "config.json"


def _default_assets_dir() -> Path:
    if _is_frozen_runtime():
        return _get_user_data_root() / "assets"
    return Path(__file__).parent / "assets"


def _default_output_dir() -> Path:
    if _is_frozen_runtime():
        return _get_user_data_root() / "output_preventivi"
    return Path("output_preventivi")


def _default_runtime_temp_dir() -> Path:
    if _is_frozen_runtime():
        return _get_user_data_root() / ".quintoquote_tmp"
    return Path(__file__).resolve().parent / ".quintoquote_tmp"


def _resolve_pdf_template_dir() -> Path:
    candidates: list[Path] = []
    for root_candidate in (
        getattr(sys, "_MEIPASS", ""),
        Path(sys.executable).resolve().parent if _is_frozen_runtime() else None,
        Path(__file__).parent,
    ):
        if not root_candidate:
            continue
        root = Path(root_candidate)
        candidates.append(root / "docs")
    candidates.append(Path(__file__).parent / "docs")
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return Path(__file__).parent / "docs"


_CONFIG_PATH = _default_config_path()
_ASSETS_DIR = _default_assets_dir()
_config_lock = threading.Lock()
_cached_profile: Optional[BrandingProfile] = None
_cached_profile_path: Optional[Path] = None
_DEFAULT_PRIMARY_COLOR = "#0a1628"
_DEFAULT_ACCENT_COLOR = "#c9a227"
_HEX_COLOR_RE = re.compile(r"^#[0-9a-fA-F]{6}$")
_ALLOWED_IMAGE_EXTS = {"jpg", "jpeg", "png"}
_ALLOWED_DOWNLOAD_EXTS = {"pdf"}
_EDITABLE_FIELDS = {"cliente_nome", "note", "disclaimer", "closing"}


def _ensure_assets_dir() -> Path:
    _ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    return _ASSETS_DIR


def configure_runtime_paths(config_path: Optional[Path] = None, assets_dir: Optional[Path] = None) -> None:
    global _CONFIG_PATH, _ASSETS_DIR, _cached_profile, _cached_profile_path
    if config_path is not None:
        _CONFIG_PATH = config_path.expanduser().resolve()
    if assets_dir is not None:
        _ASSETS_DIR = assets_dir.expanduser().resolve()
    _cached_profile = None
    _cached_profile_path = None


def normalize_cli_argv(argv: list[str]) -> tuple[list[str], bool]:
    if len(argv) > 1 and argv[1].strip().lower() == "start":
        return [argv[0], "--web", *argv[2:]], True
    if _is_frozen_runtime() and len(argv) == 1:
        return [argv[0], "--web"], True
    return argv, False


def _pick_available_port(host: str, preferred_port: int, max_tries: int = 20) -> int:
    host = host or "127.0.0.1"
    for port in range(preferred_port, preferred_port + max_tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind((host, port))
            except OSError:
                continue
            return port
    return preferred_port


def sanitize_hex_color(raw: str, fallback: str) -> str:
    candidate = (raw or "").strip()
    if _HEX_COLOR_RE.fullmatch(candidate):
        return candidate.lower()
    return fallback


def sanitize_asset_filename(raw: str) -> str:
    if not raw:
        return ""
    name = Path(str(raw)).name
    if "." not in name:
        return ""
    stem, ext = name.rsplit(".", 1)
    ext = ext.lower()
    if ext not in _ALLOWED_IMAGE_EXTS:
        return ""
    safe_stem = sanitize_filename(stem)
    return f"{safe_stem}.{ext}" if safe_stem else ""


def normalize_profile(profile: BrandingProfile) -> BrandingProfile:
    profile.nome_agente = (profile.nome_agente or "").strip()
    profile.rete_mandante = (profile.rete_mandante or "").strip()
    profile.codice_oam = (profile.codice_oam or "").strip()
    profile.telefono = (profile.telefono or "").strip()
    profile.logo_path = sanitize_asset_filename(profile.logo_path)
    profile.bollino_path = sanitize_asset_filename(profile.bollino_path)
    profile.colore_primario = sanitize_hex_color(profile.colore_primario, _DEFAULT_PRIMARY_COLOR)
    profile.colore_accento = sanitize_hex_color(profile.colore_accento, _DEFAULT_ACCENT_COLOR)
    return profile


def sanitize_text_overrides(raw: object) -> dict[str, str]:
    if not isinstance(raw, dict):
        return {}
    clean: dict[str, str] = {}
    for field in _EDITABLE_FIELDS:
        value = raw.get(field)
        if not isinstance(value, str):
            continue
        normalized = value.replace("\r\n", "\n").replace("\r", "\n").strip()
        if normalized:
            clean[field] = normalized[:2000]
    return clean


def escape_preview_text(value: object, preserve_line_breaks: bool = False) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    escaped = html.escape(text, quote=True)
    if preserve_line_breaks:
        return escaped.replace("\n", "<br/>")
    return escaped


def resolve_child_file(base_dir: Path, filename: str, allowed_exts: set[str]) -> Optional[Path]:
    if not filename or Path(filename).name != filename:
        return None
    candidate = Path(filename)
    ext = candidate.suffix.lower().lstrip(".")
    if ext not in allowed_exts:
        return None
    base_resolved = base_dir.resolve()
    resolved = (base_dir / candidate.name).resolve()
    try:
        resolved.relative_to(base_resolved)
    except ValueError:
        return None
    return resolved


def load_profile(path: Optional[Path] = None) -> BrandingProfile:
    global _cached_profile, _cached_profile_path
    cfg = (path or _CONFIG_PATH).expanduser().resolve()
    with _config_lock:
        try:
            data = json.loads(cfg.read_text(encoding="utf-8"))
            _cached_profile = BrandingProfile(**{
                k: v for k, v in data.items()
                if k in BrandingProfile.__dataclass_fields__
            })
        except (FileNotFoundError, json.JSONDecodeError, TypeError):
            _cached_profile = BrandingProfile()
        _cached_profile = normalize_profile(_cached_profile)
        _cached_profile_path = cfg
        return _cached_profile


def save_profile(profile: BrandingProfile, path: Optional[Path] = None) -> None:
    global _cached_profile, _cached_profile_path
    cfg = (path or _CONFIG_PATH).expanduser().resolve()
    with _config_lock:
        profile = normalize_profile(profile)
        cfg.parent.mkdir(parents=True, exist_ok=True)
        tmp = cfg.with_suffix(".tmp")
        tmp.write_text(json.dumps(asdict(profile), indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(str(tmp), str(cfg))
        _cached_profile = profile
        _cached_profile_path = cfg


def get_cached_profile() -> BrandingProfile:
    global _cached_profile, _cached_profile_path
    current_cfg = _CONFIG_PATH.expanduser().resolve()
    if _cached_profile is None or _cached_profile_path != current_cfg:
        return load_profile()
    return _cached_profile


def _lighten_hex(hex_color: str, factor: float = 0.45) -> str:
    """Blend a hex color toward white by *factor* (0=original, 1=white)."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"#{r:02x}{g:02x}{b:02x}"


def get_design_tokens(profile: Optional[BrandingProfile] = None) -> dict:
    p = profile or get_cached_profile()
    return {
        "navy": p.colore_primario,
        "navy_light": _lighten_hex(p.colore_primario, 0.30),
        "gold": p.colore_accento,
        "gold_light": _lighten_hex(p.colore_accento, 0.45),
        "warm_white": "#faf9f7",
        "gray_warm": "#6b7280",
        "line_soft": "#e7e2d7",
    }


# =========================
#  PDF GENERATION (ReportLab)
# =========================

DESIGN_TOKENS = {
    "navy": "#0a1628",
    "navy_light": "#1e3a5f",
    "gold": "#c9a227",
    "gold_light": "#e8d5a3",
    "warm_white": "#faf9f7",
    "gray_warm": "#6b7280",
    "line_soft": "#e7e2d7",
}

DISPLAY_FONT = "Helvetica-Bold"
BODY_FONT = "Helvetica"
MONO_FONT = "Courier"

# _OAM_IMG_PATH rimosso: il path bollino è ora in BrandingProfile.bollino_path


class _ClickableImage(Flowable):
    """Immagine cliccabile: disegna il file e sovrappone un'annotazione URL."""

    def __init__(self, path: Path, width: float, height: float, url: str) -> None:
        super().__init__()
        self.path = path
        self.width = width
        self.height = height
        self.url = url
        self.hAlign = "CENTER"

    def wrap(self, avail_w: float, avail_h: float):
        return self.width, self.height

    def draw(self) -> None:
        self.canv.drawImage(
            str(self.path), 0, 0,
            width=self.width, height=self.height,
            preserveAspectRatio=True, mask="auto",
        )
        self.canv.linkURL(
            self.url,
            (0, 0, self.width, self.height),
            relative=1,
        )


def add_months(d: date, months: int) -> date:
    total = (d.year * 12 + d.month - 1) + months
    year = total // 12
    month = total % 12 + 1
    day = min(d.day, monthrange(year, month)[1])
    return date(year, month, day)


class HeroFinanceFlowable(Flowable):
    def __init__(
        self,
        width: float,
        rata: float,
        durata_mesi: int,
        durata_anni: str,
        taeg: float,
        tan: float,
        tokens: Optional[dict] = None,
    ):
        super().__init__()
        self.width = width
        self.height = 62 * mm
        self.rata = rata
        self.durata_mesi = durata_mesi
        self.durata_anni = durata_anni
        self.taeg = taeg
        self.tan = tan
        self.tokens = tokens or DESIGN_TOKENS

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def _draw_check_icon(self, c, x: float, y: float, r: float) -> None:
        c.saveState()
        c.setFillColor(colors.HexColor(self.tokens["gold"]))
        c.circle(x, y, r, stroke=0, fill=1)
        c.setStrokeColor(colors.HexColor(self.tokens["warm_white"]))
        c.setLineWidth(1.4)
        c.line(x - r * 0.50, y - r * 0.05, x - r * 0.12, y - r * 0.38)
        c.line(x - r * 0.12, y - r * 0.38, x + r * 0.52, y + r * 0.38)
        c.restoreState()

    def draw(self):
        c = self.canv
        x = 0
        y = 0
        h = self.height

        gap = 5 * mm
        left_w = self.width * 0.62
        right_w = self.width - left_w - gap

        navy = colors.HexColor(self.tokens["navy"])
        navy_light = colors.HexColor(self.tokens["navy_light"])
        gold = colors.HexColor(self.tokens["gold"])

        # Left rate card shadow + body
        c.saveState()
        if hasattr(c, "setFillAlpha"):
            c.setFillAlpha(0.22)
        c.setFillColor(colors.HexColor("#1b2f49"))
        c.roundRect(x + 1.5, y - 1.5, left_w, h, 9, stroke=0, fill=1)
        if hasattr(c, "setFillAlpha"):
            c.setFillAlpha(1)

        c.setFillColor(navy)
        c.roundRect(x, y, left_w, h, 9, stroke=0, fill=1)

        # subtle overlay to emulate gradient
        if hasattr(c, "setFillAlpha"):
            c.setFillAlpha(0.18)
        c.setFillColor(navy_light)
        c.roundRect(x + 2, y + h * 0.42, left_w - 4, h * 0.52, 8, stroke=0, fill=1)
        if hasattr(c, "setFillAlpha"):
            c.setFillAlpha(1)

        label_y = y + h - 10 * mm
        c.setFillColor(colors.HexColor("#d8dee8"))
        c.setFont(BODY_FONT, 8)
        c.drawString(x + 6 * mm, label_y, "RATA MENSILE")

        rate_str = f"{self.rata:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        integer_part, decimal_part = rate_str.split(",")

        base_y = y + h * 0.40
        c.setFillColor(colors.white)
        c.setFont(DISPLAY_FONT, 43)
        c.drawString(x + 6 * mm, base_y, f"{integer_part},")
        int_width = c.stringWidth(f"{integer_part},", DISPLAY_FONT, 43)

        c.setFont(DISPLAY_FONT, 26)
        c.drawString(x + 6 * mm + int_width + 1.5, base_y + 2, decimal_part)
        dec_width = c.stringWidth(decimal_part, DISPLAY_FONT, 26)

        c.setFillColor(colors.HexColor("#d8dee8"))
        c.setFont(BODY_FONT, 15)
        c.drawString(x + 6 * mm + int_width + dec_width + 8, base_y + 4, "EUR")

        c.setFont(BODY_FONT, 9)
        c.setFillColor(colors.HexColor("#cbd5df"))
        c.drawString(
            x + 6 * mm,
            y + 7.2 * mm,
            f"Per {self.durata_mesi} mesi ({self.durata_anni})",
        )
        c.restoreState()

        # Right side: TAEG badge + TAN quick ref
        right_x = x + left_w + gap
        quick_h = 14 * mm
        badge_h = h - quick_h - 3 * mm

        c.saveState()
        if hasattr(c, "setFillAlpha"):
            c.setFillAlpha(0.26)
        c.setFillColor(gold)
        c.roundRect(right_x + 1.2, y + quick_h + 1.8, right_w, badge_h, 8, stroke=0, fill=1)
        if hasattr(c, "setFillAlpha"):
            c.setFillAlpha(1)

        c.setFillColor(navy)
        c.roundRect(right_x, y + quick_h + 3 * mm, right_w, badge_h, 8, stroke=0, fill=1)
        c.setStrokeColor(gold)
        c.setLineWidth(1.25)
        c.roundRect(right_x, y + quick_h + 3 * mm, right_w, badge_h, 8, stroke=1, fill=0)

        check_cx = right_x + 6 * mm
        check_cy = y + quick_h + badge_h + 1 * mm - 6.4 * mm
        self._draw_check_icon(c, check_cx, check_cy, 2.6 * mm)

        c.setFillColor(colors.HexColor("#d8dee8"))
        c.setFont(BODY_FONT, 7.4)
        c.drawString(right_x + 10 * mm, check_cy - 1.0, "COSTO DEL FINANZIAMENTO")

        taeg_str = f"{self.taeg:.3f}".replace(".", ",")
        c.setFillColor(gold)
        c.setFont(DISPLAY_FONT, 30)
        c.drawString(right_x + 6 * mm, y + quick_h + badge_h * 0.48, taeg_str)
        c.setFont(DISPLAY_FONT, 17)
        taeg_w = c.stringWidth(taeg_str, DISPLAY_FONT, 30)
        c.drawString(right_x + 6 * mm + taeg_w + 2, y + quick_h + badge_h * 0.48 + 2.5, "%")

        c.setFillColor(colors.HexColor("#e3e8ef"))
        c.setFont(MONO_FONT, 7.3)
        c.drawString(right_x + 6 * mm, y + quick_h + badge_h * 0.30, "TAEG (Tasso Annuo Effettivo Globale)")

        c.setStrokeColor(colors.HexColor("#536b8a"))
        c.setLineWidth(0.6)
        c.line(right_x + 6 * mm, y + quick_h + badge_h * 0.23, right_x + right_w - 6 * mm, y + quick_h + badge_h * 0.23)
        c.setFillColor(colors.HexColor("#c8d1de"))
        c.setFont(BODY_FONT, 6.8)
        c.drawString(
            right_x + 6 * mm,
            y + quick_h + badge_h * 0.14,
            "Include tutti i costi: interessi, imposte e spese",
        )

        c.setFillColor(colors.white)
        c.roundRect(right_x, y, right_w, quick_h, 6, stroke=0, fill=1)
        c.setStrokeColor(colors.HexColor("#d9dde4"))
        c.setLineWidth(0.8)
        c.roundRect(right_x, y, right_w, quick_h, 6, stroke=1, fill=0)

        c.setFillColor(colors.HexColor(self.tokens["gray_warm"]))
        c.setFont(BODY_FONT, 8)
        c.drawString(right_x + 4 * mm, y + quick_h - 8.9, "TAN nominale")

        tan_str = f"{self.tan:.3f}%".replace(".", ",")
        c.setFillColor(navy)
        c.setFont(MONO_FONT, 9.4)
        c.drawRightString(right_x + right_w - 4 * mm, y + quick_h - 8.8, tan_str)
        c.restoreState()


class TimelineFlowable(Flowable):
    def __init__(self, width: float, netto_erogato: float, durata_mesi: int, end_date: date, tokens: Optional[dict] = None):
        super().__init__()
        self.width = width
        self.height = 38 * mm
        self.netto_erogato = netto_erogato
        self.durata_mesi = durata_mesi
        self.end_date = end_date
        self.tokens = tokens or DESIGN_TOKENS

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def draw(self):
        c = self.canv
        w = self.width
        h = self.height
        navy = colors.HexColor(self.tokens["navy"])
        gold = colors.HexColor(self.tokens["gold"])
        gray = colors.HexColor("#b1b8c2")

        line_y = h - 12 * mm
        x0 = 8 * mm
        x2 = w - 8 * mm
        x1 = (x0 + x2) / 2

        c.saveState()
        c.setStrokeColor(navy)
        c.setLineWidth(1.1)
        c.line(x0, line_y, x2, line_y)

        c.setStrokeColor(gold)
        c.setLineWidth(1.6)
        c.line(x0 + (x2 - x0) * 0.28, line_y, x0 + (x2 - x0) * 0.72, line_y)

        dot_r = 2.4 * mm
        c.setFillColor(gold)
        c.circle(x0, line_y, dot_r, stroke=0, fill=1)
        c.setFillColor(gray)
        c.circle(x1, line_y, dot_r, stroke=0, fill=1)
        c.setFillColor(gold)
        c.circle(x2, line_y, dot_r, stroke=0, fill=1)

        c.setFillColor(navy)
        c.setFont(DISPLAY_FONT, 9)
        c.drawCentredString(x0, line_y - 12, "Oggi")
        c.drawCentredString(x2, line_y - 12, str(self.end_date.year))

        half_rates = max(1, self.durata_mesi // 2)
        c.setFillColor(colors.HexColor(self.tokens["gray_warm"]))
        c.setFont(DISPLAY_FONT, 8.2)
        c.drawCentredString(x1, line_y - 12, f"Mese {half_rates}")

        c.setFillColor(colors.HexColor(self.tokens["gray_warm"]))
        c.setFont(MONO_FONT, 6.8)
        c.drawCentredString(x0, line_y - 20, "Erogazione")
        c.drawCentredString(x1, line_y - 20, "Met\u00e0 percorso")
        c.drawCentredString(x2, line_y - 20, "Fine percorso")

        c.setFont(MONO_FONT, 6.8)
        c.setFillColor(gold)
        c.drawCentredString(x0, line_y - 28, euro(self.netto_erogato).replace("EUR ", "EUR "))

        c.setFillColor(colors.HexColor(self.tokens["gray_warm"]))
        c.drawCentredString(x1, line_y - 28, f"{half_rates} rate pagate")

        c.setFillColor(colors.HexColor("#1f9d67"))
        c.drawCentredString(x2, line_y - 28, "Obiettivo raggiunto")
        c.restoreState()


# =========================
#  PDF GENERATION (ReportLab)
# =========================

def build_styles(tokens: Optional[dict] = None):
    """Design system styles (display/body/mono) for premium PDF layout."""
    _t = tokens or DESIGN_TOKENS
    navy = colors.HexColor(_t["navy"])
    gold = colors.HexColor(_t["gold"])

    return {
        "title": ParagraphStyle(
            "Title",
            fontSize=26,
            textColor=navy,
            fontName=DISPLAY_FONT,
            alignment=TA_CENTER,
            leading=30,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            fontSize=8,
            textColor=colors.HexColor("#8d96a4"),
            fontName=BODY_FONT,
            alignment=TA_CENTER,
            leading=10,
        ),
        "section": ParagraphStyle(
            "Section",
            fontSize=8.2,
            textColor=colors.white,
            fontName=DISPLAY_FONT,
            leading=10,
        ),
        "field_label": ParagraphStyle(
            "FieldLabel",
            fontSize=7,
            textColor=colors.HexColor("#7f8894"),
            fontName=BODY_FONT,
            leading=9,
        ),
        "field_value": ParagraphStyle(
            "FieldValue",
            fontSize=10,
            textColor=navy,
            fontName=DISPLAY_FONT,
            leading=12,
        ),
        "kpi_label": ParagraphStyle(
            "KpiLabel",
            fontSize=7.2,
            textColor=colors.HexColor("#818998"),
            fontName=BODY_FONT,
            alignment=TA_CENTER,
            leading=9,
        ),
        "kpi_value": ParagraphStyle(
            "KpiValue",
            fontSize=12,
            textColor=navy,
            fontName=DISPLAY_FONT,
            alignment=TA_CENTER,
            leading=14,
        ),
        "table_label": ParagraphStyle(
            "TableLabel",
            fontSize=8.5,
            textColor=colors.HexColor("#4f5a67"),
            fontName=BODY_FONT,
            leading=11,
        ),
        "table_value": ParagraphStyle(
            "TableValue",
            fontSize=8.8,
            textColor=navy,
            fontName=DISPLAY_FONT,
            alignment=TA_RIGHT,
            leading=11,
        ),
        "table_taeg_label": ParagraphStyle(
            "TableTaegLabel",
            fontSize=9.2,
            textColor=colors.white,
            fontName=DISPLAY_FONT,
            leading=12,
        ),
        "table_taeg_value": ParagraphStyle(
            "TableTaegValue",
            fontSize=12.2,
            textColor=gold,
            fontName=DISPLAY_FONT,
            alignment=TA_RIGHT,
            leading=13,
        ),
        "table_total_label": ParagraphStyle(
            "TableTotalLabel",
            fontSize=9.4,
            textColor=navy,
            fontName=DISPLAY_FONT,
            leading=12,
        ),
        "table_total_value": ParagraphStyle(
            "TableTotalValue",
            fontSize=12,
            textColor=navy,
            fontName=DISPLAY_FONT,
            leading=13,
            alignment=TA_RIGHT,
        ),
        "cond": ParagraphStyle(
            "Cond",
            fontSize=8.3,
            textColor=colors.HexColor("#4f5a67"),
            fontName=BODY_FONT,
            leading=11.3,
        ),
        "note": ParagraphStyle(
            "Note",
            fontSize=8,
            textColor=colors.HexColor("#e7edf5"),
            fontName=BODY_FONT,
            alignment=TA_CENTER,
            leading=11,
        ),
        "meta_badge": ParagraphStyle(
            "MetaBadge",
            fontSize=7.6,
            textColor=gold,
            fontName=MONO_FONT,
            alignment=TA_CENTER,
            leading=9.4,
        ),
        "timeline_label": ParagraphStyle(
            "TimelineLabel",
            fontSize=7.5,
            textColor=colors.HexColor("#7f8894"),
            fontName=BODY_FONT,
            leading=10,
            alignment=TA_LEFT,
        ),
    }


def draw_fixed_header_footer(canvas, doc, styles, profile: Optional[BrandingProfile] = None):
    """Premium page chrome with warm background, texture, watermark and minimal footer."""
    prof = profile or get_cached_profile()
    _t = get_design_tokens(prof)
    canvas.saveState()
    page_w, page_h = A4

    navy = colors.HexColor(_t["navy"])
    gold = colors.HexColor(_t["gold"])
    warm_white = colors.HexColor(_t["warm_white"])
    gray_warm = colors.HexColor(_t["gray_warm"])

    # base background
    canvas.setFillColor(warm_white)
    canvas.rect(0, 0, page_w, page_h, stroke=0, fill=1)

    # subtle paper-like texture
    if hasattr(canvas, "setFillAlpha"):
        canvas.setFillAlpha(0.03)
    canvas.setFillColor(navy)
    step = 13 * mm
    y = 0.0
    while y < page_h:
        x = 0.0
        while x < page_w:
            canvas.circle(x + 0.9, y + 0.9, 0.25, stroke=0, fill=1)
            x += step
        y += step
    if hasattr(canvas, "setFillAlpha"):
        canvas.setFillAlpha(1)

    # top gold line
    canvas.setFillColor(gold)
    canvas.rect(0, page_h - 3.0, page_w, 3.0, stroke=0, fill=1)

    # corner detail top-left
    canvas.saveState()
    if hasattr(canvas, "setStrokeAlpha"):
        canvas.setStrokeAlpha(0.45)
    canvas.setStrokeColor(gold)
    canvas.setLineWidth(1.0)
    m = 8 * mm
    s = 13 * mm
    canvas.line(m, page_h - m, m + s, page_h - m)
    canvas.line(m, page_h - m, m, page_h - m - s)
    canvas.line(page_w - m, m, page_w - m - s, m)
    canvas.line(page_w - m, m, page_w - m, m + s)
    if hasattr(canvas, "setStrokeAlpha"):
        canvas.setStrokeAlpha(1)
    canvas.restoreState()

    lx = doc.leftMargin
    rx = page_w - doc.rightMargin
    top_y = page_h - 15.0 * mm

    if prof.nome_agente:
        canvas.setFillColor(navy)
        canvas.setFont(DISPLAY_FONT, 11)
        canvas.drawString(lx, top_y, prof.nome_agente)

    sub_parts = []
    if prof.rete_mandante:
        sub_parts.append(f"Agente {prof.rete_mandante}")
    if prof.codice_oam:
        sub_parts.append(f"OAM {prof.codice_oam}")
    if sub_parts:
        canvas.setFillColor(gray_warm)
        canvas.setFont(MONO_FONT, 6.8)
        canvas.drawString(lx, top_y - 8.5, " | ".join(sub_parts))

    if prof.telefono:
        canvas.setFillColor(colors.HexColor("#5d6673"))
        canvas.setFont(BODY_FONT, 7.6)
        canvas.drawRightString(rx, top_y, f"Tel: {prof.telefono}")

    if prof.codice_oam:
        badge_txt = f"Registrazione OAM {prof.codice_oam}"
        canvas.setFont(MONO_FONT, 6.4)
        bw = canvas.stringWidth(badge_txt, MONO_FONT, 6.4) + 9
        bx = rx - bw
        by = top_y - 10.5
        canvas.setFillColor(navy)
        canvas.roundRect(bx, by, bw, 9.0, 2.4, stroke=0, fill=1)
        canvas.setFillColor(gold)
        canvas.drawString(bx + 4.5, by + 2.2, badge_txt)

    # Logo agente (se configurato)
    if prof.logo_path:
        logo_file = Path(prof.logo_path)
        if not logo_file.is_absolute():
            logo_file = _ASSETS_DIR / prof.logo_path
        if logo_file.exists():
            logo_sz = 15 * mm
            canvas.drawImage(
                str(logo_file), rx - logo_sz, top_y - 20,
                width=logo_sz, height=logo_sz,
                preserveAspectRatio=True, mask="auto",
            )

    fy = doc.bottomMargin - 7
    canvas.setStrokeColor(colors.HexColor(_t["line_soft"]))
    canvas.setLineWidth(0.75)
    canvas.line(doc.leftMargin, fy + 14, page_w - doc.rightMargin, fy + 14)

    canvas.setFillColor(colors.HexColor("#7c8490"))
    canvas.setFont(MONO_FONT, 6.5)
    canvas.drawString(
        doc.leftMargin,
        fy + 3,
        "Documento informativo non vincolante. Verifica precontrattuale soggetta ad approvazione.",
    )

    canvas.setFillColor(navy)
    canvas.drawRightString(rx, fy + 3, "organismo-am.it")
    canvas.linkURL("https://www.organismo-am.it/", (rx - 50, fy, rx, fy + 12), relative=0)

    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    canvas.setFillColor(colors.HexColor("#9099a6"))
    canvas.setFont(MONO_FONT, 6.1)
    canvas.drawRightString(rx, fy - 6.7, f"Generato il {ts}")

    canvas.restoreState()


def _append_preventivo_elements(
    elements,
    p: Preventivo,
    styles,
    W,
    scenario_index: int = 1,
    scenario_total: int = 1,
    profile: Optional[BrandingProfile] = None,
    text_overrides: Optional[dict] = None,
) -> None:
    prof = profile or get_cached_profile()
    _t = get_design_tokens(prof)
    _ovr = text_overrides or {}

    p.compute()
    p.validate()

    try:
        data_doc = datetime.strptime(p.data_preventivo, "%d/%m/%Y")
        data_doc_date = data_doc.date()
        scadenza = (data_doc + timedelta(days=30)).strftime("%d/%m/%Y")
    except Exception:
        data_doc_date = date.today()
        scadenza = "30 giorni dalla data emissione"

    anni = p.durata_mesi / 12
    if p.durata_mesi % 12 == 0:
        durata_anni = f"{int(anni)} anni"
    else:
        durata_anni = f"{anni:.1f}".replace(".", ",") + " anni"

    indice_netto = (p.importo_erogato / p.montante * 100) if p.montante else 0.0
    end_date = add_months(data_doc_date, p.durata_mesi)

    def section_bar(text: str) -> Table:
        t = Table([[Paragraph(text, styles["section"])]], colWidths=[W])
        t.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(_t["navy"])),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 9),
                    ("LINEBELOW", (0, 0), (-1, 0), 1.2, colors.HexColor(_t["gold"])),
                ]
            )
        )
        return t

    def kpi_box(label: str, value: str, accent: bool = False) -> Table:
        t = Table(
            [[Paragraph(label.upper(), styles["kpi_label"])], [Paragraph(value, styles["kpi_value"])]],
            colWidths=[54 * mm],
        )
        border_color = colors.HexColor(_t["gold"] if accent else "#d6dce5")
        bg_color = colors.HexColor("#fffefc" if accent else "#ffffff")
        t.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), bg_color),
                    ("BOX", (0, 0), (-1, -1), 0.9 if accent else 0.7, border_color),
                    ("TOPPADDING", (0, 0), (-1, 0), 7),
                    ("BOTTOMPADDING", (0, 1), (-1, 1), 8),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ROUNDEDCORNERS", [5, 5, 5, 5]),
                ]
            )
        )
        return t

    elements.append(Spacer(1, 2 * mm))
    if scenario_total > 1:
        elements.append(Paragraph(f"Scenario {scenario_index} di {scenario_total}", styles["subtitle"]))
        elements.append(Spacer(1, 1.0 * mm))

    elements.append(Paragraph("Riepilogo illustrativo della simulazione", styles["subtitle"]))
    elements.append(Spacer(1, 1.0 * mm))
    elements.append(Paragraph(p.tipo_finanziamento.upper(), styles["title"]))
    elements.append(Spacer(1, 2.2 * mm))

    meta_tbl = Table(
        [[Paragraph(f"Emissione: {p.data_preventivo}", styles["meta_badge"]), Paragraph(f"Validità: fino al {scadenza}", styles["meta_badge"])]],
        colWidths=[W / 2, W / 2],
    )
    meta_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f7f3e7")),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(_t["gold"])),
                ("LINEAFTER", (0, 0), (0, -1), 0.5, colors.HexColor("#e4d7a7")),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("ROUNDEDCORNERS", [5, 5, 5, 5]),
            ]
        )
    )
    elements.append(meta_tbl)
    elements.append(Spacer(1, 2 * mm))
    _disc_who = f"da {prof.nome_agente}, Agente in attività finanziaria iscritto OAM {prof.codice_oam}" if prof.nome_agente and prof.codice_oam else "dall'agente"
    _disclaimer_default = (
        f"Documento predisposto {_disc_who}, "
        "operante in qualità di agente diretto del mandante. I valori economici riportati derivano dalla "
        "simulazione effettuata sul portale ufficiale FEEVO alla data di elaborazione. Il presente documento "
        "ha finalità illustrativa e non sostituisce il SECCI né la restante documentazione ufficiale "
        "precontrattuale e contrattuale del finanziatore."
    )
    elements.append(Paragraph(_ovr.get("disclaimer", _disclaimer_default), styles["cond"]))
    elements.append(Spacer(1, 3.2 * mm))

    elements.append(section_bar("PROFILO CLIENTE"))
    cliente_rows = [
        [
            Paragraph("CLIENTE", styles["field_label"]),
            Paragraph(_ovr.get("cliente_nome", p.cliente), styles["field_value"]),
            Paragraph("DATA DI NASCITA", styles["field_label"]),
            Paragraph(f"{p.data_nascita} ({p.eta} anni)", styles["field_value"]),
        ],
        [
            Paragraph("CATEGORIA LAVORATIVA", styles["field_label"]),
            Paragraph(p.tipo_lavoro, styles["field_value"]),
            Paragraph("PROVINCIA", styles["field_label"]),
            Paragraph(p.provincia, styles["field_value"]),
        ],
    ]
    _note_text = _ovr.get("note", p.note.strip())
    if _note_text:
        cliente_rows.append([Paragraph("NOTE", styles["field_label"]), Paragraph(_note_text, styles["field_value"]), "", ""])

    cliente_tbl = Table(cliente_rows, colWidths=[30 * mm, 55 * mm, 30 * mm, 55 * mm])
    cliente_style = [
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ffffff")),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#dde2ea")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("LINEBELOW", (0, 0), (-1, 0), 0.45, colors.HexColor("#eceff4")),
        ("LINEAFTER", (1, 0), (1, -1), 0.4, colors.HexColor("#eceff4")),
    ]
    if p.note.strip():
        note_row = len(cliente_rows) - 1
        cliente_style.extend([("SPAN", (1, note_row), (3, note_row)), ("LINEBELOW", (0, 1), (-1, 1), 0.45, colors.HexColor("#eceff4"))])
    cliente_tbl.setStyle(TableStyle(cliente_style))
    elements.append(cliente_tbl)
    elements.append(Spacer(1, 3.8 * mm))

    elements.append(HeroFinanceFlowable(W, p.importo_rata, p.durata_mesi, durata_anni, p.taeg, p.tan, tokens=_t))
    elements.append(Spacer(1, 3.0 * mm))

    kpi_tbl = Table(
        [[kpi_box("Netto erogato", euro(p.importo_erogato)), kpi_box("Montante totale", euro(p.montante)), kpi_box("Interessi complessivi", euro(p.interessi), accent=True)]],
        colWidths=[56 * mm, 56 * mm, 56 * mm],
    )
    kpi_tbl.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    elements.append(kpi_tbl)
    elements.append(Spacer(1, 3.6 * mm))

    elements.append(Paragraph("IL TUO PERCORSO FINANZIARIO", styles["timeline_label"]))
    elements.append(Spacer(1, 0.8 * mm))
    elements.append(TimelineFlowable(W, p.importo_erogato, p.durata_mesi, end_date, tokens=_t))
    elements.append(Spacer(1, 2.6 * mm))

    elements.append(section_bar("DETTAGLIO ECONOMICO"))
    dettaglio_rows = [
        [Paragraph("Tipologia prodotto", styles["table_label"]), Paragraph(p.tipo_finanziamento, styles["table_value"])],
        [Paragraph("Capitale netto erogato", styles["table_label"]), Paragraph(euro(p.importo_erogato), styles["table_value"])],
        [Paragraph("Interessi complessivi", styles["table_label"]), Paragraph(euro(p.interessi), styles["table_value"])],
        [Paragraph("Durata contrattuale", styles["table_label"]), Paragraph(f"{p.durata_mesi} mesi ({durata_anni})", styles["table_value"])],
        [Paragraph("TAN nominale", styles["table_label"]), Paragraph(f"{p.tan:.3f}%", styles["table_value"])],
        [Paragraph("TAEG (Tasso Annuo Effettivo Globale)", styles["table_taeg_label"]), Paragraph(f"{p.taeg:.3f}%", styles["table_taeg_value"])],
        [Paragraph("TOTALE DA RIMBORSARE", styles["table_total_label"]), Paragraph(euro(p.montante), styles["table_total_value"])],
    ]

    dettaglio_tbl = Table(dettaglio_rows, colWidths=[112 * mm, 58 * mm])
    taeg_row = 5
    total_row = len(dettaglio_rows) - 1

    dettaglio_style = [
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#d8dee7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4.8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.8),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("LINEBELOW", (0, 0), (-1, -2), 0.35, colors.HexColor("#eceff4")),
        ("BACKGROUND", (0, taeg_row), (-1, taeg_row), colors.HexColor(_t["navy"])),
        ("LINEABOVE", (0, taeg_row), (-1, taeg_row), 1.4, colors.HexColor(_t["gold"])),
        ("LINEBELOW", (0, taeg_row), (-1, taeg_row), 1.4, colors.HexColor(_t["gold"])),
        ("LINEABOVE", (0, total_row), (-1, total_row), 1.1, colors.HexColor(_t["navy"])),
        ("BACKGROUND", (0, total_row), (-1, total_row), colors.HexColor("#f8f5ec")),
    ]
    for i in range(len(dettaglio_rows) - 1):
        if i in (taeg_row, total_row):
            continue
        bg = colors.HexColor("#ffffff") if i % 2 else colors.HexColor("#fcfcfd")
        dettaglio_style.append(("BACKGROUND", (0, i), (-1, i), bg))

    dettaglio_tbl.setStyle(TableStyle(dettaglio_style))
    elements.append(dettaglio_tbl)
    elements.append(Spacer(1, 3.5 * mm))

    elements.append(section_bar("INFORMAZIONI IMPORTANTI"))
    condizioni = [
        f"Erogazione su conto intestato a {p.cliente}.",
        "Il rimborso avviene automaticamente tramite trattenuta mensile in busta paga.",
        "Le polizze assicurative obbligatorie sono già incluse nel TAEG indicato — nessuna sorpresa.",
        "Puoi estinguere il finanziamento in anticipo: ti verrà rimborsata la quota interessi non maturata.",
        "Non sono previsti costi o spese ulteriori a carico del cliente oltre a quelli eventualmente già ricompresi nei valori risultanti dalla simulazione ufficiale.",
    ]
    cond_tbl = Table([[Paragraph(f"<font color='{_t['gold']}'>•</font> {c}", styles["cond"])] for c in condizioni], colWidths=[W])
    cond_tbl.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#d8dee7")),
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ffffff")),
                ("TOPPADDING", (0, 0), (-1, -1), 4.6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4.6),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.HexColor("#eceff4")),
            ]
        )
    )
    elements.append(cond_tbl)
    elements.append(Spacer(1, 3.2 * mm))

    _contact_parts = []
    if prof.nome_agente:
        _contact_parts.append(f"<b>{prof.nome_agente}</b>")
    if prof.telefono:
        _contact_parts.append(f"al {prof.telefono}")
    _contact_str = " ".join(_contact_parts) if _contact_parts else "l'agente di riferimento"
    _closing_default = (
        f"Simulazione elaborata in data <font color='{_t['gold']}'>{p.data_preventivo}</font>. "
        f"Per procedere o richiedere chiarimenti, contatta {_contact_str}."
    )
    closing_text = _ovr.get("closing", _closing_default)
    closing = Table([[Paragraph(closing_text, styles["note"])]], colWidths=[W])
    closing.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(_t["navy"])),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(_t["gold"])),
                ("TOPPADDING", (0, 0), (-1, -1), 6.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6.5),
                ("ROUNDEDCORNERS", [6, 6, 6, 6]),
            ]
        )
    )
    elements.append(closing)

    # ── Bollino OAM gigante (riempie lo spazio vuoto di pagina 2) ──────────
    bollino_file = None
    if prof.bollino_path:
        _bf = Path(prof.bollino_path)
        if not _bf.is_absolute():
            candidate = _ASSETS_DIR / prof.bollino_path
            if candidate.exists():
                _bf = candidate
            else:
                _bf = Path(__file__).parent / prof.bollino_path
        if _bf.exists():
            bollino_file = _bf
    if bollino_file:
        badge_size = 64 * mm
        elements.append(Spacer(1, 12 * mm))
        elements.append(
            _ClickableImage(
                bollino_file,
                width=badge_size,
                height=badge_size,
                url="https://www.organismo-am.it/",
            )
        )
        caption_style = ParagraphStyle(
            "oam_caption",
            fontName=MONO_FONT,
            fontSize=7,
            textColor=colors.HexColor(_t["gray_warm"]),
            alignment=TA_CENTER,
            spaceAfter=0,
        )
        elements.append(Spacer(1, 3 * mm))
        elements.append(
            Paragraph(
                f"Iscritto all'Organismo degli Agenti e Mediatori{' — ' + prof.codice_oam if prof.codice_oam else ''}",
                caption_style,
            )
        )


def crea_preventivi_pdf(
    preventivi: list[Preventivo],
    output_path: Path,
    profile: Optional[BrandingProfile] = None,
    text_overrides: Optional[dict] = None,
) -> Path:
    if not preventivi:
        raise ValueError("Nessun preventivo da generare.")

    prof = profile or load_profile()
    tokens = get_design_tokens(prof)
    styles = build_styles(tokens)
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        topMargin=3.0 * cm,
        bottomMargin=2.25 * cm,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
    )
    W = doc.width
    elements = []

    total = len(preventivi)
    for idx, prev in enumerate(preventivi, start=1):
        _append_preventivo_elements(elements, prev, styles, W, idx, total, profile=prof, text_overrides=text_overrides)
        if idx < total:
            elements.append(PageBreak())

    doc.build(
        elements,
        onFirstPage=lambda c, d: draw_fixed_header_footer(c, d, styles, prof),
        onLaterPages=lambda c, d: draw_fixed_header_footer(c, d, styles, prof),
    )
    return output_path


def crea_preventivo_pdf(p: Preventivo, output_path: Path, profile: Optional[BrandingProfile] = None, text_overrides: Optional[dict] = None) -> Path:
    return crea_preventivi_pdf([p], output_path, profile=profile, text_overrides=text_overrides)


# =========================
#  CLI: RACCOLTA DATI
# =========================

def collect_from_cli(args) -> Preventivo:
    if args.non_interactive:
        missing: list[str] = []
        for flag, value in (
            ("--cliente", args.cliente),
            ("--data-nascita", args.data_nascita),
            ("--tipo-lavoro", args.tipo_lavoro),
            ("--provincia", args.provincia),
            ("--tipo-finanziamento", args.tipo_finanziamento),
        ):
            if not str(value or "").strip():
                missing.append(flag)
        for flag, value in (
            ("--importo-rata", args.importo_rata),
            ("--durata-mesi", args.durata_mesi),
            ("--tan", args.tan),
            ("--taeg", args.taeg),
            ("--importo-erogato", args.importo_erogato),
        ):
            if value is None:
                missing.append(flag)
        if missing:
            raise ValueError(
                "In modalità --non-interactive mancano i parametri obbligatori: "
                + ", ".join(missing)
            )
        p = Preventivo(
            cliente=args.cliente,
            data_nascita=args.data_nascita,
            tipo_lavoro=args.tipo_lavoro,
            provincia=args.provincia,
            note=args.note or "",
            tipo_finanziamento=args.tipo_finanziamento,
            importo_rata=args.importo_rata,
            durata_mesi=args.durata_mesi,
            tan=args.tan,
            taeg=args.taeg,
            importo_erogato=args.importo_erogato,
        )
        p.compute()
        p.validate()
        return p

    # Interattivo (guidato)
    print("\n=== Generatore Preventivo (CLI guidata) ===\n")
    cliente = ask("Cliente (Nome e Cognome)", args.cliente or "")
    data_nascita = ask("Data di nascita (GG/MM/AAAA)", args.data_nascita or "")
    tipo_lavoro = ask("Qualifica / Tipo lavoro", args.tipo_lavoro or "Dipendente Statale")
    provincia = ask("Provincia / Sede lavorativa", args.provincia or "Roma")
    note = ask("Note (opzionale)", args.note or "")

    tipo_finanziamento = ask_choice("Tipo finanziamento", ["Cessione del Quinto", "Delega di Pagamento"],
                                    args.tipo_finanziamento or "Cessione del Quinto")

    importo_rata = ask_float("Rata mensile", args.importo_rata if args.importo_rata is not None else 350.00, min_value=1.0)
    durata_mesi = ask_int("Durata (mesi)", args.durata_mesi if args.durata_mesi is not None else 120, min_value=1)
    tan = ask_float("TAN (%)", args.tan if args.tan is not None else 4.500, min_value=0.0)
    taeg = ask_float("TAEG (%)", args.taeg if args.taeg is not None else 4.750, min_value=0.0)
    importo_erogato = ask_float("Importo netto erogato", args.importo_erogato if args.importo_erogato is not None else 30000.00, min_value=1.0)

    p = Preventivo(
        cliente=cliente,
        data_nascita=data_nascita,
        tipo_lavoro=tipo_lavoro,
        provincia=provincia,
        note=note,
        tipo_finanziamento=tipo_finanziamento,
        importo_rata=importo_rata,
        durata_mesi=durata_mesi,
        tan=tan,
        taeg=taeg,
        importo_erogato=importo_erogato,
    )
    p.compute()
    return p

def build_output_path(cliente: str, out_dir: Path, data_str: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = sanitize_filename(cliente)
    return out_dir / f"Preventivo_{safe}_{data_str.replace('/', '-')}.pdf"

def build_output_path_multi(cliente: str, out_dir: Path, data_str: str, count: int) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = sanitize_filename(cliente)
    return out_dir / f"Preventivi_{safe}_{data_str.replace('/', '-')}_{count}opzioni.pdf"

def parse_decimal_loose(s: str) -> float:
    if isinstance(s, (int, float)):
        return float(s)
    raw = str(s).strip().replace("EUR", "").replace("€", "").replace(" ", "")
    if "." in raw and "," in raw:
        raw = raw.replace(".", "").replace(",", ".")
    elif "." in raw and "," not in raw:
        # Treat 30.000 / 1.250.000 as thousands separators, but keep 4.75 as decimal.
        if re.fullmatch(r"\d{1,3}(?:\.\d{3})+", raw):
            raw = raw.replace(".", "")
    elif "," in raw:
        raw = raw.replace(",", ".")
    return float(raw)

def normalize_tipo_finanziamento(raw: str) -> str:
    m = {
        "cessione del quinto": "Cessione del Quinto",
        "cessione": "Cessione del Quinto",
        "delega di pagamento": "Delega di Pagamento",
        "delega": "Delega di Pagamento",
    }
    key = raw.strip().lower()
    if key not in m:
        raise ValueError(
            "Tipo finanziamento scenario non valido. Usa 'Cessione del Quinto' o 'Delega di Pagamento'."
        )
    return m[key]

def parse_scenario_line(raw: str, idx: int) -> dict:
    # Formato: rata;durata;tan;taeg;importo_erogato[;tipo_finanziamento]
    parts = [x.strip() for x in raw.split(";")]
    if len(parts) not in (5, 6):
        raise ValueError(
            f"Scenario {idx}: formato non valido. Usa rata;durata;tan;taeg;importo_erogato[;tipo_finanziamento]"
        )
    try:
        rata = parse_decimal_loose(parts[0])
        durata = int(parts[1])
        tan = parse_decimal_loose(parts[2])
        taeg = parse_decimal_loose(parts[3])
        erogato = parse_decimal_loose(parts[4])
    except Exception as exc:
        raise ValueError(
            f"Scenario {idx}: numeri non validi. Usa separatore ';' e numeri tipo 350,00;120;4,5;4,75;30000,00"
        ) from exc

    if durata <= 0 or rata <= 0 or erogato <= 0:
        raise ValueError(f"Scenario {idx}: rata, durata e importo erogato devono essere > 0.")
    if not (0 <= tan < 50) or not (0 <= taeg < 50):
        raise ValueError(f"Scenario {idx}: TAN/TAEG fuori range plausibile.")

    tipo = None
    if len(parts) == 6 and parts[5]:
        tipo = normalize_tipo_finanziamento(parts[5])

    return {
        "importo_rata": rata,
        "durata_mesi": durata,
        "tan": tan,
        "taeg": taeg,
        "importo_erogato": erogato,
        "tipo_finanziamento": tipo,
    }

def parse_scenari_text(raw: str) -> list[str]:
    if not raw.strip():
        return []
    normalized = raw.replace("|", "\n")
    lines = [ln.strip() for ln in normalized.splitlines() if ln.strip()]
    return lines

def build_extra_preventivi(base: Preventivo, scenario_lines: list[str]) -> list[Preventivo]:
    out: list[Preventivo] = []
    for i, line in enumerate(scenario_lines, start=1):
        s = parse_scenario_line(line, i)
        tipo_fin = s["tipo_finanziamento"] or base.tipo_finanziamento
        p = Preventivo(
            cliente=base.cliente,
            data_nascita=base.data_nascita,
            tipo_lavoro=base.tipo_lavoro,
            provincia=base.provincia,
            note=base.note,
            tipo_finanziamento=tipo_fin,
            importo_rata=s["importo_rata"],
            durata_mesi=s["durata_mesi"],
            tan=s["tan"],
            taeg=s["taeg"],
            importo_erogato=s["importo_erogato"],
        )
        p.compute()
        p.validate()
        out.append(p)
    return out


# =========================
#  PDF TEMPLATE COMPILER (MEF)
# =========================

@dataclass(frozen=True)
class PdfFieldDef:
    name: str
    widget: str
    label: str
    placeholder: str = ""
    help_text: str = ""
    full_width: bool = False
    font_size: float = 10.0
    default: str = ""
    overlay_rect: Optional[tuple[float, float, float, float]] = None
    overlay_page: int = 0


@dataclass(frozen=True)
class PdfSectionDef:
    title: str
    icon: str
    description: str
    fields: tuple[PdfFieldDef, ...]


@dataclass(frozen=True)
class PdfTemplateSpec:
    key: str
    slug: str
    label: str
    description: str
    summary: str
    template_name: str
    output_prefix: str
    filename_field: str
    button_label: str
    sections: tuple[PdfSectionDef, ...]


_PDF_TEMPLATE_DIR = _resolve_pdf_template_dir()
_DYNAMIC_DEFAULT_TODAY = "__TODAY__"

ALLEGATO_C_SPEC = PdfTemplateSpec(
    key="allegato_c",
    slug="allegato-c",
    label="Allegato C Flussi Finanziari MEF",
    description="Dichiarazione di finanziamento cessione per i flussi finanziari MEF / NoiPA.",
    summary="20 campi mappati sul template originale Creditonet.",
    template_name="Allegato C Creditonet (1).pdf",
    output_prefix="AllegatoC",
    filename_field="dipendente_nome_cognome",
    button_label="Scarica Allegato C compilato",
    sections=(
        PdfSectionDef(
            title="Sezione Istituto Mutuante",
            icon="🏦",
            description="Campi economici e dati dell'eventuale finanziamento da estinguere.",
            fields=(
                PdfFieldDef("importo_erogato", "Testo87", "Importo erogato", "Es. 30.000,00", overlay_rect=(154, 205, 273, 218)),
                PdfFieldDef("importo_globale_ceduto", "Testo88", "Importo globale ceduto", "Es. 42.000,00", overlay_rect=(441, 205, 561, 218)),
                PdfFieldDef("spese_complessive", "Testo89", "Spese complessive", "Es. 1.250,00", overlay_rect=(154, 224, 273, 236)),
                PdfFieldDef("interessi_complessivi", "Testo90", "Interessi complessivi", "Es. 10.750,00", overlay_rect=(441, 224, 561, 236)),
                PdfFieldDef("tan", "Testo91", "TAN", "Es. 4,500", overlay_rect=(64, 242, 138, 255)),
                PdfFieldDef("isc_taeg", "Testo92", "ISC / TAEG", "Es. 5,120", overlay_rect=(350, 242, 426, 255)),
                PdfFieldDef("numero_rate_estinguibili", "Testo93", "Estinguibile in n°", "Es. 120", overlay_rect=(154, 260, 273, 271)),
                PdfFieldDef("importo_rata_estinzione", "Testo94", "Rate mensili di euro", "Es. 350,00", overlay_rect=(441, 260, 561, 271)),
                PdfFieldDef(
                    "garanzia_assicurativa",
                    "Testo95",
                    "Garanzia assicurativa",
                    "Es. Polizza n. 12345 del 15/04/2026",
                    full_width=True,
                    font_size=9.0,
                    overlay_rect=(190, 294, 562, 308),
                ),
                PdfFieldDef("revoca_finanziamento_importo", "Testo2", "Revoca altro finanziamento: importo", "Es. 220,00", overlay_rect=(297, 322, 390, 334)),
                PdfFieldDef("revoca_finanziamento_scadenza", "Testo3", "Revoca altro finanziamento: scadenza", "Es. 31/12/2030", overlay_rect=(459, 322, 561, 334)),
                PdfFieldDef(
                    "revoca_finanziamento_contratto",
                    "Testo8",
                    "Revoca altro finanziamento: contratto con",
                    "Es. Banca XYZ - pratica 123456",
                    full_width=True,
                    font_size=9.0,
                    overlay_rect=(100, 340, 562, 352),
                ),
            ),
        ),
        PdfSectionDef(
            title="Sezione Dipendente",
            icon="👤",
            description="Dati anagrafici e amministrativi del dipendente / istante.",
            fields=(
                PdfFieldDef(
                    "dipendente_nome_cognome",
                    "Testo96",
                    "Cognome e Nome",
                    "Es. Rossi Mario",
                    full_width=True,
                    font_size=9.5,
                    overlay_rect=(135, 414, 562, 427),
                ),
                PdfFieldDef("dipendente_nascita_luogo", "Testo97", "Nato/a a", "Es. Roma", overlay_rect=(81, 444, 319, 456)),
                PdfFieldDef("dipendente_nascita_provincia", "Testo98", "Prov.", "Es. RM", font_size=9.5, overlay_rect=(369, 444, 408, 456)),
                PdfFieldDef("dipendente_nascita_data", "Testo99", "Il", "GG/MM/AAAA", overlay_rect=(441, 444, 562, 456)),
                PdfFieldDef(
                    "dipendente_codice_fiscale",
                    "Testo100",
                    "Codice Fiscale",
                    "Es. RSSMRA80A01H501Z",
                    full_width=True,
                    font_size=9.5,
                    overlay_rect=(127, 468, 562, 481),
                ),
                PdfFieldDef(
                    "dipendente_in_servizio_presso",
                    "Testo4",
                    "In servizio presso",
                    "Es. Ministero dell'Economia e delle Finanze",
                    full_width=True,
                    font_size=9.0,
                    overlay_rect=(127, 497, 562, 510),
                ),
                PdfFieldDef(
                    "dipendente_ente_appartenenza",
                    "Testo5",
                    "Ente di appartenenza",
                    "Es. Ragioneria Territoriale dello Stato di Roma",
                    full_width=True,
                    font_size=9.0,
                    overlay_rect=(154, 525, 562, 539),
                ),
                PdfFieldDef("data_modulo", "Testo6", "Data", "GG/MM/AAAA", default=_DYNAMIC_DEFAULT_TODAY, overlay_rect=(28, 794, 147, 816)),
            ),
        ),
    ),
)

ALLEGATO_E_SPEC = PdfTemplateSpec(
    key="allegato_e",
    slug="allegato-e",
    label="Allegato E Delega MEF",
    description="Istanza di delegazione di pagamento con parte riservata all'istituto delegatario.",
    summary="43 campi mappati sui widget delle prime 2 pagine del template originale.",
    template_name="AllegatoE determina positiva ATC (1) (3).pdf",
    output_prefix="AllegatoE",
    filename_field="istante_nome_completo",
    button_label="Scarica Allegato E compilato",
    sections=(
        PdfSectionDef(
            title="Intestazione",
            icon="📍",
            description="Le tre righe in alto a destra del modulo.",
            fields=(
                PdfFieldDef("destinatario_riga_1", "Testo1", "Riga 1 destinatario", "Es. Direzione dei Servizi del Tesoro", full_width=True, font_size=9.0),
                PdfFieldDef("destinatario_riga_2", "Testo2", "Riga 2 destinatario", "Es. Ufficio Stipendi Centrali", full_width=True, font_size=9.0),
                PdfFieldDef("destinatario_riga_3", "Testo3", "Riga 3 destinatario", "Es. Roma", full_width=True, font_size=9.0),
            ),
        ),
        PdfSectionDef(
            title="Anagrafica Istante",
            icon="👤",
            description="Dati personali riportati nella prima pagina.",
            fields=(
                PdfFieldDef("istante_nome_completo", "Testo4", "Il/La sottoscritto/a", "Es. Rossi Mario", full_width=True, font_size=9.5),
                PdfFieldDef("nascita_comune", "Testo5", "Nato/a a", "Es. Roma", full_width=True),
                PdfFieldDef("nascita_provincia", "Testo6", "Provincia di nascita", "Es. Roma"),
                PdfFieldDef("nascita_sigla_provincia", "Testo7", "Sigla prov. nascita", "Es. RM", font_size=9.5),
                PdfFieldDef("nascita_data", "Testo8", "Data di nascita", "GG/MM/AAAA"),
                PdfFieldDef("codice_fiscale", "Testo9", "Codice fiscale", "Es. RSSMRA80A01H501Z"),
                PdfFieldDef("partita_stipendiale", "Testo10", "Partita stipendiale n.", "Es. 1234567"),
                PdfFieldDef("residenza_comune", "Testo11", "Residente a", "Es. Roma", full_width=True),
                PdfFieldDef("residenza_provincia", "Testo12", "Provincia di residenza", "Es. Roma"),
                PdfFieldDef("residenza_sigla_provincia", "Testo13", "Sigla prov. residenza", "Es. RM", font_size=9.5),
                PdfFieldDef("residenza_cap", "Testo14", "CAP", "Es. 00100"),
                PdfFieldDef("residenza_via", "Testo15", "Via/Piazza", "Es. Via Nazionale", full_width=True),
                PdfFieldDef("residenza_numero", "Testo16", "N.", "Es. 10"),
                PdfFieldDef("telefono", "Testo17", "Telefono", "Es. 06 1234567"),
                PdfFieldDef("fax", "Testo18", "Fax", "Opzionale"),
                PdfFieldDef("email_local_part", "Testo19", "Email - parte prima della @", "Es. mario.rossi"),
                PdfFieldDef("email_domain", "Testo20", "Email - dominio", "Es. pec.it"),
            ),
        ),
        PdfSectionDef(
            title="Richiesta di Delega",
            icon="💸",
            description="Campi nella parte dispositiva della prima pagina.",
            fields=(
                PdfFieldDef(
                    "istituto_delegatario",
                    "Testo21",
                    "Ha chiesto un finanziamento a",
                    "Es. Banca XYZ S.p.A.",
                    full_width=True,
                    font_size=9.0,
                ),
                PdfFieldDef("importo_trattenuta_mensile", "Testo22", "Importo di euro da trattenere", "Es. 350,00"),
                PdfFieldDef(
                    "iban_istituto_delegatario",
                    "Testo23",
                    "IBAN / coordinate conto istituto delegatario",
                    "Es. IT60X0542811101000000123456",
                    full_width=True,
                    font_size=8.5,
                ),
            ),
        ),
        PdfSectionDef(
            title="Parte Riservata all'Istituto Delegatario",
            icon="🏦",
            description="Campi economici presenti nella seconda pagina del modulo.",
            fields=(
                PdfFieldDef("importo_finanziamento_cifre", "Testo24", "Importo finanziamento (in cifre)", "Es. 30.000,00"),
                PdfFieldDef("importo_finanziamento_lettere", "Testo25", "Importo finanziamento (in lettere)", "Es. trentamila/00", full_width=True, font_size=8.8),
                PdfFieldDef("importo_globale_ceduto_cifre", "Testo26", "Importo globale ceduto (in cifre)", "Es. 42.000,00"),
                PdfFieldDef("importo_globale_ceduto_lettere", "Testo27", "Importo globale ceduto (in lettere)", "Es. quarantaduemila/00", full_width=True, font_size=8.8),
                PdfFieldDef("spese_complessive_cifre", "Testo28", "Spese complessive euro", "Es. 1.250,00"),
                PdfFieldDef("interessi_complessivi_cifre", "Testo29", "Interessi complessivi euro", "Es. 10.750,00"),
                PdfFieldDef("tan", "Testo30", "TAN", "Es. 4,500"),
                PdfFieldDef("taeg", "Testo31", "TAEG", "Es. 5,120"),
                PdfFieldDef("teg", "Testo32", "TEG", "Es. 4,980"),
                PdfFieldDef("numero_rate_da_estinguere", "Testo33", "Finanziamento da estinguere in n.", "Es. 120"),
                PdfFieldDef("importo_rata_estinzione", "Testo34", "Rate mensili di euro", "Es. 350,00"),
                PdfFieldDef("garanzia_prestito", "Testo35", "Garanzia del prestito", "Es. Polizza n. 12345 del 15/04/2026", full_width=True, font_size=8.8),
                PdfFieldDef(
                    "estinzione_altro_finanziamento_istituto",
                    "Testo36",
                    "Altro finanziamento in corso, contratto con",
                    "Es. Banca XYZ S.p.A.",
                    full_width=True,
                    font_size=8.8,
                ),
                PdfFieldDef("estinzione_altro_finanziamento_rata", "Testo37", "Per euro mensili", "Es. 220,00"),
                PdfFieldDef("estinzione_altro_finanziamento_scadenza", "Testo38", "Scadenza", "Es. 31/12/2030"),
                PdfFieldDef("luogo_timbro_istituto", "Testo39", "Luogo", "Es. Roma"),
                PdfFieldDef("data_timbro_istituto", "Testo40", "Data", "GG/MM/AAAA", default=_DYNAMIC_DEFAULT_TODAY),
            ),
        ),
        PdfSectionDef(
            title="Autentica di Firma",
            icon="🖊️",
            description="Parte bassa della seconda pagina, da compilare se necessario.",
            fields=(
                PdfFieldDef(
                    "documento_identificazione",
                    "Testo41",
                    "Identificata a mezzo",
                    "Es. carta d'identità n. AA123456",
                    full_width=True,
                    font_size=8.8,
                ),
                PdfFieldDef("luogo_autentica", "Testo42", "Luogo", "Es. Roma"),
                PdfFieldDef("data_autentica", "Testo43", "Data", "GG/MM/AAAA", default=_DYNAMIC_DEFAULT_TODAY),
            ),
        ),
    ),
)

_PDF_TEMPLATE_SPECS = {
    ALLEGATO_C_SPEC.key: ALLEGATO_C_SPEC,
    ALLEGATO_E_SPEC.key: ALLEGATO_E_SPEC,
    ALLEGATO_C_SPEC.slug: ALLEGATO_C_SPEC,
    ALLEGATO_E_SPEC.slug: ALLEGATO_E_SPEC,
}


def get_pdf_template_spec(key_or_slug: str) -> PdfTemplateSpec:
    spec = _PDF_TEMPLATE_SPECS.get((key_or_slug or "").strip().lower())
    if spec is None:
        raise KeyError(f"Template PDF non supportato: {key_or_slug}")
    return spec


def iter_pdf_fields(spec: PdfTemplateSpec):
    for section in spec.sections:
        for field in section.fields:
            yield field


def sanitize_pdf_text(raw: object) -> str:
    text = str(raw or "").replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s{2,}", " ", text)
    return text


def default_pdf_form_values(spec: PdfTemplateSpec) -> dict[str, str]:
    values: dict[str, str] = {}
    for field in iter_pdf_fields(spec):
        if field.default:
            values[field.name] = (
                datetime.now().strftime("%d/%m/%Y")
                if field.default == _DYNAMIC_DEFAULT_TODAY
                else field.default
            )
    return values


def sanitize_pdf_form_payload(spec: PdfTemplateSpec, raw: dict[str, object]) -> dict[str, str]:
    values = default_pdf_form_values(spec)
    has_any_value = False
    for field in iter_pdf_fields(spec):
        value = sanitize_pdf_text(raw.get(field.name, ""))
        if value:
            values[field.name] = value
            has_any_value = True
        else:
            values.setdefault(field.name, "")
    if not has_any_value:
        raise ValueError("Compila almeno un campo prima di generare il PDF.")
    return values


def require_pymupdf():
    if fitz is None:
        raise RuntimeError("PyMuPDF non installato. Esegui: pip install pymupdf")
    return fitz


def build_pdf_template_output_path(spec: PdfTemplateSpec, values: dict[str, str], out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    identifier = sanitize_filename(values.get(spec.filename_field, "")) or spec.output_prefix
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    return out_dir / f"{spec.output_prefix}_{identifier}_{stamp}.pdf"


def render_pdf_template(spec: PdfTemplateSpec, values: dict[str, str], output_path: Path) -> Path:
    fitz_mod = require_pymupdf()
    template_path = _PDF_TEMPLATE_DIR / spec.template_name
    if not template_path.exists():
        raise FileNotFoundError(f"Template non trovato: {template_path.name}")

    field_map = {field.widget: field for field in iter_pdf_fields(spec)}
    doc = fitz_mod.open(template_path)
    doc.need_appearances(True)
    seen_widgets: set[str] = set()
    seen_overlays: set[str] = set()
    try:
        for page_index, page in enumerate(doc):
            for widget in page.widgets() or []:
                field = field_map.get(widget.field_name)
                if field is None:
                    continue
                widget.field_value = sanitize_pdf_text(values.get(field.name, ""))
                if hasattr(widget, "text_font"):
                    widget.text_font = "Helv"
                if hasattr(widget, "text_fontsize"):
                    widget.text_fontsize = field.font_size
                widget.update()
                seen_widgets.add(widget.field_name)

            for field in iter_pdf_fields(spec):
                if field.widget in seen_widgets:
                    continue
                if field.overlay_rect is None or field.overlay_page != page_index:
                    continue
                field_value = sanitize_pdf_text(values.get(field.name, ""))
                if not field_value:
                    seen_overlays.add(field.widget)
                    continue
                rect = fitz_mod.Rect(*field.overlay_rect)
                page.insert_text(
                    fitz_mod.Point(rect.x0 + 2, rect.y1 - 3),
                    field_value,
                    fontname="helv",
                    fontsize=field.font_size,
                    color=(0, 0, 0),
                )
                seen_overlays.add(field.widget)

        missing_widgets = sorted(set(field_map) - seen_widgets - seen_overlays)
        if missing_widgets:
            raise RuntimeError(
                "Nel template PDF mancano alcuni widget attesi: " + ", ".join(missing_widgets)
            )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
    finally:
        doc.close()
    return output_path


def generate_pdf_template(spec_key: str, raw_values: dict[str, object], out_dir: Path) -> tuple[Path, dict[str, str], PdfTemplateSpec]:
    spec = get_pdf_template_spec(spec_key)
    values = sanitize_pdf_form_payload(spec, raw_values)
    output_path = build_pdf_template_output_path(spec, values, out_dir)
    render_pdf_template(spec, values, output_path)
    return output_path, values, spec


# =========================
#  DOSSIER ENGINE (NO AI)
# =========================

@dataclass(frozen=True)
class CaseFieldDef:
    name: str
    label: str
    category: str
    placeholder: str = ""


@dataclass(frozen=True)
class DocumentTypeDef:
    key: str
    label: str
    keywords: tuple[str, ...]
    description: str = ""
    helper_text: str = ""
    user_selectable: bool = False


@dataclass
class DocumentExtractionResult:
    filename: str
    document_key: str
    document_label: str
    page_count: int
    text_length: int
    keyword_hits: int
    extracted_fields: dict[str, str]
    warnings: list[str]


@dataclass(frozen=True)
class AggregatedCaseField:
    name: str
    label: str
    value: str
    source_filename: str
    source_document_label: str


def _document_result_to_dict(result: DocumentExtractionResult) -> dict[str, object]:
    return {
        "filename": result.filename,
        "document_key": result.document_key,
        "document_label": result.document_label,
        "page_count": result.page_count,
        "text_length": result.text_length,
        "keyword_hits": result.keyword_hits,
        "extracted_fields": dict(result.extracted_fields),
        "warnings": list(result.warnings),
    }


def _document_result_from_dict(data: dict[str, object]) -> DocumentExtractionResult:
    return DocumentExtractionResult(
        filename=sanitize_pdf_text(data.get("filename", "")),
        document_key=sanitize_pdf_text(data.get("document_key", "")),
        document_label=sanitize_pdf_text(data.get("document_label", "")),
        page_count=int(data.get("page_count", 0) or 0),
        text_length=int(data.get("text_length", 0) or 0),
        keyword_hits=int(data.get("keyword_hits", 0) or 0),
        extracted_fields={
            sanitize_pdf_text(name): sanitize_pdf_text(value)
            for name, value in dict(data.get("extracted_fields", {}) or {}).items()
            if sanitize_pdf_text(name) and sanitize_pdf_text(value)
        },
        warnings=[sanitize_pdf_text(item) for item in list(data.get("warnings", []) or []) if sanitize_pdf_text(item)],
    )


_MONEY_CAPTURE = r"((?:EUR\s*)?(?:€\s*)?\d{1,3}(?:[.\s]\d{3})*(?:,\d{2,3})|(?:EUR\s*)?(?:€\s*)?\d+(?:,\d{2,3})?)"
_DATE_CAPTURE = r"([0-3]?\d[\/\.-][01]?\d[\/\.-](?:\d{4}|\d{2}))"
_DOSSIER_PDF_EXTENSIONS = {".pdf"}
_DOSSIER_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
_DOSSIER_SUPPORTED_EXTENSIONS = _DOSSIER_PDF_EXTENSIONS | _DOSSIER_IMAGE_EXTENSIONS
_OCR_TRIGGER_TEXT_LENGTH = 80
_LOW_TEXT_WARNING_LENGTH = 50
_ALLOWED_INSTALLMENT_COUNTS = {48, 60, 72, 84, 96, 108, 120}
_TESSERACT_CANDIDATE_PATHS = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    str(Path.home() / "AppData" / "Local" / "Programs" / "Tesseract-OCR" / "tesseract.exe"),
)
_LOCAL_RUNTIME_TEMP_DIR = _default_runtime_temp_dir()


def _iter_bundle_runtime_roots() -> list[Path]:
    roots: list[Path] = []
    for candidate in (
        getattr(sys, "_MEIPASS", ""),
        Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else None,
        Path(__file__).resolve().parent,
    ):
        if not candidate:
            continue
        path = Path(candidate)
        if path.exists() and path not in roots:
            roots.append(path)
    return roots

CASE_FIELD_DEFS = (
    CaseFieldDef("full_name", "Nome e cognome", "anagrafica", "Es. Mario Rossi"),
    CaseFieldDef("birth_place", "Luogo di nascita", "anagrafica", "Es. Roma"),
    CaseFieldDef("birth_date", "Data di nascita", "anagrafica", "GG/MM/AAAA"),
    CaseFieldDef("birth_province_name", "Provincia di nascita", "anagrafica", "Es. Roma"),
    CaseFieldDef("birth_province_code", "Sigla provincia di nascita", "anagrafica", "Es. RM"),
    CaseFieldDef("tax_code", "Codice fiscale", "anagrafica", "Es. RSSMRA80A01H501Z"),
    CaseFieldDef("payroll_number", "Partita stipendiale", "lavoro", "Es. 12345678"),
    CaseFieldDef("residence_city", "Comune di residenza", "residenza", "Es. Roma"),
    CaseFieldDef("residence_province_name", "Provincia di residenza", "residenza", "Es. Roma"),
    CaseFieldDef("residence_province_code", "Sigla provincia di residenza", "residenza", "Es. RM"),
    CaseFieldDef("residence_cap", "CAP", "residenza", "Es. 00100"),
    CaseFieldDef("residence_street", "Via / Piazza", "residenza", "Es. Via Roma"),
    CaseFieldDef("residence_number", "Numero civico", "residenza", "Es. 10"),
    CaseFieldDef("phone", "Telefono", "contatti", "Es. 061234567"),
    CaseFieldDef("fax", "Fax", "contatti", "Opzionale"),
    CaseFieldDef("email", "Email", "contatti", "Es. nome@pec.it"),
    CaseFieldDef("service_office", "In servizio presso", "lavoro", "Es. I.C. Faenza San Rocco"),
    CaseFieldDef("employer_entity", "Ente di appartenenza", "lavoro", "Es. Ministero dell'Istruzione"),
    CaseFieldDef("lender_name", "Istituto / finanziaria", "finanza", "Es. Istituto delegatario"),
    CaseFieldDef("iban", "IBAN istituto delegatario", "bancario", "Es. IT60X0542811101000000123456"),
    CaseFieldDef("borrower_iban", "IBAN personale / accredito stipendio", "bancario", "Es. IT60X0542811101000000123456"),
    CaseFieldDef("loan_amount", "Importo finanziamento", "finanza", "Es. 30.000,00"),
    CaseFieldDef("net_disbursed", "Importo erogato", "finanza", "Es. 30.000,00"),
    CaseFieldDef("total_ceded", "Importo globale ceduto / montante", "finanza", "Es. 42.000,00"),
    CaseFieldDef("fees_total", "Spese complessive", "finanza", "Es. 0,00"),
    CaseFieldDef("interest_total", "Interessi complessivi", "finanza", "Es. 10.750,00"),
    CaseFieldDef("monthly_installment", "Rata / trattenuta mensile", "finanza", "Es. 350,00"),
    CaseFieldDef("salary_fifth", "Quinto cedibile", "finanza", "Es. 339,51"),
    CaseFieldDef("installment_count", "Numero rate / durata mesi", "finanza", "48, 60, 72, 84, 96, 108, 120"),
    CaseFieldDef("tan", "TAN", "finanza", "Es. 4,72"),
    CaseFieldDef("taeg", "TAEG", "finanza", "Es. 4,83"),
    CaseFieldDef("teg", "TEG", "finanza", "Es. 4,80"),
    CaseFieldDef("insurance", "Garanzia / polizza", "finanza", "Es. Polizza n. ..."),
    CaseFieldDef("other_financing_lender", "Altro finanziamento: istituto", "finanza", "Es. Banca XYZ"),
    CaseFieldDef("other_financing_installment", "Altro finanziamento: rata", "finanza", "Es. 220,00"),
    CaseFieldDef("other_financing_expiry", "Altro finanziamento: scadenza", "finanza", "GG/MM/AAAA"),
)

_CASE_FIELD_DEF_MAP = {field.name: field for field in CASE_FIELD_DEFS}
_CASE_CATEGORY_LABELS = {
    "anagrafica": "Anagrafica",
    "residenza": "Residenza",
    "contatti": "Contatti",
    "lavoro": "Lavoro",
    "finanza": "Finanza",
    "bancario": "Coordinate",
}

DOCUMENT_TYPE_DEFS = (
    DocumentTypeDef(
        "cedolino_noipa",
        "Busta paga NoiPA",
        ("noipa", "cedolino", "id cedolino", "anagrafica del dipendente", "ufficio servizio"),
        "Estrae anagrafica lavorativa, ente, partita stipendiale, quinto cedibile e coordinate presenti nel cedolino.",
        "Meglio PDF testuale; se carichi una scansione usa OCR locale.",
        True,
    ),
    DocumentTypeDef(
        "contratto_finanziamento",
        "Contratto di finanziamento",
        ("informazioni europee di base sul credito ai consumatori", "prestito con delegazione di pagamento", "numero rate mensili da pagare", "importo rata mensile", "tasso annuo effettivo globale"),
        "Estrae i dati del prestito e del richiedente dal contratto.",
        "Su alcuni layout i dati principali si trovano nelle pagine interne, spesso attorno a pagina 5.",
        True,
    ),
    DocumentTypeDef(
        "carta_identita",
        "Carta di identità",
        ("carta d'identità", "carta di identita", "repubblica italiana", "luogo di nascita", "cognome"),
        "Estrae nome, cognome, luogo e data di nascita e altri dati leggibili dal documento.",
        "Funziona meglio con scansione o foto nitida del documento completo.",
        True,
    ),
    DocumentTypeDef(
        "tessera_sanitaria",
        "Tessera sanitaria",
        ("tessera sanitaria", "codice fiscale", "tessera europea", "cognome", "nome"),
        "Estrae soprattutto codice fiscale e dati anagrafici di base.",
        "Meglio foto o scansione nitida della tessera completa.",
        True,
    ),
    DocumentTypeDef("documento_finanziario", "Documento finanziario", ("tan", "taeg", "importo erogato", "importo globale ceduto", "montante", "totale da rimborsare", "durata", "rata")),
    DocumentTypeDef("documento_anagrafico", "Documento anagrafico", ("codice fiscale", "data di nascita", "nato", "residente", "cognome")),
    DocumentTypeDef("coordinate_bancarie", "Coordinate bancarie", ("iban", "conto corrente", "istituto delegatario")),
    DocumentTypeDef("modulo_mef", "Modulo MEF", ("allegato e", "allegato c", "partita stipendiale")),
)

_DOCUMENT_TYPE_DEF_MAP = {doc_type.key: doc_type for doc_type in DOCUMENT_TYPE_DEFS}
_DOSSIER_UPLOAD_DOCUMENT_TYPES = tuple(doc_type for doc_type in DOCUMENT_TYPE_DEFS if doc_type.user_selectable)
_DOSSIER_UPLOAD_DOCUMENT_KEYS = {doc_type.key for doc_type in _DOSSIER_UPLOAD_DOCUMENT_TYPES}

FIELD_PATTERNS: dict[str, tuple[str, ...]] = {
    "full_name": (
        r"(?im)(?:cognome\s+e\s+nome|nome\s+e\s+cognome|nominativo)\s*[:\-]?\s*([^\n]{3,100})",
        r"(?im)il\/la\s+sottoscritto\/a\s*([^\n]{3,100})",
    ),
    "birth_place": (
        r"(?im)(?:luogo|comune)\s+di\s+nascita\s*[:\-]?\s*([^\n]{2,80})",
        r"(?is)nato\/?a?\s+a\s+([A-ZÀ-ÿ' ]{2,80}?)(?:\s+(?:provincia|prov\.|il)\b)",
    ),
    "birth_date": (
        rf"(?im)data\s+di\s+nascita\s*[:\-]?\s*{_DATE_CAPTURE}",
        rf"(?is)nato\/?a?.{{0,120}}?\bil\s+{_DATE_CAPTURE}",
    ),
    "birth_province_name": (
        r"(?im)provincia\s+di\s+nascita\s*[:\-]?\s*([^\n]{2,80})",
        r"(?is)nato\/?a?.{0,120}?provincia\s+di\s+([A-ZÀ-ÿ' ]{2,80}?)(?:\s+\(|\s+\bil\b)",
    ),
    "birth_province_code": (
        r"(?im)prov(?:incia)?\.?\s+di\s+nascita\s*[:\-]?\s*\(?([A-Z]{2})\)?",
        r"(?is)nato\/?a?.{0,120}?\(([A-Z]{2})\)\s+\bil\b",
    ),
    "tax_code": (
        r"(?im)(?:codice\s+fiscale|cod\.?\s*fiscale|cf)\s*[:\-]?\s*([A-Z0-9]{16})",
    ),
    "payroll_number": (
        r"(?im)partita\s+stipendiale(?:\s*n\.?)?\s*[:\-]?\s*([A-Z0-9\/\-]{4,20})",
    ),
    "residence_city": (
        r"(?im)(?:residente\s+a|comune\s+di\s+residenza)\s*[:\-]?\s*([^\n,]{2,80})",
    ),
    "residence_province_name": (
        r"(?im)provincia\s+di\s+residenza\s*[:\-]?\s*([^\n]{2,80})",
        r"(?is)residente\s+a.{0,120}?provincia\s+di\s+([A-ZÀ-ÿ' ]{2,80}?)(?:\s+\(|\s+cap\b)",
    ),
    "residence_province_code": (
        r"(?im)prov(?:incia)?\.?\s+di\s+residenza\s*[:\-]?\s*\(?([A-Z]{2})\)?",
        r"(?is)residente\s+a.{0,120}?\(([A-Z]{2})\)\s+cap\b",
    ),
    "residence_cap": (
        r"(?im)\bcap\b\s*[:\-]?\s*(\d{5})",
    ),
    "residence_street": (
        r"(?im)(?:via\/piazza|indirizzo(?:\s+di\s+residenza)?|residente\s+in\s+via)\s*[:\-]?\s*([^\n]{3,120})",
    ),
    "residence_number": (
        r"(?im)\bn\.?\s*[:\-]?\s*([A-Z0-9\/\-]{1,10})",
    ),
    "phone": (
        r"(?im)(?:telefono|tel\.?)\s*[:\-]?\s*(\+?[0-9][0-9\/\-\s]{5,24})",
    ),
    "fax": (
        r"(?im)\bfax\b\s*[:\-]?\s*(\+?[0-9][0-9\/\-\s]{5,24})",
    ),
    "email": (
        r"([A-Za-z0-9._%+\-]+)\s*@\s*([A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
    ),
    "service_office": (
        r"(?im)(?:in\s+servizio\s+presso|ufficio\s+di\s+servizio|servizio\s+presso)\s*[:\-]?\s*([^\n]{3,120})",
    ),
    "employer_entity": (
        r"(?im)(?:ente\s+di\s+appartenenza|amministrazione\s+di\s+appartenenza)\s*[:\-]?\s*([^\n]{3,120})",
    ),
    "lender_name": (
        r"(?im)(?:istituto\s+delegatario|istituto\s+mutuante|societ[aà]\s+finanziaria)\s*[:\-]?\s*([^\n]{3,120})",
        r"(?im)ha\s+chiesto\s+un\s+finanziamento\s+a\s*([^\n]{3,120})",
    ),
    "iban": (
        r"\b(IT\d{2}(?:\s?[A-Z0-9]){23})\b",
    ),
    "loan_amount": (
        rf"(?im)importo\s+finanziamento(?:\s*\(in\s+cifre\))?\s*(?:euro)?\s*[:\-]?\s*{_MONEY_CAPTURE}",
        rf"(?im)(?:importo\s+finanziato|capitale\s+finanziato)\s*[:\-]?\s*{_MONEY_CAPTURE}",
    ),
    "net_disbursed": (
        rf"(?im)(?:importo\s+erogato|netto\s+erogato)\s*[:\-]?\s*{_MONEY_CAPTURE}",
        rf"(?im)capitale\s+netto\s+erogato\s*[:\-]?\s*{_MONEY_CAPTURE}",
    ),
    "total_ceded": (
        rf"(?im)importo\s+globale\s+ceduto\s*[:\-]?\s*{_MONEY_CAPTURE}",
        rf"(?im)montante(?:\s+totale)?\s*[:\-]?\s*{_MONEY_CAPTURE}",
        rf"(?im)totale\s+da\s+rimborsare\s*[:\-]?\s*{_MONEY_CAPTURE}",
        rf"(?im)importo\s+totale\s+dovuto\s*[:\-]?\s*{_MONEY_CAPTURE}",
    ),
    "fees_total": (
        rf"(?im)spese\s+complessive(?:\s+euro)?\s*[:\-]?\s*{_MONEY_CAPTURE}",
        rf"(?im)costi(?!\s+del\s+finanziamento)\s*[:=\-]?\s*{_MONEY_CAPTURE}",
    ),
    "interest_total": (
        rf"(?im)interessi\s+complessivi(?:\s+euro)?\s*[:\-]?\s*{_MONEY_CAPTURE}",
    ),
    "monthly_installment": (
        rf"(?im)(?:rata\s+mensile|rate?\s+mensili\s+di(?:\s+euro)?|importo\s+di\s+euro\s+da\s+trattenere)\s*[:\-]?\s*{_MONEY_CAPTURE}",
        rf"(?im)\brata\b\s*[:\-]?\s*{_MONEY_CAPTURE}",
    ),
    "installment_count": (
        r"(?im)(?:n[°o]\s*rate|numero\s+rate|estinguibile\s+in\s+n[°o])\s*[:\-]?\s*(\d{1,3})",
        r"(?im)durata(?:\s+contrattuale)?\s*[:\-]?\s*(\d{1,3})(?:\s*mesi)?",
        r"(?im)per\s+(\d{2,3})\s*mesi",
    ),
    "tan": (
        r"(?im)\bTAN\b\s*[:\-]?\s*([0-9]+(?:[.,][0-9]{1,3})?)",
    ),
    "taeg": (
        r"(?im)\bTAEG\b\s*[:\-]?\s*([0-9]+(?:[.,][0-9]{1,3})?)",
        r"(?im)\bISC\/TAEG\b\s*[:\-]?\s*([0-9]+(?:[.,][0-9]{1,3})?)",
    ),
    "teg": (
        r"(?im)\bTEG\b\s*[:\-]?\s*([0-9]+(?:[.,][0-9]{1,3})?)",
    ),
    "insurance": (
        r"(?im)(?:garanzia\s+assicurativa|garanzia\s+del\s+prestito)\s*[:\-]?\s*([^\n]{3,120})",
    ),
    "other_financing_lender": (
        r"(?im)(?:estinzione\s+dell[’']eventuale\s+altro\s+finanziamento\s+in\s+corso,\s+contratto\s+con|revoca\s+altro\s+finanziamento.*?contratto\s+con)\s*([^\n]{3,120})",
    ),
    "other_financing_installment": (
        rf"(?im)(?:per\s+euro\s+mensili|in\s+corso\s+di\s+€)\s*[:\-]?\s*{_MONEY_CAPTURE}",
    ),
    "other_financing_expiry": (
        rf"(?im)(?:avente\s+scadenza|scadenza)\s*[:\-]?\s*{_DATE_CAPTURE}",
    ),
}

_DOC_CATEGORY_PRIORITY = {
    "documento_anagrafico": {"anagrafica": 95, "residenza": 92, "contatti": 70, "lavoro": 40, "finanza": 20, "bancario": 10},
    "cedolino_noipa": {"anagrafica": 72, "residenza": 35, "contatti": 25, "lavoro": 96, "finanza": 45, "bancario": 10},
    "contratto_finanziamento": {"anagrafica": 68, "residenza": 64, "contatti": 56, "lavoro": 62, "finanza": 98, "bancario": 40},
    "carta_identita": {"anagrafica": 98, "residenza": 78, "contatti": 10, "lavoro": 5, "finanza": 5, "bancario": 5},
    "tessera_sanitaria": {"anagrafica": 99, "residenza": 5, "contatti": 5, "lavoro": 5, "finanza": 5, "bancario": 5},
    "documento_finanziario": {"anagrafica": 35, "residenza": 15, "contatti": 20, "lavoro": 25, "finanza": 96, "bancario": 85},
    "coordinate_bancarie": {"anagrafica": 10, "residenza": 10, "contatti": 25, "lavoro": 10, "finanza": 45, "bancario": 98},
    "modulo_mef": {"anagrafica": 85, "residenza": 82, "contatti": 65, "lavoro": 88, "finanza": 88, "bancario": 90},
    "documento_generico": {"anagrafica": 45, "residenza": 45, "contatti": 45, "lavoro": 45, "finanza": 45, "bancario": 45},
}

_REJECT_VALUE_PATTERNS = (
    re.compile(r"^[_\-\s./()]+$"),
    re.compile(r"^(?:nome|cognome|prov(?:incia)?|cap|fax|telefono|email|euro|data|firma)$", re.IGNORECASE),
    re.compile(r"^(?:ente di appartenenza|in servizio presso|nato\/?a a|importo erogato:? ?€?)$", re.IGNORECASE),
)


def _normalize_search_text(text: str) -> str:
    text = (text or "").replace("\u00a0", " ").replace("\x00", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def is_supported_dossier_file(filename: str) -> bool:
    return Path(filename or "").suffix.lower() in _DOSSIER_SUPPORTED_EXTENSIONS


def _is_pdf_dossier_file(filename: str) -> bool:
    return Path(filename or "").suffix.lower() in _DOSSIER_PDF_EXTENSIONS


def _is_image_dossier_file(filename: str) -> bool:
    return Path(filename or "").suffix.lower() in _DOSSIER_IMAGE_EXTENSIONS


def find_tesseract_executable() -> str:
    configured = sanitize_pdf_text(os.getenv("QUINTOQUOTE_TESSERACT_PATH") or os.getenv("TESSERACT_CMD"))
    candidates = [configured] if configured else []
    resolved = shutil.which("tesseract")
    if resolved:
        candidates.append(resolved)
    for root in _iter_bundle_runtime_roots():
        candidates.extend(
            [
                str(root / "tesseract.exe"),
                str(root / "tesseract" / "tesseract.exe"),
                str(root / "Tesseract-OCR" / "tesseract.exe"),
                str(root / "vendor" / "tesseract" / "tesseract.exe"),
                str(root / "ocr" / "tesseract.exe"),
            ]
        )
    candidates.extend(_TESSERACT_CANDIDATE_PATHS)
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    return ""


def find_tesseract_data_dir(tesseract_cmd: Optional[str] = None) -> str:
    executable = Path(tesseract_cmd or find_tesseract_executable())
    if not executable.exists():
        return ""
    candidates = [
        executable.parent / "tessdata",
        executable.parent.parent / "tessdata",
        executable.parent / "ocr" / "tessdata",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return ""


def get_ocr_status_message() -> str:
    tesseract_cmd = find_tesseract_executable()
    if tesseract_cmd:
        return "OCR locale attivo: PDF scannerizzati e immagini JPG/PNG vengono letti con Tesseract, senza AI."
    return (
        "OCR pronto ma non attivo: per leggere scansioni e screenshot installa Tesseract OCR "
        "oppure imposta QUINTOQUOTE_TESSERACT_PATH."
    )


def _get_runtime_temp_dir() -> Path:
    _LOCAL_RUNTIME_TEMP_DIR.mkdir(parents=True, exist_ok=True)
    return _LOCAL_RUNTIME_TEMP_DIR


def _new_runtime_temp_path(suffix: str) -> Path:
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    return _get_runtime_temp_dir() / f"qq_{uuid.uuid4().hex}{suffix}"


def _get_dossier_state_dir() -> Path:
    path = _get_runtime_temp_dir() / "dossier_sessions"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _get_dossier_state_path(state_id: str) -> Path:
    safe_state_id = re.sub(r"[^a-zA-Z0-9_-]", "", state_id or "")
    if not safe_state_id:
        raise ValueError("Identificativo dossier non valido.")
    return _get_dossier_state_dir() / f"{safe_state_id}.json"


def load_dossier_state(state_id: str) -> tuple[list[DocumentExtractionResult], dict[str, str]]:
    path = _get_dossier_state_path(state_id)
    if not path.exists():
        return [], {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    results = [_document_result_from_dict(item) for item in list(payload.get("results", []) or [])]
    manual_values = {
        field_def.name: sanitize_pdf_text(dict(payload.get("manual_values", {}) or {}).get(field_def.name, ""))
        for field_def in CASE_FIELD_DEFS
        if sanitize_pdf_text(dict(payload.get("manual_values", {}) or {}).get(field_def.name, ""))
    }
    return results, manual_values


def save_dossier_state(state_id: str, results: list[DocumentExtractionResult], manual_values: dict[str, str]) -> None:
    path = _get_dossier_state_path(state_id)
    payload = {
        "results": [_document_result_to_dict(result) for result in results],
        "manual_values": {
            field_def.name: sanitize_pdf_text(manual_values.get(field_def.name, ""))
            for field_def in CASE_FIELD_DEFS
            if sanitize_pdf_text(manual_values.get(field_def.name, ""))
        },
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def clear_dossier_state(state_id: str) -> None:
    path = _get_dossier_state_path(state_id)
    if path.exists():
        path.unlink(missing_ok=True)


def _search_first_group(pattern: str, text: str, flags: int = re.IGNORECASE | re.MULTILINE | re.DOTALL) -> str:
    match = re.search(pattern, text, flags=flags)
    if not match:
        return ""
    return sanitize_pdf_text(match.group(match.lastindex or 1))


def _normalize_person_name(*parts: str) -> str:
    cleaned_parts = [sanitize_pdf_text(part) for part in parts if sanitize_pdf_text(part)]
    if not cleaned_parts:
        return ""
    normalized: list[str] = []
    for part in cleaned_parts:
        normalized.append(" ".join(token.capitalize() for token in part.split()))
    return " ".join(normalized)


def _looks_like_cedolino_noipa(text: str) -> bool:
    haystack = text.lower()
    anchors = (
        "id cedolino",
        "anagrafica del dipendente",
        "amm.ne appartenenza",
        "ufficio servizio",
        "coord. iban",
    )
    return sum(1 for anchor in anchors if anchor in haystack) >= 4


def _extract_cedolino_noipa_fields(text: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    surname = _search_first_group(r"(?im)^Cognome:\s*([^\n]+)$", text)
    name = _search_first_group(r"(?im)^Nome:\s*([^\n]+)$", text)
    full_name = _normalize_person_name(surname, name)
    if full_name:
        fields["full_name"] = full_name

    mappings = {
        "tax_code": r"(?im)^Codice fiscale:\s*([A-Z0-9]{16})$",
        "birth_date": rf"(?im)^Data di nascita:\s*{_DATE_CAPTURE}$",
        "residence_city": r"(?im)^Domicilio fiscale:\s*([^\n]+)$",
        "payroll_number": r"(?im)^N(?:°|\s)?\s*partita:\s*([A-Z0-9]+)$",
        "employer_entity": r"(?im)^Amm\.ne appartenenza:\s*([^\n]+)$",
        "service_office": r"(?im)^Ufficio servizio:\s*([^\n]+)$",
        "borrower_iban": r"(?im)^Coord\.\s*IBAN:\s*(?:\n\s*)?([A-Z0-9 ]{27,34})$",
        "salary_fifth": rf"(?im)^Quinto cedibile:\s*(?:\n\s*)?{_MONEY_CAPTURE}$",
    }
    for field_name, pattern in mappings.items():
        raw = _search_first_group(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
        cleaned = _clean_extracted_value(field_name, raw)
        if cleaned and not _reject_extracted_value(field_name, cleaned):
            fields[field_name] = cleaned
    return fields


def _looks_like_standard_financing_contract(text: str) -> bool:
    haystack = text.lower()
    anchors = (
        "informazioni europee di base sul credito ai consumatori",
        "prestito con delegazione di pagamento",
        "numero rate mensili da pagare",
        "importo rata mensile",
        "tasso annuo effettivo globale",
    )
    return sum(1 for anchor in anchors if anchor in haystack) >= 4


def _extract_page_rows(page: object, *, min_x: float = 0.0, max_x: float = 10000.0, y_tolerance: float = 1.6) -> list[tuple[float, str]]:
    rows: list[tuple[float, list[tuple[float, str]]]] = []
    for word in page.get_text("words", sort=True):
        x0, y0, x1, y1, text, *_ = word
        if x0 < min_x or x0 >= max_x:
            continue
        cleaned = sanitize_pdf_text(text)
        if not cleaned or cleaned.startswith("[["):
            continue
        if rows and abs(y0 - rows[-1][0]) <= y_tolerance:
            rows[-1][1].append((float(x0), cleaned))
        else:
            rows.append((float(y0), [(float(x0), cleaned)]))
    normalized_rows: list[tuple[float, str]] = []
    for y0, parts in rows:
        line = sanitize_pdf_text(" ".join(part for _, part in sorted(parts, key=lambda item: item[0])))
        if line and not line.startswith("[["):
            normalized_rows.append((y0, line))
    return normalized_rows


def _search_row_group(
    rows: list[tuple[float, str]],
    min_y: float,
    max_y: float,
    pattern: str,
    *,
    exclude_tokens: tuple[str, ...] = (),
) -> str:
    compiled = re.compile(pattern, re.IGNORECASE)
    for y0, line in rows:
        if not (min_y <= y0 <= max_y):
            continue
        lower_line = line.lower()
        if any(token in lower_line for token in exclude_tokens):
            continue
        match = compiled.search(line)
        if match:
            return sanitize_pdf_text(match.group(match.lastindex or 1))
    return ""


def _extract_standard_financing_contract_fields(pdf_bytes: bytes, text: str) -> dict[str, str]:
    fitz_mod = require_pymupdf()
    doc = fitz_mod.open(stream=pdf_bytes, filetype="pdf")
    try:
        fields: dict[str, str] = {}

        summary_page = doc[0]
        summary_text = _normalize_search_text(summary_page.get_text("text"))
        lender_name = _search_first_group(r"(?im)^Finanziatore\s*(?:\n\s*)?([^\n]+)$", summary_text, flags=re.IGNORECASE | re.MULTILINE)
        lender_name = _clean_extracted_value("lender_name", lender_name)
        if lender_name and not _reject_extracted_value("lender_name", lender_name):
            fields["lender_name"] = lender_name

        if doc.page_count >= 5:
            detail_page = doc[4]
            detail_rows = [
                (y0, sanitize_pdf_text(re.sub(r"_+", " ", line)))
                for y0, line in _extract_page_rows(detail_page, max_x=335)
            ]

            first_name = _search_row_group(detail_rows, 112, 117, r"\bNome\b\s+([A-ZÀ-Ÿ' ]{2,79})$")
            last_name = _search_row_group(detail_rows, 124, 129, r"\bCognome\b\s+([A-ZÀ-Ÿ' ]{2,79})$")
            full_name = _normalize_person_name(first_name, last_name)
            if full_name and not _reject_extracted_value("full_name", full_name):
                fields["full_name"] = full_name

            specialized_map = {
                "birth_place": _search_row_group(detail_rows, 137, 140, r"Nata\/o a\s+([A-ZÀ-Ÿ' ]{2,79})\s+Prov"),
                "birth_province_code": _search_row_group(detail_rows, 137, 140, r"\bProv\.?\s+([A-Z]{2})\b"),
                "birth_date": _search_row_group(detail_rows, 148, 150.5, _DATE_CAPTURE),
                "tax_code": _search_row_group(detail_rows, 148, 150.5, r"\b([A-Z0-9]{16})\b"),
                "residence_city": _search_row_group(detail_rows, 160, 161.5, r"Resident\s*e?\s+a\s+([A-ZÀ-Ÿ' ]{2,79})\s+Prov"),
                "residence_province_code": _search_row_group(detail_rows, 160, 161.5, r"\bProv\.?\s+([A-Z]{2})\b"),
                "residence_cap": _search_row_group(detail_rows, 171, 172.5, r"\bCAP\b\s+(\d{5})\b"),
                "residence_street": _search_row_group(detail_rows, 171, 172.5, r"\bVia\b\s+(.+?)\s+\bCAP\b"),
                "employer_entity": _search_row_group(detail_rows, 183, 184.5, r"\bDipendente di\b\s+(.+)$"),
                "phone": _search_row_group(detail_rows, 194, 196, r"\bTel\.?\s+([0-9 ]{5,24})$"),
                "email": _search_row_group(detail_rows, 205, 206.5, r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})"),
                "installment_count": _search_row_group(detail_rows, 239, 242, r"\bNumero rate:\s*(48|60|72|84|96|108|120)\b"),
                "monthly_installment": _search_row_group(detail_rows, 239, 241, rf"\bImporto rata:\s*{_MONEY_CAPTURE}"),
                "taeg": _search_row_group(detail_rows, 261, 263, r"\bTAEG\).*?([0-9]+(?:[.,][0-9]{1,3})?)$"),
                "teg": _search_row_group(detail_rows, 271, 274.5, r"([0-9]+(?:[.,][0-9]{1,3})?)"),
                "total_ceded": _search_row_group(detail_rows, 283, 286, rf"\bImporto Totale dovuto.*?{_MONEY_CAPTURE}"),
                "loan_amount": _search_row_group(detail_rows, 294, 297, rf"\bCapitale finanziato:\s*{_MONEY_CAPTURE}"),
                "tan": _search_row_group(detail_rows, 315, 316, r"\bTAN\).*?([0-9]+(?:[.,][0-9]{1,3})?)$"),
                "interest_total": _search_row_group(detail_rows, 334, 336, rf"\bInteressi complessivi pari a:\s*{_MONEY_CAPTURE}"),
                "net_disbursed": _search_row_group(detail_rows, 364, 370, rf"\bImporto Totale del credito:\s*{_MONEY_CAPTURE}"),
                "borrower_iban": _search_row_group(detail_rows, 461, 462.8, r"\bCodice IBAN\b\s+(IT\d{2}[A-Z0-9]{23})\b"),
            }

            for field_name, raw in specialized_map.items():
                cleaned = _clean_extracted_value(field_name, raw)
                if cleaned and not _reject_extracted_value(field_name, cleaned):
                    fields[field_name] = cleaned

            if fields.get("loan_amount") and not fields.get("net_disbursed"):
                fields["net_disbursed"] = fields["loan_amount"]
            if fields.get("net_disbursed") and not fields.get("loan_amount"):
                fields["loan_amount"] = fields["net_disbursed"]
            if fields.get("birth_place") and not fields.get("birth_province_name"):
                fields["birth_province_name"] = fields["birth_place"]
            if fields.get("residence_city") and not fields.get("residence_province_name"):
                fields["residence_province_name"] = fields["residence_city"]

        blocks = summary_page.get_text("blocks")
        numeric_blocks: list[tuple[float, float, str]] = []
        for block in blocks:
            x0, y0, x1, y1, block_text, *_ = block
            cleaned = sanitize_pdf_text(block_text).replace(" ", "")
            if x0 < 280 or not (430 <= y0 <= 760):
                continue
            if re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3}", cleaned):
                numeric_blocks.append((y0, x0, cleaned))

        def nearest_value(min_y: float, max_y: float) -> str:
            candidates = [(y0, x0, value) for y0, x0, value in numeric_blocks if min_y <= y0 <= max_y]
            if not candidates:
                return ""
            candidates.sort(key=lambda item: (item[0], item[1]))
            return candidates[0][2]

        summary_map = {
            "loan_amount": nearest_value(438, 452),
            "net_disbursed": nearest_value(454, 474),
            "installment_count": nearest_value(548, 562),
            "monthly_installment": nearest_value(562, 580),
            "total_ceded": nearest_value(638, 656),
            "tan": nearest_value(694, 712),
            "taeg": nearest_value(734, 752),
        }
        durata = nearest_value(526, 545)
        if durata and not summary_map["installment_count"]:
            summary_map["installment_count"] = durata

        if summary_map["loan_amount"] and not summary_map["net_disbursed"]:
            summary_map["net_disbursed"] = summary_map["loan_amount"]
        if summary_map["net_disbursed"] and not summary_map["loan_amount"]:
            summary_map["loan_amount"] = summary_map["net_disbursed"]

        for field_name, raw in summary_map.items():
            if field_name in fields and fields[field_name]:
                continue
            cleaned = _clean_extracted_value(field_name, raw)
            if cleaned and not _reject_extracted_value(field_name, cleaned):
                fields[field_name] = cleaned
        return fields
    finally:
        doc.close()


def _looks_like_carta_identita(text: str) -> bool:
    haystack = text.lower()
    anchors = (
        "carta d'identità",
        "carta di identita",
        "repubblica italiana",
        "luogo di nascita",
        "cognome",
        "nome",
    )
    return sum(1 for anchor in anchors if anchor in haystack) >= 2 and ("cognome" in haystack or "nome" in haystack)


def _extract_carta_identita_fields(text: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    surname = _search_first_group(r"(?im)^\s*cognome\s*[:\-]?\s*([^\n]{2,80})$", text)
    name = _search_first_group(r"(?im)^\s*nome\s*[:\-]?\s*([^\n]{2,80})$", text)
    full_name = _normalize_person_name(surname, name)
    if full_name and not _reject_extracted_value("full_name", full_name):
        fields["full_name"] = full_name

    for field_name in (
        "birth_place",
        "birth_date",
        "birth_province_name",
        "birth_province_code",
        "tax_code",
        "residence_city",
        "residence_province_name",
        "residence_province_code",
        "residence_cap",
        "residence_street",
    ):
        value = _extract_first_pattern_value(field_name, text)
        if value:
            fields[field_name] = value
    return fields


def _looks_like_tessera_sanitaria(text: str) -> bool:
    haystack = text.lower()
    anchors = (
        "tessera sanitaria",
        "codice fiscale",
        "tessera europea",
        "cognome",
        "nome",
    )
    return sum(1 for anchor in anchors if anchor in haystack) >= 2 and ("codice fiscale" in haystack or re.search(r"\b[A-Z0-9]{16}\b", text) is not None)


def _extract_tessera_sanitaria_fields(text: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    surname = _search_first_group(r"(?im)^\s*cognome\s*[:\-]?\s*([^\n]{2,80})$", text)
    name = _search_first_group(r"(?im)^\s*nome\s*[:\-]?\s*([^\n]{2,80})$", text)
    full_name = _normalize_person_name(surname, name)
    if full_name and not _reject_extracted_value("full_name", full_name):
        fields["full_name"] = full_name

    for field_name in ("tax_code", "birth_date", "birth_place"):
        value = _extract_first_pattern_value(field_name, text)
        if value:
            fields[field_name] = value
    return fields


def _extract_specialized_document_fields(
    filename: str,
    pdf_bytes: bytes,
    normalized_text: str,
    forced_document_key: Optional[str] = None,
) -> tuple[Optional[str], Optional[str], dict[str, str]]:
    if forced_document_key == "cedolino_noipa":
        return forced_document_key, _DOCUMENT_TYPE_DEF_MAP[forced_document_key].label, _extract_cedolino_noipa_fields(normalized_text)
    if forced_document_key == "contratto_finanziamento" and _is_pdf_dossier_file(filename):
        return forced_document_key, _DOCUMENT_TYPE_DEF_MAP[forced_document_key].label, _extract_standard_financing_contract_fields(pdf_bytes, normalized_text)
    if forced_document_key == "carta_identita":
        return forced_document_key, _DOCUMENT_TYPE_DEF_MAP[forced_document_key].label, _extract_carta_identita_fields(normalized_text)
    if forced_document_key == "tessera_sanitaria":
        return forced_document_key, _DOCUMENT_TYPE_DEF_MAP[forced_document_key].label, _extract_tessera_sanitaria_fields(normalized_text)

    if _looks_like_cedolino_noipa(normalized_text):
        return "cedolino_noipa", _DOCUMENT_TYPE_DEF_MAP["cedolino_noipa"].label, _extract_cedolino_noipa_fields(normalized_text)
    if _is_pdf_dossier_file(filename) and _looks_like_standard_financing_contract(normalized_text):
        return "contratto_finanziamento", _DOCUMENT_TYPE_DEF_MAP["contratto_finanziamento"].label, _extract_standard_financing_contract_fields(pdf_bytes, normalized_text)
    if _looks_like_carta_identita(normalized_text):
        return "carta_identita", _DOCUMENT_TYPE_DEF_MAP["carta_identita"].label, _extract_carta_identita_fields(normalized_text)
    if _looks_like_tessera_sanitaria(normalized_text):
        return "tessera_sanitaria", _DOCUMENT_TYPE_DEF_MAP["tessera_sanitaria"].label, _extract_tessera_sanitaria_fields(normalized_text)
    return None, None, {}


def _clean_extracted_value(field_name: str, value: str) -> str:
    value = (value or "").strip(" \t\r\n:;-")
    value = re.sub(r"\s{2,}", " ", value)
    if field_name == "email" and "@" not in value:
        parts = re.split(r"\s*@\s*", value)
        if len(parts) == 2:
            value = f"{parts[0]}@{parts[1]}"
    if field_name in {"tax_code", "iban", "borrower_iban", "birth_province_code", "residence_province_code"}:
        value = value.replace(" ", "").upper()
    if field_name in {"tan", "taeg", "teg"}:
        value = value.replace(" ", "").replace("%", "").replace(".", ",")
    if field_name in {
        "loan_amount",
        "net_disbursed",
        "total_ceded",
        "fees_total",
        "interest_total",
        "monthly_installment",
        "salary_fifth",
        "other_financing_installment",
    }:
        value = value.replace("EUR", "").replace("€", "").strip()
    if field_name.endswith("_date") or field_name == "other_financing_expiry":
        value = value.replace(".", "/").replace("-", "/")
        parts = value.split("/")
        if len(parts) == 3 and all(part.isdigit() for part in parts):
            gg, mm, aa = parts
            if len(gg) == 1:
                gg = gg.zfill(2)
            if len(mm) == 1:
                mm = mm.zfill(2)
            if len(aa) == 2:
                aa = ("20" if int(aa) < 40 else "19") + aa
            value = f"{gg}/{mm}/{aa}"
    return value.strip()


def _reject_extracted_value(field_name: str, value: str) -> bool:
    if not value:
        return True
    for pattern in _REJECT_VALUE_PATTERNS:
        if pattern.fullmatch(value):
            return True
    if value.count("_") >= max(3, len(value) // 2):
        return True
    if field_name in {"tax_code"} and not re.fullmatch(r"[A-Z0-9]{16}", value):
        return True
    if field_name in {"birth_province_code", "residence_province_code"} and not re.fullmatch(r"[A-Z]{2}", value):
        return True
    if field_name == "residence_number" and not re.search(r"\d", value):
        return True
    if field_name in {"iban", "borrower_iban"} and not re.fullmatch(r"IT\d{2}[A-Z0-9]{23}", value):
        return True
    if field_name in {"loan_amount", "net_disbursed", "total_ceded", "fees_total", "interest_total", "monthly_installment", "salary_fifth", "other_financing_installment"} and not re.search(r"\d", value):
        return True
    if field_name in {"tan", "taeg", "teg"} and not re.fullmatch(r"\d+(?:,\d{1,3})?", value):
        return True
    if field_name in {"birth_date", "other_financing_expiry"} and not re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        return True
    if field_name == "email" and "@" not in value:
        return True
    lower_value = value.lower()
    if field_name == "full_name" and any(token in lower_value for token in ("nato", "cognome", "nome", "codice fiscale")):
        return True
    if field_name in {"lender_name", "employer_entity", "service_office", "insurance", "other_financing_lender"}:
        noisy_tokens = ("sarà cura", "autorizza", "che ha compilato", "importo erogato", "revoca altro finanziamento")
        if any(token in lower_value for token in noisy_tokens):
            return True
    return False


def _extract_first_pattern_value(field_name: str, text: str) -> str:
    for pattern in FIELD_PATTERNS.get(field_name, ()):
        match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if not match:
            continue
        if field_name == "email" and match.lastindex == 2:
            raw_value = f"{match.group(1)}@{match.group(2)}"
        else:
            raw_value = match.group(match.lastindex or 1)
        cleaned = _clean_extracted_value(field_name, raw_value)
        if not _reject_extracted_value(field_name, cleaned):
            return cleaned
    return ""


def _extract_rate_duration_expression_fields(text: str) -> dict[str, str]:
    rate_capture = r"(\d{1,3}(?:[.\s]\d{3})*(?:,\d{2,3})|\d+(?:,\d{2,3})?)"
    duration_capture = r"(48|60|72|84|96|108|120)"
    patterns = (
        rf"(?im)\b{rate_capture}\s*(?:euro|eur|€)?\s*[x×]\s*{duration_capture}\s*(?:mesi|rate)?\b",
        rf"(?im)\b(?:rata(?:\s+mensile)?\s*(?:di)?\s*)?{rate_capture}\s*(?:euro|eur|€)?\s*[x×]\s*{duration_capture}\s*(?:mesi|rate)?\b",
        rf"(?im)\b{rate_capture}\s*(?:euro|eur|€)?\s*per\s*{duration_capture}\s*(?:mesi|rate)\b",
    )
    extracted: dict[str, str] = {}
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
        if not match:
            continue
        rate_value = _clean_extracted_value("monthly_installment", match.group(1))
        duration_value = _clean_extracted_value("installment_count", match.group(2))
        try:
            rate_value = _format_money_it(round(parse_decimal_loose(rate_value), 2))
        except Exception:
            pass
        try:
            duration_int = int(re.sub(r"[^\d]", "", duration_value))
        except Exception:
            duration_int = 0
        if rate_value and duration_int in _ALLOWED_INSTALLMENT_COUNTS:
            extracted["monthly_installment"] = rate_value
            extracted["installment_count"] = str(duration_int)
            break
    return extracted


def _run_tesseract_ocr(input_path: Path) -> str:
    tesseract_cmd = find_tesseract_executable()
    if not tesseract_cmd:
        raise RuntimeError(
            "OCR non disponibile: installa Tesseract OCR oppure imposta QUINTOQUOTE_TESSERACT_PATH."
        )
    tessdata_dir = find_tesseract_data_dir(tesseract_cmd)

    attempts = (
        ("ita+eng", "6"),
        ("ita+eng", "11"),
        ("ita", "6"),
        ("eng", "6"),
        ("", "6"),
    )
    errors: list[str] = []
    for language, psm in attempts:
        command = [tesseract_cmd, str(input_path), "stdout", "--psm", psm, "--oem", "1"]
        if tessdata_dir:
            command.extend(["--tessdata-dir", tessdata_dir])
        if language:
            command.extend(["-l", language])
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
        text = _normalize_search_text(completed.stdout)
        if completed.returncode == 0 and text:
            return text
        error_text = sanitize_pdf_text(completed.stderr)
        if error_text:
            errors.append(error_text)

    if errors:
        raise RuntimeError(f"OCR non riuscito: {errors[-1]}")
    raise RuntimeError("OCR non riuscito: nessun testo riconosciuto dal file.")


def _safe_ocr_image_text_from_bytes(data: bytes, suffix: str) -> str:
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    input_path = _new_runtime_temp_path(suffix)
    try:
        input_path.write_bytes(data)
        return _run_tesseract_ocr(input_path)
    finally:
        if input_path.exists():
            input_path.unlink(missing_ok=True)


def _safe_ocr_pdf_text_from_bytes(data: bytes) -> tuple[str, int]:
    fitz_mod = require_pymupdf()
    doc = fitz_mod.open(stream=data, filetype="pdf")
    try:
        texts: list[str] = []
        temp_paths: list[Path] = []
        try:
            for index, page in enumerate(doc, start=1):
                pixmap = page.get_pixmap(matrix=fitz_mod.Matrix(2.5, 2.5), alpha=False)
                image_path = _new_runtime_temp_path(".png")
                temp_paths.append(image_path)
                pixmap.save(str(image_path))
                page_text = _run_tesseract_ocr(image_path)
                if page_text:
                    texts.append(page_text)
        finally:
            for temp_path in temp_paths:
                if temp_path.exists():
                    temp_path.unlink(missing_ok=True)
        return "\n\n".join(texts), doc.page_count
    finally:
        doc.close()


def _extract_text_from_supported_document(filename: str, data: bytes) -> tuple[str, int, list[str]]:
    warnings: list[str] = []
    if _is_pdf_dossier_file(filename):
        text, page_count = _safe_pdf_text_from_bytes(data)
        normalized = _normalize_search_text(text)
        if len(normalized) < _OCR_TRIGGER_TEXT_LENGTH:
            try:
                ocr_text, _ = _safe_ocr_pdf_text_from_bytes(data)
                ocr_normalized = _normalize_search_text(ocr_text)
                if len(ocr_normalized) > len(normalized):
                    text = "\n\n".join(part for part in (text, ocr_text) if _normalize_search_text(part))
                    warnings.append("OCR locale applicato: PDF scannerizzato o poco testuale.")
                    normalized = _normalize_search_text(text)
            except RuntimeError as exc:
                warnings.append(str(exc))
        if len(normalized) < _LOW_TEXT_WARNING_LENGTH:
            warnings.append("Testo disponibile molto limitato: controlla i dati estratti e integra i campi mancanti.")
        return text, page_count, warnings

    if _is_image_dossier_file(filename):
        try:
            text = _safe_ocr_image_text_from_bytes(data, Path(filename).suffix.lower())
            warnings.append("OCR locale applicato: immagine analizzata con Tesseract.")
            normalized = _normalize_search_text(text)
            if len(normalized) < _LOW_TEXT_WARNING_LENGTH:
                warnings.append("Testo OCR limitato: verifica i valori prima di usare il prefill.")
            return text, 1, warnings
        except RuntimeError as exc:
            warnings.append(str(exc))
            warnings.append("Immagine non analizzata: senza OCR installato lo screenshot non e leggibile.")
            return "", 1, warnings

    raise ValueError("Formato non supportato nel Dossier.")


def _split_email(value: str) -> tuple[str, str]:
    value = (value or "").strip()
    if "@" not in value:
        return "", ""
    local, domain = value.split("@", 1)
    return local.strip(), domain.strip()


def _safe_pdf_text_from_bytes(data: bytes) -> tuple[str, int]:
    fitz_mod = require_pymupdf()
    doc = fitz_mod.open(stream=data, filetype="pdf")
    try:
        texts: list[str] = []
        for page in doc:
            texts.append(page.get_text("text"))
        return "\n".join(texts), doc.page_count
    finally:
        doc.close()


def _safe_pdf_widget_values_from_bytes(data: bytes) -> dict[str, str]:
    fitz_mod = require_pymupdf()
    doc = fitz_mod.open(stream=data, filetype="pdf")
    try:
        values: dict[str, str] = {}
        for page in doc:
            for widget in page.widgets() or []:
                name = sanitize_pdf_text(getattr(widget, "field_name", ""))
                value = sanitize_pdf_text(getattr(widget, "field_value", ""))
                if name and value:
                    values[name] = value
        return values
    finally:
        doc.close()


def _case_data_from_known_template_values(spec_key: str, values: dict[str, str]) -> dict[str, str]:
    if spec_key == ALLEGATO_E_SPEC.key:
        email_local = values.get("email_local_part", "")
        email_domain = values.get("email_domain", "")
        email = ""
        if email_local and email_domain:
            email = f"{email_local}@{email_domain}"
        return {
            "full_name": values.get("istante_nome_completo", ""),
            "birth_place": values.get("nascita_comune", ""),
            "birth_province_name": values.get("nascita_provincia", ""),
            "birth_province_code": values.get("nascita_sigla_provincia", ""),
            "birth_date": values.get("nascita_data", ""),
            "tax_code": values.get("codice_fiscale", ""),
            "payroll_number": values.get("partita_stipendiale", ""),
            "residence_city": values.get("residenza_comune", ""),
            "residence_province_name": values.get("residenza_provincia", ""),
            "residence_province_code": values.get("residenza_sigla_provincia", ""),
            "residence_cap": values.get("residenza_cap", ""),
            "residence_street": values.get("residenza_via", ""),
            "residence_number": values.get("residenza_numero", ""),
            "phone": values.get("telefono", ""),
            "fax": values.get("fax", ""),
            "email": email,
            "lender_name": values.get("istituto_delegatario", ""),
            "iban": values.get("iban_istituto_delegatario", ""),
            "loan_amount": values.get("importo_finanziamento_cifre", ""),
            "total_ceded": values.get("importo_globale_ceduto_cifre", ""),
            "fees_total": values.get("spese_complessive_cifre", ""),
            "interest_total": values.get("interessi_complessivi_cifre", ""),
            "monthly_installment": values.get("importo_trattenuta_mensile", values.get("importo_rata_estinzione", "")),
            "installment_count": values.get("numero_rate_da_estinguere", ""),
            "tan": values.get("tan", ""),
            "taeg": values.get("taeg", ""),
            "teg": values.get("teg", ""),
            "insurance": values.get("garanzia_prestito", ""),
            "other_financing_lender": values.get("estinzione_altro_finanziamento_istituto", ""),
            "other_financing_installment": values.get("estinzione_altro_finanziamento_rata", ""),
            "other_financing_expiry": values.get("estinzione_altro_finanziamento_scadenza", ""),
        }
    if spec_key == ALLEGATO_C_SPEC.key:
        return {
            "net_disbursed": values.get("importo_erogato", ""),
            "total_ceded": values.get("importo_globale_ceduto", ""),
            "fees_total": values.get("spese_complessive", ""),
            "interest_total": values.get("interessi_complessivi", ""),
            "tan": values.get("tan", ""),
            "taeg": values.get("isc_taeg", ""),
            "installment_count": values.get("numero_rate_estinguibili", ""),
            "monthly_installment": values.get("importo_rata_estinzione", ""),
            "insurance": values.get("garanzia_assicurativa", ""),
            "other_financing_installment": values.get("revoca_finanziamento_importo", ""),
            "other_financing_expiry": values.get("revoca_finanziamento_scadenza", ""),
            "other_financing_lender": values.get("revoca_finanziamento_contratto", ""),
            "full_name": values.get("dipendente_nome_cognome", ""),
            "birth_place": values.get("dipendente_nascita_luogo", ""),
            "birth_province_code": values.get("dipendente_nascita_provincia", ""),
            "birth_date": values.get("dipendente_nascita_data", ""),
            "tax_code": values.get("dipendente_codice_fiscale", ""),
            "service_office": values.get("dipendente_in_servizio_presso", ""),
            "employer_entity": values.get("dipendente_ente_appartenenza", ""),
        }
    return {}


def _extract_known_template_fields_from_widgets(widget_values: dict[str, str]) -> tuple[Optional[str], Optional[str], dict[str, str]]:
    best_spec: Optional[PdfTemplateSpec] = None
    best_hits = 0
    best_logical_values: dict[str, str] = {}
    for spec in (ALLEGATO_E_SPEC, ALLEGATO_C_SPEC):
        widget_to_name = {field.widget: field.name for field in iter_pdf_fields(spec)}
        logical_values = {
            widget_to_name[widget_name]: sanitize_pdf_text(widget_value)
            for widget_name, widget_value in widget_values.items()
            if widget_name in widget_to_name and sanitize_pdf_text(widget_value)
        }
        hits = len(logical_values)
        if hits > best_hits:
            best_spec = spec
            best_hits = hits
            best_logical_values = logical_values
    if best_spec is None or best_hits < 3:
        return None, None, {}
    case_data = {
        name: value
        for name, value in _case_data_from_known_template_values(best_spec.key, best_logical_values).items()
        if sanitize_pdf_text(value)
    }
    return best_spec.key, best_spec.label, case_data


def _classify_document_text(text: str, filename: str) -> tuple[str, str, int]:
    haystack = f"{filename}\n{text}".lower()
    best_key = "documento_generico"
    best_label = "Documento generico"
    best_hits = 0
    for doc_type in DOCUMENT_TYPE_DEFS:
        hits = sum(1 for keyword in doc_type.keywords if keyword in haystack)
        if hits > best_hits:
            best_key = doc_type.key
            best_label = doc_type.label
            best_hits = hits
    return best_key, best_label, best_hits


def _count_document_keyword_hits(text: str, filename: str, document_key: str) -> int:
    doc_type = _DOCUMENT_TYPE_DEF_MAP.get(document_key)
    if doc_type is None:
        return 0
    haystack = f"{filename}\n{text}".lower()
    return sum(1 for keyword in doc_type.keywords if keyword in haystack)


def extract_document_result(filename: str, raw_bytes: bytes, expected_document_key: str = "") -> DocumentExtractionResult:
    text, page_count, warnings = _extract_text_from_supported_document(filename, raw_bytes)
    normalized = _normalize_search_text(text)
    detected_key, detected_label, detected_hits = _classify_document_text(normalized, filename)
    expected_document_key = sanitize_pdf_text(expected_document_key)
    expected_document_type = _DOCUMENT_TYPE_DEF_MAP.get(expected_document_key)
    document_key = expected_document_type.key if expected_document_type else detected_key
    document_label = expected_document_type.label if expected_document_type else detected_label
    keyword_hits = _count_document_keyword_hits(normalized, filename, document_key) if expected_document_type else detected_hits
    extracted: dict[str, str] = {}
    widget_key = widget_label = special_key = special_label = None
    widget_fields: dict[str, str] = {}
    special_fields: dict[str, str] = {}
    if _is_pdf_dossier_file(filename) and not expected_document_type:
        widget_values = _safe_pdf_widget_values_from_bytes(raw_bytes)
        widget_key, widget_label, widget_fields = _extract_known_template_fields_from_widgets(widget_values)
    special_key, special_label, special_fields = _extract_specialized_document_fields(
        filename,
        raw_bytes,
        normalized,
        forced_document_key=document_key if expected_document_type else None,
    )

    if widget_fields:
        extracted.update(widget_fields)
        if not expected_document_type:
            document_key = widget_key or document_key
            document_label = widget_label or document_label
    if special_fields:
        extracted.update(special_fields)
        if not expected_document_type:
            document_key = special_key or document_key
            document_label = special_label or document_label

    expression_fields = _extract_rate_duration_expression_fields(normalized)
    if expression_fields:
        for field_name, value in expression_fields.items():
            extracted.setdefault(field_name, value)
        if not expected_document_type and document_key == "documento_generico":
            document_key = "documento_finanziario"
            document_label = "Documento finanziario"

    for field in CASE_FIELD_DEFS:
        if field.name in extracted:
            continue
        value = _extract_first_pattern_value(field.name, normalized)
        if value:
            extracted[field.name] = value

    if expected_document_type and page_count < 5 and document_key == "contratto_finanziamento":
        warnings.append("Per i contratti completi servono anche le pagine interne: su alcuni layout i dati principali sono attorno a pagina 5.")

    if expected_document_type and detected_key not in {document_key, "documento_generico"} and detected_hits >= max(2, keyword_hits + 1):
        warnings.append(
            f"Il file è stato caricato come {document_label}, ma il contenuto assomiglia di più a {detected_label.lower()}."
        )
    elif expected_document_type and keyword_hits == 0 and len(extracted) < 2:
        warnings.append(
            f"Il file è stato trattato come {document_label}, ma contiene pochi indicatori tipici di questa tipologia."
        )

    if not extracted:
        warnings.append("Nessun campo utile estratto con le regole disponibili.")

    return DocumentExtractionResult(
        filename=filename,
        document_key=document_key,
        document_label=document_label,
        page_count=page_count,
        text_length=len(normalized),
        keyword_hits=keyword_hits,
        extracted_fields=extracted,
        warnings=warnings,
    )


def aggregate_document_results(results: list[DocumentExtractionResult]) -> list[AggregatedCaseField]:
    best: dict[str, tuple[float, AggregatedCaseField]] = {}
    for result in results:
        priorities = _DOC_CATEGORY_PRIORITY.get(result.document_key, _DOC_CATEGORY_PRIORITY["documento_generico"])
        for field_name, value in result.extracted_fields.items():
            field_def = _CASE_FIELD_DEF_MAP[field_name]
            score = priorities.get(field_def.category, 45) + min(len(value), 40) / 100.0
            candidate = AggregatedCaseField(
                name=field_name,
                label=field_def.label,
                value=value,
                source_filename=result.filename,
                source_document_label=result.document_label,
            )
            current = best.get(field_name)
            if current is None or score > current[0]:
                best[field_name] = (score, candidate)
    aggregated: list[AggregatedCaseField] = []
    for field_def in CASE_FIELD_DEFS:
        item = best.get(field_def.name)
        if item:
            aggregated.append(item[1])
    return infer_aggregated_fields(aggregated)


def aggregated_fields_to_dict(fields: list[AggregatedCaseField]) -> dict[str, str]:
    return {field.name: field.value for field in fields}


def extract_manual_case_values(raw: dict[str, object]) -> dict[str, str]:
    return {
        field_def.name: sanitize_pdf_text(raw.get(field_def.name, ""))
        for field_def in CASE_FIELD_DEFS
        if sanitize_pdf_text(raw.get(field_def.name, ""))
    }


def merge_reviewed_case_fields(
    results: list[DocumentExtractionResult],
    manual_values: Optional[dict[str, str]] = None,
    aggregated_fields: Optional[list[AggregatedCaseField]] = None,
) -> list[AggregatedCaseField]:
    by_name = {
        field.name: field
        for field in (aggregated_fields if aggregated_fields is not None else aggregate_document_results(results))
    }
    for field_def in CASE_FIELD_DEFS:
        manual_value = sanitize_pdf_text((manual_values or {}).get(field_def.name, ""))
        if not manual_value:
            continue
        by_name[field_def.name] = AggregatedCaseField(
            name=field_def.name,
            label=field_def.label,
            value=manual_value,
            source_filename="Revisione utente",
            source_document_label="Dato confermato/modificato",
        )
    return infer_aggregated_fields([by_name[field_def.name] for field_def in CASE_FIELD_DEFS if field_def.name in by_name])


def _parse_decimal_maybe(raw: str) -> Optional[float]:
    try:
        return parse_decimal_loose(raw)
    except Exception:
        return None


def _format_money_it(amount: float) -> str:
    return f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def infer_aggregated_fields(fields: list[AggregatedCaseField]) -> list[AggregatedCaseField]:
    by_name = {field.name: field for field in fields}

    def add_inferred(name: str, value: str, reason: str) -> None:
        if name in by_name or not value:
            return
        by_name[name] = AggregatedCaseField(
            name=name,
            label=_CASE_FIELD_DEF_MAP[name].label,
            value=value,
            source_filename="Sistema",
            source_document_label=reason,
        )

    installment = _parse_decimal_maybe(by_name["monthly_installment"].value) if "monthly_installment" in by_name else None
    total_ceded = _parse_decimal_maybe(by_name["total_ceded"].value) if "total_ceded" in by_name else None
    loan_amount = _parse_decimal_maybe(by_name["loan_amount"].value) if "loan_amount" in by_name else None
    net_disbursed = _parse_decimal_maybe(by_name["net_disbursed"].value) if "net_disbursed" in by_name else None

    count = None
    if "installment_count" in by_name:
        try:
            count = int(re.sub(r"[^\d]", "", by_name["installment_count"].value))
        except Exception:
            count = None

    if installment is not None and count in _ALLOWED_INSTALLMENT_COUNTS and "total_ceded" not in by_name:
        add_inferred("total_ceded", _format_money_it(round(installment * count, 2)), "Derivato da rata x durata")

    if installment is not None and total_ceded is not None and "installment_count" not in by_name and installment > 0:
        derived_count = round(total_ceded / installment)
        if derived_count in _ALLOWED_INSTALLMENT_COUNTS and abs((installment * derived_count) - total_ceded) <= 0.5:
            add_inferred("installment_count", str(int(derived_count)), "Derivato da montante / rata")

    if total_ceded is not None and count in _ALLOWED_INSTALLMENT_COUNTS and "monthly_installment" not in by_name and count > 0:
        add_inferred("monthly_installment", _format_money_it(round(total_ceded / count, 2)), "Derivato da montante / durata")

    if loan_amount is not None and "net_disbursed" not in by_name:
        add_inferred("net_disbursed", _format_money_it(round(loan_amount, 2)), "Derivato da importo finanziamento")

    if net_disbursed is not None and "loan_amount" not in by_name:
        add_inferred("loan_amount", _format_money_it(round(net_disbursed, 2)), "Derivato da importo erogato")

    return [by_name[field_def.name] for field_def in CASE_FIELD_DEFS if field_def.name in by_name]


def build_review_sections(fields: list[AggregatedCaseField]) -> list[dict]:
    values = aggregated_fields_to_dict(fields)
    sources = {field.name: field for field in fields}
    sections: list[dict] = []
    for category in ("anagrafica", "residenza", "contatti", "lavoro", "finanza", "bancario"):
        section_fields = []
        for field_def in CASE_FIELD_DEFS:
            if field_def.category != category:
                continue
            source = sources.get(field_def.name)
            section_fields.append({
                "name": field_def.name,
                "label": field_def.label,
                "placeholder": field_def.placeholder,
                "value": values.get(field_def.name, ""),
                "source_filename": source.source_filename if source else "",
                "source_label": source.source_document_label if source else "",
            })
        sections.append({
            "title": _CASE_CATEGORY_LABELS.get(category, category.title()),
            "fields": section_fields,
        })
    return sections


def build_prefill_for_template(spec_key: str, case_data: dict[str, str]) -> dict[str, str]:
    spec = get_pdf_template_spec(spec_key)
    data = {k: sanitize_pdf_text(v) for k, v in case_data.items() if sanitize_pdf_text(v)}
    email_local, email_domain = _split_email(data.get("email", ""))

    if spec.key == ALLEGATO_E_SPEC.key:
        mapped = {
            "istante_nome_completo": data.get("full_name", ""),
            "nascita_comune": data.get("birth_place", ""),
            "nascita_provincia": data.get("birth_province_name", data.get("birth_province_code", "")),
            "nascita_sigla_provincia": data.get("birth_province_code", ""),
            "nascita_data": data.get("birth_date", ""),
            "codice_fiscale": data.get("tax_code", ""),
            "partita_stipendiale": data.get("payroll_number", ""),
            "residenza_comune": data.get("residence_city", ""),
            "residenza_provincia": data.get("residence_province_name", data.get("residence_province_code", "")),
            "residenza_sigla_provincia": data.get("residence_province_code", ""),
            "residenza_cap": data.get("residence_cap", ""),
            "residenza_via": data.get("residence_street", ""),
            "residenza_numero": data.get("residence_number", ""),
            "telefono": data.get("phone", ""),
            "fax": data.get("fax", ""),
            "email_local_part": email_local,
            "email_domain": email_domain,
            "istituto_delegatario": data.get("lender_name", ""),
            "importo_trattenuta_mensile": data.get("monthly_installment", ""),
            "iban_istituto_delegatario": data.get("iban", ""),
            "importo_finanziamento_cifre": data.get("loan_amount", data.get("net_disbursed", "")),
            "importo_globale_ceduto_cifre": data.get("total_ceded", ""),
            "spese_complessive_cifre": data.get("fees_total", ""),
            "interessi_complessivi_cifre": data.get("interest_total", ""),
            "tan": data.get("tan", ""),
            "taeg": data.get("taeg", ""),
            "teg": data.get("teg", ""),
            "numero_rate_da_estinguere": data.get("installment_count", ""),
            "importo_rata_estinzione": data.get("monthly_installment", ""),
            "garanzia_prestito": data.get("insurance", ""),
            "estinzione_altro_finanziamento_istituto": data.get("other_financing_lender", ""),
            "estinzione_altro_finanziamento_rata": data.get("other_financing_installment", ""),
            "estinzione_altro_finanziamento_scadenza": data.get("other_financing_expiry", ""),
        }
    elif spec.key == ALLEGATO_C_SPEC.key:
        mapped = {
            "importo_erogato": data.get("net_disbursed", data.get("loan_amount", "")),
            "importo_globale_ceduto": data.get("total_ceded", ""),
            "spese_complessive": data.get("fees_total", ""),
            "interessi_complessivi": data.get("interest_total", ""),
            "tan": data.get("tan", ""),
            "isc_taeg": data.get("taeg", ""),
            "numero_rate_estinguibili": data.get("installment_count", ""),
            "importo_rata_estinzione": data.get("monthly_installment", ""),
            "garanzia_assicurativa": data.get("insurance", ""),
            "revoca_finanziamento_importo": data.get("other_financing_installment", ""),
            "revoca_finanziamento_scadenza": data.get("other_financing_expiry", ""),
            "revoca_finanziamento_contratto": data.get("other_financing_lender", ""),
            "dipendente_nome_cognome": data.get("full_name", ""),
            "dipendente_nascita_luogo": data.get("birth_place", ""),
            "dipendente_nascita_provincia": data.get("birth_province_code", data.get("birth_province_name", "")),
            "dipendente_nascita_data": data.get("birth_date", ""),
            "dipendente_codice_fiscale": data.get("tax_code", ""),
            "dipendente_in_servizio_presso": data.get("service_office", ""),
            "dipendente_ente_appartenenza": data.get("employer_entity", ""),
        }
    else:
        mapped = {}

    clean = default_pdf_form_values(spec)
    clean.update({k: v for k, v in mapped.items() if v})
    return clean

# =========================
#  WEB (OPZIONALE) - localhost
# =========================

def run_web(
    out_dir: Path,
    host: str = "127.0.0.1",
    port: int = 5000,
    open_browser: bool = False,
    show_start_hint: bool = False,
):
    try:
        from flask import Flask, request, render_template_string, send_file, redirect, url_for, jsonify, session
    except Exception as e:
        print("Flask non installato. Per usare --web fai: pip install flask")
        raise
    import json as _json

    app = Flask(__name__)
    app.secret_key = os.environ.get("QUINTOQUOTE_SESSION_SECRET") or f"quintoquote-{uuid.uuid4().hex}"

    # ─── CSS Design System ───
    CSS = """
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    :root {
      --bg-deep: #050a18;
      --bg-card: rgba(15, 23, 42, 0.65);
      --bg-card-hover: rgba(15, 23, 42, 0.80);
      --glass-border: rgba(56, 189, 248, 0.15);
      --glass-border-hover: rgba(56, 189, 248, 0.35);
      --accent: #0ea5e9;
      --accent-light: #38bdf8;
      --accent-glow: rgba(14, 165, 233, 0.25);
      --text-primary: #f1f5f9;
      --text-secondary: #94a3b8;
      --text-muted: #64748b;
      --success: #10b981;
      --error: #ef4444;
      --warning: #f59e0b;
      --input-bg: rgba(30, 41, 59, 0.6);
      --input-border: rgba(100, 116, 139, 0.3);
      --input-focus: rgba(14, 165, 233, 0.5);
      --radius: 14px;
      --radius-sm: 10px;
      --shadow-lg: 0 25px 50px rgba(0,0,0,0.4);
      --transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
      font-family: 'Inter', system-ui, -apple-system, sans-serif;
      background: var(--bg-deep);
      color: var(--text-primary);
      min-height: 100vh;
      overflow-x: hidden;
    }

    /* Animated background */
    body::before {
      content: '';
      position: fixed;
      top: -50%; left: -50%;
      width: 200%; height: 200%;
      background: radial-gradient(ellipse at 20% 50%, rgba(14, 165, 233, 0.08) 0%, transparent 50%),
                  radial-gradient(ellipse at 80% 20%, rgba(56, 189, 248, 0.06) 0%, transparent 50%),
                  radial-gradient(ellipse at 50% 80%, rgba(6, 182, 212, 0.04) 0%, transparent 50%);
      animation: bgFloat 20s ease-in-out infinite;
      z-index: -1;
    }

    @keyframes bgFloat {
      0%, 100% { transform: translate(0, 0) rotate(0deg); }
      33% { transform: translate(2%, -2%) rotate(1deg); }
      66% { transform: translate(-1%, 1%) rotate(-0.5deg); }
    }

    .container {
      max-width: 900px;
      margin: 0 auto;
      padding: 20px 16px 40px;
    }

    /* ─── Header ─── */
    .header {
      text-align: center;
      padding: 32px 0 24px;
      position: relative;
    }

    .header-badge {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: rgba(14, 165, 233, 0.1);
      border: 1px solid rgba(14, 165, 233, 0.2);
      color: var(--accent-light);
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 1.5px;
      text-transform: uppercase;
      padding: 6px 16px;
      border-radius: 50px;
      margin-bottom: 16px;
    }

    .header h1 {
      font-size: clamp(24px, 5vw, 36px);
      font-weight: 800;
      background: linear-gradient(135deg, #f1f5f9, #38bdf8);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
      line-height: 1.2;
      margin-bottom: 8px;
    }

    .header p {
      color: var(--text-secondary);
      font-size: 14px;
      font-weight: 400;
    }

    .header .agent-info {
      display: flex;
      justify-content: center;
      gap: 20px;
      margin-top: 14px;
      flex-wrap: wrap;
    }

    .agent-info span {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: var(--text-muted);
    }

    .agent-info span .dot {
      width: 6px; height: 6px;
      border-radius: 50%;
      background: var(--success);
      box-shadow: 0 0 6px var(--success);
    }

    /* ─── Card ─── */
    .card {
      background: var(--bg-card);
      backdrop-filter: blur(24px);
      -webkit-backdrop-filter: blur(24px);
      border: 1px solid var(--glass-border);
      border-radius: var(--radius);
      padding: 32px;
      box-shadow: var(--shadow-lg);
      transition: var(--transition);
      position: relative;
      overflow: hidden;
    }

    .card::before {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 3px;
      background: linear-gradient(90deg, transparent, var(--accent), transparent);
      opacity: 0.6;
    }

    .card:hover {
      border-color: var(--glass-border-hover);
    }

    .section-title {
      font-size: 13px;
      font-weight: 700;
      color: var(--accent-light);
      text-transform: uppercase;
      letter-spacing: 1.5px;
      margin-bottom: 16px;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .section-title::after {
      content: '';
      flex: 1;
      height: 1px;
      background: linear-gradient(90deg, var(--glass-border), transparent);
    }

    /* ─── Form ─── */
    .form-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }

    .form-group {
      position: relative;
    }

    .form-group.full-width {
      grid-column: 1 / -1;
    }

    .form-group label {
      display: block;
      font-size: 12px;
      font-weight: 500;
      color: var(--text-secondary);
      margin-bottom: 6px;
      transition: var(--transition);
    }

    .form-group label .req {
      color: var(--accent);
      margin-left: 2px;
    }

    .form-group input,
    .form-group select,
    .form-group textarea {
      width: 100%;
      padding: 12px 14px;
      background: var(--input-bg);
      border: 1px solid var(--input-border);
      border-radius: var(--radius-sm);
      color: var(--text-primary);
      font-family: 'Inter', sans-serif;
      font-size: 14px;
      font-weight: 400;
      transition: var(--transition);
      outline: none;
    }

    .form-group input:focus,
    .form-group select:focus,
    .form-group textarea:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px var(--accent-glow);
      background: rgba(30, 41, 59, 0.9);
    }

    .form-group input::placeholder { color: var(--text-muted); }

    .form-group textarea { resize: vertical; min-height: 70px; }

    .form-group select {
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%2394a3b8' d='M6 8L1 3h10z'/%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 14px center;
      padding-right: 36px;
      cursor: pointer;
    }

    .form-group .error-msg {
      font-size: 11px;
      color: var(--error);
      margin-top: 4px;
      display: none;
      align-items: center;
      gap: 4px;
    }

    .form-group.has-error input,
    .form-group.has-error select {
      border-color: var(--error);
      box-shadow: 0 0 0 3px rgba(239, 68, 68, 0.15);
    }

    .form-group.has-error .error-msg { display: flex; }

    .form-group.valid input,
    .form-group.valid select {
      border-color: var(--success);
    }

    /* ─── Button ─── */
    .btn-primary {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 10px;
      width: 100%;
      padding: 14px 24px;
      background: linear-gradient(135deg, #0369a1, #0ea5e9);
      color: white;
      font-family: 'Inter', sans-serif;
      font-size: 15px;
      font-weight: 600;
      border: none;
      border-radius: var(--radius-sm);
      cursor: pointer;
      transition: var(--transition);
      position: relative;
      overflow: hidden;
      margin-top: 8px;
    }

    .btn-primary:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 25px rgba(14, 165, 233, 0.35);
    }

    .btn-primary:active { transform: translateY(0); }

    .btn-primary.loading {
      pointer-events: none;
      opacity: 0.8;
    }

    .btn-secondary {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      padding: 12px 24px;
      background: transparent;
      color: var(--accent-light);
      font-family: 'Inter', sans-serif;
      font-size: 14px;
      font-weight: 500;
      border: 1px solid var(--glass-border);
      border-radius: var(--radius-sm);
      cursor: pointer;
      transition: var(--transition);
      text-decoration: none;
    }

    .btn-secondary:hover {
      background: rgba(14, 165, 233, 0.1);
      border-color: var(--accent);
    }

    /* ─── Spinner ─── */
    .spinner {
      width: 18px; height: 18px;
      border: 2px solid rgba(255,255,255,0.3);
      border-top-color: white;
      border-radius: 50%;
      animation: spin 0.7s linear infinite;
      display: none;
    }

    .loading .spinner { display: inline-block; }
    .loading .btn-text { display: none; }

    @keyframes spin { to { transform: rotate(360deg); } }

    /* ─── Toast ─── */
    .toast {
      position: fixed;
      bottom: 24px;
      right: 24px;
      padding: 14px 20px;
      border-radius: var(--radius-sm);
      font-size: 13px;
      font-weight: 500;
      color: white;
      box-shadow: 0 10px 30px rgba(0,0,0,0.3);
      transform: translateY(100px);
      opacity: 0;
      transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
      z-index: 1000;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .toast.show { transform: translateY(0); opacity: 1; }
    .toast.success { background: linear-gradient(135deg, #059669, #10b981); }
    .toast.error { background: linear-gradient(135deg, #dc2626, #ef4444); }

    /* ─── Preview Section ─── */
    .preview-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      margin-bottom: 20px;
    }

    .preview-item {
      padding: 14px;
      background: rgba(30, 41, 59, 0.4);
      border-radius: var(--radius-sm);
      border: 1px solid rgba(100, 116, 139, 0.15);
    }

    .preview-item .label {
      font-size: 11px;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 1px;
      margin-bottom: 4px;
    }

    .preview-item .value {
      font-size: 16px;
      font-weight: 600;
      color: var(--text-primary);
    }

    .preview-item.highlight {
      background: linear-gradient(135deg, rgba(14, 165, 233, 0.15), rgba(6, 182, 212, 0.1));
      border-color: rgba(14, 165, 233, 0.3);
    }

    .preview-item.highlight .value {
      font-size: 22px;
      color: var(--accent-light);
    }

    .preview-actions {
      display: flex;
      gap: 12px;
      margin-top: 20px;
    }

    .preview-actions .btn-primary { flex: 2; margin-top: 0; }
    .preview-actions .btn-secondary { flex: 1; }

    /* ─── History ─── */
    .history-list { list-style: none; }

    .history-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 14px 16px;
      background: rgba(30, 41, 59, 0.3);
      border: 1px solid rgba(100, 116, 139, 0.12);
      border-radius: var(--radius-sm);
      margin-bottom: 8px;
      transition: var(--transition);
    }

    .history-item:hover {
      background: rgba(30, 41, 59, 0.5);
      border-color: var(--glass-border);
    }

    .history-item .file-name {
      font-size: 14px;
      font-weight: 500;
      color: var(--text-primary);
    }

    .history-item .file-meta {
      font-size: 12px;
      color: var(--text-muted);
      margin-top: 2px;
    }

    .history-item .download-btn {
      padding: 8px 16px;
      background: rgba(14, 165, 233, 0.1);
      border: 1px solid rgba(14, 165, 233, 0.25);
      border-radius: 8px;
      color: var(--accent-light);
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      transition: var(--transition);
      text-decoration: none;
    }

    .history-item .download-btn:hover {
      background: rgba(14, 165, 233, 0.2);
    }

    .empty-state {
      text-align: center;
      padding: 40px;
      color: var(--text-muted);
    }

    .empty-state .icon { font-size: 40px; margin-bottom: 12px; }

    /* ─── Module Compiler ─── */
    .module-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }

    .module-card {
      padding: 22px;
      background: rgba(30, 41, 59, 0.34);
      border: 1px solid rgba(100, 116, 139, 0.16);
      border-radius: var(--radius);
    }

    .module-card h3 {
      font-size: 18px;
      margin-bottom: 8px;
      color: var(--text-primary);
    }

    .module-card p {
      font-size: 13px;
      color: var(--text-secondary);
      line-height: 1.6;
      margin-bottom: 16px;
    }

    .module-meta {
      display: inline-flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 16px;
    }

    .module-meta span {
      padding: 5px 10px;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 600;
      color: var(--accent-light);
      background: rgba(14, 165, 233, 0.08);
      border: 1px solid rgba(14, 165, 233, 0.18);
    }

    .helper-text {
      margin-top: 6px;
      font-size: 11px;
      color: var(--text-muted);
      line-height: 1.5;
    }

    .info-banner {
      padding: 12px 16px;
      margin-bottom: 18px;
      border-radius: 10px;
      background: rgba(14, 165, 233, 0.08);
      border: 1px solid rgba(14, 165, 233, 0.18);
      color: var(--text-secondary);
      font-size: 13px;
      line-height: 1.6;
    }

    .upload-panel {
      padding: 18px;
      border-radius: var(--radius-sm);
      border: 1px dashed rgba(14, 165, 233, 0.28);
      background: rgba(14, 165, 233, 0.04);
      margin-bottom: 18px;
    }

    .upload-panel input[type="file"] {
      width: 100%;
      margin-top: 10px;
      color: var(--text-secondary);
    }

    .dossier-list {
      display: flex;
      flex-direction: column;
      gap: 12px;
      margin-top: 14px;
    }

    .dossier-item {
      padding: 16px;
      border-radius: 12px;
      background: rgba(30, 41, 59, 0.34);
      border: 1px solid rgba(100, 116, 139, 0.16);
    }

    .dossier-item-head {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: flex-start;
      flex-wrap: wrap;
      margin-bottom: 8px;
    }

    .dossier-item h4 {
      font-size: 15px;
      margin: 0;
      color: var(--text-primary);
    }

    .dossier-badges {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }

    .dossier-badges span {
      padding: 4px 9px;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 600;
      color: var(--accent-light);
      background: rgba(14, 165, 233, 0.08);
      border: 1px solid rgba(14, 165, 233, 0.18);
    }

    .field-summary {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin-top: 12px;
    }

    .field-summary .preview-item {
      margin: 0;
    }

    .warning-list {
      margin-top: 10px;
      color: #facc15;
      font-size: 12px;
      line-height: 1.5;
    }

    /* ─── Nav ─── */
    .nav {
      display: flex;
      gap: 8px;
      margin-bottom: 24px;
    }

    .nav a {
      padding: 8px 18px;
      font-size: 13px;
      font-weight: 500;
      color: var(--text-secondary);
      text-decoration: none;
      border-radius: 8px;
      transition: var(--transition);
    }

    .nav a:hover { color: var(--text-primary); background: rgba(255,255,255,0.05); }
    .nav a.active { color: var(--accent-light); background: rgba(14, 165, 233, 0.12); }

    /* ─── Scenario Builder ─── */
    .scenarios-list {
      display: flex;
      flex-direction: column;
      gap: 16px;
      margin-bottom: 20px;
    }

    .scenario-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 16px;
      padding-bottom: 12px;
      border-bottom: 1px solid rgba(100, 116, 139, 0.1);
    }

    .scenario-number {
      font-size: 12px;
      font-weight: 700;
      color: var(--accent-light);
      text-transform: uppercase;
    }

    .btn-remove {
      background: rgba(239, 68, 68, 0.1);
      color: var(--error);
      border: 1px solid rgba(239, 68, 68, 0.2);
      padding: 4px 10px;
      border-radius: 6px;
      font-size: 11px;
      font-weight: 600;
      cursor: pointer;
      transition: var(--transition);
    }

    .btn-remove:hover {
      background: var(--error);
      color: white;
    }

    .btn-add-scenario {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      width: 100%;
      padding: 12px;
      background: rgba(14, 165, 233, 0.05);
      border: 1px dashed var(--accent);
      border-radius: var(--radius-sm);
      color: var(--accent-light);
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: var(--transition);
      margin-bottom: 10px;
    }

    .btn-add-scenario:hover {
      background: rgba(14, 165, 233, 0.1);
      border-style: solid;
      transform: translateY(-1px);
    }

    /* ─── Footer ─── */
    .footer {
      text-align: center;
      padding: 24px 0;
      margin-top: 32px;
      border-top: 1px solid rgba(100, 116, 139, 0.1);
    }

    .footer p { font-size: 12px; color: var(--text-muted); margin-bottom: 4px; }

    .footer a {
      color: var(--accent);
      text-decoration: none;
      font-weight: 500;
    }

    .footer a:hover { color: var(--accent-light); text-decoration: underline; }

    /* ─── Split Layout ─── */
    .split-layout { display: flex; gap: 24px; align-items: flex-start; }
    .split-form { flex: 1; min-width: 0; }
    .split-preview { flex: 1; min-width: 0; position: sticky; top: 20px; max-height: 92vh; display: flex; flex-direction: column; }
    .preview-panel-header {
      display: flex; align-items: center; justify-content: space-between;
      padding: 12px 16px; border-radius: 12px 12px 0 0;
      background: rgba(14, 165, 233, 0.08); border: 1px solid rgba(14, 165, 233, 0.2);
      border-bottom: none;
    }
    .preview-panel-header h3 { margin: 0; font-size: 14px; color: var(--accent); font-weight: 600; }
    .preview-scroll {
      flex: 1; overflow-y: auto; border: 1px solid rgba(255,255,255,.1);
      border-radius: 0 0 12px 12px; background: #fff;
    }
    .preview-empty {
      display: flex; align-items: center; justify-content: center;
      min-height: 400px; color: #999; font-size: 14px; text-align: center; padding: 40px;
    }
    .preview-download-bar {
      margin-top: 12px; display: flex; gap: 10px;
    }
    .preview-download-bar .btn-primary { flex: 1; }

    /* ─── PDF Preview Styles ─── */
    .pdf-preview { font-family: Helvetica, Arial, sans-serif; color: #1a1a1a; padding: 20px; font-size: 11px; }
    .pdf-preview * { box-sizing: border-box; }
    .pdf-gold-top { height: 3px; border-radius: 2px 2px 0 0; }
    .pdf-header { display: flex; justify-content: space-between; align-items: flex-start; padding: 10px 0; margin-bottom: 8px; }
    .pdf-header-left h2 { margin: 0 0 2px; font-size: 14px; letter-spacing: .5px; }
    .pdf-header-left .sub { font-family: monospace; font-size: 9px; opacity: .7; }
    .pdf-header-right { text-align: right; font-size: 10px; }
    .pdf-header-right .oam-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-family: monospace; font-size: 8.5px; }
    .pdf-section-bar { padding: 5px 10px; font-size: 10px; font-weight: 700; color: #fff; letter-spacing: 1px; border-radius: 3px; margin: 10px 0 6px; }
    .pdf-meta { display: flex; gap: 0; margin-bottom: 10px; border-radius: 4px; overflow: hidden; }
    .pdf-meta span { flex: 1; text-align: center; padding: 5px; font-family: monospace; font-size: 9.5px; }
    .pdf-disclaimer { font-size: 9.5px; line-height: 1.5; margin-bottom: 10px; color: #444; }
    .pdf-client-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; border: 1px solid #dde2ea; }
    .pdf-client-table td { padding: 5px 8px; font-size: 10px; border-bottom: 1px solid #eceff4; }
    .pdf-client-table .lbl { color: #6b7280; font-size: 8.5px; text-transform: uppercase; width: 22%; }
    .pdf-client-table .val { font-weight: 700; }
    .pdf-hero { display: flex; gap: 8px; margin: 10px 0; }
    .pdf-hero-left { flex: 0.62; border-radius: 9px; padding: 14px; color: #fff; position: relative; overflow: hidden; }
    .pdf-hero-left .label { font-size: 9px; opacity: .8; margin-bottom: 4px; }
    .pdf-hero-left .rata { font-size: 34px; font-weight: 800; line-height: 1.1; }
    .pdf-hero-left .rata .dec { font-size: 20px; }
    .pdf-hero-left .rata .eur { font-size: 14px; opacity: .7; margin-left: 4px; }
    .pdf-hero-left .dur { font-size: 10px; opacity: .75; margin-top: 6px; }
    .pdf-hero-right { flex: 0.38; display: flex; flex-direction: column; gap: 5px; }
    .pdf-taeg-badge { flex: 1; border-radius: 8px; padding: 10px; text-align: left; border-width: 1.5px; border-style: solid; }
    .pdf-taeg-badge .cost-label { font-size: 8px; opacity: .75; color: #ddd; }
    .pdf-taeg-badge .taeg-val { font-size: 24px; font-weight: 800; }
    .pdf-taeg-badge .taeg-pct { font-size: 14px; font-weight: 700; }
    .pdf-taeg-badge .taeg-sub { font-family: monospace; font-size: 8px; opacity: .7; color: #ddd; margin-top: 2px; }
    .pdf-taeg-badge .taeg-info { font-size: 8px; opacity: .65; color: #ccc; margin-top: 4px; }
    .pdf-tan-box { background: #fff; border: 1px solid #d9dde4; border-radius: 6px; padding: 6px 8px; display: flex; justify-content: space-between; align-items: center; }
    .pdf-tan-box .tan-label { color: #6b7280; font-size: 9px; }
    .pdf-tan-box .tan-val { font-family: monospace; font-weight: 700; font-size: 10px; }
    .pdf-kpis { display: flex; gap: 6px; margin: 10px 0; }
    .pdf-kpis .kpi { flex: 1; text-align: center; padding: 8px 4px; border-radius: 6px; border: 1px solid #d6dce5; background: #fff; }
    .pdf-kpis .kpi.accent { border-color: var(--pdf-gold); }
    .pdf-kpis .kpi .kpi-label { font-size: 8px; text-transform: uppercase; color: #6b7280; margin-bottom: 3px; }
    .pdf-kpis .kpi .kpi-val { font-size: 12px; font-weight: 700; }
    .pdf-timeline { display: flex; align-items: flex-start; position: relative; padding: 8px 0 16px; margin: 8px 0; }
    .pdf-timeline .tl-line { position: absolute; top: 14px; left: 10%; right: 10%; height: 2px; }
    .pdf-timeline .tl-point { flex: 1; text-align: center; position: relative; z-index: 1; }
    .pdf-timeline .tl-dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; margin-bottom: 4px; border: 2px solid; }
    .pdf-timeline .tl-title { display: block; font-weight: 700; font-size: 9px; }
    .pdf-timeline .tl-sub { display: block; font-family: monospace; font-size: 8px; color: #6b7280; }
    .pdf-timeline .tl-val { display: block; font-family: monospace; font-size: 8.5px; margin-top: 2px; }
    .pdf-detail-table { width: 100%; border-collapse: collapse; margin: 6px 0 10px; border: 1px solid #d8dee7; }
    .pdf-detail-table td { padding: 6px 10px; font-size: 10px; }
    .pdf-detail-table .row-label { width: 65%; }
    .pdf-detail-table .row-value { text-align: right; font-weight: 700; }
    .pdf-detail-table tr:nth-child(odd) { background: #fcfcfd; }
    .pdf-detail-table tr:nth-child(even) { background: #fff; }
    .pdf-detail-table tr.taeg-row td { color: #fff; font-weight: 700; }
    .pdf-detail-table tr.total-row td { font-weight: 800; font-size: 11px; }
    .pdf-conditions { border: 1px solid #d8dee7; border-radius: 4px; padding: 6px 10px; margin: 8px 0; background: #fff; }
    .pdf-conditions li { list-style: none; padding: 4px 0; font-size: 9.5px; line-height: 1.5; }
    .pdf-conditions li::before { content: '•'; margin-right: 6px; font-weight: 700; }
    .pdf-closing { border-radius: 6px; padding: 10px 14px; text-align: center; font-size: 10px; color: #fff; margin: 10px 0; border-width: 1px; border-style: solid; }
    .pdf-oam-badge { text-align: center; margin: 12px 0; }
    .pdf-oam-badge img { max-width: 120px; border-radius: 4px; }
    .pdf-oam-badge .caption { font-family: monospace; font-size: 8.5px; color: #6b7280; margin-top: 4px; }
    .pdf-page-sep { border: none; border-top: 2px dashed rgba(0,0,0,.15); margin: 20px 0; }

    /* Editable elements in preview */
    .editable { cursor: text; border-bottom: 1px dashed transparent; transition: border-color .2s, background .2s; }
    .editable:hover { border-bottom-color: rgba(201,162,39,.5); }
    .editable:focus { outline: none; background: rgba(201,162,39,.08); border-bottom-color: rgba(201,162,39,.8); }

    /* ─── Responsive ─── */
    @media (max-width: 900px) {
      .split-layout { flex-direction: column; }
      .split-preview { position: static; max-height: none; }
    }
    @media (max-width: 640px) {
      .form-grid { grid-template-columns: 1fr; }
      .module-grid { grid-template-columns: 1fr; }
      .field-summary { grid-template-columns: 1fr; }
      .preview-grid { grid-template-columns: 1fr; }
      .card { padding: 20px; }
      .preview-actions { flex-direction: column; }
      .agent-info { flex-direction: column; gap: 6px; }
    }
    """

    # ─── Base HTML Template ───
    BASE_HTML = """
    <!doctype html>
    <html lang="it">
    <head>
      <meta charset="utf-8"/>
      <meta name="viewport" content="width=device-width, initial-scale=1"/>
      <title>{{ title }} — {{ prof_nome }}</title>
      <style>""" + CSS + """</style>
    </head>
    <body>
      <div class="container">
        <header class="header">
          {% if prof_rete %}<div class="header-badge">💼 Agente {{ prof_rete }}</div>{% endif %}
          <h1>QuintoQuote — Generatore Preventivi</h1>
          <p>Cessione del Quinto &amp; Delega di Pagamento</p>
          <div class="agent-info">
            {% if prof_nome %}<span><span class="dot"></span> {{ prof_nome }}</span>{% endif %}
            {% if prof_tel %}<span>📞 {{ prof_tel }}</span>{% endif %}
            {% if prof_oam %}<span>🆔 OAM {{ prof_oam }}</span>{% endif %}
          </div>
        </header>

        <nav class="nav">
          <a href="/" class="{{ 'active' if page == 'form' }}">✏️ Nuovo</a>
          <a href="/dossier" class="{{ 'active' if page == 'dossier' }}">🗂️ Dossier</a>
          <a href="/moduli" class="{{ 'active' if page == 'modules' }}">📎 Moduli</a>
          <a href="/storico" class="{{ 'active' if page == 'history' }}">📂 Storico</a>
          <a href="/impostazioni" class="{{ 'active' if page == 'settings' }}">⚙️ Impostazioni</a>
          {% if desktop_app %}<a href="/esci">⏻ Chiudi App</a>{% endif %}
        </nav>

        {{ content|safe }}

        <footer class="footer">
          <p>{% if prof_nome %}{{ prof_nome }}{% endif %}{% if prof_rete %} — Agente {{ prof_rete }}{% endif %}{% if prof_oam %} — OAM {{ prof_oam }}{% endif %}{% if not prof_nome and not prof_rete %}QuintoQuote{% endif %}</p>
          <p><a href="https://www.organismo-am.it/" target="_blank" rel="noopener">Verifica autorizzazione OAM ↗</a></p>
        </footer>
      </div>
      <div class="toast" id="toast"></div>
      {{ scripts|safe }}
    </body>
    </html>
    """

    # ─── Form Page ───
    FORM_CONTENT = """
    <div class="split-layout">
    <div class="split-form">
    <div class="card">
      <form id="prevForm" method="post" action="/genera">

        <div class="section-title">👤 Dati Anagrafici</div>
        <div class="form-grid">
          <div class="form-group">
            <label>Cliente <span class="req">*</span></label>
            <input name="cliente" id="f_cliente" required placeholder="Nome e Cognome"/>
            <div class="error-msg">⚠ Inserisci il nome del cliente</div>
          </div>
          <div class="form-group">
            <label>Data di nascita <span class="req">*</span></label>
            <input name="data_nascita" id="f_data_nascita" required placeholder="GG/MM/AAAA" maxlength="10"/>
            <span id="ageHelper" style="display:none;font-size:11px;font-weight:600;margin-left:6px"></span>
            <div class="error-msg">⚠ Formato: GG/MM/AAAA</div>
          </div>
          <div class="form-group">
            <label>Qualifica / Tipo lavoro <span class="req">*</span></label>
            <input name="tipo_lavoro" id="f_tipo_lavoro" required value="Dipendente Statale"/>
            <div class="error-msg">⚠ Campo obbligatorio</div>
          </div>
          <div class="form-group">
            <label>Provincia / Sede lavorativa <span class="req">*</span></label>
            <input name="provincia" id="f_provincia" required value="Roma"/>
            <div class="error-msg">⚠ Campo obbligatorio</div>
          </div>
        </div>

        <div style="height: 20px"></div>
        <div class="section-title">💰 Dati Finanziari</div>
        <div class="form-grid">
          <div class="form-group full-width">
            <label>Tipo finanziamento</label>
            <select name="tipo_finanziamento" id="f_tipo_fin">
              <option value="Cessione del Quinto">Cessione del Quinto</option>
              <option value="Delega di Pagamento">Delega di Pagamento</option>
            </select>
          </div>
          <div class="form-group">
            <label>Rata mensile (€) <span class="req">*</span></label>
            <input name="importo_rata" id="f_rata" required placeholder="350,00"/>
            <div class="error-msg">⚠ Importo non valido</div>
          </div>
          <div class="form-group">
            <label>Durata (mesi) <span class="req">*</span></label>
            <input name="durata_mesi" id="f_durata" required type="number" min="1" placeholder="120"/>
            <div class="error-msg">⚠ Min 1 mese</div>
          </div>
          <div class="form-group">
            <label>TAN (%) <span class="req">*</span></label>
            <input name="tan" id="f_tan" required placeholder="4,500"/>
            <div class="error-msg">⚠ Valore 0-50</div>
          </div>
          <div class="form-group">
            <label>TAEG (%) <span class="req">*</span></label>
            <input name="taeg" id="f_taeg" required placeholder="4,750"/>
            <div class="error-msg">⚠ Valore 0-50</div>
          </div>
          <div class="form-group full-width">
            <label>Importo netto erogato (€) <span class="req">*</span></label>
            <input name="importo_erogato" id="f_erogato" required placeholder="30.000,00"/>
            <div class="error-msg">⚠ Importo non valido</div>
          </div>
        </div>
        <div id="calcHelper" style="display:none;padding:8px 12px;margin-top:4px;background:rgba(14,165,233,.08);border:1px solid rgba(14,165,233,.2);border-radius:8px;font-size:12px;color:var(--text-secondary)"></div>

        <div style="height: 20px"></div>
        <div class="section-title">📊 Scenari Multipli</div>
        <div id="scenarios-container" class="scenarios-list">
          <!-- Dinamicamente popolato -->
        </div>
        <button type="button" class="btn-add-scenario" id="addScenarioBtn">
          <span>➕ Aggiungi Scenario Opzionale</span>
        </button>
        <p style="font-size: 11px; color: var(--text-muted); margin-top: 8px;">
          Puoi creare più opzioni (es. diverse rate o durate) che verranno stampate su pagine separate nel PDF.
        </p>

        <input type="hidden" name="scenari_json" id="f_scenari_json" value="[]">


        <div style="height: 20px"></div>
        <div class="section-title">Note</div>
        <div class="form-grid">
          <div class="form-group full-width">
            <label>Note aggiuntive (opzionale)</label>
            <textarea name="note" id="f_note" rows="3" placeholder="Eventuali note per il preventivo..."></textarea>
          </div>
        </div>

        <div style="height: 12px"></div>
        <button type="submit" class="btn-primary" id="submitBtn" style="display:none">
          <span class="btn-text">📄 Genera Preventivo PDF</span>
          <span class="spinner"></span>
        </button>
      </form>
    </div>
    </div><!-- /split-form -->

    <div class="split-preview">
      <div class="preview-panel-header">
        <h3>Anteprima Live</h3>
        <span id="previewStatus" style="font-size:11px;color:var(--text-muted)"></span>
      </div>
      <div class="preview-scroll" id="previewScroll">
        <div id="previewContent">
          <div class="preview-empty">Compila il form per vedere l'anteprima del preventivo</div>
        </div>
      </div>
      <div class="preview-download-bar">
        <button type="button" class="btn-primary" id="downloadPdfBtn" disabled>📥 Scarica PDF</button>
      </div>
    </div>
    </div><!-- /split-layout -->
    """

    FORM_SCRIPTS = r"""
    <script>
    (function(){
      const form = document.getElementById('prevForm');
      const btn = document.getElementById('submitBtn');
      const toast = document.getElementById('toast');
      const scenariosContainer = document.getElementById('scenarios-container');
      const addScenarioBtn = document.getElementById('addScenarioBtn');
      const scenariJsonInput = document.getElementById('f_scenari_json');
      const previewContent = document.getElementById('previewContent');
      const downloadBtn = document.getElementById('downloadPdfBtn');
      const previewStatus = document.getElementById('previewStatus');

      let scenarioCount = 0;
      let debounceTimer = null;
      const textOverrides = {};

      function parseIt(s) {
        if (!s) return NaN;
        return parseFloat(s.toString().replace(/EUR|€/g,'').replace(/\./g,'').replace(',','.').trim());
      }

      function validateDate(s) {
        const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        if (!m) return false;
        const d = new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
        return d && d.getDate() === parseInt(m[1]);
      }

      function setError(el, show) {
        const group = el.closest('.form-group');
        if (!group) return;
        if (show) { group.classList.add('has-error'); group.classList.remove('valid'); }
        else { group.classList.remove('has-error'); group.classList.add('valid'); }
      }

      function addScenario() {
        scenarioCount++;
        const id = `scenario_${scenarioCount}`;
        const html = `
          <div class="scenario-card" id="${id}">
            <div class="scenario-header">
              <span class="scenario-number">Scenario #${scenarioCount}</span>
              <button type="button" class="btn-remove" onclick="removeScenario('${id}')">Rimuovi</button>
            </div>
            <div class="form-grid">
              <div class="form-group full-width">
                <label>Tipo (Opzionale, default base)</label>
                <select class="s_tipo">
                  <option value="">Usa base</option>
                  <option value="Cessione del Quinto">Cessione del Quinto</option>
                  <option value="Delega di Pagamento">Delega di Pagamento</option>
                </select>
              </div>
              <div class="form-group">
                <label>Rata (€)</label>
                <input class="s_rata" placeholder="320,00" required>
              </div>
              <div class="form-group">
                <label>Durata (mesi)</label>
                <input class="s_durata" type="number" value="120" required>
              </div>
              <div class="form-group">
                <label>TAN (%)</label>
                <input class="s_tan" placeholder="4,500" required>
              </div>
              <div class="form-group">
                <label>TAEG (%)</label>
                <input class="s_taeg" placeholder="4,750" required>
              </div>
              <div class="form-group full-width">
                <label>Netto Erogato (€)</label>
                <input class="s_erogato" placeholder="28.000,00" required>
              </div>
            </div>
          </div>
        `;
        scenariosContainer.insertAdjacentHTML('beforeend', html);
        const card = document.getElementById(id);
        card.querySelectorAll('input, select').forEach(input => {
          input.addEventListener('input', schedulePreview);
          input.addEventListener('blur', () => validateField(input));
        });
        schedulePreview();
      }

      window.removeScenario = function(id) {
        const el = document.getElementById(id);
        if (el) el.remove();
        document.querySelectorAll('.scenario-number').forEach((span, i) => {
          span.textContent = `Scenario #${i+1}`;
        });
        schedulePreview();
      };

      function validateField(input) {
        const v = parseIt(input.value);
        let valid = true;
        if (input.classList.contains('s_rata') || input.classList.contains('s_erogato')) {
          valid = !isNaN(v) && v > 0;
        } else if (input.classList.contains('s_tan') || input.classList.contains('s_taeg')) {
          valid = !isNaN(v) && v >= 0 && v < 50;
        }
        setError(input, !valid);
        return valid;
      }

      addScenarioBtn.addEventListener('click', addScenario);

      // ─── Live Validation ───
      document.getElementById('f_data_nascita').addEventListener('input', function(e) {
        let v = e.target.value.replace(/[^0-9\/]/g, '');
        if (v.length === 2 && !v.includes('/')) v += '/';
        if (v.length === 5 && v.split('/').length === 2) v += '/';
        e.target.value = v;
        if (v.length === 10) setError(this, !validateDate(v));
        schedulePreview();
      });

      ['f_rata','f_erogato','f_tan','f_taeg'].forEach(id => {
        const el = document.getElementById(id);
        el.addEventListener('blur', () => {
          const v = parseIt(el.value);
          let valid = true;
          if (id==='f_rata' || id==='f_erogato') valid = !isNaN(v) && v > 0;
          else valid = !isNaN(v) && v >= 0 && v < 50;
          setError(el, !valid);
        });
      });

      // ─── Live Preview ───
      function collectFormData() {
        const params = new URLSearchParams();
        params.set('cliente', document.getElementById('f_cliente').value);
        params.set('data_nascita', document.getElementById('f_data_nascita').value);
        params.set('tipo_lavoro', document.getElementById('f_tipo_lavoro').value);
        params.set('provincia', document.getElementById('f_provincia').value);
        params.set('note', document.getElementById('f_note').value);
        params.set('tipo_finanziamento', document.getElementById('f_tipo_fin').value);
        params.set('importo_rata', String(parseIt(document.getElementById('f_rata').value) || 0));
        params.set('durata_mesi', document.getElementById('f_durata').value || '0');
        params.set('tan', String(parseIt(document.getElementById('f_tan').value) || 0));
        params.set('taeg', String(parseIt(document.getElementById('f_taeg').value) || 0));
        params.set('importo_erogato', String(parseIt(document.getElementById('f_erogato').value) || 0));

        // Scenarios
        const scenarios = [];
        document.querySelectorAll('.scenario-card').forEach(card => {
          scenarios.push({
            tipo_finanziamento: card.querySelector('.s_tipo').value || null,
            importo_rata: parseIt(card.querySelector('.s_rata').value) || 0,
            durata_mesi: parseInt(card.querySelector('.s_durata').value) || 0,
            tan: parseIt(card.querySelector('.s_tan').value) || 0,
            taeg: parseIt(card.querySelector('.s_taeg').value) || 0,
            importo_erogato: parseIt(card.querySelector('.s_erogato').value) || 0,
          });
        });
        if (scenarios.length > 0) params.set('scenari_json', JSON.stringify(scenarios));

        // Text overrides from inline editing
        if (Object.keys(textOverrides).length > 0) {
          params.set('text_overrides_json', JSON.stringify(textOverrides));
        }
        return params;
      }

      function schedulePreview() {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(updatePreview, 400);
      }

      function updatePreview() {
        const rata = parseIt(document.getElementById('f_rata').value);
        const durata = parseInt(document.getElementById('f_durata').value);
        const erogato = parseIt(document.getElementById('f_erogato').value);
        if (isNaN(rata) || rata <= 0 || !durata || durata <= 0 || isNaN(erogato) || erogato <= 0) {
          return; // Not enough data for preview
        }
        if (previewStatus) previewStatus.textContent = 'Aggiornamento...';
        const qs = collectFormData();
        fetch('/preview?' + qs.toString())
          .then(r => { if (!r.ok) throw new Error(r.statusText); return r.text(); })
          .then(html => {
            previewContent.innerHTML = html;
            downloadBtn.disabled = false;
            if (previewStatus) previewStatus.textContent = '';
          })
          .catch(() => {
            if (previewStatus) previewStatus.textContent = '';
          });
      }

      // Attach preview trigger to all form inputs
      document.querySelectorAll('#prevForm input, #prevForm select, #prevForm textarea').forEach(el => {
        el.addEventListener('input', schedulePreview);
      });

      // ─── Inline Editing ───
      previewContent.addEventListener('blur', function(e) {
        if (e.target.classList.contains('editable')) {
          const field = e.target.dataset.field;
          const text = e.target.textContent.trim();
          if (text) textOverrides[field] = text;
        }
      }, true);

      // ─── Download PDF ───
      downloadBtn.addEventListener('click', function() {
        const rata = parseIt(document.getElementById('f_rata').value);
        const durata = parseInt(document.getElementById('f_durata').value);
        if (isNaN(rata) || !durata) {
          showToast('Compila i campi finanziari prima di scaricare', 'error');
          return;
        }

        // Collect scenarios
        const scenarios = [];
        document.querySelectorAll('.scenario-card').forEach(card => {
          scenarios.push({
            tipo_finanziamento: card.querySelector('.s_tipo').value || null,
            importo_rata: parseIt(card.querySelector('.s_rata').value) || 0,
            durata_mesi: parseInt(card.querySelector('.s_durata').value) || 0,
            tan: parseIt(card.querySelector('.s_tan').value) || 0,
            taeg: parseIt(card.querySelector('.s_taeg').value) || 0,
            importo_erogato: parseIt(card.querySelector('.s_erogato').value) || 0,
          });
        });

        // Build base preventivo list
        const preventivi = [{
          cliente: document.getElementById('f_cliente').value,
          data_nascita: document.getElementById('f_data_nascita').value,
          tipo_lavoro: document.getElementById('f_tipo_lavoro').value,
          provincia: document.getElementById('f_provincia').value,
          note: document.getElementById('f_note').value,
          tipo_finanziamento: document.getElementById('f_tipo_fin').value,
          importo_rata: parseIt(document.getElementById('f_rata').value),
          durata_mesi: parseInt(document.getElementById('f_durata').value),
          tan: parseIt(document.getElementById('f_tan').value),
          taeg: parseIt(document.getElementById('f_taeg').value),
          importo_erogato: parseIt(document.getElementById('f_erogato').value),
        }];
        scenarios.forEach(s => {
          preventivi.push({
            cliente: preventivi[0].cliente,
            data_nascita: preventivi[0].data_nascita,
            tipo_lavoro: preventivi[0].tipo_lavoro,
            provincia: preventivi[0].provincia,
            note: preventivi[0].note,
            tipo_finanziamento: s.tipo_finanziamento || preventivi[0].tipo_finanziamento,
            importo_rata: s.importo_rata,
            durata_mesi: s.durata_mesi,
            tan: s.tan,
            taeg: s.taeg,
            importo_erogato: s.importo_erogato,
          });
        });

        // POST via hidden form
        const f = document.createElement('form');
        f.method = 'POST';
        f.action = '/download';
        f.style.display = 'none';

        const dj = document.createElement('input');
        dj.name = 'data_json';
        dj.value = JSON.stringify(preventivi);
        f.appendChild(dj);

        const oj = document.createElement('input');
        oj.name = 'text_overrides_json';
        oj.value = JSON.stringify(textOverrides);
        f.appendChild(oj);

        document.body.appendChild(f);
        f.submit();
        document.body.removeChild(f);
      });

      // ─── Form submit (hidden, kept for fallback) ───
      form.addEventListener('submit', function(e) {
        e.preventDefault();
      });

      function showToast(msg, type) {
        toast.className = 'toast ' + type + ' show';
        toast.textContent = msg;
        setTimeout(function(){ toast.classList.remove('show'); }, 4000);
      }

      document.getElementById('f_data_nascita').addEventListener('keydown', function(e) {
        const v = this.value;
        if (e.key === 'Backspace' && (v.length === 3 || v.length === 6) && v.endsWith('/')) {
          e.preventDefault();
          this.value = v.slice(0, -1);
        }
      });

      // ─── Assisted Calculator ───
      function updateCalc() {
        const rata = parseIt(document.getElementById('f_rata').value);
        const durata = parseInt(document.getElementById('f_durata').value);
        const erogato = parseIt(document.getElementById('f_erogato').value);
        const el = document.getElementById('calcHelper');
        if (!el) return;
        if (!isNaN(rata) && rata > 0 && durata > 0) {
          const montante = rata * durata;
          const mFmt = montante.toLocaleString('it-IT', {minimumFractionDigits:2, maximumFractionDigits:2});
          let html = 'Montante: <strong>' + mFmt + ' EUR</strong>';
          if (!isNaN(erogato) && erogato > 0) {
            const interessi = montante - erogato;
            const iFmt = interessi.toLocaleString('it-IT', {minimumFractionDigits:2, maximumFractionDigits:2});
            html += ' | Interessi: <strong>' + iFmt + ' EUR</strong>';
          }
          el.innerHTML = html;
          el.style.display = 'block';
        } else {
          el.style.display = 'none';
        }
      }
      ['f_rata','f_durata','f_erogato'].forEach(id => {
        document.getElementById(id).addEventListener('input', updateCalc);
      });

      // ─── Age Calculator ───
      function updateAge() {
        const v = document.getElementById('f_data_nascita').value;
        const el = document.getElementById('ageHelper');
        if (!el) return;
        if (validateDate(v)) {
          const parts = v.split('/');
          const born = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
          const today = new Date();
          let age = today.getFullYear() - born.getFullYear();
          const m = today.getMonth() - born.getMonth();
          if (m < 0 || (m === 0 && today.getDate() < born.getDate())) age--;
          el.textContent = age + ' anni';
          el.style.display = 'inline';
          if (age < 18 || age > 80) { el.style.color = '#ef4444'; } else { el.style.color = '#22c55e'; }
        } else {
          el.style.display = 'none';
        }
      }
      document.getElementById('f_data_nascita').addEventListener('input', updateAge);
    })();
    </script>
    """

    # ─── Preview Page ───
    PREVIEW_CONTENT = """
    <div class="card">
      <div class="section-title">📄 Riepilogo Preventivo</div>

      <div class="preview-grid" style="margin-bottom: 24px;">
        <div class="preview-item">
          <div class="label">Cliente</div>
          <div class="value">{{ p.cliente }}</div>
        </div>
        <div class="preview-item">
          <div class="label">Data di nascita</div>
          <div class="value">{{ p.data_nascita }} ({{ p.eta }} anni)</div>
        </div>
        <div class="preview-item">
          <div class="label">Qualifica</div>
          <div class="value">{{ p.tipo_lavoro }}</div>
        </div>
        <div class="preview-item">
          <div class="label">Provincia</div>
          <div class="value">{{ p.provincia }}</div>
        </div>
      </div>

      {% for prev in all_preventivi %}
      <div class="section-title">💰 {{ 'Preventivo Base' if loop.first else 'Scenario #' ~ loop.index0 }}</div>
      <div class="preview-grid" style="margin-bottom: 20px;">
        <div class="preview-item highlight">
          <div class="label">Rata Mensile</div>
          <div class="value">{{ euro_web(prev.importo_rata) }}</div>
        </div>
        <div class="preview-item highlight">
          <div class="label">Importo Netto Erogato</div>
          <div class="value">{{ euro_web(prev.importo_erogato) }}</div>
        </div>
        <div class="preview-item">
          <div class="label">Durata</div>
          <div class="value">{{ prev.durata_mesi }} mesi ({{ prev.durata_mesi // 12 }} anni)</div>
        </div>
        <div class="preview-item">
          <div class="label">Tipo Finanziamento</div>
          <div class="value">{{ prev.tipo_finanziamento }}</div>
        </div>
        <div class="preview-item">
          <div class="label">TAN / TAEG</div>
          <div class="value">{{ "%.3f"|format(prev.tan) }}% / {{ "%.3f"|format(prev.taeg) }}%</div>
        </div>
        <div class="preview-item">
          <div class="label">Montante</div>
          <div class="value">{{ euro_web(prev.montante) }}</div>
        </div>
      </div>
      {% endfor %}

      {% if p.note %}
      <div class="section-title">📝 Note</div>
      <div class="preview-item" style="grid-column: 1/-1; margin-bottom: 16px;">
        <div class="value" style="font-size: 14px; font-weight: 400;">{{ p.note }}</div>
      </div>
      {% endif %}

      <p style="font-size:0.82em;color:#666;margin-bottom:0.7em;">I valori riportati derivano dalla simulazione effettuata sul portale ufficiale FEEVO e hanno finalità illustrativa.</p>
      <div class="preview-actions">
        <form method="post" action="/download" style="flex:2;display:flex;">
          <input type="hidden" name="data_json" value='{{ data_json|e }}'/>
          <button type="submit" class="btn-primary" style="flex:1;margin:0;">
            <span class="btn-text">⬇️ Scarica PDF ({{ num_preventivi }} pagin{{ 'a' if num_preventivi == 1 else 'e' }})</span>
            <span class="spinner"></span>
          </button>
        </form>
        <a href="/" class="btn-secondary" style="flex:1;display:flex;align-items:center;justify-content:center;">✏️ Modifica</a>
      </div>
    </div>
    """

    # ─── History Page ───
    HISTORY_CONTENT = """
    <div class="card">
      <div class="section-title">📂 Storico PDF</div>
      {% if files %}
      <ul class="history-list">
        {% for f in files %}
        <li class="history-item">
          <div>
            <div class="file-name">{{ f.name }}</div>
            <div class="file-meta">{{ f.size }} KB — {{ f.date }}</div>
          </div>
          <a href="/download-file/{{ f.name }}" class="download-btn">⬇ Scarica</a>
        </li>
        {% endfor %}
      </ul>
      {% else %}
      <div class="empty-state">
        <div class="icon">📭</div>
        <p>Nessun PDF generato ancora.</p>
        <p style="margin-top:8px;"><a href="/" style="color:var(--accent-light);">Crea il primo preventivo →</a></p>
      </div>
      {% endif %}
    </div>
    """

    DOSSIER_CONTENT = """
    <div class="card">
      <div class="section-title">🗂️ Dossier Documenti</div>
      <div class="info-banner">
        Engine deterministico, senza AI: accetta <strong>PDF, JPG e PNG</strong>, estrae i campi con regex, anchor text,
        parser dedicati e OCR locale quando disponibile. Per aumentare l'affidabilita, ogni upload richiede la
        <strong>tipologia documento</strong>.
        {{ ocr_status_message }}
      </div>
      <div class="helper-text" style="margin:12px 0 18px">
        Puoi costruire il dossier poco alla volta: scegli una tipologia, carica un documento, poi un altro tipo documento, salva la revisione
        e continua finché non premi <strong>Nuova analisi</strong>.
      </div>

      {% if notice %}
      <div style="padding:12px 16px;margin-bottom:16px;border-radius:10px;background:rgba(16,185,129,.12);border:1px solid rgba(16,185,129,.3);color:#d1fae5;font-size:13px">
        {{ notice }}
      </div>
      {% endif %}

      {% if error %}
      <div style="padding:12px 16px;margin-bottom:16px;border-radius:10px;background:rgba(239,68,68,.12);border:1px solid rgba(239,68,68,.3);color:#fecaca;font-size:13px">
        {{ error }}
      </div>
      {% endif %}

      <form method="post" action="/dossier" enctype="multipart/form-data">
        <div class="upload-panel">
          <strong>1. Scegli la tipologia documento</strong>
          <div class="helper-text">
            Per ogni upload scegli una sola tipologia e carica uno o piu file dello stesso tipo.
          </div>
          <div class="module-grid" style="margin-top:14px">
            {% for doc_type in upload_document_types %}
            <label class="module-card" style="cursor:pointer;border:1px solid {{ 'rgba(201,162,39,.55)' if selected_document_key == doc_type.key else 'rgba(255,255,255,.08)' }};box-shadow:{{ '0 0 0 1px rgba(201,162,39,.28) inset' if selected_document_key == doc_type.key else 'none' }};">
              <div style="display:flex;gap:12px;align-items:flex-start">
                <input type="radio" name="document_type" value="{{ doc_type.key }}" {% if selected_document_key == doc_type.key %}checked{% endif %} style="margin-top:4px;accent-color:var(--accent-light)" />
                <div>
                  <h3 style="margin:0 0 8px">{{ doc_type.label }}</h3>
                  <p style="margin:0 0 8px">{{ doc_type.description }}</p>
                  <div class="helper-text">{{ doc_type.helper_text }}</div>
                </div>
              </div>
            </label>
            {% endfor %}
          </div>
          <strong style="display:block;margin-top:18px">2. Carica i documenti sorgente</strong>
          <div class="helper-text">
            Se vuoi analizzare tipi diversi, aggiungili in passaggi separati: ad esempio prima il contratto, poi la tessera sanitaria.
          </div>
          <input type="file" name="documenti" accept="application/pdf,.pdf,image/png,.png,image/jpeg,.jpg,.jpeg" multiple {% if not results %}required{% endif %}/>
        </div>
        <div class="preview-actions" style="margin-top:14px">
          <button type="submit" class="btn-primary" style="flex:1;margin:0">Analizza e aggiungi documenti</button>
          {% if results %}
          <button type="submit" formaction="/dossier/review" class="btn-secondary" style="flex:1;margin:0">Salva revisione</button>
          {% endif %}
        </div>

        {% if results %}
        <div class="section-title" style="margin-top:24px">📄 File Analizzati</div>
        <div class="dossier-list">
          {% for result in results %}
          <div class="dossier-item">
            <div class="dossier-item-head">
              <div>
                <h4>{{ result.filename }}</h4>
                <div class="helper-text">{{ result.document_label }}</div>
              </div>
              <div class="dossier-badges">
                <span>{{ result.page_count }} pag.</span>
                <span>{{ result.text_length }} caratteri</span>
                <span>{{ result.keyword_hits }} indicatori</span>
                <span>{{ result.extracted_fields|length }} campi</span>
              </div>
            </div>
            {% if result.warnings %}
            <div class="warning-list">
              {% for warning in result.warnings %}
              <div>• {{ warning }}</div>
              {% endfor %}
            </div>
            {% endif %}
          </div>
          {% endfor %}
        </div>

        <div class="section-title" style="margin-top:24px">🧾 Revisione Dati</div>
        <div class="info-banner">
          Qui puoi controllare i dati estratti, correggerli, integrare quelli mancanti, aggiungere altri documenti e poi aprire il modulo gia precompilato.
        </div>
        {% for section in review_sections %}
        <div class="section-title" style="margin-top:{{ '0' if loop.first else '20px' }}">🧩 {{ section.title }}</div>
        <div class="form-grid">
          {% for field in section.fields %}
          <div class="form-group">
            <label>{{ field.label }}</label>
            <input name="{{ field.name }}" value="{{ field.value }}" placeholder="{{ field.placeholder }}" autocomplete="off"/>
            {% if field.source_filename %}
            <div class="helper-text">Fonte: {{ field.source_filename }} · {{ field.source_label }}</div>
            {% else %}
            <div class="helper-text">Non trovato automaticamente: puoi inserirlo a mano.</div>
            {% endif %}
          </div>
          {% endfor %}
        </div>
        {% endfor %}

        <div class="section-title" style="margin-top:24px">📎 Apri Modulo</div>
        <div class="preview-actions">
          <button type="submit" formaction="/moduli/allegato-e/prefill" class="btn-primary" style="flex:1;margin:0">Apri Allegato E precompilato</button>
          <button type="submit" formaction="/moduli/allegato-c/prefill" class="btn-primary" style="flex:1;margin:0">Apri Allegato C precompilato</button>
        </div>
        {% endif %}
      </form>

      {% if results %}
      <form method="post" action="/dossier/reset" style="margin-top:12px">
        <button type="submit" class="btn-secondary" style="width:100%;margin:0">Nuova analisi</button>
      </form>
      {% endif %}
    </div>
    """

    MODULI_HOME_CONTENT = """
    <div class="card">
      <div class="section-title">📎 Compilatore Modulistica MEF</div>
      <div class="info-banner">
        I due template PDF presenti in <code>docs/</code> sono stati mappati sui rispettivi widget:
        puoi compilare i campi con etichette leggibili e scaricare direttamente il PDF finale.
      </div>

      <div class="module-grid">
        {% for spec in specs %}
        <div class="module-card">
          <h3>{{ spec.label }}</h3>
          <p>{{ spec.description }}</p>
          <div class="module-meta">
            <span>{{ spec.summary }}</span>
            <span>{{ spec.template_name }}</span>
          </div>
          <a href="/moduli/{{ spec.slug }}" class="btn-primary" style="margin-top:0;text-decoration:none">
            Apri compilatore
          </a>
        </div>
        {% endfor %}
      </div>
    </div>
    """

    MODULO_FORM_CONTENT = """
    <div class="card">
      <div class="section-title">📄 {{ spec.label }}</div>
      <div class="info-banner">
        {{ spec.description }}<br/>
        Template di origine: <code>{{ spec.template_name }}</code>. {{ spec.summary }}
      </div>

      {% if error %}
      <div style="padding:12px 16px;margin-bottom:16px;border-radius:10px;background:rgba(239,68,68,.12);border:1px solid rgba(239,68,68,.3);color:#fecaca;font-size:13px">
        {{ error }}
      </div>
      {% endif %}

      <form method="post" action="/moduli/{{ spec.slug }}/download">
        {% for section in spec.sections %}
        <div class="section-title" style="margin-top:{{ '0' if loop.first else '20px' }}">{{ section.icon }} {{ section.title }}</div>
        <p style="font-size:12px;color:var(--text-muted);margin:-6px 0 14px">{{ section.description }}</p>
        <div class="form-grid">
          {% for field in section.fields %}
          <div class="form-group {{ 'full-width' if field.full_width else '' }}">
            <label>{{ field.label }}</label>
            <input
              name="{{ field.name }}"
              value="{{ values.get(field.name, '') }}"
              placeholder="{{ field.placeholder }}"
              autocomplete="off"
            />
            {% if field.help_text %}
            <div class="helper-text">{{ field.help_text }}</div>
            {% endif %}
          </div>
          {% endfor %}
        </div>
        {% endfor %}

        <div style="display:flex;gap:12px;margin-top:24px;flex-wrap:wrap">
          <button type="submit" class="btn-primary" style="flex:2;min-width:240px">{{ spec.button_label }}</button>
          <a href="/moduli" class="btn-secondary" style="flex:1;min-width:180px;display:flex;align-items:center;justify-content:center">↩ Torna ai moduli</a>
        </div>
      </form>
    </div>
    """
    def parse_float_web(s: str) -> float:
        return parse_decimal_loose(s)

    def euro_web(amount: float) -> str:
        return euro(amount)

    def render_page(content_tpl, scripts_tpl="", **kwargs):
        """Two-stage render: first render content/scripts, then inject into base."""
        prof = get_cached_profile()
        kwargs.setdefault("prof_nome", prof.nome_agente)
        kwargs.setdefault("prof_rete", prof.rete_mandante)
        kwargs.setdefault("prof_oam", prof.codice_oam)
        kwargs.setdefault("prof_tel", prof.telefono)
        kwargs.setdefault("desktop_app", _is_frozen_runtime())
        content_html = render_template_string(content_tpl, **kwargs)
        scripts_html = render_template_string(scripts_tpl, **kwargs) if scripts_tpl else ""
        return render_template_string(BASE_HTML, content=content_html, scripts=scripts_html, **kwargs)

    def render_modulo_form(spec_key: str, values: Optional[dict[str, str]] = None, error: str = ""):
        spec = get_pdf_template_spec(spec_key)
        merged_values = default_pdf_form_values(spec)
        if values:
            merged_values.update(values)
        return render_page(
            MODULO_FORM_CONTENT,
            page="modules",
            title=spec.label,
            spec=spec,
            values=merged_values,
            error=error,
        )

    def get_dossier_state_id(create: bool = False) -> str:
        state_id = sanitize_pdf_text(session.get("dossier_state_id", ""))
        if state_id:
            return state_id
        if not create:
            return ""
        state_id = uuid.uuid4().hex
        session["dossier_state_id"] = state_id
        return state_id

    def load_current_dossier_state() -> tuple[list[DocumentExtractionResult], dict[str, str]]:
        state_id = get_dossier_state_id(create=False)
        if not state_id:
            return [], {}
        try:
            return load_dossier_state(state_id)
        except Exception:
            return [], {}

    def save_current_dossier_state(results: list[DocumentExtractionResult], manual_values: dict[str, str]) -> None:
        state_id = get_dossier_state_id(create=True)
        save_dossier_state(state_id, results, manual_values)

    def clear_current_dossier_state() -> None:
        state_id = get_dossier_state_id(create=False)
        if state_id:
            clear_dossier_state(state_id)
        session.pop("dossier_state_id", None)

    def get_last_dossier_document_type() -> str:
        document_key = sanitize_pdf_text(session.get("dossier_last_document_type", ""))
        return document_key if document_key in _DOSSIER_UPLOAD_DOCUMENT_KEYS else ""

    def set_last_dossier_document_type(document_key: str) -> None:
        clean_key = sanitize_pdf_text(document_key)
        if clean_key in _DOSSIER_UPLOAD_DOCUMENT_KEYS:
            session["dossier_last_document_type"] = clean_key

    def render_dossier_page(
        results: Optional[list[DocumentExtractionResult]] = None,
        manual_values: Optional[dict[str, str]] = None,
        error: str = "",
        notice: str = "",
        selected_document_key: str = "",
    ):
        reviewed_fields = merge_reviewed_case_fields(results or [], manual_values or {})
        return render_page(
            DOSSIER_CONTENT,
            page="dossier",
            title="Dossier",
            results=results or [],
            aggregated_fields=reviewed_fields,
            review_sections=build_review_sections(reviewed_fields),
            upload_document_types=_DOSSIER_UPLOAD_DOCUMENT_TYPES,
            selected_document_key=selected_document_key or get_last_dossier_document_type(),
            ocr_status_message=get_ocr_status_message(),
            error=error,
            notice=notice,
        )

    def _render_preview_html(preventivi: list, prof: BrandingProfile, text_overrides: Optional[dict] = None) -> str:
        """Render an HTML fragment that mirrors the PDF layout (navy/gold theme)."""
        _t = get_design_tokens(prof)
        _ovr = sanitize_text_overrides(text_overrides)
        navy = _t["navy"]
        gold = _t["gold"]
        gold_light = _t["gold_light"]
        parts = []
        total = len(preventivi)
        prof_nome = escape_preview_text(prof.nome_agente)
        prof_rete = escape_preview_text(prof.rete_mandante)
        prof_oam = escape_preview_text(prof.codice_oam)
        prof_tel = escape_preview_text(prof.telefono)
        bollino_name = sanitize_asset_filename(prof.bollino_path)

        for idx, p in enumerate(preventivi, start=1):
            p.compute()
            try:
                data_doc = datetime.strptime(p.data_preventivo, "%d/%m/%Y")
                scadenza = (data_doc + timedelta(days=30)).strftime("%d/%m/%Y")
            except Exception:
                scadenza = "30 giorni dalla data emissione"

            anni = p.durata_mesi / 12
            durata_anni = f"{int(anni)} anni" if p.durata_mesi % 12 == 0 else f"{anni:.1f}".replace(".", ",") + " anni"

            rata_str = f"{p.importo_rata:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            rata_int, rata_dec = rata_str.split(",")

            taeg_str = f"{p.taeg:.3f}".replace(".", ",")
            tan_str = f"{p.tan:.3f}%".replace(".", ",")

            half_rates = max(1, p.durata_mesi // 2)
            try:
                end_date = add_months(datetime.strptime(p.data_preventivo, "%d/%m/%Y").date(), p.durata_mesi)
                end_year = str(end_date.year)
            except Exception:
                end_year = "—"

            cliente_display_raw = _ovr.get("cliente_nome", p.cliente)
            note_display_raw = _ovr.get("note", p.note.strip())

            _disc_who = f"da {prof.nome_agente}, Agente in attività finanziaria iscritto OAM {prof.codice_oam}" if prof.nome_agente and prof.codice_oam else "dall'agente"
            _disclaimer_default = (
                f"Documento predisposto {_disc_who}, "
                "operante in qualità di agente diretto del mandante. I valori economici riportati derivano dalla "
                "simulazione effettuata sul portale ufficiale FEEVO alla data di elaborazione."
            )
            disclaimer_text_raw = _ovr.get("disclaimer", _disclaimer_default)

            _contact_parts = []
            if prof.nome_agente:
                _contact_parts.append(prof.nome_agente)
            if prof.telefono:
                _contact_parts.append(f"al {prof.telefono}")
            _contact_str = " ".join(_contact_parts) if _contact_parts else "l'agente di riferimento"
            _closing_default = (
                f"Simulazione elaborata in data {p.data_preventivo}. "
                f"Per procedere o richiedere chiarimenti, contatta {_contact_str}."
            )
            closing_text_raw = _ovr.get("closing", _closing_default)

            cliente_display = escape_preview_text(cliente_display_raw, preserve_line_breaks=True)
            note_display = escape_preview_text(note_display_raw, preserve_line_breaks=True)
            disclaimer_text = escape_preview_text(disclaimer_text_raw, preserve_line_breaks=True)
            closing_text = escape_preview_text(closing_text_raw, preserve_line_breaks=True)
            tipo_finanziamento_display = escape_preview_text(p.tipo_finanziamento)
            tipo_finanziamento_title = escape_preview_text(p.tipo_finanziamento.upper())
            tipo_lavoro_display = escape_preview_text(p.tipo_lavoro)
            provincia_display = escape_preview_text(p.provincia)
            data_nascita_display = escape_preview_text(p.data_nascita)

            _h_name = f'<h2 style="color:{navy}">{prof_nome}</h2>' if prof_nome else ''
            _h_sub_parts = []
            if prof_rete:
                _h_sub_parts.append(f"Agente {prof_rete}")
            if prof_oam:
                _h_sub_parts.append(f"OAM {prof_oam}")
            _h_sub = f'<div class="sub">{" | ".join(_h_sub_parts)}</div>' if _h_sub_parts else ''
            _h_tel = f'<div style="color:#5d6673">Tel: {prof_tel}</div>' if prof_tel else ''
            _h_badge = f'<span class="oam-badge" style="background:{navy};color:{gold}">Registrazione OAM {prof_oam}</span>' if prof_oam else ''

            note_row = ""
            if note_display_raw.strip():
                note_row = (
                    "<tr><td class='lbl'>NOTE</td><td colspan='3' class='val editable' "
                    f"data-field='note' contenteditable='true'>{note_display}</td></tr>"
                )

            if idx > 1:
                parts.append('<hr class="pdf-page-sep"/>')

            parts.append(f'''
<div class="pdf-preview" style="--pdf-navy:{navy};--pdf-gold:{gold};--pdf-gold-light:{gold_light}">
  <div class="pdf-gold-top" style="background:{gold}"></div>

  <!-- Header -->
  <div class="pdf-header">
    <div class="pdf-header-left">
      {_h_name}
      {_h_sub}
    </div>
    <div class="pdf-header-right">
      {_h_tel}
      {_h_badge}
    </div>
  </div>

  {"<div style='text-align:center;font-size:9px;color:#6b7280;margin-bottom:4px'>Scenario " + str(idx) + " di " + str(total) + "</div>" if total > 1 else ""}
  <div style="text-align:center;font-size:9px;color:#6b7280;margin-bottom:4px">Riepilogo illustrativo della simulazione</div>
  <div style="text-align:center;font-size:18px;font-weight:800;color:{navy};margin-bottom:10px;letter-spacing:.5px">{tipo_finanziamento_title}</div>

  <!-- Meta -->
  <div class="pdf-meta" style="border:1px solid {gold};background:#f7f3e7">
    <span style="color:{gold};border-right:1px solid #e4d7a7">Emissione: {p.data_preventivo}</span>
    <span style="color:{gold}">Validità: fino al {scadenza}</span>
  </div>

  <!-- Disclaimer -->
  <div class="pdf-disclaimer editable" data-field="disclaimer" contenteditable="true">{disclaimer_text}</div>

  <!-- Client Profile -->
  <div class="pdf-section-bar" style="background:{navy};border-bottom:2px solid {gold}">PROFILO CLIENTE</div>
  <table class="pdf-client-table">
    <tr><td class="lbl">CLIENTE</td><td class="val editable" data-field="cliente_nome" contenteditable="true">{cliente_display}</td>
        <td class="lbl">DATA DI NASCITA</td><td class="val">{data_nascita_display} ({p.eta} anni)</td></tr>
    <tr><td class="lbl">CATEGORIA</td><td class="val">{tipo_lavoro_display}</td>
        <td class="lbl">PROVINCIA</td><td class="val">{provincia_display}</td></tr>
    {note_row}
  </table>

  <!-- Hero Finance Card -->
  <div class="pdf-hero">
    <div class="pdf-hero-left" style="background:{navy}">
      <div class="label">RATA MENSILE</div>
      <div class="rata">{rata_int},<span class="dec">{rata_dec}</span><span class="eur">EUR</span></div>
      <div class="dur">Per {p.durata_mesi} mesi ({durata_anni})</div>
    </div>
    <div class="pdf-hero-right">
      <div class="pdf-taeg-badge" style="background:{navy};border-color:{gold}">
        <div class="cost-label">COSTO DEL FINANZIAMENTO</div>
        <div><span class="taeg-val" style="color:{gold}">{taeg_str}</span><span class="taeg-pct" style="color:{gold}">%</span></div>
        <div class="taeg-sub">TAEG (Tasso Annuo Effettivo Globale)</div>
        <div class="taeg-info">Include tutti i costi: interessi, imposte e spese</div>
      </div>
      <div class="pdf-tan-box">
        <span class="tan-label">TAN nominale</span>
        <span class="tan-val" style="color:{navy}">{tan_str}</span>
      </div>
    </div>
  </div>

  <!-- KPIs -->
  <div class="pdf-kpis">
    <div class="kpi"><div class="kpi-label">Netto erogato</div><div class="kpi-val" style="color:{navy}">{euro(p.importo_erogato)}</div></div>
    <div class="kpi"><div class="kpi-label">Montante totale</div><div class="kpi-val" style="color:{navy}">{euro(p.montante)}</div></div>
    <div class="kpi accent" style="border-color:{gold}"><div class="kpi-label">Interessi</div><div class="kpi-val" style="color:{navy}">{euro(p.interessi)}</div></div>
  </div>

  <!-- Timeline -->
  <div style="font-size:8.5px;color:#7f8894;margin-bottom:2px">IL TUO PERCORSO FINANZIARIO</div>
  <div class="pdf-timeline">
    <div class="tl-line" style="background:{navy}"></div>
    <div class="tl-point"><span class="tl-dot" style="background:{gold};border-color:{gold}"></span><span class="tl-title" style="color:{navy}">Oggi</span><span class="tl-sub">Erogazione</span><span class="tl-val" style="color:{gold}">{euro(p.importo_erogato)}</span></div>
    <div class="tl-point"><span class="tl-dot" style="background:#b1b8c2;border-color:#b1b8c2"></span><span class="tl-title">Mese {half_rates}</span><span class="tl-sub">Metà percorso</span><span class="tl-val" style="color:#6b7280">{half_rates} rate pagate</span></div>
    <div class="tl-point"><span class="tl-dot" style="background:{gold};border-color:{gold}"></span><span class="tl-title" style="color:{navy}">{end_year}</span><span class="tl-sub">Fine percorso</span><span class="tl-val" style="color:#1f9d67">Obiettivo raggiunto</span></div>
  </div>

  <!-- Detail Table -->
  <div class="pdf-section-bar" style="background:{navy};border-bottom:2px solid {gold}">DETTAGLIO ECONOMICO</div>
  <table class="pdf-detail-table">
    <tr><td class="row-label">Tipologia prodotto</td><td class="row-value">{tipo_finanziamento_display}</td></tr>
    <tr><td class="row-label">Capitale netto erogato</td><td class="row-value">{euro(p.importo_erogato)}</td></tr>
    <tr><td class="row-label">Interessi complessivi</td><td class="row-value">{euro(p.interessi)}</td></tr>
    <tr><td class="row-label">Durata contrattuale</td><td class="row-value">{p.durata_mesi} mesi ({durata_anni})</td></tr>
    <tr><td class="row-label">TAN nominale</td><td class="row-value">{p.tan:.3f}%</td></tr>
    <tr class="taeg-row" style="background:{navy}"><td class="row-label" style="border-top:2px solid {gold};border-bottom:2px solid {gold}">TAEG</td><td class="row-value" style="color:{gold};border-top:2px solid {gold};border-bottom:2px solid {gold}">{p.taeg:.3f}%</td></tr>
    <tr class="total-row" style="background:#f8f5ec"><td class="row-label" style="border-top:1.5px solid {navy}">TOTALE DA RIMBORSARE</td><td class="row-value" style="border-top:1.5px solid {navy}">{euro(p.montante)}</td></tr>
  </table>

  <!-- Important Info -->
  <div class="pdf-section-bar" style="background:{navy};border-bottom:2px solid {gold}">INFORMAZIONI IMPORTANTI</div>
  <ul class="pdf-conditions">
    <li style="color:{gold}"><span style="color:#333">Erogazione su conto intestato a {cliente_display}.</span></li>
    <li style="color:{gold}"><span style="color:#333">Il rimborso avviene automaticamente tramite trattenuta mensile in busta paga.</span></li>
    <li style="color:{gold}"><span style="color:#333">Le polizze assicurative obbligatorie sono già incluse nel TAEG indicato.</span></li>
    <li style="color:{gold}"><span style="color:#333">Puoi estinguere il finanziamento in anticipo: ti verrà rimborsata la quota interessi non maturata.</span></li>
    <li style="color:{gold}"><span style="color:#333">Non sono previsti costi ulteriori a carico del cliente.</span></li>
  </ul>

  <!-- Closing -->
  <div class="pdf-closing editable" data-field="closing" contenteditable="true" style="background:{navy};border-color:{gold}">{closing_text}</div>

  <!-- OAM Badge -->
  <div class="pdf-oam-badge">''')

            if bollino_name:
                parts.append(f'<img src="/assets/{bollino_name}" alt="Bollino OAM" onerror="this.style.display=\'none\'"/>')
            _oam_caption = f"Iscritto all'Organismo degli Agenti e Mediatori{' — ' + prof.codice_oam if prof.codice_oam else ''}"
            if bollino_name or prof.codice_oam:
                parts.append(f'<div class="caption">{escape_preview_text(_oam_caption)}</div>')
            parts.append('''
  </div>
</div>''')

        return "\n".join(parts)

    @app.get("/preview")
    def preview():
        try:
            prof = get_cached_profile()
            p = Preventivo(
                cliente=request.args.get("cliente", "Cliente"),
                data_nascita=request.args.get("data_nascita", "01/01/1980"),
                tipo_lavoro=request.args.get("tipo_lavoro", "—"),
                provincia=request.args.get("provincia", "—"),
                note=request.args.get("note", ""),
                tipo_finanziamento=request.args.get("tipo_finanziamento", "Cessione del Quinto"),
                importo_rata=parse_float_web(request.args.get("importo_rata", "0")),
                durata_mesi=int(request.args.get("durata_mesi", 0) or 0),
                tan=parse_float_web(request.args.get("tan", "0")),
                taeg=parse_float_web(request.args.get("taeg", "0")),
                importo_erogato=parse_float_web(request.args.get("importo_erogato", "0")),
            )
            preventivi = [p]

            # Parse extra scenarios
            scenari_raw = request.args.get("scenari_json", "")
            if scenari_raw:
                scenari_data = _json.loads(scenari_raw)
                if not isinstance(scenari_data, list):
                    raise ValueError("Scenari non validi.")
                for s in scenari_data:
                    if not isinstance(s, dict):
                        raise ValueError("Scenario non valido.")
                    sp = Preventivo(
                        cliente=p.cliente,
                        data_nascita=p.data_nascita,
                        tipo_lavoro=p.tipo_lavoro,
                        provincia=p.provincia,
                        note=p.note,
                        tipo_finanziamento=s.get("tipo_finanziamento") or p.tipo_finanziamento,
                        importo_rata=parse_float_web(s.get("importo_rata", 0)),
                        durata_mesi=int(s.get("durata_mesi", 0)),
                        tan=parse_float_web(s.get("tan", 0)),
                        taeg=parse_float_web(s.get("taeg", 0)),
                        importo_erogato=parse_float_web(s.get("importo_erogato", 0)),
                    )
                    preventivi.append(sp)

            text_overrides: dict[str, str] = {}
            ovr_raw = request.args.get("text_overrides_json", "")
            if ovr_raw:
                text_overrides = sanitize_text_overrides(_json.loads(ovr_raw))

            return _render_preview_html(preventivi, prof, text_overrides)
        except Exception:
            return '<div class="preview-empty" style="color:#ef4444">Dati insufficienti per l\'anteprima</div>'

    SETUP_BANNER = """
    {% if not prof_nome %}
    <div style="padding:12px 16px;margin-bottom:16px;border-radius:10px;background:rgba(201,162,39,.12);border:1px solid rgba(201,162,39,.3);color:#e8d5a3;font-size:13px">
      ⚙️ <strong>Benvenuto in QuintoQuote!</strong> Configura il tuo profilo agente nelle <a href="/impostazioni" style="color:#c9a227;font-weight:600">Impostazioni</a> per personalizzare i preventivi.
    </div>
    {% endif %}
    """

    @app.get("/")
    def home():
        return render_page(SETUP_BANNER + FORM_CONTENT, FORM_SCRIPTS, page="form", title="Nuovo Preventivo")

    @app.get("/dossier")
    def dossier_get():
        results, manual_values = load_current_dossier_state()
        return render_dossier_page(results=results, manual_values=manual_values)

    @app.post("/dossier")
    def dossier_post():
        existing_results, manual_values = load_current_dossier_state()
        manual_values.update(extract_manual_case_values(request.form.to_dict(flat=True)))
        selected_document_key = sanitize_pdf_text(request.form.get("document_type", ""))
        uploaded_files = [file for file in request.files.getlist("documenti") if file and file.filename]
        if not uploaded_files:
            return render_dossier_page(
                results=existing_results,
                manual_values=manual_values,
                error="Carica almeno un PDF, JPG o PNG da analizzare.",
                selected_document_key=selected_document_key,
            ), 400
        if selected_document_key not in _DOSSIER_UPLOAD_DOCUMENT_KEYS:
            return render_dossier_page(
                results=existing_results,
                manual_values=manual_values,
                error="Seleziona la tipologia documento prima di analizzare il file.",
                selected_document_key=selected_document_key,
            ), 400

        new_results: list[DocumentExtractionResult] = []
        errors: list[str] = []
        for file in uploaded_files:
            safe_name = Path(file.filename).name
            if not is_supported_dossier_file(safe_name):
                errors.append(f"{safe_name}: formato non supportato, usa PDF, JPG o PNG.")
                continue
            data = file.read()
            if not data:
                errors.append(f"{safe_name}: file vuoto.")
                continue
            try:
                new_results.append(extract_document_result(safe_name, data, expected_document_key=selected_document_key))
            except Exception as exc:
                errors.append(f"{safe_name}: {exc}")

        if not new_results and errors and not existing_results:
            return render_dossier_page(error=" ".join(errors), selected_document_key=selected_document_key), 400

        all_results = existing_results + new_results
        save_current_dossier_state(all_results, manual_values)
        set_last_dossier_document_type(selected_document_key)
        error_text = " ".join(errors) if errors else ""
        notice = f"Dossier aggiornato: {len(all_results)} documenti analizzati." if new_results else ""
        return render_dossier_page(
            results=all_results,
            manual_values=manual_values,
            error=error_text,
            notice=notice,
            selected_document_key=selected_document_key,
        )

    @app.post("/dossier/review")
    def dossier_review_post():
        results, manual_values = load_current_dossier_state()
        manual_values.update(extract_manual_case_values(request.form.to_dict(flat=True)))
        save_current_dossier_state(results, manual_values)
        return render_dossier_page(results=results, manual_values=manual_values, notice="Revisione salvata.")

    @app.post("/dossier/reset")
    def dossier_reset_post():
        clear_current_dossier_state()
        return render_dossier_page(notice="Nuova analisi avviata: dossier azzerato.")

    @app.get("/moduli")
    def moduli_home():
        specs = (ALLEGATO_E_SPEC, ALLEGATO_C_SPEC)
        return render_page(MODULI_HOME_CONTENT, page="modules", title="Modulistica MEF", specs=specs)

    @app.get("/moduli/<slug>")
    def modulo_form(slug):
        try:
            return render_modulo_form(slug)
        except KeyError:
            return "Modulo non trovato", 404

    @app.post("/moduli/<slug>/prefill")
    def modulo_prefill(slug):
        try:
            spec = get_pdf_template_spec(slug)
            raw_values = request.form.to_dict(flat=True)
            if any(name in raw_values for name in _CASE_FIELD_DEF_MAP):
                results, manual_values = load_current_dossier_state()
                manual_values.update(extract_manual_case_values(raw_values))
                save_current_dossier_state(results, manual_values)
                case_values = aggregated_fields_to_dict(merge_reviewed_case_fields(results, manual_values))
                values = build_prefill_for_template(spec.key, case_values)
            else:
                values = raw_values
            return render_modulo_form(slug, values=values)
        except KeyError:
            return "Modulo non trovato", 404

    @app.post("/moduli/<slug>/download")
    def modulo_download(slug):
        try:
            spec = get_pdf_template_spec(slug)
            output_path, values, _ = generate_pdf_template(spec.key, request.form.to_dict(flat=True), out_dir)
            return send_file(output_path, as_attachment=True, download_name=output_path.name)
        except KeyError:
            return "Modulo non trovato", 404
        except Exception as exc:
            try:
                return render_modulo_form(slug, values=request.form.to_dict(flat=True), error=str(exc)), 400
            except KeyError:
                return f"Errore generazione modulo: {exc}", 400

    @app.post("/genera")
    def genera():
        form = request.form
        try:
            p = Preventivo(
                cliente=form["cliente"],
                data_nascita=form["data_nascita"],
                tipo_lavoro=form["tipo_lavoro"],
                provincia=form["provincia"],
                note=form.get("note", ""),
                tipo_finanziamento=form["tipo_finanziamento"],
                importo_rata=parse_float_web(form["importo_rata"]),
                durata_mesi=int(form["durata_mesi"]),
                tan=parse_float_web(form["tan"]),
                taeg=parse_float_web(form["taeg"]),
                importo_erogato=parse_float_web(form["importo_erogato"]),
            )
            p.compute()
            p.validate()

            scenari_data = _json.loads(form.get("scenari_json", "[]"))
            preventivi = [p]
            for s in scenari_data:
                tipo_f = s.get("tipo_finanziamento") or p.tipo_finanziamento
                sp = Preventivo(
                    cliente=p.cliente,
                    data_nascita=p.data_nascita,
                    tipo_lavoro=p.tipo_lavoro,
                    provincia=p.provincia,
                    note=p.note,
                    tipo_finanziamento=tipo_f,
                    importo_rata=s["importo_rata"],
                    durata_mesi=s["durata_mesi"],
                    tan=s["tan"],
                    taeg=s["taeg"],
                    importo_erogato=s["importo_erogato"]
                )
                sp.compute()
                sp.validate()
                preventivi.append(sp)
        except (ValueError, KeyError) as exc:
            return render_page(FORM_CONTENT, FORM_SCRIPTS, page="form",
                              title="Errore", error=str(exc)), 400

        # Build JSON list for hidden field (supports single + multi scenario)
        data_json = _json.dumps([
            {
                "cliente": x.cliente,
                "data_nascita": x.data_nascita,
                "tipo_lavoro": x.tipo_lavoro,
                "provincia": x.provincia,
                "note": x.note,
                "tipo_finanziamento": x.tipo_finanziamento,
                "importo_rata": x.importo_rata,
                "durata_mesi": x.durata_mesi,
                "tan": x.tan,
                "taeg": x.taeg,
                "importo_erogato": x.importo_erogato,
            }
            for x in preventivi
        ])

        return render_page(
            PREVIEW_CONTENT,
            page="form",
            title="Riepilogo",
            p=p,
            all_preventivi=preventivi,
            euro_web=euro_web,
            data_json=data_json,
            num_preventivi=len(preventivi),
        )

    @app.post("/download")
    def download():
        try:
            raw = _json.loads(request.form["data_json"])
            items = raw if isinstance(raw, list) else [raw]
            preventivi: list[Preventivo] = []
            for data in items:
                if not isinstance(data, dict):
                    raise ValueError("Dati preventivo non validi.")
                p = Preventivo(
                    cliente=data["cliente"],
                    data_nascita=data["data_nascita"],
                    tipo_lavoro=data["tipo_lavoro"],
                    provincia=data["provincia"],
                    note=data.get("note", ""),
                    tipo_finanziamento=data["tipo_finanziamento"],
                    importo_rata=parse_float_web(data["importo_rata"]),
                    durata_mesi=int(data["durata_mesi"]),
                    tan=parse_float_web(data["tan"]),
                    taeg=parse_float_web(data["taeg"]),
                    importo_erogato=parse_float_web(data["importo_erogato"]),
                )
                p.compute()
                p.validate()
                preventivi.append(p)
            text_overrides = sanitize_text_overrides(
                _json.loads(request.form.get("text_overrides_json", "{}"))
            )
        except (KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
            return f"Dati non validi: {exc}", 400

        prof = load_profile()

        base = preventivi[0]
        if len(preventivi) == 1:
            out = build_output_path(base.cliente, out_dir, base.data_preventivo)
            crea_preventivo_pdf(base, out, profile=prof, text_overrides=text_overrides)
        else:
            out = build_output_path_multi(base.cliente, out_dir, base.data_preventivo, len(preventivi))
            crea_preventivi_pdf(preventivi, out, profile=prof, text_overrides=text_overrides)
        return send_file(out, as_attachment=True, download_name=out.name)

    @app.get("/storico")
    def storico():
        out_dir.mkdir(parents=True, exist_ok=True)
        files = []
        for f in sorted(out_dir.glob("*.pdf"), key=lambda x: x.stat().st_mtime, reverse=True):
            st = f.stat()
            files.append({
                "name": f.name,
                "size": f"{st.st_size // 1024}",
                "date": datetime.fromtimestamp(st.st_mtime).strftime("%d/%m/%Y %H:%M"),
            })
        return render_page(HISTORY_CONTENT, page="history",
                          title="Storico", files=files)

    @app.get("/download-file/<filename>")
    def download_file(filename):
        fpath = resolve_child_file(out_dir, filename, _ALLOWED_DOWNLOAD_EXTS)
        if fpath is None or not fpath.exists():
            return "File non trovato", 404
        return send_file(fpath, as_attachment=True, download_name=fpath.name)

    # ─── Settings Page ───
    SETTINGS_CONTENT = """
    <div class="card">
      <div class="section-title">⚙️ Profilo Agente (Branding)</div>
      <form id="settingsForm" method="post" action="/impostazioni">
        <div class="form-grid">
          <div class="form-group">
            <label>Nome Agente</label>
            <input name="nome_agente" value="{{ s_nome }}" required/>
          </div>
          <div class="form-group">
            <label>Rete Mandante</label>
            <input name="rete_mandante" value="{{ s_rete }}" required/>
          </div>
          <div class="form-group">
            <label>Codice OAM</label>
            <input name="codice_oam" value="{{ s_oam }}" required/>
          </div>
          <div class="form-group">
            <label>Telefono</label>
            <input name="telefono" value="{{ s_tel }}" required/>
          </div>
          <div class="form-group">
            <label>Colore Primario</label>
            <div style="display:flex;gap:8px;align-items:center">
              <input type="color" name="colore_primario" value="{{ s_col1 }}" style="width:48px;height:36px;padding:2px;border:1px solid rgba(255,255,255,.15);border-radius:6px;background:transparent;cursor:pointer"/>
              <code style="color:var(--text-secondary)">{{ s_col1 }}</code>
            </div>
          </div>
          <div class="form-group">
            <label>Colore Accento</label>
            <div style="display:flex;gap:8px;align-items:center">
              <input type="color" name="colore_accento" value="{{ s_col2 }}" style="width:48px;height:36px;padding:2px;border:1px solid rgba(255,255,255,.15);border-radius:6px;background:transparent;cursor:pointer"/>
              <code style="color:var(--text-secondary)">{{ s_col2 }}</code>
            </div>
          </div>
        </div>

        <div class="section-title" style="margin-top:20px">🖼️ Immagini</div>
        <div class="form-grid">
          <div class="form-group">
            <label>Bollino OAM</label>
            <input type="file" accept="image/jpeg,image/png" id="bollino_upload"
                   style="padding:8px;background:rgba(255,255,255,.06);border:1px dashed rgba(255,255,255,.18);border-radius:8px;color:var(--text-primary)"/>
            {% if s_bollino %}
            <img src="/assets/{{ s_bollino }}" alt="Bollino OAM" id="bollino_preview"
                 style="margin-top:8px;max-height:80px;border-radius:6px;border:1px solid rgba(255,255,255,.1)"/>
            {% else %}
            <img id="bollino_preview" style="margin-top:8px;max-height:80px;border-radius:6px;display:none"/>
            {% endif %}
          </div>
          <div class="form-group">
            <label>Logo Agente (opzionale)</label>
            <input type="file" accept="image/jpeg,image/png" id="logo_upload"
                   style="padding:8px;background:rgba(255,255,255,.06);border:1px dashed rgba(255,255,255,.18);border-radius:8px;color:var(--text-primary)"/>
            {% if s_logo %}
            <img src="/assets/{{ s_logo }}" alt="Logo" id="logo_preview"
                 style="margin-top:8px;max-height:80px;border-radius:6px;border:1px solid rgba(255,255,255,.1)"/>
            {% else %}
            <img id="logo_preview" style="margin-top:8px;max-height:80px;border-radius:6px;display:none"/>
            {% endif %}
          </div>
        </div>

        <button type="submit" class="btn btn-primary" style="margin-top:20px;width:100%">Salva Impostazioni</button>
      </form>
    </div>
    """

    SETTINGS_SCRIPTS = """
    <script>
    function uploadImage(inputEl, field, previewEl) {
      inputEl.addEventListener('change', function() {
        const file = this.files[0];
        if (!file) return;
        const fd = new FormData();
        fd.append('file', file);
        fd.append('field', field);
        fetch('/upload-image', {method:'POST', body:fd})
          .then(r => r.json())
          .then(data => {
            if (data.url) {
              previewEl.src = data.url;
              previewEl.style.display = 'block';
              showToast(field === 'bollino' ? 'Bollino aggiornato' : 'Logo aggiornato', 'success');
            }
          })
          .catch(() => showToast('Errore upload', 'error'));
      });
    }
    const bInput = document.getElementById('bollino_upload');
    const bPrev = document.getElementById('bollino_preview');
    const lInput = document.getElementById('logo_upload');
    const lPrev = document.getElementById('logo_preview');
    if (bInput && bPrev) uploadImage(bInput, 'bollino', bPrev);
    if (lInput && lPrev) uploadImage(lInput, 'logo', lPrev);

    function showToast(msg, type) {
      const t = document.getElementById('toast');
      if (!t) return;
      t.textContent = msg;
      t.className = 'toast show ' + (type || '');
      setTimeout(() => t.className = 'toast', 3000);
    }
    {% if saved %}showToast('Impostazioni salvate con successo', 'success');{% endif %}
    </script>
    """

    @app.get("/impostazioni")
    def impostazioni_get():
        prof = load_profile()
        bollino_name = Path(prof.bollino_path).name if prof.bollino_path else ""
        logo_name = Path(prof.logo_path).name if prof.logo_path else ""
        return render_page(SETTINGS_CONTENT, SETTINGS_SCRIPTS, page="settings",
                          title="Impostazioni",
                          s_nome=prof.nome_agente, s_rete=prof.rete_mandante,
                          s_oam=prof.codice_oam, s_tel=prof.telefono,
                          s_col1=prof.colore_primario, s_col2=prof.colore_accento,
                          s_bollino=bollino_name, s_logo=logo_name, saved=False)

    @app.post("/impostazioni")
    def impostazioni_post():
        form = request.form
        prof = load_profile()
        prof.nome_agente = form.get("nome_agente", prof.nome_agente)
        prof.rete_mandante = form.get("rete_mandante", prof.rete_mandante)
        prof.codice_oam = form.get("codice_oam", prof.codice_oam)
        prof.telefono = form.get("telefono", prof.telefono)
        prof.colore_primario = form.get("colore_primario", prof.colore_primario)
        prof.colore_accento = form.get("colore_accento", prof.colore_accento)
        save_profile(prof)
        bollino_name = Path(prof.bollino_path).name if prof.bollino_path else ""
        logo_name = Path(prof.logo_path).name if prof.logo_path else ""
        return render_page(SETTINGS_CONTENT, SETTINGS_SCRIPTS, page="settings",
                          title="Impostazioni",
                          s_nome=prof.nome_agente, s_rete=prof.rete_mandante,
                          s_oam=prof.codice_oam, s_tel=prof.telefono,
                          s_col1=prof.colore_primario, s_col2=prof.colore_accento,
                          s_bollino=bollino_name, s_logo=logo_name, saved=True)

    @app.post("/upload-image")
    def upload_image():
        field = request.form.get("field", "")
        file = request.files.get("file")
        if not file or field not in ("bollino", "logo"):
            return jsonify({"error": "Parametri non validi"}), 400
        ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
        if ext not in _ALLOWED_IMAGE_EXTS:
            return jsonify({"error": "Formato non supportato. Usa JPG o PNG."}), 400
        _ensure_assets_dir()
        safe_name = f"bollino_oam.{ext}" if field == "bollino" else f"logo_agente.{ext}"
        save_path = _ASSETS_DIR / safe_name
        file.save(str(save_path))
        prof = load_profile()
        if field == "bollino":
            prof.bollino_path = safe_name
        else:
            prof.logo_path = safe_name
        save_profile(prof)
        return jsonify({"url": f"/assets/{safe_name}", "field": field})

    @app.get("/assets/<path:filename>")
    def serve_asset(filename):
        fpath = resolve_child_file(_ASSETS_DIR, filename, _ALLOWED_IMAGE_EXTS)
        if fpath is None or not fpath.exists():
            return "File non trovato", 404
        return send_file(fpath)

    @app.get("/esci")
    def app_exit():
        if not _is_frozen_runtime():
            return redirect(url_for("home"))
        threading.Timer(0.4, lambda: os._exit(0)).start()
        return (
            "<html><head><meta charset='utf-8'><title>QuintoQuote</title></head>"
            "<body style='font-family:Segoe UI,Arial,sans-serif;padding:32px;background:#0f172a;color:#e5e7eb'>"
            "<h2 style='margin-top:0'>QuintoQuote si sta chiudendo...</h2>"
            "<p>Puoi chiudere anche questa scheda del browser.</p>"
            "</body></html>"
        )

    display_host = "127.0.0.1" if host in ("0.0.0.0", "::") else host
    app_url = f"http://{display_host}:{port}"
    if open_browser:
        def _open_browser() -> None:
            try:
                webbrowser.open(app_url, new=2)
            except Exception:
                pass
        threading.Timer(0.8, _open_browser).start()
        print("\nQuintoQuote pronto.")
        print(f"Clicca qui per l'applicazione: {app_url}")
        print("Se il browser non si apre da solo, copia/incolla il link sopra.")
        print("Premi CTRL+C per chiudere il server.\n")
    else:
        print("\nQuintoQuote pronto.")
        print(f"Apri l'applicazione qui: {app_url}")
        if show_start_hint:
            print("Suggerimento: dopo pip install -e . puoi avviare tutto con QuintoQuote start.")
        print("Premi CTRL+C per chiudere il server.\n")
    app.run(host=host, port=port, debug=False)

# =========================
#  MAIN
# =========================

def main():
    ap = argparse.ArgumentParser(
        description="Generatore Preventivi PDF (CLI o Web localhost).",
        epilog="Avvio rapido: QuintoQuote start | quintoquote start | python .\\quintoquote.py start",
    )
    ap.add_argument("--web", action="store_true", help="Avvia la GUI web in localhost (richiede flask).")
    ap.add_argument("--host", default="127.0.0.1", help="Host bind per la web UI.")
    ap.add_argument("--port", type=int, default=5000, help="Porta della web UI.")
    ap.add_argument("--config-path", default="", help="Percorso del file config JSON da usare.")
    ap.add_argument("--assets-dir", default="", help="Cartella assets da usare per upload logo/bollino.")

    ap.add_argument("--out-dir", default=str(_default_output_dir()), help="Cartella di output PDF.")
    ap.add_argument("--non-interactive", action="store_true", help="Non chiedere input: usa solo gli argomenti.")

    # argomenti (facoltativi in interattivo, obbligatori in non-interactive)
    ap.add_argument("--cliente", default="")
    ap.add_argument("--data-nascita", default="")
    ap.add_argument("--tipo-lavoro", default="")
    ap.add_argument("--provincia", default="")
    ap.add_argument("--note", default="")
    ap.add_argument("--tipo-finanziamento", default="Cessione del Quinto")

    ap.add_argument("--importo-rata", type=float, default=None)
    ap.add_argument("--durata-mesi", type=int, default=None)
    ap.add_argument("--tan", type=float, default=None)
    ap.add_argument("--taeg", type=float, default=None)
    ap.add_argument("--importo-erogato", type=float, default=None)
    ap.add_argument(
        "--scenario",
        action="append",
        default=[],
        help=(
            "Scenario aggiuntivo nel formato "
            "rata;durata;tan;taeg;importo_erogato[;tipo_finanziamento]. "
            "Ripeti --scenario per aggiungerne altri."
        ),
    )

    argv, used_start_alias = normalize_cli_argv(sys.argv)
    args = ap.parse_args(argv[1:])
    out_dir = Path(args.out_dir)
    if not (1 <= args.port <= 65535):
        ap.exit(2, "Errore: --port deve essere compresa tra 1 e 65535.\n")

    config_path = Path(args.config_path) if args.config_path else None
    assets_dir = Path(args.assets_dir) if args.assets_dir else None
    configure_runtime_paths(config_path=config_path, assets_dir=assets_dir)

    try:
        if args.web:
            selected_port = _pick_available_port(args.host, args.port) if (used_start_alias or _is_frozen_runtime()) else args.port
            run_web(
                out_dir,
                host=args.host,
                port=selected_port,
                open_browser=used_start_alias,
                show_start_hint=not used_start_alias,
            )
            return

        # CLI
        p = collect_from_cli(args)
        extra_lines: list[str] = []
        for raw in args.scenario:
            extra_lines.extend(parse_scenari_text(raw))

        preventivi = [p] + build_extra_preventivi(p, extra_lines)
        if len(preventivi) == 1:
            out = build_output_path(p.cliente, out_dir, p.data_preventivo)
            crea_preventivo_pdf(p, out)
        else:
            out = build_output_path_multi(p.cliente, out_dir, p.data_preventivo, len(preventivi))
            crea_preventivi_pdf(preventivi, out)

        print("\nPreventivo generato con successo")
        print(f"  File: {out}")
        print(f"  Cliente: {p.cliente}")
        if len(preventivi) == 1:
            print(f"  Rata: {euro(p.importo_rata)} x {p.durata_mesi} mesi")
            print(f"  Montante: {euro(p.montante)}")
            print(f"  Interessi: {euro(p.interessi)}")
            print(f"  TAEG: {p.taeg:.3f}%")
        else:
            print(f"  Preventivi inclusi nel PDF: {len(preventivi)}")
    except KeyboardInterrupt:
        ap.exit(130, "\nOperazione annullata.\n")
    except ValueError as exc:
        ap.exit(2, f"Errore: {exc}\n")

if __name__ == "__main__":
    main()

