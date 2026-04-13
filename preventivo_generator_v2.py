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
import threading
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

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

_CONFIG_PATH = Path(__file__).parent / "config.json"
_ASSETS_DIR = Path(__file__).parent / "assets"
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
    _ASSETS_DIR.mkdir(exist_ok=True)
    return _ASSETS_DIR


def configure_runtime_paths(config_path: Optional[Path] = None, assets_dir: Optional[Path] = None) -> None:
    global _CONFIG_PATH, _ASSETS_DIR, _cached_profile, _cached_profile_path
    if config_path is not None:
        _CONFIG_PATH = config_path.expanduser().resolve()
    if assets_dir is not None:
        _ASSETS_DIR = assets_dir.expanduser().resolve()
    _cached_profile = None
    _cached_profile_path = None


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
#  WEB (OPZIONALE) - localhost
# =========================

def run_web(out_dir: Path, host: str = "127.0.0.1", port: int = 5000):
    try:
        from flask import Flask, request, render_template_string, send_file, redirect, url_for, jsonify
    except Exception as e:
        print("Flask non installato. Per usare --web fai: pip install flask")
        raise
    import json as _json

    app = Flask(__name__)

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
          <a href="/storico" class="{{ 'active' if page == 'history' }}">📂 Storico</a>
          <a href="/impostazioni" class="{{ 'active' if page == 'settings' }}">⚙️ Impostazioni</a>
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
      <div class="section-title">📂 Storico Preventivi</div>
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
        <p>Nessun preventivo generato ancora.</p>
        <p style="margin-top:8px;"><a href="/" style="color:var(--accent-light);">Crea il primo preventivo →</a></p>
      </div>
      {% endif %}
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
        content_html = render_template_string(content_tpl, **kwargs)
        scripts_html = render_template_string(scripts_tpl, **kwargs) if scripts_tpl else ""
        return render_template_string(BASE_HTML, content=content_html, scripts=scripts_html, **kwargs)

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

    print(f"\nWeb UI attiva su http://{host}:{port}\n")
    app.run(host=host, port=port, debug=False)

# =========================
#  MAIN
# =========================

def main():
    ap = argparse.ArgumentParser(description="Generatore Preventivi PDF (CLI o Web localhost).")
    ap.add_argument("--web", action="store_true", help="Avvia la GUI web in localhost (richiede flask).")
    ap.add_argument("--host", default="127.0.0.1", help="Host bind per la web UI.")
    ap.add_argument("--port", type=int, default=5000, help="Porta della web UI.")
    ap.add_argument("--config-path", default="", help="Percorso del file config JSON da usare.")
    ap.add_argument("--assets-dir", default="", help="Cartella assets da usare per upload logo/bollino.")

    ap.add_argument("--out-dir", default="output_preventivi", help="Cartella di output PDF.")
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

    args = ap.parse_args()
    out_dir = Path(args.out_dir)
    if not (1 <= args.port <= 65535):
        ap.exit(2, "Errore: --port deve essere compresa tra 1 e 65535.\n")

    config_path = Path(args.config_path) if args.config_path else None
    assets_dir = Path(args.assets_dir) if args.assets_dir else None
    configure_runtime_paths(config_path=config_path, assets_dir=assets_dir)

    try:
        if args.web:
            run_web(out_dir, host=args.host, port=args.port)
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

