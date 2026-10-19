"""
Reporte diario Ajover EXPO Vacíos — Interno
Modeled after enviar_reporte.py (llenos).
.env: EMAIL_FROM, EMAIL_PASSWORD, EMAIL_TO_INTERNO, EMAIL_CC_INTERNO
"""
import os, json, smtplib, sys, base64
from collections import Counter, defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

BASE = Path(__file__).parent
load_dotenv(BASE / ".env")

# ── Firma embebida ────────────────────────────────────────────────────────────
_SIG_DIR = Path(os.environ.get("APPDATA","")) / "Microsoft" / "Signatures" / "jeffer arias1_archivos"

def _sig_b64(name):
    try: return base64.b64encode((_SIG_DIR / name).read_bytes()).decode()
    except: return ""

def _imgtag(b64, w, h, alt=""):
    if not b64: return f'<span style="display:inline-block;width:{w}px;height:{h}px"></span>'
    return f'<img src="data:image/png;base64,{b64}" width="{w}" height="{h}" alt="{alt}" style="display:block;border:0">'

_b64_logo = _sig_b64("image001.png")
_b64_li   = _sig_b64("image002.png")
_b64_ig   = _sig_b64("image003.png")
_b64_fb   = _sig_b64("image004.png")

FIRMA_HTML = f"""
<table cellpadding="0" cellspacing="0" style="margin:0">
  <tr>
    <td style="padding-right:14px;border-right:2px solid #F5811E;vertical-align:top">
      {_imgtag(_b64_logo, 120, 120, "TC Logistix")}
    </td>
    <td style="padding-left:14px;vertical-align:top">
      <p style="margin:0 0 2px;font-family:Helvetica,Arial,sans-serif;font-weight:700;color:#ED7D31;font-size:13px">Jeffer Antonio Arias Ordosgoitia</p>
      <p style="margin:0 0 2px;font-family:Arial,sans-serif;font-weight:700;font-style:italic;color:#002060;font-size:10px">Analista de Datos</p>
      <p style="margin:0;font-family:Arial,sans-serif;font-style:italic;color:#002060;font-size:9px">Dirección: Manga Cra. 25 No. 24A-16 Piso 17</p>
      <p style="margin:0;font-family:Arial,sans-serif;font-style:italic;color:#002060;font-size:9px">Edificio Twin Bays, Torre Bancolombia</p>
      <p style="margin:0;font-family:Arial,sans-serif;font-style:italic;color:#091140;font-size:10px">Tel: 3115993117</p>
      <p style="margin:0;font-family:Arial,sans-serif;font-style:italic;color:#091140;font-size:10px">Cartagena, Colombia</p>
      <p style="margin:0;font-size:10px;font-family:Arial,sans-serif">Email: <a href="mailto:jarias@tractocar.com" style="color:blue;font-style:italic">jarias@tractocar.com</a></p>
      <p style="margin:0"><a href="http://www.tractocar.com/" style="font-family:Arial,sans-serif;font-weight:700;font-style:italic;color:#002060;font-size:10px;text-decoration:none">www.tractocar.com</a></p>
      <table cellpadding="0" cellspacing="0" style="margin-top:5px"><tr>
        <td style="padding-right:5px"><a href="https://www.linkedin.com/company/tractocar-logistics-sas/">{_imgtag(_b64_li, 19, 19, "LinkedIn")}</a></td>
        <td style="padding-right:5px">{_imgtag(_b64_ig, 21, 21, "Instagram")}</td>
        <td>{_imgtag(_b64_fb, 10, 21, "Facebook")}</td>
      </tr></table>
    </td>
  </tr>
</table>
"""

EMAIL_FROM        = os.environ.get("EMAIL_FROM", "jarias@tractocar.com")
EMAIL_PASSWORD    = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_TO_INTERNO  = [e.strip() for e in os.environ.get("EMAIL_TO_INTERNO", EMAIL_FROM).split(",") if e.strip()]
EMAIL_CC_INTERNO  = [e.strip() for e in os.environ.get("EMAIL_CC_INTERNO","").split(",") if e.strip()]
SMTP_HOST         = os.environ.get("SMTP_HOST","smtp.office365.com")
SMTP_PORT         = int(os.environ.get("SMTP_PORT","587"))

MESES_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
            'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
MESES_C  = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
DIAS_ES  = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']

def _fecha_es(dt):
    return f"{DIAS_ES[dt.weekday()]} {dt.day:02d} de {MESES_ES[dt.month-1].lower()} de {dt.year}"

def _ml(iso, c=False):
    try:
        p = iso.split("-"); n = int(p[1])-1
        return f"{(MESES_C if c else MESES_ES)[n]} {p[0]}"
    except: return iso

def _pct(a, b): return round(a/b*100, 1) if b else 0.0

def _col(pct):
    if pct >= 80: return {"bar":"#22c55e","light":"#dcfce7","text":"#15803d","dark":"#14532d"}
    if pct >= 60: return {"bar":"#f59e0b","light":"#fef3c7","text":"#b45309","dark":"#78350f"}
    return              {"bar":"#ef4444","light":"#fee2e2","text":"#b91c1c","dark":"#7f1d1d"}

# ── Helpers de fila ───────────────────────────────────────────────────────────
def _acumple(r):   return (r.get("cumpl") or "") == "A tiempo"
def _tarde(r):     return (r.get("cumpl") or "").startswith("Tarde")
def _sf(r):        return not (r.get("cumpl") or "").strip() or (r.get("cumpl") or "") in ("Sin fecha","Pendiente")
def _acumple_a(r): return (r.get("cumpl_aten") or "") == "A tiempo"
def _tarde_a(r):   return (r.get("cumpl_aten") or "").startswith("Tarde")
def _sf_a(r):      return not (r.get("cumpl_aten") or "").strip() or (r.get("cumpl_aten") or "") in ("Sin fecha","Pendiente","")

def _pill_cumpl(val):
    if not val or val in ("Sin fecha","Pendiente",""):
        return '<span style="background:#f1f5f9;color:#94a3b8;padding:2px 8px;border-radius:6px;font-size:10px;border:1px solid #e2e8f0">Sin fecha</span>'
    ok = val == "A tiempo"
    if ok:
        return f'<span style="background:#dcfce7;color:#15803d;padding:2px 8px;border-radius:6px;font-size:10px;font-weight:700;border:1px solid #86efac">&#10003; {val}</span>'
    return f'<span style="background:#fee2e2;color:#b91c1c;padding:2px 8px;border-radius:6px;font-size:10px;font-weight:700;border:1px solid #fca5a5">&#10007; {val}</span>'

# ── Barra vertical ────────────────────────────────────────────────────────────
def _vbar(pct, col, active, max_h=110, w=38):
    bh = max(4, int(pct / 100 * max_h))
    sh = max_h - bh
    border = f'border-left:2px solid {col};border-right:2px solid {col};border-top:2px solid {col};' if active else ''
    return f'''<table cellpadding="0" cellspacing="0" align="center" style="margin:0 auto">
      <tr><td width="{w}" height="{sh}" style="font-size:0;line-height:0"> </td></tr>
      <tr><td width="{w}" height="{bh}" bgcolor="{col}" style="background:{col};{border}font-size:0;line-height:0"> </td></tr>
    </table>'''

def _fmt_dt(v):
    """Formatea minutos signed como +2h 5min / -30min."""
    if v is None: return "—"
    h = int(abs(v) // 60); mn = round(abs(v) % 60)
    txt = (v <= 0 and "-" or "+") + (f"{h}h {mn}min" if h else f"{mn}min")
    col = "#22c55e" if v <= 0 else ("#f59e0b" if v <= 30 else "#ef4444")
    return f'<span style="color:{col};font-weight:700">{txt}</span>'

def _tc(col="#334155", mono=False, bold=False, wrap=True):
    ff = "font-family:monospace;" if mono else ""
    fw = "font-weight:700;" if bold else ""
    ws = "" if wrap else "white-space:nowrap;"
    return f'<td style="padding:5px 8px;font-size:10px;color:{col};border-bottom:1px solid #f1f5f9;vertical-align:top;{ff}{fw}{ws}">'


# ─────────────────────────────────────────────────────────────────────────────
def build_html(datos):
    rows     = datos["ajover"]["vacios"]["rows"]
    hoy      = datetime.now()
    generado = datos.get("generado","")[:10]

    corte       = hoy - timedelta(days=1)
    mes_actual  = corte.strftime("%Y-%m")
    corte_label = f"{corte.day} de {MESES_ES[corte.month-1].lower()} de {corte.year}"

    def _parse_dt(s):
        s = (s or "").strip()
        if not s: return None
        try: return datetime.strptime(s, "%d-%m-%Y %H:%M")
        except: return None

    def _r_fprog(r):
        try:
            f = (r.get("fprog","") or "").strip().split(" ")[0]
            d, m, y = f.split("-")
            return datetime(int(y), int(m), int(d)).date()
        except: return None

    # ── Mes actual
    mes_rows = [r for r in rows
                if r.get("mes") == mes_actual
                and (_r_fprog(r) is None or _r_fprog(r) <= corte.date())]
    total_m   = len(mes_rows)
    sf_m      = sum(1 for r in mes_rows if _sf(r))
    base_m    = total_m - sf_m
    cumpl_m   = sum(1 for r in mes_rows if _acumple(r))
    tarde_m   = sum(1 for r in mes_rows if _tarde(r))
    pct_m     = _pct(cumpl_m, base_m)
    col_m     = _col(pct_m)

    sf_a_m    = sum(1 for r in mes_rows if _sf_a(r))
    base_a_m  = total_m - sf_a_m
    cumpl_a_m = sum(1 for r in mes_rows if _acumple_a(r))
    pct_a_m   = _pct(cumpl_a_m, base_a_m)
    col_a_m   = _col(pct_a_m)

    # Motivos mes
    mot_m = Counter((r.get("motivo_incumpl","") or "Sin motivo").strip() or "Sin motivo"
                    for r in mes_rows if _tarde(r))

    # ── Tendencia últimos 7 meses (solo meses con ≥5 registros con cita)
    meses_todos = sorted(set(r["mes"] for r in rows if r.get("mes") and r["mes"] <= mes_actual))
    tend = []
    for mes in meses_todos:
        mr = mes_rows if mes == mes_actual else [r for r in rows if r.get("mes") == mes]
        sf = sum(1 for r in mr if _sf(r))
        base = len(mr) - sf
        if base < 5: continue   # skip meses sin suficiente cita (Excel histórico)
        cumpl = sum(1 for r in mr if _acumple(r))
        tend.append({"mes": mes, "total": len(mr), "pct": _pct(cumpl, base)})
    tend = tend[-7:]  # últimos 7 con cita

    # Gráfico tendencia
    f_pct = f_bar = f_lbl = ""
    for t in tend:
        c      = _col(t["pct"])
        active = t["mes"] == mes_actual
        lbl    = _ml(t["mes"], True).split(" ")
        mes_c  = lbl[0] + ("'" + lbl[1][2:] if len(lbl) > 1 else "")
        fw_a   = "font-weight:800;" if active else "font-weight:600;"
        f_pct += f'<td width="76" style="text-align:center;padding-bottom:6px;font-size:{"11" if active else "10"}px;{fw_a}color:{c["text"]}">{t["pct"]}%</td>'
        f_bar += f'<td width="76" style="text-align:center;vertical-align:bottom">{_vbar(t["pct"], c["bar"], active)}</td>'
        f_lbl += f'<td width="76" style="text-align:center;padding-top:6px;border-top:2px solid #e2e8f0"><span style="font-size:{"11" if active else "9"}px;{fw_a}color:{"#1e293b" if active else "#64748b"}">{mes_c}</span><br><span style="font-size:9px;color:#94a3b8">{t["total"]} vaciós</span></td>'

    chart_html = f'''<table cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse">
      <tr>{f_pct}</tr>
      <tr>{f_bar}</tr>
      <tr>{f_lbl}</tr>
    </table>'''

    # ── Ayer (busca último día con registros por fprog)
    dia_rows, dia_dt = [], None
    for delta in range(1, 9):
        dt  = hoy - timedelta(days=delta)
        ds  = dt.strftime("%d-%m-%Y")
        found = [r for r in rows if (r.get("fprog","") or "").startswith(ds)]
        if found: dia_rows, dia_dt = found, dt; break

    dia_label = _fecha_es(dia_dt) if dia_dt else "Sin datos recientes"
    dia_total = len(dia_rows)
    dia_sf    = sum(1 for r in dia_rows if _sf(r))
    dia_cumpl = sum(1 for r in dia_rows if _acumple(r))
    dia_tarde = sum(1 for r in dia_rows if _tarde(r))
    dia_pct   = _pct(dia_cumpl, dia_total - dia_sf)
    col_d     = _col(dia_pct)

    dia_sf_a    = sum(1 for r in dia_rows if _sf_a(r))
    dia_cumpl_a = sum(1 for r in dia_rows if _acumple_a(r))
    dia_tarde_a = sum(1 for r in dia_rows if _tarde_a(r))
    dia_pct_a   = _pct(dia_cumpl_a, dia_total - dia_sf_a)

    # ── Alertas: contenedores sin fcita_retiro que tenían registro
    mes_sin_cita = [r for r in mes_rows
                    if not (r.get("fcita_retiro","") or "").strip()
                    and (r.get("fllpatio","") or "").strip()]

    def _alerta_card(icon, label, n, desc, bg, fg, border, rows_list=None):
        rows_html = ""
        if rows_list:
            for r in rows_list[:5]:
                rows_html += f'<tr><td style="padding:2px 0 2px 12px;font-size:9px;color:#475569;font-family:monospace">{r.get("cont","—")} &nbsp;<span style="color:#94a3b8">{r.get("fprog","") or ""}</span></td></tr>'
        rows_block = f'<table cellpadding="0" cellspacing="0" style="margin-top:4px">{rows_html}</table>' if rows_html else ""
        return f'''<table cellpadding="0" cellspacing="0" width="100%" style="margin-bottom:8px">
          <tr>
            <td width="4" bgcolor="{fg}" style="background:{fg};border-radius:4px 0 0 4px;font-size:0"> </td>
            <td bgcolor="{bg}" style="background:{bg};padding:10px 14px;border:1px solid {border};border-left:none;border-radius:0 8px 8px 0">
              <table width="100%" cellpadding="0" cellspacing="0"><tr>
                <td><span style="font-size:10px;font-weight:700;color:{fg}">{icon} {label}</span>
                <p style="margin:2px 0 0;font-size:9px;color:#64748b">{desc}</p>{rows_block}</td>
                <td align="right" valign="top"><span style="font-size:22px;font-weight:900;color:{fg}">{n}</span></td>
              </tr></table>
            </td>
          </tr>
        </table>'''

    alertas_items = ""
    n_alertas = 0
    if mes_sin_cita:
        n_alertas += len(mes_sin_cita)
        alertas_items += _alerta_card("&#9888;", f"{len(mes_sin_cita)} contenedor{'es' if len(mes_sin_cita)!=1 else ''} sin cita en patio registrada",
                                      len(mes_sin_cita), "Tienen llegada a patio pero sin fecha de cita en patio.",
                                      "#fef9c3","#b45309","#fde68a", mes_sin_cita)
    if not alertas_items:
        alertas_items = _alerta_card("&#10003;", "Sin alertas operativas", "",
                                     "Todos los registros tienen cita en patio registrada.",
                                     "#dcfce7","#15803d","#86efac")

    # ── Detalle ayer — ordenado: tarde > sin fecha > a tiempo
    def _sort_key(r):
        if _tarde(r):   return 0
        if _sf(r):      return 1
        return 2

    TH_S = 'style="padding:7px 10px;font-size:9px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.06em;text-align:left;background:#f8fafc;border-bottom:2px solid #e2e8f0"'
    GRUPO_LABEL = {
        0: ("Tarde — no cumplió cita", "#fee2e2", "#dc2626"),
        1: ("Sin fecha de cita en patio", "#f1f5f9", "#64748b"),
        2: ("A tiempo", "#dcfce7", "#15803d"),
    }
    det_ayer = ""
    prev_g   = None
    NCOLS    = 9
    for r in sorted(dia_rows, key=_sort_key):
        g = _sort_key(r)
        if g != prev_g:
            lbl, bg, fg = GRUPO_LABEL[g]
            det_ayer += f'<tr><td colspan="{NCOLS}" bgcolor="{bg}" style="background:{bg};padding:4px 10px;font-size:9px;font-weight:700;color:{fg};text-transform:uppercase;letter-spacing:.1em;border-bottom:1px solid #e2e8f0">{lbl}</td></tr>'
            prev_g = g
        dt1_html = _fmt_dt(r.get("dt1"))
        dt2_html = _fmt_dt(r.get("dt2"))
        mot  = (r.get("motivo_incumpl","") or "—")
        resp = (r.get("resp_incumpl","")  or "—")
        cumpl_a_html = _pill_cumpl(r.get("cumpl_aten",""))
        det_ayer += f'''<tr>
          {_tc(mono=True)}              {r.get("cont","—")}</td>
          {_tc(col="#64748b")}          {r.get("patio","—")}</td>
          {_tc(col="#64748b",wrap=False)}{r.get("fcita_retiro","") or "—"}</td>
          {_tc(col="#64748b",wrap=False)}{r.get("fllpatio","") or "—"}</td>
          <td style="padding:5px 8px;font-size:10px;border-bottom:1px solid #f1f5f9">{dt1_html}</td>
          {_tc(col="#64748b",wrap=False)}{r.get("faten","") or "—"}</td>
          <td style="padding:5px 8px;font-size:10px;border-bottom:1px solid #f1f5f9">{dt2_html}</td>
          {_tc(col="#b45309")}          {mot}</td>
          {_tc(col="#64748b")}          {resp}</td>
        </tr>'''

    # ── Motivos mes (tardíos)
    mot_rows_html = ""
    if mot_m:
        max_m = max(mot_m.values(), default=1)
        BAR_W = 200
        for mot, cnt in mot_m.most_common(8):
            w = max(4, int(cnt / max_m * BAR_W))
            e = BAR_W - w
            ep = f'<td width="{e}" height="9" bgcolor="#f1f5f9" style="background:#f1f5f9;font-size:0"> </td>' if e else ""
            mot_rows_html += f'''<tr>
              <td style="padding:5px 10px 5px 0;font-size:10px;color:#475569;width:160px;vertical-align:middle">{mot}</td>
              <td style="padding:5px 0;vertical-align:middle">
                <table cellpadding="0" cellspacing="0"><tr>
                  <td width="{w}" height="9" bgcolor="#f59e0b" style="background:#f59e0b;font-size:0"> </td>
                  {ep}
                </tr></table>
              </td>
              <td style="padding:5px 0 5px 10px;font-size:11px;font-weight:700;color:#b45309;white-space:nowrap;vertical-align:middle">{cnt}</td>
            </tr>'''

    SEC = lambda label, col="#2563eb": f'<p style="margin:0 0 12px;font-size:8px;font-weight:700;color:{col};text-transform:uppercase;letter-spacing:.18em">{label}</p>'

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Indicador Cumplimiento Citas — Vacíos</title></head>
<body style="margin:0;padding:20px 0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif">
<table width="760" cellpadding="0" cellspacing="0" align="center" style="max-width:760px">

  <!-- ═══ HEADER ═══════════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#071626" style="background:linear-gradient(160deg,#0a1f35 0%,#071626 60%,#050f1a 100%);border-radius:14px 14px 0 0;padding:28px 32px 0">
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px"><tr>
      <td>
        <table cellpadding="0" cellspacing="0"><tr>
          <td bgcolor="#1e3a5f" style="background:#1e3a5f;border-radius:8px;padding:5px 14px;border:1px solid #2563eb22">
            <span style="font-size:14px;font-weight:900;color:#fff;letter-spacing:-.5px">TC</span>
            <span style="font-size:9px;color:#60a5fa;margin-left:8px;letter-spacing:.06em;font-weight:600">TRACTOCAR LOGISTICS</span>
          </td>
        </tr></table>
        <p style="margin:18px 0 4px;font-size:8px;font-weight:700;color:#3b82f6;text-transform:uppercase;letter-spacing:.22em">Reporte Ejecutivo &nbsp;·&nbsp; Ajover EXPO</p>
        <p style="margin:0 0 6px;font-size:24px;font-weight:900;color:#f8fafc;line-height:1.15;letter-spacing:-.03em">Indicador de Cumplimiento de Citas <span style="color:#38bdf8">—</span> Vacíos</p>
        <p style="margin:0;font-size:11px;color:rgba(255,255,255,0.38);letter-spacing:.01em">{_ml(mes_actual)} &nbsp;·&nbsp; A corte: {corte_label} &nbsp;·&nbsp; {total_m} vaciós</p>
      </td>
      <td align="right" valign="top" style="font-size:10px;color:rgba(255,255,255,0.2);white-space:nowrap;padding-left:16px;padding-top:4px">{generado}</td>
    </tr></table>
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="40" height="2" bgcolor="#38bdf8" style="background:#38bdf8;font-size:0;border-radius:2px"> </td>
      <td height="2" bgcolor="#1e3a5f22" style="background:#1e3a5f22;font-size:0"> </td>
    </tr></table>
  </td></tr>

  <!-- ═══ ALERTAS OPERATIVAS (solo interno) ══════════════════════════════ -->
  <tr><td bgcolor="{'#fff7ed' if n_alertas else '#f0fdf4'}" style="background:{'#fff7ed' if n_alertas else '#f0fdf4'};padding:14px 32px;border-left:4px solid {'#f97316' if n_alertas else '#16a34a'};border-bottom:2px solid {'#f97316' if n_alertas else '#16a34a'}">
    <p style="margin:0 0 4px;font-size:10px;font-weight:700;color:{'#9a3412' if n_alertas else '#15803d'};text-transform:uppercase;letter-spacing:.08em">{'⚠️ Alertas del mes' if n_alertas else '✅ Sin alertas'} — Para: comex2@tractocar.com</p>
    <p style="margin:0 0 10px;font-size:10px;color:#64748b">{'Por favor completar en la web app los siguientes registros que tienen información pendiente:' if n_alertas else 'Todos los registros del mes están completos.'}</p>
    {alertas_items}
  </td></tr>

  <!-- ═══ 4 KPI CARDS ═════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#0d2137" style="background:#0d2137;padding:0 32px 28px">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <!-- Total -->
        <td width="25%" style="padding-right:5px;vertical-align:top">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td height="3" bgcolor="#475569" style="background:#475569;border-radius:3px 3px 0 0;font-size:0"> </td></tr>
            <tr><td bgcolor="#0f1e2e" style="background:#0f1e2e;padding:14px 14px 16px;border:1px solid #1e2d3d;border-top:none;border-radius:0 0 10px 10px">
              <p style="margin:0 0 10px;font-size:8px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.12em">Total Vacíos</p>
              <p style="margin:0;font-size:34px;font-weight:900;color:#f1f5f9;line-height:1;letter-spacing:-.02em">{total_m}</p>
              <p style="margin:10px 0 0;font-size:9px;color:#475569;border-top:1px solid #1e2d3d;padding-top:8px">del mes</p>
            </td></tr>
          </table>
        </td>
        <!-- Cumpl. Vehículo -->
        <td width="25%" style="padding:0 3px;vertical-align:top">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td height="3" bgcolor="{col_m['bar']}" style="background:{col_m['bar']};border-radius:3px 3px 0 0;font-size:0"> </td></tr>
            <tr><td bgcolor="#0a1f12" style="background:#0a1f12;padding:14px 14px 16px;border:1px solid #14532d;border-top:none;border-radius:0 0 10px 10px">
              <p style="margin:0 0 10px;font-size:8px;font-weight:700;color:#16a34a;text-transform:uppercase;letter-spacing:.12em">Cumpl. Vehículo</p>
              <p style="margin:0;font-size:34px;font-weight:900;color:{col_m['bar']};line-height:1;letter-spacing:-.02em">{pct_m}%</p>
              <p style="margin:10px 0 0;font-size:9px;color:#166534;border-top:1px solid #14532d;padding-top:8px">{cumpl_m} ok &nbsp;·&nbsp; {tarde_m} tarde</p>
            </td></tr>
          </table>
        </td>
        <!-- Cumpl. Atención -->
        <td width="25%" style="padding:0 3px;vertical-align:top">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td height="3" bgcolor="{col_a_m['bar']}" style="background:{col_a_m['bar']};border-radius:3px 3px 0 0;font-size:0"> </td></tr>
            <tr><td bgcolor="#0c1929" style="background:#0c1929;padding:14px 14px 16px;border:1px solid #1e3a5f;border-top:none;border-radius:0 0 10px 10px">
              <p style="margin:0 0 10px;font-size:8px;font-weight:700;color:#60a5fa;text-transform:uppercase;letter-spacing:.12em">Cumpl. Atención</p>
              <p style="margin:0;font-size:34px;font-weight:900;color:{col_a_m['bar']};line-height:1;letter-spacing:-.02em">{pct_a_m}%</p>
              <p style="margin:10px 0 0;font-size:9px;color:#3b82f6;border-top:1px solid #1e3a5f;padding-top:8px">{cumpl_a_m} ok &nbsp;·&nbsp; {sum(1 for r in mes_rows if _tarde_a(r))} tarde</p>
            </td></tr>
          </table>
        </td>
        <!-- Sin fecha -->
        <td width="25%" style="padding-left:5px;vertical-align:top">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td height="3" bgcolor="#818cf8" style="background:#818cf8;border-radius:3px 3px 0 0;font-size:0"> </td></tr>
            <tr><td bgcolor="#0f0d1f" style="background:#0f0d1f;padding:14px 14px 16px;border:1px solid #312e6e;border-top:none;border-radius:0 0 10px 10px">
              <p style="margin:0 0 10px;font-size:8px;font-weight:700;color:#818cf8;text-transform:uppercase;letter-spacing:.12em">Sin Cita</p>
              <p style="margin:0;font-size:34px;font-weight:900;color:#a5b4fc;line-height:1;letter-spacing:-.02em">{sf_m}</p>
              <p style="margin:10px 0 0;font-size:9px;color:#4c1d95;border-top:1px solid #312e6e;padding-top:8px">sin cita en patio</p>
            </td></tr>
          </table>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- ═══ TENDENCIA MES A MES ═════════════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:24px 32px;border-top:1px solid #e2e8f0">
    {SEC("Cumpl. vehículo — mes a mes")}
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f8fafc" style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px">
      <tr><td style="padding:16px 8px 10px">{chart_html}</td></tr>
    </table>
    <p style="margin:10px 0 0;font-size:9px;color:#94a3b8">
      <span style="color:#22c55e">&#9646;</span> &ge;80%&nbsp;&nbsp;
      <span style="color:#f59e0b">&#9646;</span> 60–80%&nbsp;&nbsp;
      <span style="color:#ef4444">&#9646;</span> &lt;60%&nbsp;&nbsp;&nbsp;
      Barra más resaltada = mes actual
    </p>
  </td></tr>

  <!-- ═══ MOTIVOS INCUMPLIMIENTO MES ══════════════════════════════════════ -->
  {('<tr><td bgcolor="#ffffff" style="background:#fff;padding:20px 32px;border-top:1px solid #e2e8f0">' + SEC("Distribución de motivos de incumplimiento", "#dc2626") + '<table cellpadding="0" cellspacing="4">' + mot_rows_html + '</table></td></tr>') if mot_rows_html else ""}

  <!-- ═══ DIVIDER AYER ════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#0f3460" style="background:#0f3460;padding:11px 32px">
    <p style="margin:0 0 2px;font-size:8px;font-weight:700;color:#7dd3fc;text-transform:uppercase;letter-spacing:.2em">Detalle del día anterior</p>
    <p style="margin:0;font-size:14px;font-weight:700;color:#fff">{dia_label}</p>
  </td></tr>

  <!-- ═══ KPIs AYER ════════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:20px 32px;border-top:1px solid #e2e8f0">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td style="text-align:center;padding-right:12px;border-right:1px solid #f1f5f9;vertical-align:middle">
          <table cellpadding="0" cellspacing="0" align="center">
            <tr><td bgcolor="{col_d['light']}" style="background:{col_d['light']};border-radius:10px;padding:14px 18px;border:2px solid {col_d['bar']}">
              <p style="margin:0;font-size:28px;font-weight:900;color:{col_d['text']};line-height:1;text-align:center">{dia_pct}%</p>
              <p style="margin:4px 0 0;font-size:7px;font-weight:700;color:{col_d['text']};text-transform:uppercase;letter-spacing:.1em;opacity:.7;text-align:center">cumpl. vehículo</p>
            </td></tr>
          </table>
        </td>
        <td style="padding-left:20px;vertical-align:middle">
          <table cellpadding="0" cellspacing="0"><tr>
            <td style="padding:0 14px 0 0;text-align:center;border-right:1px solid #e2e8f0">
              <p style="margin:0;font-size:26px;font-weight:800;color:#1e293b;line-height:1">{dia_total}</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">Total</p>
            </td>
            <td style="padding:0 14px;text-align:center;border-right:1px solid #e2e8f0">
              <p style="margin:0;font-size:26px;font-weight:800;color:#16a34a;line-height:1">{dia_cumpl}</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">A tiempo</p>
            </td>
            <td style="padding:0 14px;text-align:center;border-right:1px solid #e2e8f0">
              <p style="margin:0;font-size:26px;font-weight:800;color:#dc2626;line-height:1">{dia_tarde}</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">Tarde</p>
            </td>
            <td style="padding:0 0 0 14px;text-align:center">
              <p style="margin:0;font-size:26px;font-weight:800;color:#38bdf8;line-height:1">{dia_pct_a}%</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">Cumpl. Aten.</p>
            </td>
          </tr></table>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- ═══ DETALLE CONTENEDORES AYER ═══════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:0 32px 24px;border-top:1px solid #e2e8f0">
    {SEC("Todos los vaciós de ayer")}
    {"<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #f1f5f9;border-radius:8px;overflow:hidden'><thead><tr><th " + TH_S + ">Contenedor</th><th " + TH_S + ">Patio</th><th " + TH_S + ">Cita patio</th><th " + TH_S + ">Llegada</th><th " + TH_S + ">Δ Cita→Llegada</th><th " + TH_S + ">Atención</th><th " + TH_S + ">Δ Cita→Aten.</th><th " + TH_S + ">Motivo incumpl.</th><th " + TH_S + ">Resp.</th></tr></thead><tbody>" + det_ayer + "</tbody></table>" if det_ayer else "<p style='margin:0;font-size:12px;color:#94a3b8;text-align:center;padding:12px'>Sin registros para este día</p>"}
  </td></tr>

  <!-- ═══ FIRMA ═══════════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:20px 32px 24px;border-top:1px solid #e2e8f0">
    <p style="margin:0 0 16px;font-family:Calibri,Arial,sans-serif;font-size:12pt;font-style:italic;font-weight:700;color:#1e293b">Cordialmente,</p>
    {FIRMA_HTML}
  </td></tr>

  <!-- ═══ FOOTER ════════════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#f8fafc" style="background:#f8fafc;border-radius:0 0 12px 12px;padding:13px 32px;border-top:1px solid #e2e8f0">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td style="font-size:10px;color:#94a3b8">
        <span style="color:#1d4ed8;font-weight:700">TC</span> Tractocar Logistics &nbsp;·&nbsp; Reporte automático diario
      </td>
      <td align="right" style="font-size:10px;color:#94a3b8">Generado {generado}</td>
    </tr></table>
  </td></tr>

</table>
</body></html>"""


def send_report():
    datos_path = BASE / "datos_bot.json"
    log_path   = BASE / "sent_log.json"
    if not datos_path.exists():
        print("[ERROR] No existe datos_bot.json"); sys.exit(1)

    from datetime import date
    hoy      = date.today().isoformat()
    mod_date = datetime.fromtimestamp(datos_path.stat().st_mtime).date().isoformat()
    if mod_date < hoy:
        print(f"[ERROR] datos_bot.json fue actualizado el {mod_date}, no hoy ({hoy}). Corre procesar.py primero.")
        sys.exit(1)

    # Verificar si ya se envió hoy
    log = {}
    if log_path.exists():
        try: log = json.loads(log_path.read_text(encoding="utf-8"))
        except: pass
    if log.get("vacios_interno") == hoy:
        print(f"[SKIP] Correo interno vacíos ya enviado hoy ({hoy}).")
        sys.exit(0)

    with open(datos_path, encoding="utf-8") as f:
        datos = json.load(f)
    if not EMAIL_PASSWORD:
        print("[ERROR] Falta EMAIL_PASSWORD en .env"); sys.exit(1)

    html    = build_html(datos)
    subject = f"[TC-Interno] Indicador Cumplimiento Citas Vacíos — {datos.get('generado','')[:10]}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = EMAIL_FROM
    msg["To"]      = ", ".join(EMAIL_TO_INTERNO)
    if EMAIL_CC_INTERNO: msg["Cc"] = ", ".join(EMAIL_CC_INTERNO)
    msg.attach(MIMEText(html, "html", "utf-8"))

    dest_all = EMAIL_TO_INTERNO + EMAIL_CC_INTERNO
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo(); smtp.starttls()
            smtp.login(EMAIL_FROM, EMAIL_PASSWORD)
            smtp.sendmail(EMAIL_FROM, dest_all, msg.as_string())
        print(f"[OK] Vacíos INTERNO — enviado a {', '.join(dest_all)}")
        log["vacios_interno"] = hoy
        log_path.write_text(json.dumps(log, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[ERROR] {e}"); sys.exit(1)


if __name__ == "__main__":
    if "--preview" in sys.argv:
        with open(BASE / "datos_bot.json", encoding="utf-8") as f:
            datos = json.load(f)
        out = BASE / "preview_vacios_interno.html"
        out.write_text(build_html(datos), encoding="utf-8")
        print(f"[Preview] {out}")
    else:
        send_report()
