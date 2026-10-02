"""
Reporte diario Ajover EXPO Llenos — diseño ejecutivo con SVG
.env: EMAIL_FROM, EMAIL_PASSWORD, EMAIL_TO, EMAIL_CC (opcional)
"""
import os, json, smtplib, sys, math
from collections import Counter, defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

BASE = Path(__file__).parent
load_dotenv(BASE / ".env")

EMAIL_FROM     = os.environ.get("EMAIL_FROM", "jarias@tractocar.com")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_TO       = [e.strip() for e in os.environ.get("EMAIL_TO", EMAIL_FROM).split(",") if e.strip()]
EMAIL_CC       = [e.strip() for e in os.environ.get("EMAIL_CC","").split(",") if e.strip()]
SMTP_HOST      = os.environ.get("SMTP_HOST","smtp.office365.com")
SMTP_PORT      = int(os.environ.get("SMTP_PORT","587"))

MESES_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
            'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
MESES_C  = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

def _mes_label(iso, corto=False):
    try:
        p = iso.split("-")
        return f"{(MESES_C if corto else MESES_ES)[int(p[1])-1]} {p[0]}"
    except: return iso

def _pct(a, b): return round(a/b*100, 1) if b else 0.0
def _pct_str(a, b): return f"{_pct(a,b)}%" if b else "—"

def _norm_mot(mot):
    m = (mot or "").lower()
    if not m: return "Sin motivo"
    if "escolta" in m:             return "Problema de escolta"
    if "sello" in m:               return "Demoras en sello"
    if "diferencia de peso" in m or "contenedor con diferencia" in m: return "Diferencia de peso"
    if "lluvia" in m or "clima" in m or "alto trafi" in m: return "Condiciones climáticas"
    if "tráfico" in m or "trafico" in m or "vía cerrada" in m or "via cerrada" in m: return "Tráfico / vía cerrada"
    if "báscula" in m or "bascula" in m: return "Daños en báscula"
    if "mecán" in m or "mecan" in m or "llanta" in m or "falla mec" in m: return "Falla mecánica"
    if "motonave" in m or "buque" in m:  return "Cambio fecha motonave"
    if "programac" in m:                 return "Cambio en programación"
    if "rndc" in m or "manifest" in m:   return "Trámites RNDC"
    if "seguridad" in m:                 return "Retrasos con seguridad"
    if "solicitud" in m or "cliente" in m: return "Solicitud del cliente"
    if "peso" in m:                      return "Diferencia de peso"
    return mot[:38] + ("…" if len(mot)>38 else "")

def _cumple(r):
    cc = r.get("cumpl_cita","") or ""
    return cc.startswith("A tiempo") or "externo" in cc or "reprog Ajover" in cc

def _tarde(r):
    return (r.get("cumpl_cita","") or "").startswith("Tarde")

def _resp_cat(r):
    cc = r.get("cumpl_cita","") or ""
    if "externo" in cc: return "Externo"
    if "reprog Ajover" in cc: return "Ajover"
    return r.get("resp_repr") or "Tractocar"

# ── SVG Gauge (donut semicircular) ────────────────────────────────────────────
def svg_gauge(pct, size=160):
    """Gauge semicircular con el % en el centro."""
    r = size * 0.38
    cx = size / 2
    cy = size * 0.56
    sw = size * 0.1
    # Ángulos: -180° a 0° (semicírculo superior)
    a_start = math.pi          # izquierda
    a_end   = 0                # derecha
    # Arco de fondo
    def arc_path(a0, a1):
        x0 = cx + r * math.cos(a0); y0 = cy + r * math.sin(a0)
        x1 = cx + r * math.cos(a1); y1 = cy + r * math.sin(a1)
        return f"M{x0:.1f},{y0:.1f} A{r:.1f},{r:.1f} 0 0,1 {x1:.1f},{y1:.1f}"

    fill_angle = math.pi - (pct / 100) * math.pi  # de izq a der según %
    col = "#22c55e" if pct >= 90 else "#f59e0b" if pct >= 75 else "#ef4444"
    label_col = col

    bg_d   = arc_path(math.pi, 0)
    fill_d = arc_path(math.pi, fill_angle) if pct > 0 else ""

    fill_path = f'<path d="{fill_d}" stroke="{col}" stroke-width="{sw}" fill="none" stroke-linecap="round"/>' if fill_d else ""

    # Pequeños ticks en 75% y 90%
    def tick(pct_t):
        a = math.pi - (pct_t/100)*math.pi
        x0 = cx + (r - sw/2 - 4)*math.cos(a); y0 = cy + (r - sw/2 - 4)*math.sin(a)
        x1 = cx + (r + sw/2 + 4)*math.cos(a); y1 = cy + (r + sw/2 + 4)*math.sin(a)
        return f'<line x1="{x0:.1f}" y1="{y0:.1f}" x2="{x1:.1f}" y2="{y1:.1f}" stroke="#1e293b" stroke-width="1.5"/>'

    return f'''<svg width="{size}" height="{int(size*0.65)}" viewBox="0 0 {size} {int(size*0.65)}" xmlns="http://www.w3.org/2000/svg">
  <path d="{bg_d}" stroke="#1e3a5f" stroke-width="{sw}" fill="none" stroke-linecap="round"/>
  {fill_path}
  {tick(75)}{tick(90)}
  <text x="{cx}" y="{cy-4}" text-anchor="middle" font-size="{int(size*0.22)}" font-weight="800" fill="{label_col}" font-family="Arial,sans-serif">{pct:.1f}%</text>
  <text x="{cx}" y="{cy+14}" text-anchor="middle" font-size="{int(size*0.08)}" fill="#94a3b8" font-family="Arial,sans-serif">CUMPLIMIENTO</text>
  <text x="{cx-r-4}" y="{cy+18}" text-anchor="middle" font-size="{int(size*0.065)}" fill="#475569" font-family="Arial">0%</text>
  <text x="{cx+r+4}" y="{cy+18}" text-anchor="middle" font-size="{int(size*0.065)}" fill="#475569" font-family="Arial">100%</text>
</svg>'''

# ── SVG mini barra horizontal ─────────────────────────────────────────────────
def bar_row(label, val, total, col, width=320):
    pct = val/total*100 if total else 0
    bw  = int(pct/100*width)
    return f'''<tr>
      <td style="padding:4px 0;font-size:11px;color:#cbd5e1;width:130px;white-space:nowrap">{label}</td>
      <td style="padding:4px 8px">
        <div style="background:#1e3a5f;border-radius:4px;height:12px;width:{width}px;overflow:hidden">
          <div style="background:{col};width:{bw}px;height:12px;border-radius:4px"></div>
        </div>
      </td>
      <td style="padding:4px 0 4px 8px;font-size:11px;color:#e2e8f0;font-weight:700;white-space:nowrap">{val} <span style="color:#64748b;font-weight:400">({pct:.0f}%)</span></td>
    </tr>'''

# ── Pill cumplimiento ─────────────────────────────────────────────────────────
def pill(txt, ok=None):
    if not txt or txt == "Sin fecha":
        return '<span style="background:#1e293b;color:#64748b;padding:2px 9px;border-radius:10px;font-size:10px;font-weight:600">Sin fecha</span>'
    if ok is None: ok = txt.startswith("A tiempo") or "externo" in txt or "reprog Ajover" in txt
    if ok:
        return f'<span style="background:#064e3b;color:#6ee7b7;padding:2px 9px;border-radius:10px;font-size:10px;font-weight:700">✓ {txt}</span>'
    return f'<span style="background:#450a0a;color:#fca5a5;padding:2px 9px;border-radius:10px;font-size:10px;font-weight:700">✗ {txt}</span>'

# ── SVG Gráfico de barras + línea (tendencia mes a mes) ───────────────────────
def svg_trend_chart(tend, mes_actual, w=620, h=240):
    """
    Gráfico de barras con línea encima, igual al del tablero.
    tend: lista de dicts {mes, total, pct, nc}
    """
    if not tend: return ""
    n      = len(tend)
    PAD_L  = 44   # eje Y
    PAD_R  = 16
    PAD_T  = 36   # espacio para etiquetas de %
    PAD_B  = 52   # espacio para etiquetas mes + llenos
    CHART_W = w - PAD_L - PAD_R
    CHART_H = h - PAD_T - PAD_B
    bar_w   = max(20, int(CHART_W / n * 0.55))
    gap     = CHART_W / n

    def cx(i): return PAD_L + gap * i + gap / 2   # centro de cada barra
    def bar_y(pct): return PAD_T + CHART_H * (1 - pct / 100)

    # Líneas de guía horizontales
    guides = ""
    for pct_g in [0, 25, 50, 75, 90, 100]:
        y = bar_y(pct_g)
        dash = "stroke-dasharray='4,4'" if pct_g not in (0, 100) else ""
        col_g = "#1e3a5f" if pct_g not in (75, 90) else ("#f59e0b44" if pct_g == 75 else "#22c55e44")
        guides += f'<line x1="{PAD_L}" y1="{y:.1f}" x2="{w-PAD_R}" y2="{y:.1f}" stroke="{col_g}" stroke-width="1" {dash}/>'
        guides += f'<text x="{PAD_L-6}" y="{y+4:.1f}" text-anchor="end" font-size="9" fill="#475569" font-family="Arial">{pct_g}%</text>'

    # Barras
    bars = ""
    for i, t in enumerate(tend):
        x    = cx(i) - bar_w / 2
        yb   = bar_y(t["pct"])
        bar_h = PAD_T + CHART_H - yb
        active = t["mes"] == mes_actual
        col  = "#22c55e" if t["pct"] >= 90 else "#f59e0b" if t["pct"] >= 75 else "#ef4444"
        alpha = "ff" if active else "99"
        bars += f'<rect x="{x:.1f}" y="{yb:.1f}" width="{bar_w}" height="{bar_h:.1f}" rx="3" fill="{col}{alpha}"/>'
        # Etiqueta % encima de la barra
        lbl_col = col if active else (col + "cc")
        fw = "800" if active else "700"
        bars += f'<text x="{cx(i):.1f}" y="{yb-5:.1f}" text-anchor="middle" font-size="{"11" if active else "10"}" font-weight="{fw}" fill="{lbl_col}" font-family="Arial">{t["pct"]}%</text>'

    # Línea de tendencia
    pts = [(cx(i), bar_y(t["pct"])) for i, t in enumerate(tend)]
    line_d = "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    last_col = "#22c55e" if tend[-1]["pct"] >= 90 else "#f59e0b" if tend[-1]["pct"] >= 75 else "#ef4444"
    dots = ""
    for i, (px, py) in enumerate(pts):
        active = tend[i]["mes"] == mes_actual
        r = 5 if active else 3.5
        dots += f'<circle cx="{px:.1f}" cy="{py:.1f}" r="{r}" fill="{last_col}" stroke="#0a1628" stroke-width="1.5"/>'

    # Etiquetas eje X
    xlabels = ""
    for i, t in enumerate(tend):
        active = t["mes"] == mes_actual
        lbl = _mes_label(t["mes"], True)          # "Oct 2026" → recortamos
        parts = lbl.split(" ")
        mes_c = parts[0]
        anio  = "'" + parts[1][2:] if len(parts) > 1 else ""
        fw = "700" if active else "400"
        fc = "#e2e8f0" if active else "#64748b"
        y_m = PAD_T + CHART_H + 16
        xlabels += f'<text x="{cx(i):.1f}" y="{y_m}" text-anchor="middle" font-size="10" font-weight="{fw}" fill="{fc}" font-family="Arial">{mes_c}&#x2019;{anio}</text>'
        xlabels += f'<text x="{cx(i):.1f}" y="{y_m+13}" text-anchor="middle" font-size="9" fill="#475569" font-family="Arial">{t["total"]} llenos</text>'

    return f'''<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg">
  {guides}
  {bars}
  <path d="{line_d}" stroke="{last_col}" stroke-width="2" fill="none" opacity="0.9"/>
  {dots}
  {xlabels}
</svg>'''


# ── SVG mini sparkline (cabecera tendencia) ───────────────────────────────────
def sparkline(vals, w=200, h=40):
    if len(vals) < 2: return ""
    mn, mx = min(vals), max(vals)
    rng = mx - mn if mx != mn else 1
    pts = [(i*(w/(len(vals)-1)), h - (v-mn)/rng*h) for i,v in enumerate(vals)]
    path = "M" + " L".join(f"{x:.1f},{y:.1f}" for x,y in pts)
    col  = "#22c55e" if vals[-1] >= 90 else "#f59e0b" if vals[-1] >= 75 else "#ef4444"
    last = pts[-1]
    return f'''<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg">
  <path d="{path}" stroke="{col}" stroke-width="2" fill="none"/>
  <circle cx="{last[0]:.1f}" cy="{last[1]:.1f}" r="3.5" fill="{col}"/>
</svg>'''

# ── Build email ───────────────────────────────────────────────────────────────
def build_html(datos):
    rows = datos["ajover"]["llenos"]["rows"]
    hoy  = datetime.now()
    mes_actual = hoy.strftime("%Y-%m")
    generado   = datos.get("generado","")[:10]

    # Mes actual
    mes_rows = [r for r in rows if r.get("mes_iso")==mes_actual]
    total_m  = len(mes_rows)
    sf_m     = sum(1 for r in mes_rows if (r.get("cumpl_cita","") or "")=="Sin fecha")
    cumpl_m  = sum(1 for r in mes_rows if _cumple(r))
    tarde_m  = [r for r in mes_rows if _tarde(r)]
    repr_m   = [r for r in mes_rows if r.get("fcita_repr")]
    base_m   = total_m - sf_m
    pct_m    = _pct(cumpl_m, base_m)

    RESP_COL = {"Tractocar":"#60a5fa","Ajover":"#f87171","Externo":"#6ee7b7"}
    RESP_BG  = {"Tractocar":"#1e3a5f","Ajover":"#450a0a","Externo":"#064e3b"}

    inc_resp = defaultdict(list)
    for r in tarde_m: inc_resp[_resp_cat(r)].append(r)

    mot_m = defaultdict(int)
    for r in tarde_m: mot_m[_norm_mot(r.get("motivo","") or r.get("motivo_repr",""))] += 1

    # Tendencia últimos 6-7 meses
    meses_disp = sorted(set(r["mes_iso"] for r in rows if r.get("mes_iso")))[-7:]
    tend = []
    for mes in meses_disp:
        mr   = [r for r in rows if r.get("mes_iso")==mes]
        sf   = sum(1 for r in mr if (r.get("cumpl_cita","") or "")=="Sin fecha")
        base = len(mr)-sf
        nc   = sum(1 for r in mr if _tarde(r))
        pct  = _pct(len(mr)-sf-nc, base)
        tend.append({"mes":mes,"total":len(mr),"pct":pct,"nc":nc})
    spark_vals   = [t["pct"] for t in tend]
    trend_chart  = svg_trend_chart(tend, mes_actual)

    # Ayer
    dia_rows, dia_dt = [], None
    for delta in range(1,8):
        dt = hoy-timedelta(days=delta)
        ds = dt.strftime("%d-%m-%Y")
        found = [r for r in rows if (r.get("fecha","") or "").startswith(ds[:6]) and ds[6:] in (r.get("fecha","") or "")]
        if found: dia_rows, dia_dt = found, dt; break

    dia_label = dia_dt.strftime("%A %d de %B de %Y").capitalize() if dia_dt else "Sin datos"
    dia_total = len(dia_rows)
    dia_sf    = sum(1 for r in dia_rows if (r.get("cumpl_cita","") or "")=="Sin fecha")
    dia_cumpl = sum(1 for r in dia_rows if _cumple(r))
    dia_tarde = [r for r in dia_rows if _tarde(r)]
    dia_repr  = [r for r in dia_rows if r.get("fcita_repr")]
    dia_pct   = _pct(dia_cumpl, dia_total-dia_sf)

    dia_inc_resp = defaultdict(list)
    for r in dia_tarde: dia_inc_resp[_resp_cat(r)].append(r)

    # ── HTML ──────────────────────────────────────────────────────────────────
    gauge_svg = svg_gauge(pct_m, 200)
    dia_gauge  = svg_gauge(dia_pct, 160)
    spark_svg  = sparkline(spark_vals)

    # Tendencia table rows
    tend_html = ""
    for t in tend:
        active = t["mes"]==mes_actual
        bg = "background:#0f2744;" if active else ""
        fw = "font-weight:700;" if active else ""
        col_pct = "#22c55e" if t["pct"]>=90 else "#f59e0b" if t["pct"]>=75 else "#ef4444"
        tend_html += f'''<tr style="{bg}border-bottom:1px solid #1e3a5f">
          <td style="padding:8px 14px;font-size:12px;color:#e2e8f0;{fw}white-space:nowrap">{"▶ " if active else ""}{_mes_label(t["mes"],True)}</td>
          <td style="padding:8px 14px;font-size:12px;color:#94a3b8;text-align:center">{t["total"]}</td>
          <td style="padding:8px 14px;font-size:13px;font-weight:800;color:{col_pct};text-align:center">{t["pct"]}%</td>
          <td style="padding:8px 14px;font-size:12px;color:#f87171;text-align:center">{t["nc"]}</td>
        </tr>'''

    # Incumplimientos mes
    inc_mes_html = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        lst = inc_resp.get(resp,[])
        if not lst: continue
        col = RESP_COL.get(resp,"#94a3b8")
        bg  = RESP_BG.get(resp,"#1e293b")
        mots = Counter(_norm_mot(r.get("motivo","") or r.get("motivo_repr","")) for r in lst)
        top_mot = "; ".join(f"{m} ({c})" for m,c in mots.most_common(3))
        inc_mes_html += f'''<tr style="border-bottom:1px solid #1e3a5f">
          <td style="padding:10px 14px">
            <span style="display:inline-block;background:{bg};color:{col};padding:3px 12px;border-radius:12px;font-size:11px;font-weight:700">{resp}</span>
          </td>
          <td style="padding:10px 14px;font-size:20px;font-weight:800;color:{col};text-align:center">{len(lst)}</td>
          <td style="padding:10px 14px;font-size:10px;color:#64748b">{top_mot or "—"}</td>
        </tr>'''

    # Motivos mes — barras
    mot_bars_html = ""
    max_mot = max(mot_m.values()) if mot_m else 1
    for mot, cnt in sorted(mot_m.items(), key=lambda x:-x[1])[:7]:
        bw = int(cnt/max_mot*240)
        mot_bars_html += f'''<tr>
          <td style="padding:5px 0;font-size:11px;color:#cbd5e1;width:160px">{mot}</td>
          <td style="padding:5px 10px">
            <div style="background:#1e3a5f;border-radius:4px;height:10px;width:240px;overflow:hidden">
              <div style="background:#f87171;width:{bw}px;height:10px;border-radius:4px"></div>
            </div>
          </td>
          <td style="padding:5px 0;font-size:11px;color:#f87171;font-weight:700;white-space:nowrap">{cnt} &nbsp;<span style="color:#475569;font-weight:400">({cnt*100//len(tarde_m) if tarde_m else 0}%)</span></td>
        </tr>'''

    # Reprog mes
    repr_mes_html = ""
    if repr_m:
        repr_resp = defaultdict(int)
        repr_mot  = defaultdict(int)
        for r in repr_m:
            repr_resp[r.get("resp_repr") or "Sin resp."] += 1
            repr_mot[_norm_mot(r.get("motivo_repr",""))] += 1
        for resp, cnt in sorted(repr_resp.items(), key=lambda x:-x[1]):
            col = RESP_COL.get(resp,"#94a3b8")
            repr_mes_html += f'<span style="display:inline-block;margin:3px;background:#1e3a5f;padding:4px 12px;border-radius:10px;font-size:11px;color:{col};font-weight:600">{resp} — {cnt}</span>'

    # Ayer incumplimientos
    dia_inc_html = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        lst = dia_inc_resp.get(resp,[])
        if not lst: continue
        col = RESP_COL.get(resp,"#94a3b8")
        bg  = RESP_BG.get(resp,"#1e293b")
        for r in lst:
            cc  = r.get("cumpl_cita","") or ""
            mot = _norm_mot(r.get("motivo","") or r.get("motivo_repr",""))
            dia_inc_html += f'''<tr style="border-bottom:1px solid #1e3a5f">
              <td style="padding:7px 12px;font-size:11px;font-family:monospace;color:#e2e8f0">{r.get("cont","—")}</td>
              <td style="padding:7px 12px;font-size:11px">
                <span style="background:{bg};color:{col};padding:2px 10px;border-radius:8px;font-size:10px;font-weight:700">{resp}</span>
              </td>
              <td style="padding:7px 12px;font-size:11px;color:#94a3b8">{r.get("fcita","—") or "—"}</td>
              <td style="padding:7px 12px;font-size:11px;color:#f87171">{cc}</td>
              <td style="padding:7px 12px;font-size:10px;color:#64748b">{mot}</td>
            </tr>'''

    # Reprog ayer
    repr_ayer_html = ""
    for r in dia_repr:
        resp_r = r.get("resp_repr") or "Sin responsable"
        col    = RESP_COL.get(resp_r,"#94a3b8")
        bg     = RESP_BG.get(resp_r,"#1e293b")
        repr_ayer_html += f'''<tr style="border-bottom:1px solid #1e3a5f">
          <td style="padding:7px 12px;font-size:11px;font-family:monospace;color:#e2e8f0">{r.get("cont","—")}</td>
          <td style="padding:7px 12px;font-size:11px;color:#94a3b8">{r.get("fcita","—") or "—"}</td>
          <td style="padding:7px 12px;font-size:11px;color:#60a5fa">{r.get("fcita_repr","—") or "—"}</td>
          <td style="padding:7px 12px;font-size:11px">
            <span style="background:{bg};color:{col};padding:2px 10px;border-radius:8px;font-size:10px;font-weight:700">{resp_r}</span>
          </td>
          <td style="padding:7px 12px;font-size:10px;color:#64748b">{_norm_mot(r.get("motivo_repr",""))}</td>
        </tr>'''

    # Detalle todos contenedores ayer
    det_ayer_html = ""
    for r in sorted(dia_rows, key=lambda r: r.get("fcita","") or ""):
        cc = r.get("cumpl_cita","") or ""
        ok = _cumple(r)
        mot = _norm_mot(r.get("motivo","") or r.get("motivo_repr","")) if not ok and cc else "—"
        det_ayer_html += f'''<tr style="border-bottom:1px solid #1e3a5f">
          <td style="padding:7px 12px;font-size:11px;font-family:monospace;color:#e2e8f0">{r.get("cont","—")}</td>
          <td style="padding:7px 12px;font-size:11px;color:#94a3b8">{r.get("terminal","—")}</td>
          <td style="padding:7px 12px;font-size:11px;color:#94a3b8">{r.get("fcita","—") or "—"}</td>
          <td style="padding:7px 12px;font-size:11px;color:#94a3b8">{r.get("fllpuerto","—") or "—"}</td>
          <td style="padding:7px 12px">{pill(cc, ok)}</td>
          <td style="padding:7px 12px;font-size:10px;color:#64748b">{mot}</td>
        </tr>'''

    th = 'style="padding:8px 12px;font-size:9px;color:#475569;font-weight:700;text-transform:uppercase;letter-spacing:.06em;text-align:left;border-bottom:1px solid #1e3a5f"'

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Indicador Cumplimiento Citas — Llenos</title></head>
<body style="margin:0;padding:24px 0;background:#060f1e;font-family:Arial,Helvetica,sans-serif">
<div style="max-width:700px;margin:0 auto">

  <!-- ═══ HEADER ═══════════════════════════════════════════════════════════ -->
  <div style="background:linear-gradient(135deg,#0a1628 0%,#0f2744 60%,#1a3a6e 100%);border-radius:16px 16px 0 0;padding:32px 36px 28px;border-bottom:1px solid #1e3a5f">
    <!-- Logo TC + nombre -->
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td>
          <div style="display:inline-block">
            <div style="background:linear-gradient(135deg,#1d4ed8,#3b82f6);border-radius:10px;padding:8px 14px;display:inline-block">
              <span style="font-size:18px;font-weight:900;color:#fff;letter-spacing:-1px">TC</span>
            </div>
            <span style="font-size:13px;color:#64748b;margin-left:10px;vertical-align:middle">TRACTOCAR LOGISTICS</span>
          </div>
        </td>
        <td align="right" style="font-size:11px;color:#475569">{generado}</td>
      </tr>
    </table>
    <!-- Título -->
    <div style="margin-top:20px">
      <div style="font-size:9px;font-weight:700;color:#3b82f6;text-transform:uppercase;letter-spacing:.18em;margin-bottom:6px">Reporte Ejecutivo · Ajover EXPO</div>
      <div style="font-size:26px;font-weight:800;color:#f8fafc;line-height:1.1">Indicador de Cumplimiento<br>de Citas — Llenos</div>
      <div style="font-size:13px;color:#64748b;margin-top:8px">{_mes_label(mes_actual)} · {total_m} contenedores procesados</div>
    </div>
  </div>

  <!-- ═══ GAUGE + KPIs MES ══════════════════════════════════════════════════ -->
  <div style="background:#0a1628;padding:28px 36px;border-bottom:1px solid #1e3a5f">
    <div style="font-size:9px;font-weight:700;color:#3b82f6;text-transform:uppercase;letter-spacing:.15em;margin-bottom:20px">Mes actual — {_mes_label(mes_actual)}</div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center" width="220" style="vertical-align:top">
          {gauge_svg}
          <div style="font-size:10px;color:#64748b;margin-top:-4px">{cumpl_m} de {base_m} con fecha</div>
        </td>
        <td style="vertical-align:middle;padding-left:24px">
          <table cellpadding="0" cellspacing="0">
            <tr>
              <td style="padding:6px 16px 6px 0">
                <div style="font-size:28px;font-weight:800;color:#f8fafc">{total_m}</div>
                <div style="font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:.06em">Total</div>
              </td>
              <td style="padding:6px 16px 6px 0;border-left:1px solid #1e3a5f;padding-left:16px">
                <div style="font-size:28px;font-weight:800;color:#22c55e">{cumpl_m}</div>
                <div style="font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:.06em">A tiempo</div>
              </td>
              <td style="padding:6px 16px 6px 0;border-left:1px solid #1e3a5f;padding-left:16px">
                <div style="font-size:28px;font-weight:800;color:#ef4444">{len(tarde_m)}</div>
                <div style="font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:.06em">Incumpl.</div>
              </td>
              <td style="padding:6px 0 6px 0;border-left:1px solid #1e3a5f;padding-left:16px">
                <div style="font-size:28px;font-weight:800;color:#f59e0b">{len(repr_m)}</div>
                <div style="font-size:10px;color:#64748b;text-transform:uppercase;letter-spacing:.06em">Reprog.</div>
              </td>
            </tr>
          </table>
          <!-- Barra responsables -->
          {'<div style="margin-top:18px"><div style="font-size:9px;color:#475569;text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px">Incumplimientos por responsable</div><table cellpadding="0" cellspacing="0">' + "".join(bar_row(r, len(l), len(tarde_m), RESP_COL.get(r,"#94a3b8"), 200) for r,l in sorted(inc_resp.items(), key=lambda x:-len(x[1]))) + "</table></div>" if tarde_m else '<div style="margin-top:18px;font-size:12px;color:#22c55e">✓ Sin incumplimientos en el mes</div>'}
        </td>
      </tr>
    </table>
  </div>

  <!-- ═══ TENDENCIA ════════════════════════════════════════════════════════ -->
  <div style="background:#080e1c;padding:24px 36px 20px;border-bottom:1px solid #1e3a5f">
    <div style="font-size:9px;font-weight:700;color:#3b82f6;text-transform:uppercase;letter-spacing:.15em;margin-bottom:2px">Cumplimiento de cita — mes a mes</div>
    <div style="font-size:10px;color:#475569;margin-bottom:16px">Últimos {len(tend)} meses · barra resaltada = mes actual</div>
    <div style="overflow:hidden;border-radius:8px;background:#060f1e;padding:10px 4px 4px">
      {trend_chart}
    </div>
    <!-- leyenda -->
    <div style="margin-top:10px;font-size:9px;color:#475569">
      <span style="display:inline-block;width:10px;height:10px;background:#22c55e;border-radius:2px;vertical-align:middle;margin-right:4px"></span>≥ 90% &nbsp;&nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:#f59e0b;border-radius:2px;vertical-align:middle;margin-right:4px"></span>75–90% &nbsp;&nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:#ef4444;border-radius:2px;vertical-align:middle;margin-right:4px"></span>&lt; 75%
    </div>
  </div>

  <!-- ═══ INCUMPLIMIENTOS MES ══════════════════════════════════════════════ -->
  {'<div style="background:#0a1628;padding:24px 36px;border-bottom:1px solid #1e3a5f"><div style="font-size:9px;font-weight:700;color:#ef4444;text-transform:uppercase;letter-spacing:.15em;margin-bottom:16px">Incumplimientos ' + _mes_label(mes_actual, True) + ' — responsables y motivos</div><table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px"><thead><tr style="border-bottom:1px solid #1e3a5f"><th ' + th + '>Responsable</th><th ' + th + ' style="text-align:center">Casos</th><th ' + th + '>Principales motivos</th></tr></thead><tbody>' + inc_mes_html + '</tbody></table><div style="font-size:9px;color:#475569;text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px">Distribución de motivos</div><table cellpadding="0" cellspacing="0">' + mot_bars_html + '</table></div>' if tarde_m else ""}

  <!-- ═══ REPROGRAMACIONES MES ════════════════════════════════════════════ -->
  {'<div style="background:#080e1c;padding:20px 36px;border-bottom:1px solid #1e3a5f"><div style="font-size:9px;font-weight:700;color:#f59e0b;text-transform:uppercase;letter-spacing:.15em;margin-bottom:10px">Reprogramaciones ' + _mes_label(mes_actual, True) + ' — ' + str(len(repr_m)) + ' en total</div><div>' + repr_mes_html + "</div></div>" if repr_m else ""}

  <!-- ═══ DIVIDER AYER ════════════════════════════════════════════════════ -->
  <div style="background:linear-gradient(90deg,#1d4ed8,#0f172a);padding:12px 36px">
    <div style="font-size:9px;font-weight:700;color:#93c5fd;text-transform:uppercase;letter-spacing:.18em">Detalle del día anterior</div>
    <div style="font-size:14px;font-weight:700;color:#f8fafc;margin-top:2px">{dia_label}</div>
  </div>

  <!-- ═══ KPIs AYER ════════════════════════════════════════════════════════ -->
  <div style="background:#0a1628;padding:24px 36px;border-bottom:1px solid #1e3a5f">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center" width="180" style="vertical-align:top">{dia_gauge}</td>
        <td style="vertical-align:middle;padding-left:20px">
          <table cellpadding="0" cellspacing="0">
            <tr>
              <td style="padding:6px 14px 6px 0">
                <div style="font-size:24px;font-weight:800;color:#f8fafc">{dia_total}</div>
                <div style="font-size:10px;color:#64748b">Total</div>
              </td>
              <td style="padding:6px 14px;border-left:1px solid #1e3a5f">
                <div style="font-size:24px;font-weight:800;color:#22c55e">{dia_cumpl}</div>
                <div style="font-size:10px;color:#64748b">A tiempo</div>
              </td>
              <td style="padding:6px 14px;border-left:1px solid #1e3a5f">
                <div style="font-size:24px;font-weight:800;color:#ef4444">{len(dia_tarde)}</div>
                <div style="font-size:10px;color:#64748b">Tarde</div>
              </td>
              <td style="padding:6px 0 6px 14px;border-left:1px solid #1e3a5f">
                <div style="font-size:24px;font-weight:800;color:#f59e0b">{len(dia_repr)}</div>
                <div style="font-size:10px;color:#64748b">Reprog.</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>

  <!-- ═══ INCUMPLIMIENTOS AYER ═════════════════════════════════════════════ -->
  {'<div style="background:#080e1c;padding:20px 36px;border-bottom:1px solid #1e3a5f"><div style="font-size:9px;font-weight:700;color:#ef4444;text-transform:uppercase;letter-spacing:.15em;margin-bottom:14px">Incumplimientos de ayer — responsable y motivo</div><table width="100%" cellpadding="0" cellspacing="0"><thead><tr style="border-bottom:1px solid #1e3a5f"><th ' + th + '>Contenedor</th><th ' + th + '>Responsable</th><th ' + th + '>Cita</th><th ' + th + '>Cumplimiento</th><th ' + th + '>Motivo</th></tr></thead><tbody>' + dia_inc_html + '</tbody></table></div>' if dia_inc_html else '<div style="background:#080e1c;padding:16px 36px;border-bottom:1px solid #1e3a5f"><span style="font-size:12px;color:#22c55e">✓ Sin incumplimientos ayer</span></div>'}

  <!-- ═══ REPROGRAMACIONES AYER ════════════════════════════════════════════ -->
  {'<div style="background:#0a1628;padding:20px 36px;border-bottom:1px solid #1e3a5f"><div style="font-size:9px;font-weight:700;color:#f59e0b;text-transform:uppercase;letter-spacing:.15em;margin-bottom:14px">Reprogramaciones de ayer — ' + str(len(dia_repr)) + '</div><table width="100%" cellpadding="0" cellspacing="0"><thead><tr style="border-bottom:1px solid #1e3a5f"><th ' + th + '>Contenedor</th><th ' + th + '>Cita original</th><th ' + th + '>Nueva cita</th><th ' + th + '>Responsable</th><th ' + th + '>Motivo</th></tr></thead><tbody>' + repr_ayer_html + '</tbody></table></div>' if dia_repr else ''}

  <!-- ═══ DETALLE CONTENEDORES AYER ════════════════════════════════════════ -->
  <div style="background:#080e1c;padding:20px 36px;border-bottom:1px solid #1e3a5f">
    <div style="font-size:9px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.15em;margin-bottom:14px">Todos los contenedores de ayer</div>
    {'<table width="100%" cellpadding="0" cellspacing="0"><thead><tr style="border-bottom:1px solid #1e3a5f"><th ' + th + '>Contenedor</th><th ' + th + '>Terminal</th><th ' + th + '>Cita puerto</th><th ' + th + '>Llegada</th><th ' + th + '>Cumplimiento</th><th ' + th + '>Motivo</th></tr></thead><tbody>' + det_ayer_html + '</tbody></table>' if det_ayer_html else '<div style="font-size:12px;color:#475569;text-align:center;padding:12px">Sin registros para este día</div>'}
  </div>

  <!-- ═══ FOOTER ════════════════════════════════════════════════════════════ -->
  <div style="background:#060f1e;border-radius:0 0 16px 16px;padding:16px 36px;border-top:1px solid #1e3a5f">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td style="font-size:10px;color:#334155">
          <span style="color:#1d4ed8;font-weight:700">TC</span> Tractocar Logistics · Reporte automático
        </td>
        <td align="right" style="font-size:10px;color:#334155">Generado {generado}</td>
      </tr>
    </table>
  </div>

</div>
</body></html>"""


def send_report():
    datos_path = BASE / "datos_bot.json"
    if not datos_path.exists():
        print("[ERROR] No existe datos_bot.json"); sys.exit(1)
    with open(datos_path, encoding="utf-8") as f:
        datos = json.load(f)
    if not EMAIL_PASSWORD:
        print("[ERROR] Falta EMAIL_PASSWORD en .env"); sys.exit(1)

    html = build_html(datos)
    msg  = MIMEMultipart("alternative")
    msg["Subject"] = f"[TC] Indicador Cumplimiento Citas Llenos — {datos.get('generado','')[:10]}"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = ", ".join(EMAIL_TO)
    if EMAIL_CC: msg["Cc"] = ", ".join(EMAIL_CC)
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo(); smtp.starttls()
            smtp.login(EMAIL_FROM, EMAIL_PASSWORD)
            smtp.sendmail(EMAIL_FROM, EMAIL_TO + EMAIL_CC, msg.as_string())
        print(f"[OK] Correo enviado a {', '.join(EMAIL_TO + EMAIL_CC)}")
    except Exception as e:
        print(f"[ERROR] {e}"); sys.exit(1)


if __name__ == "__main__":
    send_report()
