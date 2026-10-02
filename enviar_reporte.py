"""
Reporte diario Ajover EXPO Llenos — tema claro, gráficos CSS/tabla (compatible Outlook)
.env: EMAIL_FROM, EMAIL_PASSWORD, EMAIL_TO, EMAIL_CC (opcional)
"""
import os, json, smtplib, sys
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

# ── Helpers ───────────────────────────────────────────────────────────────────
def _mes_label(iso, corto=False):
    try:
        p = iso.split("-")
        return f"{(MESES_C if corto else MESES_ES)[int(p[1])-1]} {p[0]}"
    except: return iso

def _pct(a, b): return round(a/b*100, 1) if b else 0.0

def _norm_mot(mot):
    m = (mot or "").lower()
    if not m: return "Sin motivo"
    if "escolta" in m:             return "Problema de escolta"
    if "sello" in m:               return "Demoras en sello"
    if "diferencia de peso" in m or "contenedor con diferencia" in m or "peso" in m: return "Diferencia de peso"
    if "lluvia" in m or "clima" in m: return "Condiciones climáticas"
    if "tráfico" in m or "trafico" in m or "vía cerrada" in m or "via cerrada" in m: return "Tráfico / vía cerrada"
    if "báscula" in m or "bascula" in m: return "Daños en báscula"
    if "mecán" in m or "mecan" in m or "llanta" in m or "falla mec" in m: return "Falla mecánica"
    if "motonave" in m or "buque" in m:  return "Cambio fecha motonave"
    if "programac" in m:                 return "Cambio en programación"
    if "rndc" in m or "manifest" in m:   return "Trámites RNDC"
    if "seguridad" in m:                 return "Retrasos con seguridad"
    if "solicitud" in m or "cliente" in m: return "Solicitud del cliente"
    return mot[:38] + ("…" if len(mot) > 38 else "")

def _cumple(r):
    cc = r.get("cumpl_cita","") or ""
    return cc.startswith("A tiempo") or "externo" in cc or "reprog Ajover" in cc

def _tarde(r):
    return (r.get("cumpl_cita","") or "").startswith("Tarde")

def _resp_cat(r):
    cc = r.get("cumpl_cita","") or ""
    if "externo" in cc:       return "Externo"
    if "reprog Ajover" in cc: return "Ajover"
    return r.get("resp_repr") or "Tractocar"

# ── Colores por umbral ────────────────────────────────────────────────────────
def _pct_col(pct):
    if pct >= 90: return {"bar":"#16a34a","text":"#15803d","bg":"#dcfce7","border":"#86efac"}
    if pct >= 75: return {"bar":"#d97706","text":"#b45309","bg":"#fef3c7","border":"#fcd34d"}
    return {"bar":"#dc2626","text":"#b91c1c","bg":"#fee2e2","border":"#fca5a5"}

RESP_COL = {"Tractocar":"#2563eb","Ajover":"#dc2626","Externo":"#16a34a"}
RESP_BG  = {"Tractocar":"#eff6ff","Ajover":"#fee2e2","Externo":"#dcfce7"}
RESP_BORDER = {"Tractocar":"#bfdbfe","Ajover":"#fca5a5","Externo":"#86efac"}

# ── Gráfico de barras verticales (CSS tabla — compatible Outlook) ─────────────
def html_trend_chart(tend, mes_actual, max_h=140):
    if not tend: return ""
    cells = ""
    for t in tend:
        c      = _pct_col(t["pct"])
        bh     = max(4, int(t["pct"] / 100 * max_h))
        active = t["mes"] == mes_actual
        lbl    = _mes_label(t["mes"], True)
        parts  = lbl.split(" ")
        mes_c  = parts[0]
        anio   = "'" + parts[1][2:] if len(parts) > 1 else ""
        border = f'border:2px solid {c["bar"]};' if active else "border:2px solid transparent;"
        bg_cell = f'background:{c["bg"]};' if active else ""
        cells += f'''<td style="text-align:center;vertical-align:bottom;padding:0 3px;width:70px">
          <div style="font-size:{"11" if active else "10"}px;font-weight:{"800" if active else "600"};color:{c["text"]};margin-bottom:4px">{t["pct"]}%</div>
          <div style="width:42px;height:{bh}px;background:{c["bar"]};border-radius:4px 4px 0 0;margin:0 auto;opacity:{"1" if active else "0.75"}"></div>
          <div style="border-top:2px solid #e2e8f0;padding-top:5px;margin-top:0">
            <div style="font-size:{"10" if active else "9"}px;font-weight:{"700" if active else "400"};color:{"#1e293b" if active else "#64748b"}">{mes_c}&#39;{anio}</div>
            <div style="font-size:9px;color:#94a3b8">{t["total"]}</div>
          </div>
        </td>'''
    return f'''<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%">
  <tr style="vertical-align:bottom">{cells}</tr>
</table>'''

# ── Barra horizontal de progreso ──────────────────────────────────────────────
def progress_bar(val, total, col, width=260):
    pct = val / total * 100 if total else 0
    bw  = max(0, min(width, int(pct / 100 * width)))
    return f'''<table cellpadding="0" cellspacing="0"><tr>
      <td><div style="background:#e2e8f0;border-radius:4px;height:10px;width:{width}px;overflow:hidden">
        <div style="background:{col};width:{bw}px;height:10px;border-radius:4px"></div>
      </div></td>
      <td style="padding-left:8px;font-size:11px;font-weight:700;color:{col};white-space:nowrap">{val} <span style="color:#94a3b8;font-weight:400">({pct:.0f}%)</span></td>
    </tr></table>'''

# ── KPI card ──────────────────────────────────────────────────────────────────
def kpi_card(label, val, col, sub=""):
    return f'''<td style="padding:0 8px;text-align:center;border-right:1px solid #e2e8f0">
      <div style="font-size:32px;font-weight:800;color:{col};line-height:1">{val}</div>
      <div style="font-size:9px;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em;margin-top:3px">{label}</div>
      {"<div style='font-size:9px;color:#94a3b8;margin-top:1px'>" + sub + "</div>" if sub else ""}
    </td>'''

# ── Pill estado ───────────────────────────────────────────────────────────────
def pill(txt, ok=None):
    if not txt or txt == "Sin fecha":
        return '<span style="background:#f1f5f9;color:#94a3b8;padding:2px 8px;border-radius:8px;font-size:10px;font-weight:600;border:1px solid #e2e8f0">Sin fecha</span>'
    if ok is None:
        ok = txt.startswith("A tiempo") or "externo" in txt or "reprog Ajover" in txt
    if ok:
        return f'<span style="background:#dcfce7;color:#15803d;padding:2px 8px;border-radius:8px;font-size:10px;font-weight:700;border:1px solid #86efac">✓ {txt}</span>'
    return f'<span style="background:#fee2e2;color:#b91c1c;padding:2px 8px;border-radius:8px;font-size:10px;font-weight:700;border:1px solid #fca5a5">✗ {txt}</span>'

# ── Section header ────────────────────────────────────────────────────────────
def sec(label, col="#2563eb"):
    return f'<div style="font-size:9px;font-weight:700;color:{col};text-transform:uppercase;letter-spacing:.14em;padding:0 0 10px">{label}</div>'

TH = 'style="padding:8px 14px;font-size:9px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:.06em;text-align:left;border-bottom:2px solid #f1f5f9;background:#f8fafc"'
TD = 'style="padding:8px 14px;font-size:11px;color:#334155;border-bottom:1px solid #f1f5f9"'

# ── Build HTML ────────────────────────────────────────────────────────────────
def build_html(datos):
    rows       = datos["ajover"]["llenos"]["rows"]
    hoy        = datetime.now()
    mes_actual = hoy.strftime("%Y-%m")
    generado   = datos.get("generado","")[:10]

    # ── Mes actual
    mes_rows = [r for r in rows if r.get("mes_iso") == mes_actual]
    total_m  = len(mes_rows)
    sf_m     = sum(1 for r in mes_rows if (r.get("cumpl_cita","") or "") == "Sin fecha")
    cumpl_m  = sum(1 for r in mes_rows if _cumple(r))
    tarde_m  = [r for r in mes_rows if _tarde(r)]
    repr_m   = [r for r in mes_rows if r.get("fcita_repr")]
    base_m   = total_m - sf_m
    pct_m    = _pct(cumpl_m, base_m)
    col_m    = _pct_col(pct_m)

    inc_resp = defaultdict(list)
    for r in tarde_m: inc_resp[_resp_cat(r)].append(r)

    mot_m = defaultdict(int)
    for r in tarde_m:
        mot_m[_norm_mot(r.get("motivo","") or r.get("motivo_repr",""))] += 1

    # ── Tendencia últimos 7 meses
    meses_disp = sorted(set(r["mes_iso"] for r in rows if r.get("mes_iso")))[-7:]
    tend = []
    for mes in meses_disp:
        mr   = [r for r in rows if r.get("mes_iso") == mes]
        sf   = sum(1 for r in mr if (r.get("cumpl_cita","") or "") == "Sin fecha")
        base = len(mr) - sf
        nc   = sum(1 for r in mr if _tarde(r))
        tend.append({"mes":mes,"total":len(mr),"pct":_pct(base-nc, base),"nc":nc})

    # ── Ayer (último día con datos, hasta 8 días atrás)
    dia_rows, dia_dt = [], None
    for delta in range(1, 9):
        dt  = hoy - timedelta(days=delta)
        ds  = dt.strftime("%d-%m-%Y")
        found = [r for r in rows
                 if (r.get("fecha","") or "").startswith(ds[:6])
                 and ds[6:] in (r.get("fecha","") or "")]
        if found: dia_rows, dia_dt = found, dt; break

    dia_label = dia_dt.strftime("%A %d de %B de %Y").capitalize() if dia_dt else "Sin datos recientes"
    dia_total = len(dia_rows)
    dia_sf    = sum(1 for r in dia_rows if (r.get("cumpl_cita","") or "") == "Sin fecha")
    dia_cumpl = sum(1 for r in dia_rows if _cumple(r))
    dia_tarde = [r for r in dia_rows if _tarde(r)]
    dia_repr  = [r for r in dia_rows if r.get("fcita_repr")]
    dia_pct   = _pct(dia_cumpl, dia_total - dia_sf)
    col_d     = _pct_col(dia_pct)

    dia_inc_resp = defaultdict(list)
    for r in dia_tarde: dia_inc_resp[_resp_cat(r)].append(r)

    # ── HTML de secciones ─────────────────────────────────────────────────────

    # Barras responsables mes
    resp_bars = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        lst = inc_resp.get(resp, [])
        if not lst: continue
        resp_bars += f'''<tr>
          <td style="padding:5px 0;font-size:11px;color:#334155;width:90px;white-space:nowrap">
            <span style="background:{RESP_BG[resp]};color:{RESP_COL[resp]};padding:2px 8px;border-radius:8px;font-size:10px;font-weight:700;border:1px solid {RESP_BORDER[resp]}">{resp}</span>
          </td>
          <td style="padding:5px 0 5px 10px">{progress_bar(len(lst), len(tarde_m), RESP_COL[resp], 200)}</td>
        </tr>'''

    # Motivos mes
    mot_bars = ""
    max_mot = max(mot_m.values()) if mot_m else 1
    for mot, cnt in sorted(mot_m.items(), key=lambda x: -x[1])[:7]:
        bw = int(cnt / max_mot * 220)
        mot_bars += f'''<tr>
          <td style="padding:5px 0;font-size:11px;color:#475569;width:165px;white-space:nowrap;overflow:hidden">{mot}</td>
          <td style="padding:5px 8px">
            <div style="background:#fee2e2;border-radius:4px;height:10px;width:220px;overflow:hidden">
              <div style="background:#dc2626;width:{bw}px;height:10px;border-radius:4px"></div>
            </div>
          </td>
          <td style="padding:5px 0;font-size:11px;font-weight:700;color:#dc2626;white-space:nowrap">{cnt}</td>
        </tr>'''

    # Incumplimientos mes — tabla
    inc_mes_rows = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        lst = inc_resp.get(resp, [])
        if not lst: continue
        mots = Counter(_norm_mot(r.get("motivo","") or r.get("motivo_repr","")) for r in lst)
        top  = "; ".join(f"{m} ({c})" for m,c in mots.most_common(3))
        inc_mes_rows += f'''<tr style="border-bottom:1px solid #f1f5f9">
          <td {TD}><span style="background:{RESP_BG[resp]};color:{RESP_COL[resp]};padding:3px 10px;border-radius:8px;font-size:10px;font-weight:700;border:1px solid {RESP_BORDER[resp]}">{resp}</span></td>
          <td style="padding:8px 14px;font-size:20px;font-weight:800;color:{RESP_COL[resp]};text-align:center;border-bottom:1px solid #f1f5f9">{len(lst)}</td>
          <td style="padding:8px 14px;font-size:10px;color:#64748b;border-bottom:1px solid #f1f5f9">{top or "—"}</td>
        </tr>'''

    # Reprog mes
    repr_mes_html = ""
    if repr_m:
        repr_resp = Counter(r.get("resp_repr") or "Sin resp." for r in repr_m)
        for resp, cnt in repr_resp.most_common():
            col_r = RESP_COL.get(resp,"#64748b"); bg_r = RESP_BG.get(resp,"#f1f5f9"); bd_r = RESP_BORDER.get(resp,"#e2e8f0")
            repr_mes_html += f'<span style="display:inline-block;margin:3px;background:{bg_r};color:{col_r};padding:4px 12px;border-radius:10px;font-size:11px;font-weight:600;border:1px solid {bd_r}">{resp} — {cnt}</span>'

    # Incumplimientos ayer
    dia_inc_rows = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        for r in dia_inc_resp.get(resp, []):
            cc  = r.get("cumpl_cita","") or ""
            mot = _norm_mot(r.get("motivo","") or r.get("motivo_repr",""))
            dia_inc_rows += f'''<tr style="border-bottom:1px solid #f1f5f9">
              <td {TD} style="font-family:monospace;padding:8px 14px;font-size:11px;color:#334155;border-bottom:1px solid #f1f5f9">{r.get("cont","—")}</td>
              <td style="padding:8px 14px;border-bottom:1px solid #f1f5f9"><span style="background:{RESP_BG[resp]};color:{RESP_COL[resp]};padding:2px 8px;border-radius:8px;font-size:10px;font-weight:700;border:1px solid {RESP_BORDER[resp]}">{resp}</span></td>
              <td style="padding:8px 14px;font-size:11px;color:#64748b;border-bottom:1px solid #f1f5f9">{r.get("fcita","—") or "—"}</td>
              <td style="padding:8px 14px;font-size:10px;color:#dc2626;border-bottom:1px solid #f1f5f9">{cc}</td>
              <td style="padding:8px 14px;font-size:10px;color:#64748b;border-bottom:1px solid #f1f5f9">{mot}</td>
            </tr>'''

    # Reprog ayer
    repr_ayer_rows = ""
    for r in dia_repr:
        resp_r = r.get("resp_repr") or "Sin responsable"
        col_r  = RESP_COL.get(resp_r,"#64748b"); bg_r = RESP_BG.get(resp_r,"#f1f5f9"); bd_r = RESP_BORDER.get(resp_r,"#e2e8f0")
        repr_ayer_rows += f'''<tr style="border-bottom:1px solid #f1f5f9">
          <td {TD} style="font-family:monospace;padding:8px 14px;font-size:11px;color:#334155;border-bottom:1px solid #f1f5f9">{r.get("cont","—")}</td>
          <td style="padding:8px 14px;font-size:11px;color:#64748b;border-bottom:1px solid #f1f5f9">{r.get("fcita","—") or "—"}</td>
          <td style="padding:8px 14px;font-size:11px;color:#2563eb;font-weight:600;border-bottom:1px solid #f1f5f9">{r.get("fcita_repr","—") or "—"}</td>
          <td style="padding:8px 14px;border-bottom:1px solid #f1f5f9"><span style="background:{bg_r};color:{col_r};padding:2px 8px;border-radius:8px;font-size:10px;font-weight:700;border:1px solid {bd_r}">{resp_r}</span></td>
          <td style="padding:8px 14px;font-size:10px;color:#64748b;border-bottom:1px solid #f1f5f9">{_norm_mot(r.get("motivo_repr",""))}</td>
        </tr>'''

    # Detalle contenedores ayer
    det_ayer_rows = ""
    for r in sorted(dia_rows, key=lambda r: r.get("fcita","") or ""):
        cc  = r.get("cumpl_cita","") or ""
        ok  = _cumple(r)
        mot = _norm_mot(r.get("motivo","") or r.get("motivo_repr","")) if not ok and cc else "—"
        det_ayer_rows += f'''<tr style="border-bottom:1px solid #f1f5f9">
          <td {TD} style="font-family:monospace;padding:8px 14px;font-size:11px;color:#334155;border-bottom:1px solid #f1f5f9">{r.get("cont","—")}</td>
          <td style="padding:8px 14px;font-size:11px;color:#64748b;border-bottom:1px solid #f1f5f9">{r.get("terminal","—")}</td>
          <td style="padding:8px 14px;font-size:11px;color:#64748b;border-bottom:1px solid #f1f5f9">{r.get("fcita","—") or "—"}</td>
          <td style="padding:8px 14px;font-size:11px;color:#64748b;border-bottom:1px solid #f1f5f9">{r.get("fllpuerto","—") or "—"}</td>
          <td style="padding:8px 14px;border-bottom:1px solid #f1f5f9">{pill(cc, ok)}</td>
          <td style="padding:8px 14px;font-size:10px;color:#64748b;border-bottom:1px solid #f1f5f9">{mot}</td>
        </tr>'''

    trend_html = html_trend_chart(tend, mes_actual)

    # ─────────────────────────────────────────────────────────────────────────
    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Indicador Cumplimiento Citas — Llenos</title></head>
<body style="margin:0;padding:20px 0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif">
<div style="max-width:680px;margin:0 auto">

  <!-- ═══ HEADER ══════════════════════════════════════════════════════════════ -->
  <div style="background:linear-gradient(135deg,#0f172a 0%,#1e3a6e 60%,#1d4ed8 100%);border-radius:14px 14px 0 0;padding:28px 32px 24px">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td>
          <div style="display:inline-block;background:rgba(255,255,255,0.15);border-radius:8px;padding:6px 13px">
            <span style="font-size:17px;font-weight:900;color:#fff;letter-spacing:-0.5px">TC</span>
            <span style="font-size:10px;color:rgba(255,255,255,0.6);margin-left:6px;vertical-align:middle">TRACTOCAR LOGISTICS</span>
          </div>
        </td>
        <td align="right" style="font-size:10px;color:rgba(255,255,255,0.45)">{generado}</td>
      </tr>
    </table>
    <div style="margin-top:18px">
      <div style="font-size:8px;font-weight:700;color:#93c5fd;text-transform:uppercase;letter-spacing:.2em;margin-bottom:5px">Reporte Ejecutivo · Ajover EXPO</div>
      <div style="font-size:22px;font-weight:800;color:#fff;line-height:1.2">Indicador de Cumplimiento<br>de Citas — Llenos</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.5);margin-top:7px">{_mes_label(mes_actual)} &nbsp;·&nbsp; {total_m} contenedores</div>
    </div>
  </div>

  <!-- ═══ KPIs MES ACTUAL ══════════════════════════════════════════════════════ -->
  <div style="background:#fff;padding:24px 32px;border-bottom:1px solid #e2e8f0">
    <div style="font-size:8px;font-weight:700;color:#2563eb;text-transform:uppercase;letter-spacing:.16em;margin-bottom:18px">Mes actual — {_mes_label(mes_actual)}</div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <!-- Indicador principal -->
        <td width="160" style="vertical-align:middle;text-align:center;border-right:1px solid #f1f5f9;padding-right:24px">
          <div style="background:{col_m["bg"]};border-radius:12px;padding:18px 12px;border:2px solid {col_m["border"]}">
            <div style="font-size:38px;font-weight:900;color:{col_m["text"]};line-height:1">{pct_m}%</div>
            <div style="font-size:8px;font-weight:700;color:{col_m["text"]};text-transform:uppercase;letter-spacing:.12em;margin-top:5px;opacity:.7">Cumplimiento</div>
            <!-- Barra de progreso -->
            <div style="background:rgba(0,0,0,0.08);border-radius:4px;height:6px;margin-top:10px;overflow:hidden">
              <div style="background:{col_m["bar"]};width:{min(100,pct_m):.0f}%;height:6px;border-radius:4px"></div>
            </div>
            <div style="font-size:9px;color:{col_m["text"]};opacity:.6;margin-top:5px">{cumpl_m} de {base_m} con fecha</div>
          </div>
        </td>
        <!-- KPIs numéricos -->
        <td style="vertical-align:middle;padding-left:24px">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
              {kpi_card("Total", total_m, "#1e293b")}
              {kpi_card("A tiempo", cumpl_m, "#16a34a")}
              {kpi_card("Incumpl.", len(tarde_m), "#dc2626")}
              <td style="padding:0 0 0 8px;text-align:center">
                <div style="font-size:32px;font-weight:800;color:#d97706;line-height:1">{len(repr_m)}</div>
                <div style="font-size:9px;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em;margin-top:3px">Reprog.</div>
              </td>
            </tr>
          </table>
          <!-- Responsables -->
          {"<div style='margin-top:18px;padding-top:14px;border-top:1px solid #f1f5f9'><div style='font-size:8px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px'>Incumplimientos por responsable</div><table cellpadding='0' cellspacing='4'>" + resp_bars + "</table></div>" if tarde_m else "<div style='margin-top:18px;padding:10px 14px;background:#dcfce7;border-radius:8px;border:1px solid #86efac'><span style='font-size:12px;color:#15803d;font-weight:600'>&#10003; Sin incumplimientos en el mes</span></div>"}
        </td>
      </tr>
    </table>
  </div>

  <!-- ═══ TENDENCIA MES A MES ══════════════════════════════════════════════════ -->
  <div style="background:#fff;padding:24px 32px 20px;border-bottom:1px solid #e2e8f0;margin-top:2px">
    <div style="font-size:8px;font-weight:700;color:#2563eb;text-transform:uppercase;letter-spacing:.16em;margin-bottom:4px">Cumplimiento de cita — mes a mes</div>
    <div style="font-size:10px;color:#94a3b8;margin-bottom:18px">Últimos {len(tend)} meses &nbsp;·&nbsp; barra destacada = mes actual</div>
    <div style="background:#f8fafc;border-radius:10px;padding:16px 8px 8px;border:1px solid #e2e8f0">
      {trend_html}
    </div>
    <div style="margin-top:10px;font-size:9px;color:#94a3b8">
      <span style="display:inline-block;width:10px;height:10px;background:#16a34a;border-radius:2px;vertical-align:middle;margin-right:4px"></span>&#8805; 90%&nbsp;&nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:#d97706;border-radius:2px;vertical-align:middle;margin-right:4px"></span>75–90%&nbsp;&nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:#dc2626;border-radius:2px;vertical-align:middle;margin-right:4px"></span>&lt; 75%
    </div>
  </div>

  <!-- ═══ INCUMPLIMIENTOS MES ══════════════════════════════════════════════════ -->
  {"<div style='background:#fff;padding:24px 32px;border-bottom:1px solid #e2e8f0;margin-top:2px'>" + sec("Incumplimientos " + _mes_label(mes_actual,True) + " — responsables y motivos", "#dc2626") + "<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #f1f5f9;border-radius:8px;overflow:hidden;margin-bottom:18px'><thead><tr><th " + TH + ">Responsable</th><th " + TH + " style='text-align:center'>Casos</th><th " + TH + ">Principales motivos</th></tr></thead><tbody>" + inc_mes_rows + "</tbody></table><div style='font-size:8px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px'>Distribución de motivos</div><table cellpadding='0' cellspacing='2'>" + mot_bars + "</table></div>" if tarde_m else ""}

  <!-- ═══ REPROGRAMACIONES MES ════════════════════════════════════════════════ -->
  {"<div style='background:#fff;padding:18px 32px;border-bottom:1px solid #e2e8f0;margin-top:2px'>" + sec("Reprogramaciones " + _mes_label(mes_actual,True) + " — " + str(len(repr_m)) + " en total", "#d97706") + "<div>" + repr_mes_html + "</div></div>" if repr_m else ""}

  <!-- ═══ DIVIDER AYER ════════════════════════════════════════════════════════ -->
  <div style="background:linear-gradient(90deg,#1e3a6e,#2563eb);padding:11px 32px;margin-top:2px">
    <div style="font-size:8px;font-weight:700;color:#bfdbfe;text-transform:uppercase;letter-spacing:.18em">Detalle del día anterior</div>
    <div style="font-size:14px;font-weight:700;color:#fff;margin-top:2px">{dia_label}</div>
  </div>

  <!-- ═══ KPIs AYER ════════════════════════════════════════════════════════════ -->
  <div style="background:#fff;padding:20px 32px;border-bottom:1px solid #e2e8f0">
    <table cellpadding="0" cellspacing="0">
      <tr>
        <td width="120" style="text-align:center;padding-right:20px;border-right:1px solid #f1f5f9;vertical-align:middle">
          <div style="background:{col_d["bg"]};border-radius:10px;padding:14px 10px;border:2px solid {col_d["border"]}">
            <div style="font-size:28px;font-weight:900;color:{col_d["text"]};line-height:1">{dia_pct}%</div>
            <div style="font-size:7px;font-weight:700;color:{col_d["text"]};text-transform:uppercase;letter-spacing:.1em;margin-top:4px;opacity:.7">Cumplimiento</div>
          </div>
        </td>
        <td style="vertical-align:middle;padding-left:20px">
          <table cellpadding="0" cellspacing="0"><tr>
            {kpi_card("Total", dia_total, "#1e293b")}
            {kpi_card("A tiempo", dia_cumpl, "#16a34a")}
            {kpi_card("Tarde", len(dia_tarde), "#dc2626")}
            <td style="padding:0 0 0 8px;text-align:center">
              <div style="font-size:28px;font-weight:800;color:#d97706;line-height:1">{len(dia_repr)}</div>
              <div style="font-size:9px;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em;margin-top:3px">Reprog.</div>
            </td>
          </tr></table>
        </td>
      </tr>
    </table>
  </div>

  <!-- ═══ INCUMPLIMIENTOS AYER ════════════════════════════════════════════════ -->
  {"<div style='background:#fff;padding:20px 32px;border-bottom:1px solid #e2e8f0;margin-top:2px'>" + sec("Incumplimientos de ayer", "#dc2626") + "<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #f1f5f9;border-radius:8px;overflow:hidden'><thead><tr><th " + TH + ">Contenedor</th><th " + TH + ">Responsable</th><th " + TH + ">Cita</th><th " + TH + ">Resultado</th><th " + TH + ">Motivo</th></tr></thead><tbody>" + dia_inc_rows + "</tbody></table></div>" if dia_inc_rows else "<div style='background:#fff;padding:14px 32px;border-bottom:1px solid #e2e8f0;margin-top:2px'><div style='background:#dcfce7;border-radius:8px;padding:10px 14px;border:1px solid #86efac'><span style='font-size:12px;color:#15803d;font-weight:600'>&#10003; Sin incumplimientos ayer</span></div></div>"}

  <!-- ═══ REPROGRAMACIONES AYER ════════════════════════════════════════════════ -->
  {"<div style='background:#fff;padding:20px 32px;border-bottom:1px solid #e2e8f0;margin-top:2px'>" + sec("Reprogramaciones de ayer — " + str(len(dia_repr)), "#d97706") + "<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #f1f5f9;border-radius:8px;overflow:hidden'><thead><tr><th " + TH + ">Contenedor</th><th " + TH + ">Cita original</th><th " + TH + ">Nueva cita</th><th " + TH + ">Responsable</th><th " + TH + ">Motivo</th></tr></thead><tbody>" + repr_ayer_rows + "</tbody></table></div>" if dia_repr else ""}

  <!-- ═══ DETALLE CONTENEDORES AYER ════════════════════════════════════════════ -->
  <div style="background:#fff;padding:20px 32px;border-bottom:1px solid #e2e8f0;margin-top:2px">
    {sec("Todos los contenedores de ayer")}
    {"<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #f1f5f9;border-radius:8px;overflow:hidden'><thead><tr><th " + TH + ">Contenedor</th><th " + TH + ">Terminal</th><th " + TH + ">Cita puerto</th><th " + TH + ">Llegada</th><th " + TH + ">Cumplimiento</th><th " + TH + ">Motivo</th></tr></thead><tbody>" + det_ayer_rows + "</tbody></table>" if det_ayer_rows else "<div style='font-size:12px;color:#94a3b8;text-align:center;padding:12px'>Sin registros para este día</div>"}
  </div>

  <!-- ═══ FOOTER ════════════════════════════════════════════════════════════════ -->
  <div style="background:#f8fafc;border-radius:0 0 14px 14px;padding:14px 32px;border-top:1px solid #e2e8f0">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td style="font-size:10px;color:#94a3b8">
          <span style="color:#1d4ed8;font-weight:700">TC</span> Tractocar Logistics &nbsp;·&nbsp; Reporte automático diario
        </td>
        <td align="right" style="font-size:10px;color:#94a3b8">Generado {generado}</td>
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
