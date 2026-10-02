"""
Reporte diario Ajover EXPO Llenos
Outlook-safe: bgcolor/height/width en <td>, sin divs con background
.env: EMAIL_FROM, EMAIL_PASSWORD, EMAIL_TO, EMAIL_CC (opcional)
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

# ── Firma: imágenes embebidas como base64 ─────────────────────────────────────
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
def _ml(iso, c=False):
    try:
        p = iso.split("-")
        n = int(p[1])-1
        return f"{(MESES_C if c else MESES_ES)[n]} {p[0]}"
    except: return iso

def _pct(a, b): return round(a/b*100, 1) if b else 0.0

def _col(pct):
    if pct >= 90: return {"bar":"#22c55e","light":"#dcfce7","text":"#15803d","dark":"#14532d"}
    if pct >= 75: return {"bar":"#f59e0b","light":"#fef3c7","text":"#b45309","dark":"#78350f"}
    return              {"bar":"#ef4444","light":"#fee2e2","text":"#b91c1c","dark":"#7f1d1d"}

RESP = {
    "Tractocar": {"fg":"#60a5fa","bg":"#1e3a5f","text":"#93c5fd"},
    "Ajover":    {"fg":"#f87171","bg":"#450a0a","text":"#fca5a5"},
    "Externo":   {"fg":"#34d399","bg":"#064e3b","text":"#6ee7b7"},
}
RESP_L = {  # versión clara para fondo blanco
    "Tractocar": {"fg":"#2563eb","bg":"#eff6ff","border":"#bfdbfe","text":"#1d4ed8"},
    "Ajover":    {"fg":"#dc2626","bg":"#fee2e2","border":"#fca5a5","text":"#b91c1c"},
    "Externo":   {"fg":"#16a34a","bg":"#dcfce7","border":"#86efac","text":"#15803d"},
}

def _norm_mot(mot):
    m = (mot or "").lower()
    if not m: return "Sin motivo"
    if "escolta" in m:   return "Problema de escolta"
    if "sello" in m:     return "Demoras en sello"
    if "peso" in m or "diferencia" in m: return "Diferencia de peso"
    if "lluvia" in m or "clima" in m: return "Condiciones climáticas"
    if "tráfico" in m or "trafico" in m or "vía" in m or "via" in m: return "Tráfico / vía cerrada"
    if "báscula" in m or "bascula" in m: return "Daños en báscula"
    if "mecán" in m or "mecan" in m or "llanta" in m: return "Falla mecánica"
    if "motonave" in m or "buque" in m: return "Cambio fecha motonave"
    if "programac" in m: return "Cambio en programación"
    if "rndc" in m or "manifest" in m: return "Trámites RNDC"
    if "seguridad" in m: return "Retrasos con seguridad"
    if "solicitud" in m or "cliente" in m: return "Solicitud del cliente"
    return mot[:36] + ("…" if len(mot) > 36 else "")

def _cumple(r):
    cc = r.get("cumpl_cita","") or ""
    return cc.startswith("A tiempo") or "externo" in cc or "reprog Ajover" in cc

def _tarde(r):
    return (r.get("cumpl_cita","") or "").startswith("Tarde")

def _pendiente(r):
    """Responsable asignado pero sin nueva cita aun — excluir del calculo como sin fecha."""
    return (bool((r.get("resp_repr","") or "").strip())
            and not (r.get("fcita_repr","") or "").strip()
            and not _cumple(r))

def _resp_cat(r):
    cc = r.get("cumpl_cita","") or ""
    if "externo" in cc:       return "Externo"
    if "reprog Ajover" in cc: return "Ajover"
    return r.get("resp_repr") or "Tractocar"

# ── Barra vertical (Outlook-safe: bgcolor + height en td) ─────────────────────
def _vbar(pct, col, active, max_h=110, w=38):
    bh = max(4, int(pct / 100 * max_h))
    sh = max_h - bh
    border = f'border-left:2px solid {col};border-right:2px solid {col};border-top:2px solid {col};' if active else ''
    return f'''<table cellpadding="0" cellspacing="0" align="center" style="margin:0 auto">
      <tr><td width="{w}" height="{sh}" style="font-size:0;line-height:0"> </td></tr>
      <tr><td width="{w}" height="{bh}" bgcolor="{col}" style="background:{col};{border}font-size:0;line-height:0"> </td></tr>
    </table>'''

# ── Barra horizontal (Outlook-safe) ──────────────────────────────────────────
def _hbar(val, total, col, W=200):
    p = val/total*100 if total else 0
    f = max(0, min(W, int(p/100*W)))
    e = W - f
    empty_td = f'<td width="{e}" height="8" bgcolor="#e2e8f0" style="background:#e2e8f0;font-size:0"> </td>' if e > 0 else ""
    return f'''<table cellpadding="0" cellspacing="0"><tr>
      <td width="{f}" height="8" bgcolor="{col}" style="background:{col};font-size:0"> </td>
      {empty_td}
    </tr></table>
    <span style="font-size:11px;font-weight:700;color:{col}">{val}</span>
    <span style="font-size:10px;color:#94a3b8"> ({p:.0f}%)</span>'''

# ── Pill estado ───────────────────────────────────────────────────────────────
def _pill(txt, ok=None):
    if not txt or txt == "Sin fecha":
        return '<span style="background:#f1f5f9;color:#94a3b8;padding:2px 8px;border-radius:6px;font-size:10px;border:1px solid #e2e8f0">Sin fecha</span>'
    if ok is None:
        ok = txt.startswith("A tiempo") or "externo" in txt or "reprog Ajover" in txt
    if ok:
        return f'<span style="background:#dcfce7;color:#15803d;padding:2px 8px;border-radius:6px;font-size:10px;font-weight:700;border:1px solid #86efac">&#10003; {txt}</span>'
    return f'<span style="background:#fee2e2;color:#b91c1c;padding:2px 8px;border-radius:6px;font-size:10px;font-weight:700;border:1px solid #fca5a5">&#10007; {txt}</span>'

def _resp_pill(resp):
    r = RESP_L.get(resp, {"fg":"#64748b","bg":"#f1f5f9","border":"#e2e8f0","text":"#334155"})
    return f'<span style="background:{r["bg"]};color:{r["text"]};padding:2px 10px;border-radius:6px;font-size:10px;font-weight:700;border:1px solid {r["border"]}">{resp}</span>'

# ─────────────────────────────────────────────────────────────────────────────
def build_html(datos, interno=False):
    rows    = datos["ajover"]["llenos"]["rows"]
    hoy     = datetime.now()
    generado = datos.get("generado","")[:10]

    # Fecha de corte: siempre ayer
    corte      = hoy - timedelta(days=1)
    mes_actual = corte.strftime("%Y-%m")
    corte_label = corte.strftime("%d de %B de %Y").lstrip("0")

    def _r_fecha(r):
        """Parsea r['fecha'] (DD-MM-YYYY ...) a date, o None."""
        try:
            f = (r.get("fecha","") or "").strip().split(" ")[0]
            d, m, y = f.split("-")
            return datetime(int(y), int(m), int(d)).date()
        except: return None

    # ── Mes actual — solo filas hasta el corte (ayer)
    mes_rows = [r for r in rows
                if r.get("mes_iso") == mes_actual
                and (_r_fecha(r) is None or _r_fecha(r) <= corte.date())]
    total_m  = len(mes_rows)
    sf_m     = sum(1 for r in mes_rows if (r.get("cumpl_cita","") or "") == "Sin fecha" or _pendiente(r))
    cumpl_m  = sum(1 for r in mes_rows if _cumple(r))
    tarde_m  = [r for r in mes_rows if _tarde(r) and not _pendiente(r)]
    repr_m   = [r for r in mes_rows if r.get("fcita_repr")]
    base_m   = total_m - sf_m
    pct_m    = _pct(cumpl_m, base_m)
    col_m    = _col(pct_m)

    inc_resp = defaultdict(list)
    for r in tarde_m: inc_resp[_resp_cat(r)].append(r)
    mot_m = defaultdict(int)
    for r in tarde_m: mot_m[(r.get("motivo_repr","") or "").strip() or "Sin motivo"] += 1

    # ── Tendencia últimos 7 meses (hasta el corte en el mes actual)
    meses_disp = sorted(set(r["mes_iso"] for r in rows if r.get("mes_iso")))[-7:]
    tend = []
    for mes in meses_disp:
        if mes == mes_actual:
            mr = mes_rows          # ya filtrado hasta ayer
        else:
            mr = [r for r in rows if r.get("mes_iso") == mes]
        sf    = sum(1 for r in mr if (r.get("cumpl_cita","") or "") == "Sin fecha" or _pendiente(r))
        base  = len(mr) - sf
        cumpl = sum(1 for r in mr if _cumple(r))
        nc    = sum(1 for r in mr if _tarde(r))
        tend.append({"mes":mes,"total":len(mr),"pct":_pct(cumpl, base),"nc":nc})

    # ── Ayer
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
    dia_sf    = sum(1 for r in dia_rows if (r.get("cumpl_cita","") or "") == "Sin fecha" or _pendiente(r))
    dia_cumpl = sum(1 for r in dia_rows if _cumple(r))
    dia_tarde = [r for r in dia_rows if _tarde(r) and not _cumple(r) and not _pendiente(r)]
    dia_repr  = [r for r in dia_rows if r.get("fcita_repr")]
    dia_pct   = _pct(dia_cumpl, dia_total - dia_sf)
    col_d     = _col(dia_pct)

    dia_inc_resp = defaultdict(list)
    for r in dia_tarde: dia_inc_resp[_resp_cat(r)].append(r)

    n_tc  = len(dia_inc_resp.get("Tractocar", []))
    n_ajv = len(dia_inc_resp.get("Ajover", []))
    n_ext = len(dia_inc_resp.get("Externo", []))

    # Alertas operativas del día
    dia_sin_llegada    = [r for r in dia_rows
                          if not (r.get("fllpuerto","") or "").strip()
                          and (r.get("fcita","") or "").strip()]
    dia_repr_incompl   = [r for r in dia_rows
                          if (r.get("fcita_repr","") or "").strip()
                          and (not (r.get("resp_repr","") or "").strip()
                               or not (r.get("motivo_repr","") or "").strip())]
    dia_repr_sin_fecha = [r for r in dia_rows
                          if (r.get("resp_repr","") or "").strip()
                          and not (r.get("fcita_repr","") or "").strip()]

    # ── Gráfico de tendencia ──────────────────────────────────────────────────
    # Fila 1: porcentajes
    f_pct = ""
    # Fila 2: barras (usando _vbar)
    f_bar = ""
    # Fila 3: etiquetas mes + cantidad
    f_lbl = ""
    for t in tend:
        c      = _col(t["pct"])
        active = t["mes"] == mes_actual
        lbl    = _ml(t["mes"], True).split(" ")
        mes_c  = lbl[0] + ("'" + lbl[1][2:] if len(lbl) > 1 else "")
        fw_a   = "font-weight:800;" if active else "font-weight:600;"
        f_pct += f'<td width="76" style="text-align:center;padding-bottom:6px;font-size:{"11" if active else "10"}px;{fw_a}color:{c["text"]}">{t["pct"]}%</td>'
        f_bar += f'<td width="76" style="text-align:center;vertical-align:bottom">{_vbar(t["pct"], c["bar"], active)}</td>'
        f_lbl += f'<td width="76" style="text-align:center;padding-top:6px;border-top:2px solid #e2e8f0"><span style="font-size:{"11" if active else "9"}px;{fw_a}color:{"#1e293b" if active else "#64748b"}">{mes_c}</span><br><span style="font-size:9px;color:#94a3b8">{t["total"]} llenos</span></td>'

    chart_html = f'''<table cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse">
      <tr>{f_pct}</tr>
      <tr>{f_bar}</tr>
      <tr>{f_lbl}</tr>
    </table>'''

    # ── Tablas de detalle ─────────────────────────────────────────────────────
    TH_S = 'style="padding:8px 12px;font-size:9px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.06em;text-align:left;background:#f8fafc;border-bottom:2px solid #e2e8f0"'
    TD_S = 'style="padding:8px 12px;font-size:11px;color:#334155;border-bottom:1px solid #f1f5f9"'

    # Incumplimientos mes
    inc_mes = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        lst = inc_resp.get(resp, [])
        if not lst: continue
        mots = Counter((r.get("motivo_repr","") or "").strip() or "Sin motivo" for r in lst)
        top  = "; ".join(f"{m} ({c})" for m,c in mots.most_common(3))
        r_   = RESP_L[resp]
        inc_mes += f'''<tr>
          <td {TD_S}>{_resp_pill(resp)}</td>
          <td style="padding:8px 12px;font-size:22px;font-weight:800;color:{r_["text"]};text-align:center;border-bottom:1px solid #f1f5f9">{len(lst)}</td>
          <td style="padding:8px 12px;font-size:10px;color:#64748b;border-bottom:1px solid #f1f5f9">{top or "—"}</td>
        </tr>'''

    # Motivos mes
    mot_rows = ""
    max_mot = max(mot_m.values()) if mot_m else 1
    for mot, cnt in sorted(mot_m.items(), key=lambda x:-x[1])[:8]:
        bw = max(4, int(cnt/max_mot*220))
        ew = 220 - bw
        empty_td = f'<td width="{ew}" height="8" bgcolor="#fee2e2" style="background:#fee2e2;font-size:0"> </td>' if ew > 0 else ""
        mot_rows += f'''<tr>
          <td style="padding:5px 10px 5px 0;font-size:10px;color:#475569;width:150px;vertical-align:middle">{mot}</td>
          <td style="padding:5px 0;vertical-align:middle">
            <table cellpadding="0" cellspacing="0"><tr>
              <td width="{bw}" height="8" bgcolor="#dc2626" style="background:#dc2626;font-size:0"> </td>
              {empty_td}
            </tr></table>
          </td>
          <td style="padding:5px 0 5px 10px;font-size:11px;font-weight:700;color:#dc2626;white-space:nowrap;vertical-align:middle">{cnt}</td>
        </tr>'''

    # Reprog mes pills
    repr_mes = ""
    if repr_m:
        for resp, cnt in Counter(r.get("resp_repr") or "Sin resp." for r in repr_m).most_common():
            r_ = RESP_L.get(resp, {"bg":"#f1f5f9","border":"#e2e8f0","text":"#334155"})
            repr_mes += f'<span style="display:inline-block;margin:3px;background:{r_["bg"]};color:{r_["text"]};padding:4px 12px;border-radius:8px;font-size:11px;font-weight:600;border:1px solid {r_["border"]}">{resp} — {cnt}</span>'

    # ── Helpers de celda para tablas densas
    def _tc(val, mono=False, bold=False, col="#334155", wrap=True):
        ff = "font-family:monospace;" if mono else ""
        fw = "font-weight:700;" if bold else ""
        ws = "" if wrap else "white-space:nowrap;"
        return f'<td style="padding:5px 8px;font-size:10px;color:{col};border-bottom:1px solid #f1f5f9;vertical-align:top;{ff}{fw}{ws}">'

    # Incumplimientos ayer
    dia_inc = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        for r in dia_inc_resp.get(resp, []):
            cc   = r.get("cumpl_cita","") or ""
            mot  = (r.get("motivo_repr","") or "").strip() or "—"
            obs  = (r.get("obs","") or r.get("motivo","") or "—")[:60] + ("…" if len(r.get("obs","") or r.get("motivo","") or "") > 60 else "")
            nrep = r.get("fcita_repr","") or "—"
            rresp = r.get("resp_repr","") or "—"
            dia_inc += f'''<tr>
              {_tc(None,mono=True)}  {r.get("ob","—")}</td>
              {_tc(None,mono=True)}  {r.get("cont","—")}</td>
              {_tc(None)}            {_resp_pill(resp)}</td>
              {_tc(None,col="#64748b",wrap=False)}{r.get("fcita","—") or "—"}</td>
              {_tc(None,col="#2563eb",wrap=False)}{nrep}</td>
              {_tc(None,col="#64748b")}           {_resp_pill(rresp) if rresp != "—" else "—"}</td>
              {_tc(None,col="#b91c1c")}            {cc}</td>
              {_tc(None,col="#475569")}            {mot}</td>
              {_tc(None,col="#94a3b8")}            {obs}</td>
            </tr>'''

    # Reprog ayer
    repr_ayer = ""
    for r in dia_repr:
        rr  = r.get("resp_repr") or "Sin responsable"
        mot = (r.get("motivo_repr","") or "").strip() or "—"
        obs = (r.get("obs","") or r.get("motivo","") or "—")[:60] + ("…" if len(r.get("obs","") or r.get("motivo","") or "") > 60 else "")
        repr_ayer += f'''<tr>
          {_tc(None,mono=True)}   {r.get("ob","—")}</td>
          {_tc(None,mono=True)}   {r.get("cont","—")}</td>
          {_tc(None,col="#64748b",wrap=False)}{r.get("fcita","—") or "—"}</td>
          {_tc(None,col="#2563eb",wrap=False,bold=True)}{r.get("fcita_repr","—") or "—"}</td>
          {_tc(None)}              {_resp_pill(rr)}</td>
          {_tc(None,col="#475569")}{mot}</td>
          {_tc(None,col="#94a3b8")}{obs}</td>
        </tr>'''

    # Detalle ayer completo — agrupado: reprog > tarde sin reprog > sin fecha > a tiempo
    def _sort_key(r):
        cc = r.get("cumpl_cita","") or ""
        tiene_reprog = bool(r.get("fcita_repr",""))
        if _tarde(r) and tiene_reprog and not _pendiente(r):  grupo = 0
        elif _tarde(r) and not _pendiente(r):                 grupo = 1
        elif cc == "Sin fecha" or _pendiente(r):              grupo = 2
        elif _cumple(r):                                      grupo = 3
        else:                                                 grupo = 4
        return (grupo, r.get("fcita","") or "")

    det_ayer = ""
    prev_grupo = None
    GRUPO_LABEL = {
        0: ("Incumplimientos con reprogramacion", "#fee2e2", "#dc2626"),
        1: ("Incumplimientos sin reprogramacion", "#fef3c7", "#b45309"),
        2: ("Sin fecha de cita", "#f1f5f9", "#64748b"),
        3: ("A tiempo", "#dcfce7", "#15803d"),
        4: ("Otros", "#f8fafc", "#94a3b8"),
    }
    NCOLS = 10  # número de columnas de la tabla
    for r in sorted(dia_rows, key=_sort_key):
        cc   = r.get("cumpl_cita","") or ""
        ok   = _cumple(r)
        tiene_reprog = bool(r.get("fcita_repr",""))
        if _tarde(r) and tiene_reprog and not _pendiente(r):  grupo = 0
        elif _tarde(r) and not _pendiente(r):                 grupo = 1
        elif cc == "Sin fecha" or _pendiente(r): grupo = 2
        elif ok:                               grupo = 3
        else:                                  grupo = 4

        # Separador de grupo
        if grupo != prev_grupo:
            lbl, bg, fg = GRUPO_LABEL[grupo]
            det_ayer += f'<tr><td colspan="{NCOLS}" bgcolor="{bg}" style="background:{bg};padding:5px 10px;font-size:9px;font-weight:700;color:{fg};text-transform:uppercase;letter-spacing:.1em;border-bottom:1px solid #e2e8f0">{lbl}</td></tr>'
            prev_grupo = grupo

        mot  = ((r.get("motivo_repr","") or "").strip() or "—") if (not ok and cc) else "—"
        obs  = (r.get("obs","") or r.get("motivo","") or "") if (not ok and cc) else ""
        obs_s = obs[:55] + ("…" if len(obs) > 55 else "") if obs else "—"
        nrep  = r.get("fcita_repr","") or "—"
        rresp = r.get("resp_repr","") or "—"
        det_ayer += f'''<tr>
          {_tc(None,mono=True)}              {r.get("ob","—")}</td>
          {_tc(None,mono=True)}              {r.get("cont","—")}</td>
          {_tc(None,col="#64748b")}          {r.get("terminal","—")}</td>
          {_tc(None,col="#64748b",wrap=False)}{r.get("fcita","—") or "—"}</td>
          {_tc(None,col="#64748b",wrap=False)}{r.get("fllpuerto","—") or "—"}</td>
          {_tc(None,col="#2563eb",wrap=False)}{nrep}</td>
          {_tc(None)}                         {_resp_pill(rresp) if rresp != "—" else "—"}</td>
          {_tc(None)}                         {_pill(cc, ok)}</td>
          {_tc(None,col="#475569")}           {mot}</td>
          {_tc(None,col="#94a3b8")}           {obs_s}</td>
        </tr>'''

    # ─────────────────────────────────────────────────────────────────────────
    # Barras responsables mes
    resp_bars = ""
    for resp in ["Tractocar","Ajover","Externo"]:
        lst = inc_resp.get(resp, [])
        if not lst: continue
        r_ = RESP_L[resp]
        resp_bars += f'''<tr>
          <td style="padding:5px 12px 5px 0;white-space:nowrap">{_resp_pill(resp)}</td>
          <td style="padding:5px 0">{_hbar(len(lst), len(tarde_m), r_["fg"], 180)}</td>
        </tr>'''

    # ── HTML: responsabilidad ayer ────────────────────────────────────────────
    total_inc_d = n_tc + n_ajv + n_ext
    def _resp_card(label, n, fg, bg, border):
        txt_n = str(n) if n else "0"
        sub   = "incumplimiento" + ("s" if n != 1 else "") + " imputables"
        return f'''<table width="100%" cellpadding="0" cellspacing="0">
          <tr><td>
            <table cellpadding="0" cellspacing="0" width="100%"><tr>
              <td width="4" bgcolor="{fg}" style="background:{fg};border-radius:4px 0 0 4px;font-size:0"> </td>
              <td bgcolor="{bg}" style="background:{bg};padding:12px 14px;border:1px solid {border};border-left:none;border-radius:0 8px 8px 0">
                <p style="margin:0 0 4px;font-size:8px;font-weight:700;color:{fg};text-transform:uppercase;letter-spacing:.12em">{label}</p>
                <p style="margin:0;font-size:32px;font-weight:900;color:{fg};line-height:1">{txt_n}</p>
                <p style="margin:5px 0 0;font-size:9px;color:#64748b">{sub}</p>
              </td>
            </tr></table>
          </td></tr>
        </table>'''

    _split_bar_w = 540
    def _split_seg(n, col, label, total):
        w = max(4, int(n / total * _split_bar_w)) if total else 0
        return (f'<td width="{w}" height="12" bgcolor="{col}" style="background:{col};font-size:0"> </td>', w)

    if total_inc_d:
        seg_tc,  w_tc  = _split_seg(n_tc,  "#2563eb", "Tractocar", total_inc_d)
        seg_ajv, w_ajv = _split_seg(n_ajv, "#dc2626", "Ajover",    total_inc_d)
        seg_ext, w_ext = _split_seg(n_ext, "#16a34a", "Externo",   total_inc_d)
        gap = _split_bar_w - w_tc - w_ajv - w_ext
        gap_td = f'<td width="{gap}" height="12" bgcolor="#e2e8f0" style="background:#e2e8f0;font-size:0"> </td>' if gap > 0 else ""
        resp_split_bar = f'<table cellpadding="0" cellspacing="0" style="margin-top:10px"><tr>{seg_tc}{seg_ajv}{seg_ext}{gap_td}</tr></table>'
    else:
        resp_split_bar = ""

    # ── HTML: alertas operativas ──────────────────────────────────────────────
    def _alerta_card(icon, label, n, desc, bg, fg, border, rows_list=None):
        rows_html = ""
        if rows_list:
            for r in rows_list[:4]:
                rows_html += f'<tr><td style="padding:2px 0 2px 12px;font-size:9px;color:#475569;font-family:monospace">{r.get("cont","—")} &nbsp;<span style="color:#94a3b8">{r.get("fcita","") or r.get("estado","")}</span></td></tr>'
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
    if dia_sin_llegada:
        n_alertas += len(dia_sin_llegada)
        alertas_items += _alerta_card("&#9888;", f"{len(dia_sin_llegada)} contenedor{'es' if len(dia_sin_llegada)!=1 else ''} sin fecha de llegada a puerto",
                                      len(dia_sin_llegada), "Tienen cita registrada pero sin llegada al puerto.",
                                      "#fee2e2","#dc2626","#fca5a5", dia_sin_llegada)
    if dia_repr_incompl:
        n_alertas += len(dia_repr_incompl)
        miss = []
        if any(not (r.get("resp_repr","") or "").strip() for r in dia_repr_incompl): miss.append(f"{sum(1 for r in dia_repr_incompl if not (r.get('resp_repr','') or '').strip())} responsable")
        if any(not (r.get("motivo_repr","") or "").strip() for r in dia_repr_incompl): miss.append(f"{sum(1 for r in dia_repr_incompl if not (r.get('motivo_repr','') or '').strip())} motivo")
        alertas_items += _alerta_card("&#9432;", f"{len(dia_repr_incompl)} reprogramaci{'ones' if len(dia_repr_incompl)!=1 else 'ón'} con datos incompletos",
                                      len(dia_repr_incompl), f"Faltan: {' · '.join(miss)}.",
                                      "#fef9c3","#b45309","#fde68a", dia_repr_incompl)
    if dia_repr_sin_fecha:
        n_alertas += len(dia_repr_sin_fecha)
        alertas_items += _alerta_card("&#9432;", f"{len(dia_repr_sin_fecha)} reprogramaci{'ones' if len(dia_repr_sin_fecha)!=1 else 'ón'} sin nueva cita registrada",
                                      len(dia_repr_sin_fecha), "Responsable asignado pero no se registró la nueva fecha de cita.",
                                      "#fef9c3","#b45309","#fde68a", dia_repr_sin_fecha)
    if not alertas_items:
        alertas_items = _alerta_card("&#10003;", "Sin alertas operativas", "",
                                     "Todos los registros tienen datos completos.",
                                     "#dcfce7","#15803d","#86efac")

    # =========================================================================
    # HTML EMAIL
    # =========================================================================
    SEC = lambda label, col="#2563eb": f'<p style="margin:0 0 12px;font-size:8px;font-weight:700;color:{col};text-transform:uppercase;letter-spacing:.18em">{label}</p>'

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Indicador Cumplimiento Citas — Llenos</title></head>
<body style="margin:0;padding:20px 0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif">
<table width="760" cellpadding="0" cellspacing="0" align="center" style="max-width:760px">

  <!-- ═══ HEADER ═══════════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#0d2137" style="background:#0d2137;border-radius:12px 12px 0 0;padding:26px 32px 22px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td>
        <table cellpadding="0" cellspacing="0"><tr>
          <td bgcolor="#ffffff22" style="background:rgba(255,255,255,0.12);border-radius:7px;padding:5px 12px">
            <span style="font-size:15px;font-weight:900;color:#fff;letter-spacing:-0.5px">TC</span>
            <span style="font-size:9px;color:rgba(255,255,255,0.5);margin-left:7px">TRACTOCAR LOGISTICS</span>
          </td>
        </tr></table>
        <p style="margin:14px 0 2px;font-size:8px;font-weight:700;color:#60a5fa;text-transform:uppercase;letter-spacing:.2em">Reporte Ejecutivo · Ajover EXPO</p>
        <p style="margin:0 0 4px;font-size:22px;font-weight:800;color:#fff;line-height:1.2">Indicador de Cumplimiento de Citas — Llenos</p>
        <p style="margin:0;font-size:11px;color:rgba(255,255,255,0.45)">{_ml(mes_actual)} &nbsp;·&nbsp; A corte: {corte_label} &nbsp;·&nbsp; {total_m} contenedores</p>
      </td>
      <td align="right" valign="top" style="font-size:10px;color:rgba(255,255,255,0.3);white-space:nowrap;padding-left:16px">{generado}</td>
    </tr></table>
  </td></tr>

  <!-- ═══ 4 KPI CARDS (como el tablero) ═══════════════════════════════════ -->
  <tr><td bgcolor="#0d2137" style="background:#0d2137;padding:0 32px 24px">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <!-- Total -->
        <td width="25%" style="padding-right:6px">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td bgcolor="#0f2d4a" style="background:#0f2d4a;border-radius:10px;padding:16px 14px;border:1px solid #1e3a5f">
              <p style="margin:0 0 4px;font-size:8px;color:#64748b;text-transform:uppercase;letter-spacing:.1em">Total Llenos</p>
              <p style="margin:0;font-size:30px;font-weight:800;color:#f8fafc;line-height:1">{total_m}</p>
              <p style="margin:6px 0 0;font-size:9px;color:#475569">registros</p>
            </td></tr>
          </table>
        </td>
        <!-- Cumplen cita -->
        <td width="25%" style="padding:0 3px">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td bgcolor="#0a2e1a" style="background:#0a2e1a;border-radius:10px;padding:16px 14px;border:1px solid #14532d">
              <p style="margin:0 0 4px;font-size:8px;color:#16a34a;text-transform:uppercase;letter-spacing:.1em">Cumplen Cita</p>
              <p style="margin:0;font-size:30px;font-weight:900;color:{col_m["bar"]};line-height:1">{pct_m}%</p>
              <table cellpadding="0" cellspacing="0" style="margin-top:8px"><tr>
                <td width="{min(80,int(pct_m/100*80))}" height="4" bgcolor="{col_m["bar"]}" style="background:{col_m["bar"]};font-size:0"> </td>
                {"<td width='" + str(80-min(80,int(pct_m/100*80))) + "' height='4' bgcolor='#14532d' style='background:#14532d;font-size:0'> </td>" if int(pct_m/100*80) < 80 else ""}
              </tr></table>
              <p style="margin:6px 0 0;font-size:9px;color:#166534">{cumpl_m} ok &nbsp;/&nbsp; {len(tarde_m)} no cumple</p>
            </td></tr>
          </table>
        </td>
        <!-- Reprog -->
        <td width="25%" style="padding:0 3px">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td bgcolor="#2d1b00" style="background:#2d1b00;border-radius:10px;padding:16px 14px;border:1px solid #78350f">
              <p style="margin:0 0 4px;font-size:8px;color:#d97706;text-transform:uppercase;letter-spacing:.1em">Reprogramaciones</p>
              <p style="margin:0;font-size:30px;font-weight:800;color:#f59e0b;line-height:1">{len(repr_m)}</p>
              <p style="margin:6px 0 0;font-size:9px;color:#92400e">del mes actual</p>
            </td></tr>
          </table>
        </td>
        <!-- Sin fecha -->
        <td width="25%" style="padding-left:6px">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td bgcolor="#1e1b2e" style="background:#1e1b2e;border-radius:10px;padding:16px 14px;border:1px solid #312e6e">
              <p style="margin:0 0 4px;font-size:8px;color:#818cf8;text-transform:uppercase;letter-spacing:.1em">Sin Fecha Cita</p>
              <p style="margin:0;font-size:30px;font-weight:800;color:#a5b4fc;line-height:1">{sf_m}</p>
              <p style="margin:6px 0 0;font-size:9px;color:#4c1d95">contenedores</p>
            </td></tr>
          </table>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- ═══ TENDENCIA MES A MES ═════════════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:24px 32px;border-top:1px solid #e2e8f0;margin-top:4px">
    {SEC("Cumplimiento de cita — mes a mes")}
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f8fafc" style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px">
      <tr><td style="padding:16px 8px 10px">
        {chart_html}
      </td></tr>
    </table>
    <p style="margin:10px 0 0;font-size:9px;color:#94a3b8">
      <span style="color:#22c55e">&#9646;</span> &ge;90%&nbsp;&nbsp;
      <span style="color:#f59e0b">&#9646;</span> 75–90%&nbsp;&nbsp;
      <span style="color:#ef4444">&#9646;</span> &lt;75%&nbsp;&nbsp;&nbsp;
      Barra más resaltada = mes actual
    </p>
  </td></tr>

  <!-- ═══ INCUMPLIMIENTOS MES ══════════════════════════════════════════════ -->
  {"<tr><td bgcolor='#ffffff' style='background:#fff;padding:20px 32px;border-top:1px solid #e2e8f0'>" + SEC("Distribución de motivos", "#dc2626") + "<table cellpadding='0' cellspacing='4' style='margin-bottom:18px'>" + mot_rows + "</table>" + SEC("Incumplimientos " + _ml(mes_actual,True) + " — responsables", "#dc2626") + "<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #f1f5f9;border-radius:8px;overflow:hidden'><thead><tr><th " + TH_S + ">Responsable</th><th " + TH_S + " style='text-align:center'>Casos</th><th " + TH_S + ">Principales motivos</th></tr></thead><tbody>" + inc_mes + "</tbody></table></td></tr>" if tarde_m else ""}

  <!-- ═══ CITAS PERDIDAS MES ══════════════════════════════════════════════ -->
  {"<tr><td bgcolor='#ffffff' style='background:#fff;padding:16px 32px;border-top:1px solid #e2e8f0'>" + SEC("Citas perdidas " + _ml(mes_actual,True) + " — " + str(len(repr_m)) + " en total", "#d97706") + repr_mes + "</td></tr>" if repr_m else ""}

  <!-- ═══ DIVIDER AYER ════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#1e3a6e" style="background:#1e3a6e;padding:11px 32px;margin-top:4px">
    <p style="margin:0 0 2px;font-size:8px;font-weight:700;color:#93c5fd;text-transform:uppercase;letter-spacing:.2em">Detalle del día anterior</p>
    <p style="margin:0;font-size:14px;font-weight:700;color:#fff">{dia_label}</p>
  </td></tr>

  <!-- ═══ KPIs AYER ════════════════════════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:20px 32px;border-top:1px solid #e2e8f0">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td style="text-align:center;padding-right:12px;border-right:1px solid #f1f5f9;vertical-align:middle">
          <table cellpadding="0" cellspacing="0" align="center">
            <tr><td bgcolor="{col_d["light"]}" style="background:{col_d["light"]};border-radius:10px;padding:14px 18px;border:2px solid {col_d["bar"]}">
              <p style="margin:0;font-size:28px;font-weight:900;color:{col_d["text"]};line-height:1;text-align:center">{dia_pct}%</p>
              <p style="margin:4px 0 0;font-size:7px;font-weight:700;color:{col_d["text"]};text-transform:uppercase;letter-spacing:.1em;opacity:.7;text-align:center">cumplimiento</p>
            </td></tr>
          </table>
        </td>
        <td style="padding-left:20px;vertical-align:middle">
          <table cellpadding="0" cellspacing="0"><tr>
            <td style="padding:0 16px 0 0;text-align:center;border-right:1px solid #e2e8f0">
              <p style="margin:0;font-size:26px;font-weight:800;color:#1e293b;line-height:1">{dia_total}</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">Total</p>
            </td>
            <td style="padding:0 16px;text-align:center;border-right:1px solid #e2e8f0">
              <p style="margin:0;font-size:26px;font-weight:800;color:#16a34a;line-height:1">{dia_cumpl}</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">A tiempo</p>
            </td>
            <td style="padding:0 16px;text-align:center;border-right:1px solid #e2e8f0">
              <p style="margin:0;font-size:26px;font-weight:800;color:#dc2626;line-height:1">{len(dia_tarde)}</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">Tarde</p>
            </td>
            <td style="padding:0 0 0 16px;text-align:center">
              <p style="margin:0;font-size:26px;font-weight:800;color:#d97706;line-height:1">{len(dia_repr)}</p>
              <p style="margin:3px 0 0;font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em">C. perdidas</p>
            </td>
          </tr></table>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- ═══ RESPONSABILIDAD AYER ═══════════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:0 32px 16px;border-top:1px solid #f1f5f9">
    {SEC("Análisis de responsabilidad — ayer")}
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="33%" style="padding-right:6px">{_resp_card("Tractocar", n_tc, "#2563eb", "#eff6ff", "#bfdbfe")}</td>
      <td width="33%" style="padding:0 3px">   {_resp_card("Ajover",    n_ajv,"#dc2626","#fee2e2","#fca5a5")}</td>
      <td width="33%" style="padding-left:6px">{_resp_card("Externo",   n_ext,"#16a34a","#dcfce7","#86efac")}</td>
    </tr></table>
    {resp_split_bar}
  </td></tr>

  <!-- ═══ ALERTAS OPERATIVAS AYER (solo correo interno) ═══════════════════ -->
  {"<tr><td bgcolor='#f8fafc' style='background:#f8fafc;padding:16px 32px;border-top:1px solid #e2e8f0;border-bottom:1px solid #e2e8f0'>" + SEC('Alertas operativas · ' + dia_label[:20], '#d97706' if n_alertas else '#16a34a') + alertas_items + "</td></tr>" if interno else ""}

  <!-- ═══ TODOS LOS CONTENEDORES AYER ═════════════════════════════════════ -->
  <tr><td bgcolor="#ffffff" style="background:#fff;padding:0 32px 24px;border-top:1px solid #e2e8f0">
    {SEC("Todos los contenedores de ayer")}
    {"<table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #f1f5f9;border-radius:8px;overflow:hidden'><thead><tr><th " + TH_S + ">Pedido</th><th " + TH_S + ">Contenedor</th><th " + TH_S + ">Terminal</th><th " + TH_S + ">Cita puerto</th><th " + TH_S + ">Llegada</th><th " + TH_S + ">Nueva cita</th><th " + TH_S + ">Resp. reprog.</th><th " + TH_S + ">Cumplimiento</th><th " + TH_S + ">Motivo</th><th " + TH_S + ">Observación</th></tr></thead><tbody>" + det_ayer + "</tbody></table>" if det_ayer else "<p style='margin:0;font-size:12px;color:#94a3b8;text-align:center;padding:12px'>Sin registros para este día</p>"}
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


EMAIL_TO_INTERNO = [e.strip() for e in os.environ.get("EMAIL_TO_INTERNO", EMAIL_FROM).split(",") if e.strip()]
EMAIL_CC_INTERNO = [e.strip() for e in os.environ.get("EMAIL_CC_INTERNO","").split(",") if e.strip()]


def send_report(interno=False):
    datos_path = BASE / "datos_bot.json"
    if not datos_path.exists():
        print("[ERROR] No existe datos_bot.json"); sys.exit(1)
    with open(datos_path, encoding="utf-8") as f:
        datos = json.load(f)
    if not EMAIL_PASSWORD:
        print("[ERROR] Falta EMAIL_PASSWORD en .env"); sys.exit(1)

    html      = build_html(datos, interno=interno)
    if interno:
        dest_to  = EMAIL_TO_INTERNO
        dest_cc  = EMAIL_CC_INTERNO
        tag      = "[TC-Interno]"
    else:
        dest_to  = EMAIL_TO
        dest_cc  = EMAIL_CC
        tag      = "[TC]"
    subject = f"{tag} Indicador Cumplimiento Citas Llenos — {datos.get('generado','')[:10]}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = EMAIL_FROM
    msg["To"]      = ", ".join(dest_to)
    if dest_cc: msg["Cc"] = ", ".join(dest_cc)
    msg.attach(MIMEText(html, "html", "utf-8"))

    dest_all = dest_to + dest_cc
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo(); smtp.starttls()
            smtp.login(EMAIL_FROM, EMAIL_PASSWORD)
            smtp.sendmail(EMAIL_FROM, dest_all, msg.as_string())
        modo = "INTERNO (con alertas)" if interno else "CLIENTE (sin alertas)"
        print(f"[OK] {modo} — enviado a {', '.join(dest_all)}")
    except Exception as e:
        print(f"[ERROR] {e}"); sys.exit(1)


if __name__ == "__main__":
    modo = sys.argv[1] if len(sys.argv) > 1 else "--cliente"
    if modo == "--interno":
        send_report(interno=True)
    elif modo == "--cliente":
        send_report(interno=False)
    else:
        print(f"Uso: python enviar_reporte.py [--interno | --cliente]")
        print("  --interno  8:30 AM → equipo COMEX (incluye alertas operativas)")
        print("  --cliente  9:00 AM → cliente Ajover (sin alertas operativas)")
        sys.exit(1)
