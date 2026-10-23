"""
renovar_cookie.py — renueva CF_Authorization en .env automáticamente.

Flujo:
  1. Playwright abre despachoscomex.tractocar.com (headless)
  2. Cloudflare Access envía código OTP a jarias@tractocar.com
  3. Lee el código desde Outlook local (win32com) — sin login adicional
  4. Playwright ingresa el código → Cloudflare autentica
  5. Extrae CF_Authorization via CDP y actualiza .env

Uso:
  python renovar_cookie.py            → solo renueva la cookie
  python renovar_cookie.py --full     → renueva + procesar.py + 3 correos
"""
import os, re, sys, time, subprocess
from pathlib import Path
from dotenv import load_dotenv

ENV_PATH = Path(r'C:\Users\jarias\Desktop\tractocar-ventas\.env')
load_dotenv(ENV_PATH, override=True)

# ── Leer código OTP desde Outlook local ──────────────────────────────────────
def esperar_otp_outlook(desde_ts_unix, max_espera=120):
    """Lee el último correo de Cloudflare Access llegado después de desde_ts_unix.
    Outlook/win32com devuelve ReceivedTime en hora LOCAL — usar timestamp Unix para comparar.
    desde_ts_unix = time.time() tomado justo antes de iniciar el navegador.
    """
    import win32com.client

    outlook = win32com.client.Dispatch('Outlook.Application')
    ns      = outlook.GetNamespace('MAPI')
    inbox   = ns.GetDefaultFolder(6)  # 6 = Inbox

    print('  Esperando código en Outlook', end='', flush=True)
    deadline = time.time() + max_espera
    ya_visto = set()

    while time.time() < deadline:
        items = inbox.Items
        items.Sort('[ReceivedTime]', True)  # más reciente primero
        for item in items:
            try:
                rt = item.ReceivedTime
                # mktime convierte hora local a unix timestamp (igual que time.time())
                recv_ts = time.mktime(time.struct_time(
                    (rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second, 0, 0, -1)
                ))
                # Detenerse cuando los emails son más viejos que el inicio del script
                if recv_ts < desde_ts_unix - 120:
                    break
                sender = getattr(item, 'SenderEmailAddress', '') or ''
                subj   = getattr(item, 'Subject', '') or ''
                if 'cloudflare' in sender.lower() or 'cloudflare' in subj.lower():
                    body = item.Body or ''
                    m = re.search(r'\b(\d{6})\b', body)
                    if m:
                        code = m.group(1)
                        item_id = getattr(item, 'EntryID', recv_ts)
                        if item_id not in ya_visto:
                            print(f' ✓  ({code})')
                            ya_visto.add(item_id)
                            return code
            except Exception:
                continue
        print('.', end='', flush=True)
        time.sleep(5)
    raise TimeoutError('No llegó el código OTP en Outlook en 2 minutos')

# ── Actualizar .env ───────────────────────────────────────────────────────────
def actualizar_env(cookie):
    text = ENV_PATH.read_text(encoding='utf-8')
    text = re.sub(r'^WEBAPP_COOKIE=.*$', f'WEBAPP_COOKIE={cookie}', text, flags=re.MULTILINE)
    ENV_PATH.write_text(text, encoding='utf-8')
    print('[OK] WEBAPP_COOKIE actualizado en .env')

# ── Flujo principal con Playwright ────────────────────────────────────────────
def renovar_cookie():
    from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

    print('[1/4] Iniciando navegador headless...')
    ts_antes = time.time() - 30  # unix timestamp con margen de 30s

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/120.0.0.0 Safari/537.36'
        )
        page = ctx.new_page()

        print('[2/4] Navegando al web app...')
        page.goto('https://despachoscomex.tractocar.com',
                  wait_until='domcontentloaded', timeout=30000)

        # Si no redirigió a Cloudflare ya estamos autenticados
        if 'cloudflareaccess.com' not in page.url:
            cookies = ctx.cookies(['https://despachoscomex.tractocar.com'])
            cf = next((c['value'] for c in cookies if c['name'] == 'CF_Authorization'), None)
            if cf:
                print('[INFO] Cookie válida encontrada en sesión existente.')
                browser.close()
                return cf

        # Página de login de Cloudflare
        print('     Ingresando email en Cloudflare Access...')
        page.wait_for_selector('input[type="email"], input[name="email"]', timeout=15000)
        page.fill('input[type="email"]', 'jarias@tractocar.com')
        page.click('button:has-text("Send login code"), input[type="submit"]')

        # Esperar la pantalla del código
        page.wait_for_selector('input[name="code"], input[placeholder*="000000"]', timeout=15000)

        print('[3/4] Leyendo código desde Outlook...')
        otp = esperar_otp_outlook(ts_antes)

        print('[4/4] Verificando código con Cloudflare...')
        page.fill('input[name="code"], input[placeholder*="000000"]', otp)
        page.click('button:has-text("Verify"), input[type="submit"]')

        # Esperar redirección de vuelta al web app
        try:
            page.wait_for_url('https://despachoscomex.tractocar.com/**', timeout=20000)
        except PwTimeout:
            pass

        cookies = ctx.cookies(['https://despachoscomex.tractocar.com'])
        cf_cookie = next((c['value'] for c in cookies if c['name'] == 'CF_Authorization'), None)
        browser.close()

    if not cf_cookie:
        raise RuntimeError('No se encontró CF_Authorization tras la autenticación.')
    return cf_cookie

# ── Helpers para --full ───────────────────────────────────────────────────────
def _hoy():
    """Fecha de hoy en Colombia (UTC-5) como string YYYY-MM-DD."""
    import datetime
    return (datetime.datetime.utcnow() - datetime.timedelta(hours=5)).strftime('%Y-%m-%d')

def _sent_log(base):
    import json
    p = base / 'sent_log.json'
    try:
        return json.loads(p.read_text(encoding='utf-8')) if p.exists() else {}
    except Exception:
        return {}

def _datos_frescos(base):
    """True si datos_bot.json fue guardado hoy (Colombia)."""
    import datetime
    p = base / 'datos_bot.json'
    if not p.exists():
        return False
    mtime = datetime.datetime.utcfromtimestamp(p.stat().st_mtime) - datetime.timedelta(hours=5)
    return mtime.strftime('%Y-%m-%d') == _hoy()

# ─────────────────────────────────────────────────────────────────────────────
def main():
    try:
        cookie = renovar_cookie()
        actualizar_env(cookie)
    except Exception as e:
        print(f'[ERROR renovar_cookie] {e}')
        sys.exit(1)

    if '--full' not in sys.argv:
        return

    base = Path(__file__).parent
    py   = sys.executable
    hoy  = _hoy()
    log  = _sent_log(base)

    # ── ¿Ya enviamos todo hoy? ────────────────────────────────────────────────
    pendientes = [k for k in ('llenos_interno', 'llenos_cliente', 'vacios_interno')
                  if log.get(k) != hoy]
    if not pendientes:
        print(f'[OK] Todos los correos ya fueron enviados hoy ({hoy}). Nada que hacer.')
        return

    print(f'[INFO] Pendientes hoy: {", ".join(pendientes)}')

    # ── Actualizar datos si no están frescos ─────────────────────────────────
    if _datos_frescos(base):
        print(f'[OK] datos_bot.json ya está actualizado hoy. Saltando procesar.py.')
    else:
        print('\n[AUTO] Corriendo procesar.py...')
        r = subprocess.run([py, str(base / 'procesar.py')], cwd=str(base))
        if r.returncode != 0:
            print('[ERROR] procesar.py falló — abortando envío de correos.')
            sys.exit(1)

    # ── Enviar solo los correos pendientes ────────────────────────────────────
    cmds = {
        'llenos_interno':  [py, str(base / 'enviar_reporte.py'), '--interno'],
        'llenos_cliente':  [py, str(base / 'enviar_reporte.py'), '--cliente'],
        'vacios_interno':  [py, str(base / 'enviar_reporte_vacios.py')],
    }
    for key in ('llenos_interno', 'llenos_cliente', 'vacios_interno'):
        if key not in pendientes:
            print(f'[SKIP] {key} ya enviado hoy.')
            continue
        print(f'[SEND] {key}...')
        r = subprocess.run(cmds[key], capture_output=True, text=True, cwd=str(base))
        out = (r.stdout + r.stderr).strip()
        print(f'       {out}')

if __name__ == '__main__':
    main()
