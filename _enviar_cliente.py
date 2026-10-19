import sys, os, json, subprocess
from pathlib import Path
from datetime import datetime

os.chdir(r'C:\Users\jarias\Desktop\tractocar-ventas')
ruta     = Path(r'C:\Users\jarias\Desktop\tractocar-ventas\datos_bot.json')
log_path = Path(r'C:\Users\jarias\Desktop\tractocar-ventas\sent_log.json')

hoy = datetime.now().date().isoformat()

if not ruta.exists():
    print('ERROR: datos_bot.json no existe'); sys.exit(1)

mod = datetime.fromtimestamp(ruta.stat().st_mtime).date().isoformat()
if mod < hoy:
    print(f'SKIP: datos_bot.json actualizado el {mod}, hoy es {hoy}. Actualiza el token primero.')
    sys.exit(0)

# Verificar si ya se envió hoy
log = {}
if log_path.exists():
    try: log = json.loads(log_path.read_text(encoding='utf-8'))
    except: pass

if log.get('llenos_cliente') == hoy:
    print(f'SKIP: correo cliente llenos ya enviado hoy ({hoy}).')
    sys.exit(0)

print(f'OK: enviando correo cliente llenos...')
result = subprocess.run([sys.executable, 'enviar_reporte.py', '--cliente'], capture_output=True, text=True)
print(result.stdout)
if result.returncode != 0:
    print('ERROR:', result.stderr); sys.exit(1)

log['llenos_cliente'] = hoy
log_path.write_text(json.dumps(log, indent=2), encoding='utf-8')
