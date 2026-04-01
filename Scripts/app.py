#!/usr/bin/env python3
"""
Servidor web para Generador de Reportes Telcel Empresas
Ejecutar: python app.py
Abrir: http://localhost:5000
"""

import csv
import glob
import json
import os
import queue
import sys
import threading
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, Response, abort, jsonify, render_template, request, send_file

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

from generar_reportes import ReporteLeads
import gmail_drafts

# Catálogo completo de Soluciones → Servicios (fuente: SFMC)
CATALOG = {
    'Seguridad':                  ['Control de Dispositivos MDM', 'Control Móvil Telcel', 'Norton Antivirus Empresarial', 'Lookout', 'VMware Workspace ONE'],
    'Gestión de Fuerza en Campo': ['Localización Empresarial', 'Cobranza', 'Promotoría y Merchandising', 'Servicios en Campo', 'Control de Sucursales', 'Seguridad', 'Logística'],
    'Servicios Cloud':            ['Microsoft 365', 'Hosting y Desarrollo Web', 'Claro Drive Negocio', 'Google Workspace'],
    'Business Intelligence':      ['GEODATA', 'Autenticación Móvil', 'Indicadores Móviles'],
    'Mobile Marketing':           ['Mensajes RCS', 'Internet Patrocinado', 'Distribución de Aplicaciones', 'Mensajería Masiva', 'Recompensas'],
    'Comunicación Empresarial':   ['Control de Ventas', 'Formularios Verum', 'SekurMessenger', 'Push to Talk'],
    'Conectividad Telcel':        ['eSIM', 'Conectividad Avanzada M2M', 'Redes Privadas de Datos', 'Internet Gestionado'],
    'Gestión Vehicular Telcel':   ['Video a bordo', 'Cadena de Frío', 'Hábitos de Conducción', 'Medición de Combustible', 'Módulo de Ruteo', 'Telemetría'],
    'Planes Empresariales':       ['Plan Telcel Empresa', 'Plan Telcel Ultra Empresa'],
}

# Normalización de valores que llegan en el CSV → nombre canónico
SOL_NORMALIZE = {
    'seguridad': 'Seguridad',
    'gestión de fuerza en campo': 'Gestión de Fuerza en Campo',
    'gestion de fuerza en campo': 'Gestión de Fuerza en Campo',
    'servicios cloud': 'Servicios Cloud',
    'business intelligence': 'Business Intelligence',
    'mobile marketing': 'Mobile Marketing',
    'comunicación empresarial': 'Comunicación Empresarial',
    'comunicacion empresarial': 'Comunicación Empresarial',
    'conectividad telcel': 'Conectividad Telcel',
    'gestión vehicular telcel': 'Gestión Vehicular Telcel',
    'gestion vehicular telcel': 'Gestión Vehicular Telcel',
    'planes empresariales': 'Planes Empresariales',
    # Códigos SFMC
    'seg': 'Seguridad',
    'gfc': 'Gestión de Fuerza en Campo',
    'sec': 'Servicios Cloud',
    'bui': 'Business Intelligence',
    'mom': 'Mobile Marketing',
    'coe': 'Comunicación Empresarial',
    'cov': 'Conectividad Telcel',
    'cot': 'Conectividad Telcel',
    'gvt': 'Gestión Vehicular Telcel',
    'pne': 'Planes Empresariales',
    'pte': 'Planes Empresariales',
    'teu': 'Planes Empresariales',
}

app = Flask(__name__)

_output_queue = queue.Queue()
_generation_lock = threading.Lock()

# Resultados de la última generación (para el módulo Gmail)
_last_result = {
    'archivos_por_region': {},   # { 'R1': '/ruta/R1-leads-...xlsx', ... }
    'contraseñas': {},            # { 'R1': 'R1TELXXX', ... }
    'periodo_legible': '',
}


class _StreamCapture:
    """Redirige print() al queue de SSE sin perder stdout original."""
    def __init__(self, q, original):
        self.q = q
        self.original = original

    def write(self, text):
        self.original.write(text)
        if text.strip():
            self.q.put({'type': 'log', 'text': text.rstrip('\n')})

    def flush(self):
        self.original.flush()


def _run_generation(fechas_dates):
    global _last_result
    old_stdout = sys.stdout
    sys.stdout = _StreamCapture(_output_queue, old_stdout)
    try:
        config_path = os.path.join(script_dir, 'config.json')
        generador = ReporteLeads(config_path, fechas_reporte=fechas_dates)
        fecha_hoy = datetime.now().strftime('%d%m%Y')
        carpeta_csv = generador.config['rutas']['carpeta_csv']
        csv_path = os.path.join(carpeta_csv, f'leads_diarios_{fecha_hoy}.csv')
        ok = generador.generar_todos_reportes(csv_path)

        # Guardar resultados para Gmail
        _last_result = {
            'archivos_por_region': {
                a['region']: a['ruta'] for a in generador.archivos_generados
            },
            'contraseñas': dict(generador.contraseñas_generadas),
            'periodo_legible': generador._periodo_legible(),
        }

        carpeta = Path(generador.config['rutas']['carpeta_salida'])
        archivos = sorted(
            [f for f in carpeta.glob('*') if f.suffix in ('.xlsx', '.txt')],
            key=lambda f: f.stat().st_mtime,
            reverse=True
        )
        files = [f.name for f in archivos]
        _output_queue.put({'type': 'done', 'success': ok, 'archivos': files})
    except Exception as e:
        _output_queue.put({'type': 'error', 'text': str(e)})
        _output_queue.put({'type': 'done', 'success': False, 'archivos': []})
    finally:
        sys.stdout = old_stdout
        _generation_lock.release()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generar', methods=['POST'])
def generar():
    if not _generation_lock.acquire(blocking=False):
        return jsonify({'error': 'Ya hay una generación en proceso'}), 409

    try:
        data = request.json
        fecha_inicio = datetime.strptime(data['fecha_inicio'], '%Y-%m-%d').date()
        fecha_fin = datetime.strptime(data['fecha_fin'], '%Y-%m-%d').date()

        if fecha_inicio > fecha_fin:
            _generation_lock.release()
            return jsonify({'error': 'La fecha inicio no puede ser mayor a la fecha fin'}), 400

        while not _output_queue.empty():
            try:
                _output_queue.get_nowait()
            except queue.Empty:
                break

        fechas = []
        d = fecha_inicio
        while d <= fecha_fin:
            fechas.append(d)
            d += timedelta(days=1)

        t = threading.Thread(target=_run_generation, args=(fechas,), daemon=True)
        t.start()
        return jsonify({'status': 'started'})

    except Exception as e:
        _generation_lock.release()
        return jsonify({'error': str(e)}), 500


@app.route('/stream')
def stream():
    def generate():
        while True:
            try:
                msg = _output_queue.get(timeout=60)
                yield f'data: {json.dumps(msg)}\n\n'
                if msg.get('type') in ('done', 'error'):
                    break
            except queue.Empty:
                yield 'data: {"type":"ping"}\n\n'

    return Response(
        generate(),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'}
    )


@app.route('/descargar/<path:filename>')
def descargar(filename):
    config_path = os.path.join(script_dir, 'config.json')
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    carpeta = Path(config['rutas']['carpeta_salida'])
    filepath = carpeta / filename
    if not filepath.exists() or not filepath.is_file():
        abort(404)
    return send_file(str(filepath.resolve()), as_attachment=True, download_name=filename)


@app.route('/gmail/estado')
def gmail_estado():
    """Devuelve si hay token válido guardado."""
    autorizado = gmail_drafts.token_valido(script_dir)
    tiene_datos = bool(_last_result.get('periodo_legible'))
    return jsonify({'autorizado': autorizado, 'tiene_datos': tiene_datos})


@app.route('/gmail/borradores', methods=['POST'])
def gmail_borradores():
    """Crea los 11 borradores en Gmail. Abre el navegador la primera vez para autorizar."""
    if not _last_result.get('periodo_legible'):
        return jsonify({'error': 'Primero genera los reportes.'}), 400

    try:
        resultados = gmail_drafts.crear_borradores(
            script_dir=script_dir,
            archivos_por_region=_last_result['archivos_por_region'],
            contraseñas=_last_result['contraseñas'],
            periodo_legible=_last_result['periodo_legible'],
        )
        total_ok = sum(1 for r in resultados if r['ok'])
        return jsonify({'ok': True, 'total': len(resultados), 'creados': total_ok, 'detalle': resultados})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/conteo')
def conteo():
    return render_template('conteo.html')


@app.route('/api/conteo', methods=['POST'])
def api_conteo():
    try:
        data = request.json
        fecha_inicio = datetime.strptime(data['fecha_inicio'], '%Y-%m-%d').date()
        fecha_fin    = datetime.strptime(data['fecha_fin'],    '%Y-%m-%d').date()

        config_path = os.path.join(script_dir, 'config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        carpeta_csv = config['rutas']['carpeta_csv']
        fecha_hoy = datetime.now().strftime('%d%m%Y')
        csv_path = os.path.join(carpeta_csv, f'leads_diarios_{fecha_hoy}.csv')
        if not os.path.exists(csv_path):
            return jsonify({'error': f'No se encontró el archivo leads_diarios_{fecha_hoy}.csv'}), 404

        REGIONES  = ['R1','R2','R3','R4','R5','R6','R7','R8','R9','MOM','BUI']
        FORMATOS  = [
            '%Y-%m-%d %H:%M:%S', '%m/%d/%Y %I:%M:%S %p', '%m/%d/%Y %H:%M:%S',
            '%d-%m-%Y %H:%M:%S', '%m/%d/%Y %H:%M',       '%m/%d/%Y %I:%M %p',
            '%Y-%m-%d %H:%M',
        ]
        counts    = {}
        # Inicializar desglose con el catálogo completo en 0
        breakdown = {sol: {srv: 0 for srv in srvs} for sol, srvs in CATALOG.items()}

        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    fecha_str = row.get('CreatedAt', '').strip()
                    fecha = None
                    for fmt in FORMATOS:
                        try:
                            fecha = datetime.strptime(fecha_str, fmt).date()
                            break
                        except ValueError:
                            continue
                    if not fecha or not (fecha_inicio <= fecha <= fecha_fin):
                        continue

                    region   = row.get('Region', '').upper().strip()
                    solucion = row.get('SolucionInteres', '').upper().strip()

                    if 'MOBILE MARKETING' in solucion or solucion == 'MOM':
                        bucket = 'MOM'
                    elif 'BUSINESS INTELLIGENCE' in solucion or solucion == 'BUI':
                        bucket = 'BUI'
                    elif region in ['R1','R2','R3','R4','R5','R6','R7','R8','R9']:
                        bucket = region
                    else:
                        continue

                    # Conteo por día/región
                    key = fecha.strftime('%d/%m/%Y')
                    if key not in counts:
                        counts[key] = {r: 0 for r in REGIONES}
                    counts[key][bucket] += 1

                    # Desglose: normalizar solución al nombre canónico
                    sol_raw = row.get('SolucionInteres', '').strip()
                    srv_raw = row.get('ServicioInteres', '').strip() or '(Sin especificar)'
                    sol_key = SOL_NORMALIZE.get(sol_raw.lower(), sol_raw) or '(Sin especificar)'

                    if sol_key not in breakdown:
                        breakdown[sol_key] = {}
                    if srv_raw not in breakdown[sol_key]:
                        breakdown[sol_key][srv_raw] = 0
                    breakdown[sol_key][srv_raw] += 1

                except Exception:
                    continue

        all_dates = sorted(counts.keys(), key=lambda d: datetime.strptime(d, '%d/%m/%Y'))
        rows = [{'dia': d, **counts[d]} for d in all_dates]

        # Serializar en orden del catálogo; extras al final
        catalog_order = list(CATALOG.keys())
        extras = [s for s in breakdown if s not in catalog_order]
        desglose = []
        for sol in catalog_order + extras:
            if sol not in breakdown:
                continue
            servicios_dict = breakdown[sol]
            sol_total = sum(servicios_dict.values())
            # Servicios del catálogo primero (en orden), luego extras
            cat_srvs = CATALOG.get(sol, [])
            extra_srvs = [s for s in servicios_dict if s not in cat_srvs]
            ordered_srvs = cat_srvs + extra_srvs
            desglose.append({
                'solucion': sol,
                'total': sol_total,
                'servicios': [
                    {'servicio': srv, 'total': servicios_dict.get(srv, 0)}
                    for srv in ordered_srvs
                ]
            })

        return jsonify({'rows': rows, 'desglose': desglose, 'csv': os.path.basename(csv_path)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("\n========================================")
    print("  Reportes Telcel Empresas - Servidor")
    print("========================================")
    print("  Abre en tu navegador:")
    print("  http://localhost:5000")
    print("========================================\n")
    app.run(debug=False, port=5000, host='127.0.0.1', threaded=True)
