# -*- coding: utf-8 -*-
import io
import os
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from flask import Flask, render_template_string, request, redirect, url_for, session, send_file, send_from_directory
from datetime import datetime
import openpyxl

app = Flask(__name__)
app.secret_key = 'clave_secreta_denier_inventario_maestro'

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

LOGO_DENIER_SVG = """
<svg width="60" height="60" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M50 15 L85 50 L50 85 L15 50 Z" stroke="#38bdf8" stroke-width="4" fill="none"/>
    <path d="M50 25 L75 50 L50 75 L25 50 Z" stroke="#38bdf8" stroke-width="3" fill="none"/>
    <path d="M30 30 L70 70 M70 30 L30 70" stroke="#38bdf8" stroke-width="5" stroke-linecap="round"/>
    <path d="M50 15 L50 85 M15 50 L85 50" stroke="#38bdf8" stroke-width="5" stroke-linecap="round"/>
</svg>
"""

LOGO_PRINT_SVG = """
<svg width="50" height="50" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M50 15 L85 50 L50 85 L15 50 Z" stroke="#000" stroke-width="4" fill="none"/>
    <path d="M50 25 L75 50 L50 75 L25 50 Z" stroke="#000" stroke-width="3" fill="none"/>
    <path d="M30 30 L70 70 M70 30 L30 70" stroke="#000" stroke-width="5" stroke-linecap="round"/>
    <path d="M50 15 L50 85 M15 50 L85 50" stroke="#000" stroke-width="5" stroke-linecap="round"/>
</svg>
"""

usuarios_db = [
    { 'id': 1, 'email': 'controlprod_1@deniersa.com', 'password': '123', 'nombre': 'Administrador General', 'modulos': ['admin', 'hilos', 'cruda', 'devo'] },
    { 'id': 2, 'email': 'bodega.hilos@deniersa.com', 'password': '123', 'nombre': 'Operador Hilos', 'modulos': ['hilos'] },
    { 'id': 3, 'email': 'bodega.cruda@deniersa.com', 'password': '123', 'nombre': 'Operador Tela Cruda', 'modulos': ['cruda'] },
    { 'id': 4, 'email': 'bodega.devoluciones@deniersa.com', 'password': '123', 'nombre': 'Operador Devoluciones', 'modulos': ['devo'] }
]

inventario_hilos = [
    { 'id': 1, 'tipo': 'Denier', 'oc': 'P04150', 'codigo': 'H-COT-20', 'calibre': '20/1', 'composicion': 'Algodon 100%', 'proveedor': 'Hilos del Sur', 'lote': 'L-442', 'kilos': 450.5, 'conos': 120, 'ubicacion': 'A-1' }
]
packing_history_hilos = []
contador_packing_hilos = 1

inventario_tela = [
    { 'id': 1, 'oc': 'OC-9012', 'proveedor': 'Knitopia', 'lote_tela': 'TC-101', 'rollo': 'R-001', 'kilos': 250.5, 'ubicacion': 'Bodega A-1' }
]
packing_tela_history = []
contador_packing_tela = 1

devoluciones_procesada = [
    { 'id': 1, 'cliente': 'Magic Apparel', 'lote': '20452026-8', 'tela': 'Interlock, 100% Algodón', 'color': 'White', 'rollos': 3, 'kilos': 75.50, 'yardas': 240.00, 'calidad': 'Rechazada', 'ubicacion': 'Rack 3', 'archivo': None }
]
despachos_devoluciones_history = []
contador_despacho_devo = 1


LOGIN_HTML = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>DENIER - Login</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen flex flex-col justify-between p-6">
    <div class="flex-grow flex items-center justify-center">
        <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-sm">
            <div class="text-center mb-6">
                <div class="flex justify-center mb-2">""" + LOGO_DENIER_SVG + """</div>
                <h1 class="text-2xl font-black tracking-widest text-white uppercase">DENIER</h1>
                <div class="border-b border-slate-800 my-4"></div>
            </div>
            {% if error %}<div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded-lg mb-4 text-xs font-bold text-center">{{ error }}</div>{% endif %}
            <form action="/login" method="POST" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1.5 text-slate-300">Email</label><input type="email" name="email" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="controlprod_1@deniersa.com"></div>
                <div><label class="block text-xs font-bold mb-1.5 text-slate-300">Password</label><input type="password" name="password" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="••••••••"></div>
                <button type="submit" class="w-full bg-teal-700 hover:bg-teal-600 text-white py-2.5 rounded-lg font-bold text-sm shadow">Log in</button>
            </form>
            <div class="flex justify-between items-center mt-6 text-xs text-slate-400">
                <span class="text-slate-500">¿Problemas?</span>
                <a href="/reset-password-form" class="text-teal-400 font-semibold hover:underline">Reset Password</a>
            </div>
        </div>
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider">Creado: Estuardo Alonzo</div>
</body>
</html>
"""

HOME_HTML = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>DENIER - Panel</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen p-8 flex flex-col justify-between">
    <div class="max-w-6xl mx-auto w-full">
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8 flex justify-between items-center">
            <div class="flex items-center gap-4">
                """ + LOGO_DENIER_SVG + """
                <div>
                    <h1 class="text-2xl font-black tracking-wide text-white uppercase">DENIER</h1>
                    <p class="text-sm text-slate-400">Usuario: <b class="text-teal-400">{{ session.get('nombre') }}</b></p>
                </div>
            </div>
            <a href="/logout" class="bg-red-950 text-red-300 border border-red-800 px-4 py-2 rounded-xl text-sm font-semibold">Cerrar Sesión</a>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div><h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Hilos</h2><p class="text-sm text-slate-400 mb-4">Control de inventario, packing lists y devoluciones.</p></div>
                <a href="/hilos" class="bg-teal-700 hover:bg-teal-600 text-white text-center py-3 rounded-xl font-medium shadow">Entrar a Hilos</a>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div><h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Tela Cruda</h2><p class="text-sm text-slate-400 mb-4">Control de rollos y cargas masivas Knitopia.</p></div>
                <a href="/cruda" class="bg-teal-700 hover:bg-teal-600 text-white text-center py-3 rounded-xl font-medium shadow">Entrar a Tela Cruda</a>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div><h2 class="text-xl font-bold text-rose-400 mb-2">Módulo Devoluciones</h2><p class="text-sm text-slate-400 mb-4">Control de tela devuelta y salidas por venta o reproceso.</p></div>
                <a href="/devo" class="bg-rose-700 hover:bg-rose-600 text-white text-center py-3 rounded-xl font-medium shadow">Entrar a Devoluciones</a>
            </div>
        </div>
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider mt-8">Creado: Estuardo Alonzo</div>
</body>
</html>
"""


@app.route('/')
def index():
    if 'email' not in session: return redirect(url_for('login_form'))
    return render_template_string(HOME_HTML, usuarios_db=usuarios_db)

@app.route('/login-form')
def login_form(): return render_template_string(LOGIN_HTML, error=None)

@app.route('/login', methods=['POST'])
def login():
    email = request.form.get('email', '').strip()
    password = request.form.get('password', '').strip()
    usuario = next((u for u in usuarios_db if u['email'] == email and u['password'] == password), None)
    if usuario:
        session['email'] = usuario['email']
        session['nombre'] = usuario['nombre']
        session['modulos'] = usuario['modulos']
        return redirect(url_for('index'))
    return render_template_string(LOGIN_HTML, error="Correo o contraseña incorrectos")

@app.route('/reset-password-form')
def reset_password_form():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 min-h-screen flex items-center justify-center p-6"><div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-sm"><h1 class="text-xl font-bold text-white mb-2">Restablecer Contraseña</h1><form action="/procesar-reset" method="POST" class="space-y-4"><input type="email" name="email" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="Correo"><input type="password" name="nueva_password" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="Nueva Clave"><button type="submit" class="w-full bg-teal-700 text-white py-2.5 rounded-lg font-bold text-sm">Actualizar</button></form><div class="text-center mt-6"><a href="/login-form" class="text-xs text-teal-400 font-semibold">← Volver</a></div></div></body></html>""", error=None, mensaje=None)

@app.route('/procesar-reset', methods=['POST'])
def procesar_reset():
    email = request.form.get('email', '').strip()
    nueva = request.form.get('nueva_password', '').strip()
    usuario = next((u for u in usuarios_db if u['email'] == email), None)
    if not usuario: return "Correo no encontrado. <a href='/reset-password-form'>Volver</a>"
    usuario['password'] = nueva
    return "¡Contraseña actualizada con éxito! <a href='/login-form'>Ir al Login</a>"

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_form'))


# ==========================================
# MÓDULO HILOS (/hilos)
# ==========================================
HTML_HILOS = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>DENIER - Hilos</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen p-6 flex flex-col justify-between">
    <div class="max-w-7xl mx-auto w-full">
        <div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl">
            <div class="flex items-center gap-4">
                """ + LOGO_DENIER_SVG + """
                <div>
                    <h1 class="text-2xl font-black tracking-wider uppercase text-white">DENIER - HILOS</h1>
                    <p class="text-sm text-slate-400">Control de Conos e Hilos</p>
                </div>
            </div>
        </div>
        <div class="flex flex-wrap gap-3 mb-6">
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white">← Menú DENIER</a>
            <a href="/hilos" class="px-4 py-3 rounded-xl font-semibold bg-teal-700 text-white">Inventario Activo</a>
            <a href="/hilos/historial" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Historial Packing List</a>
            <a href="/hilos/devoluciones-modulo" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Módulo Devoluciones</a>
        </div>
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <h2 class="text-xl font-bold mb-4 text-teal-400">Registrar Nuevo Hilo</h2>
            <form action="/hilos/guardar" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div><label class="block text-sm mb-1 text-slate-300">Tipo</label><select name="tipo" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"><option value="Denier">Denier</option><option value="Artitex">Artitex</option></select></div>
                <div><label class="block text-sm mb-1 text-slate-300">OC / Ingreso</label><input type="text" name="oc" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Código</label><input type="text" name="codigo" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Calibre</label><input type="text" name="calibre" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Composición</label><input type="text" name="composicion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Proveedor</label><input type="text" name="proveedor" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Lote</label><input type="text" name="lote" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Kilos</label><input type="number" step="0.01" name="kilos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Conos</label><input type="number" name="conos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Ubicación</label><input type="text" name="ubicacion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: A-1"></div>
                <div class="md:col-span-3 flex justify-end"><button type="submit" class="bg-teal-700 text-white px-6 py-2.5 rounded-xl font-medium">Guardar Hilo</button></div>
            </form>
        </div>
        <form action="/hilos/generar-packing-form" method="POST">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-6 flex justify-between items-center">
                <span class="text-sm text-slate-300">Total Kilos: <strong class="text-teal-400">{{ "%.2f"|format(total_kilos_filtrados) }} kg</strong></span>
                <button type="submit" class="bg-purple-700 text-white px-6 py-2.5 rounded-xl font-medium">Crear Packing List</button>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
                <table class="w-full text-left">
                    <thead><tr class="border-b border-slate-800 text-sm text-slate-400"><th class="p-3">Sel</th><th class="p-3">Tipo</th><th class="p-3">Lote</th><th class="p-3">Composición</th><th class="p-3">Kilos</th><th class="p-3">Conos</th><th class="p-3">Ubicación</th></tr></thead>
                    <tbody class="divide-y divide-slate-800 text-sm">
                        {% for item in resultados %}
                        <tr>
                            <td class="p-3"><input type="checkbox" name="hilos_sel" value="{{ item.id }}" class="w-4 h-4 text-purple-600 rounded"></td>
                            <td class="p-3">{{ item.tipo }}</td><td class="p-3 font-semibold">{{ item.lote }}</td><td class="p-3">{{ item.composicion }}</td>
                            <td class="p-3 text-teal-300 font-bold">{{ item.kilos }} kg</td><td class="p-3 text-purple-400 font-bold">{{ item.conos }} conos</td><td class="p-3 text-emerald-400">{{ item.ubicacion }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </form>
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider mt-8">Creado: Estuardo Alonzo</div>
</body>
</html>
"""

PACKING_HILOS_PRINT = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Packing List Hilos</title>
<style>body{font-family:monospace;padding:30px;background:#fff;color:#000;}table{width:100%;border-collapse:collapse;margin-top:20px;}th,td{border:1px solid #000;padding:8px;text-align:left;font-size:13px;}th{background:#f2f2f2;}</style>
</head>
<body>
    <div style="display:flex;justify-content:space-between;border-bottom:2px solid #000;padding-bottom:10px;">
        <div>""" + LOGO_PRINT_SVG + """<h3>DENIER S.A.</h3><p>PACKING LIST HILOS</p></div>
        <div style="text-align:right;"><h2>NO: {{ numero }}</h2><p>Cliente: {{ cliente }} | Orden: {{ orden }}</p></div>
    </div>
    <table>
        <thead><tr><th>Lote</th><th>Composición</th><th>Calibre</th><th>Ubicación</th><th>Conos</th><th>Kilos</th></tr></thead>
        <tbody>
            {% for item in seleccionados %}
            <tr><td><strong>{{ item.lote }}</strong></td><td>{{ item.composicion }}</td><td>{{ item.calibre }}</td><td>{{ item.ubicacion }}</td><td>{{ item.req_conos }}</td><td>{{ "%.2f"|format(item.req_kilos) }}</td></tr>
            {% endfor %}
        </tbody>
    </table>
    <h3 style="text-align:right;margin-top:20px;">TOTAL: {{ total_conos }} Conos / {{ "%.2f"|format(total_kilos) }} Kg</h3>
    <div style="text-align:center;margin-top:40px;"><button onclick="window.print()" style="padding:10px 20px;">🖨 Imprimir</button><br><br><a href="/hilos">Volver</a></div>
</body>
</html>
"""

@app.route('/hilos')
def hilos_index():
    if 'email' not in session or ('hilos' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])): return redirect(url_for('index'))
    activos = [i for i in inventario_hilos if i['kilos'] > 0 or i['conos'] > 0]
    return render_template_string(HTML_HILOS, resultados=activos, total_kilos_filtrados=sum(i['kilos'] for i in activos), total_conos_filtrados=sum(i['conos'] for i in activos))

@app.route('/hilos/guardar', methods=['POST'])
def hilos_guardar():
    inventario_hilos.append({
        'id': len(inventario_hilos) + 1, 'tipo': request.form.get('tipo'), 'oc': request.form.get('oc','N/A'),
        'codigo': request.form.get('codigo'), 'calibre': request.form.get('calibre'), 'composicion': request.form.get('composicion'),
        'proveedor': request.form.get('proveedor'), 'lote': request.form.get('lote'), 'kilos': float(request.form.get('kilos')),
        'conos': int(request.form.get('conos')), 'ubicacion': request.form.get('ubicacion')
    })
    return redirect(url_for('hilos_index'))

@app.route('/hilos/generar-packing-form', methods=['POST'])
def hilos_generar_packing_form():
    ids = request.form.getlist('hilos_sel')
    if not ids: return redirect(url_for('hilos_index'))
    seleccionados = [i for i in inventario_hilos if i['id'] in [int(x) for x in ids]]
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-6"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold mb-6 text-purple-400">Configurar Despacho Hilos</h1><form action="/hilos/procesar-packing-multiple" method="POST" class="space-y-6"><div class="grid grid-cols-2 gap-4 pb-4 border-b border-slate-800"><input type="text" name="cliente" class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Cliente..." required><input type="text" name="orden" class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Orden..." required></div>{% for item in seleccionados %}<input type="hidden" name="hilos_sel" value="{{ item.id }}"><div class="p-4 bg-black rounded-xl border border-slate-800"><p class="font-bold text-teal-400">Lote: {{ item.lote }} | Disp: {{ item.kilos }} kg</p><div class="grid grid-cols-2 gap-4 mt-2"><input type="number" step="0.01" name="req_kilos_{{ item.id }}" value="{{ item.kilos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white"><input type="number" name="req_conos_{{ item.id }}" value="{{ item.conos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white"></div></div>{% endfor %}<button type="submit" class="bg-purple-700 text-white px-6 py-3 rounded-xl font-medium w-full">Generar Packing List</button></form></div></body></html>""", seleccionados=seleccionados)

@app.route('/hilos/procesar-packing-multiple', methods=['POST'])
def hilos_procesar_packing_multiple():
    global contador_packing_hilos
    ids = request.form.getlist('hilos_sel')
    sel, tot_k, tot_c = [], 0, 0
    for sid in ids:
        item = next((i for i in inventario_hilos if i['id'] == int(sid)), None)
        if item:
            rk = float(request.form.get(f'req_kilos_{item["id"]}', 0))
            rc = int(request.form.get(f'req_conos_{item["id"]}', 0))
            item['kilos'] -= rk
            item['conos'] -= rc
            copia = item.copy()
            copia['req_kilos'], copia['req_conos'] = rk, rc
            sel.append(copia)
            tot_k += rk
            tot_c += rc
    reg = {'id_historial': len(packing_history_hilos)+1, 'numero': f"{contador_packing_hilos:03d}", 'cliente': request.form.get('cliente'), 'orden': request.form.get('orden'), 'total_kilos': tot_k, 'total_conos': tot_c, 'seleccionados': sel}
    packing_history_hilos.append(reg)
    contador_packing_hilos += 1
    return render_template_string(PACKING_HILOS_PRINT, numero=reg['numero'], cliente=reg['cliente'], orden=reg['orden'], seleccionados=sel, total_kilos=tot_k, total_conos=tot_c)

@app.route('/hilos/historial')
def hilos_historial():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-6xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold text-emerald-400 mb-6">Historial Packing Lists Hilos</h1><a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a><table class="w-full text-left text-sm"><thead><tr class="border-b border-slate-800"><th class="p-2">Número</th><th class="p-2">Cliente</th><th class="p-2">Totales</th><th class="p-2">Acción</th></tr></thead><tbody>{% for p in historial %}<tr class="border-b border-slate-800"><td class="p-2 font-bold">N {{ p.numero }}</td><td class="p-2">{{ p.cliente }}</td><td class="p-2">{{ p.total_kilos }} kg / {{ p.total_conos }} conos</td><td class="p-2"><a href="/hilos/ver-packing/{{ p.id_historial }}" target="_blank" class="text-teal-400 underline">Ver</a></td></tr>{% endfor %}</tbody></table></div></body></html>""", historial=packing_history_hilos)

@app.route('/hilos/ver-packing/<int:hid>')
def hilos_ver_packing(hid):
    r = next((p for p in packing_history_hilos if p['id_historial'] == hid), None)
    return render_template_string(PACKING_HILOS_PRINT, numero=r['numero'], cliente=r['cliente'], orden=r['orden'], seleccionados=r['seleccionados'], total_kilos=r['total_kilos'], total_conos=r['total_conos'])

@app.route('/hilos/devoluciones-modulo')
def hilos_devoluciones_modulo():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-5xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold text-rose-400 mb-6">Módulo Devoluciones - Hilos</h1><a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a>{% if error %}<div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded mb-4 text-xs font-bold">{{ error }}</div>{% endif %}<form action="/hilos/procesar-devolucion-parcial" method="POST" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800"><select name="historial_id" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white"><option value="">-- Elige Packing List --</option>{% for p in historial %}<option value="{{ p.id_historial }}">N {{ p.numero }} | Cliente: {{ p.cliente }} (Max: {{ p.total_kilos }} kg)</option>{% endfor %}</select><div class="grid grid-cols-2 gap-4"><input type="number" step="0.01" name="ret_kilos" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Kilos a retornar"><input type="number" name="ret_conos" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Conos a retornar"></div><input type="text" name="motivo" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Motivo..."><button type="submit" class="bg-rose-700 text-white px-6 py-2 rounded-xl font-medium">Procesar Reingreso</button></form></div></body></html>""", historial=packing_history_hilos, error=request.args.get('error'))

@app.route('/hilos/procesar-devolucion-parcial', methods=['POST'])
def hilos_procesar_devolucion_parcial():
    hid = int(request.form.get('historial_id'))
    ret_k = float(request.form.get('ret_kilos'))
    ret_c = int(request.form.get('ret_conos'))
    p = next((x for x in packing_history_hilos if x['id_historial'] == hid), None)
    if p:
        if ret_k > p['total_kilos'] or ret_c > p['total_conos']:
            return redirect(url_for('hilos_devoluciones_modulo', error="⚠ No puedes devolver más de lo despachado."))
        p['total_kilos'] -= ret_k
        p['total_conos'] -= ret_c
    return redirect(url_for('hilos_devoluciones_modulo'))


# ==========================================
# MÓDULO TELA CRUDA (/cruda)
# ==========================================
HTML_TELA = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>DENIER - Tela Cruda</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen p-6 flex flex-col justify-between">
    <div class="max-w-7xl mx-auto w-full">
        <div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl">
            <div class="flex items-center gap-4">
                """ + LOGO_DENIER_SVG + """
                <div>
                    <h1 class="text-2xl font-black tracking-wider uppercase text-white">DENIER - TELA CRUDA</h1>
                    <p class="text-sm text-slate-400">Control de Tejeduría y Rollos Crudos</p>
                </div>
            </div>
        </div>
        <div class="flex flex-wrap gap-3 mb-6">
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white">← Menú DENIER</a>
            <a href="/cruda" class="px-4 py-3 rounded-xl font-semibold bg-teal-700 text-white">Inventario Cruda</a>
            <a href="/cruda/carga-masiva-form" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Carga Masiva (Knitopia)</a>
            <a href="/cruda/historial-tela" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Historial Despachos</a>
        </div>
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <h2 class="text-xl font-bold mb-4 text-teal-400">Registrar Rollo Individual</h2>
            <form action="/cruda/guardar-tela" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div><label class="block text-sm mb-1 text-slate-300">OC</label><input type="text" name="oc" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Proveedor</label><input type="text" name="proveedor" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Lote Tela</label><input type="text" name="lote_tela" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">No. Rollo</label><input type="text" name="rollo" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Kilos</label><input type="number" step="0.01" name="kilos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Ubicación</label><input type="text" name="ubicacion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: A-1"></div>
                <div class="md:col-span-3 flex justify-end"><button type="submit" class="bg-teal-700 text-white px-6 py-2.5 rounded-xl font-medium">Guardar Rollo</button></div>
            </form>
        </div>
        <form action="/cruda/generar-packing-tela-form" method="POST">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-6 flex justify-between items-center">
                <span class="text-sm text-slate-400">Total Kilos: <strong class="text-teal-400">{{ "%.2f"|format(total_kilos_filtrados) }} kg</strong></span>
                <button type="submit" class="bg-purple-700 text-white px-6 py-2.5 rounded-xl font-medium">Crear Packing List Tela</button>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
                <table class="w-full text-left">
                    <thead><tr class="border-b border-slate-800 text-sm text-slate-400"><th class="p-3">Sel</th><th class="p-3">OC</th><th class="p-3">Proveedor</th><th class="p-3">Lote</th><th class="p-3">Rollo</th><th class="p-3">Kilos</th><th class="p-3">Ubicación</th></tr></thead>
                    <tbody class="divide-y divide-slate-800 text-sm">
                        {% for item in resultados %}
                        <tr>
                            <td class="p-3"><input type="checkbox" name="telas_sel" value="{{ item.id }}" class="w-4 h-4 text-purple-600 rounded"></td>
                            <td class="p-3 font-bold text-teal-400">{{ item.oc }}</td><td class="p-3">{{ item.proveedor }}</td><td class="p-3">{{ item.lote_tela }}</td><td class="p-3">{{ item.rollo }}</td>
                            <td class="p-3 font-bold text-teal-300">{{ item.kilos }} kg</td><td class="p-3 text-emerald-400">{{ item.ubicacion }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </form>
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider mt-8">Creado: Estuardo Alonzo</div>
</body>
</html>
"""

PACKING_TELA_PRINT = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Packing List Tela</title>
<style>body{font-family:monospace;padding:30px;background:#fff;color:#000;}table{width:100%;border-collapse:collapse;margin-top:20px;}th,td{border:1px solid #000;padding:8px;text-align:left;font-size:13px;}th{background:#f2f2f2;}</style>
</head>
<body>
    <div style="display:flex;justify-content:space-between;border-bottom:2px solid #000;padding-bottom:10px;">
        <div>""" + LOGO_PRINT_SVG + """<h3>DENIER S.A.</h3><p>DESPACHO TELA CRUDA</p></div>
        <div style="text-align:right;"><h2>NO: {{ numero }}</h2><p>Destino: {{ cliente }} | Orden: {{ orden }}</p></div>
    </div>
    <table>
        <thead><tr><th>OC</th><th>Lote Tela</th><th>Rollo</th><th>Ubicación</th><th>Kilos Despachados</th></tr></thead>
        <tbody>
            {% for item in seleccionados %}
            <tr><td>{{ item.oc }}</td><td><strong>{{ item.lote_tela }}</strong></td><td>{{ item.rollo }}</td><td>{{ item.ubicacion }}</td><td>{{ "%.2f"|format(item.req_kilos) }} kg</td></tr>
            {% endfor %}
        </tbody>
    </table>
    <h3 style="text-align:right;margin-top:20px;">TOTAL KILOS: {{ "%.2f"|format(total_kilos) }} KG</h3>
    <div style="text-align:center;margin-top:40px;"><button onclick="window.print()" style="padding:10px 20px;">🖨 Imprimir</button><br><br><a href="/cruda">Volver</a></div>
</body>
</html>
"""

@app.route('/cruda')
def cruda_index():
    if 'email' not in session or ('cruda' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])): return redirect(url_for('index'))
    activos = [i for i in inventario_tela if i['kilos'] > 0]
    return render_template_string(HTML_TELA, resultados=activos, total_kilos_filtrados=sum(i['kilos'] for i in activos))

@app.route('/cruda/guardar-tela', methods=['POST'])
def cruda_guardar_tela():
    inventario_tela.append({
        'id': len(inventario_tela) + 1, 'oc': request.form.get('oc'), 'proveedor': request.form.get('proveedor'),
        'lote_tela': request.form.get('lote_tela'), 'rollo': request.form.get('rollo'),
        'kilos': float(request.form.get('kilos')), 'ubicacion': request.form.get('ubicacion')
    })
    return redirect(url_for('cruda_index'))

@app.route('/cruda/carga-masiva-form')
def cruda_carga_masiva_form():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold text-amber-400 mb-4">Carga Masiva Knitopia</h1><a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a><form action="/cruda/procesar-carga-masiva" method="POST" enctype="multipart/form-data" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800"><input type="file" name="excel_file" accept=".xlsx" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded text-white"><button type="submit" class="bg-amber-700 text-white px-6 py-2.5 rounded-xl font-medium w-full">Cargar Excel</button></form></div></body></html>""")

@app.route('/cruda/procesar-carga-masiva', methods=['POST'])
def cruda_procesar_carga_masiva():
    archivo = request.files.get('excel_file')
    if archivo:
        wb = openpyxl.load_workbook(archivo)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[4] is not None:
                inventario_tela.append({'id': len(inventario_tela)+1, 'oc': str(row[0]), 'proveedor': str(row[1]), 'lote_tela': str(row[2]), 'rollo': str(row[3]), 'kilos': float(row[4]), 'ubicacion': str(row[5]) if len(row)>5 and row[5] else 'A-1'})
    return redirect(url_for('cruda_index'))

@app.route('/cruda/generar-packing-tela-form', methods=['POST'])
def cruda_generar_packing_tela_form():
    ids = request.form.getlist('telas_sel')
    if not ids: return redirect(url_for('cruda_index'))
    seleccionados = [i for i in inventario_tela if i['id'] in [int(x) for x in ids]]
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-6"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold mb-6 text-purple-400">Configurar Despacho Tela</h1><form action="/cruda/procesar-packing-tela-multiple" method="POST" class="space-y-6"><input type="text" name="cliente" required class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Cliente destino..."><input type="text" name="orden" required class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Orden..."><input type="hidden" name="telas_sel" value="{{ ids_str }}">{% for item in seleccionados %}<div class="p-4 bg-black rounded-xl border border-slate-800"><p class="font-bold text-teal-400">Rollo: {{ item.rollo }} | Disp: {{ item.kilos }} kg</p><input type="hidden" name="telas_sel_list" value="{{ item.id }}"><input type="number" step="0.01" name="req_kilos_{{ item.id }}" value="{{ item.kilos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white mt-2"></div>{% endfor %}<button type="submit" class="bg-purple-700 text-white px-6 py-3 rounded-xl font-medium w-full">Generar Packing List</button></form></div></body></html>""", seleccionados=seleccionados, ids_str=",".join(ids))

@app.route('/cruda/procesar-packing-tela-multiple', methods=['POST'])
def cruda_procesar_packing_tela_multiple():
    global contador_packing_tela
    ids = request.form.getlist('telas_sel_list')
    sel, tot = [], 0
    for sid in ids:
        item = next((i for i in inventario_tela if i['id'] == int(sid)), None)
        if item:
            req = float(request.form.get(f'req_kilos_{item["id"]}', 0))
            item['kilos'] -= req
            copia = item.copy()
            copia['req_kilos'] = req
            sel.append(copia)
            tot += req
    reg = {'id_historial': len(packing_tela_history)+1, 'numero': f"{contador_packing_tela:03d}", 'cliente': request.form.get('cliente'), 'orden': request.form.get('orden'), 'total_kilos': tot, 'seleccionados': sel}
    packing_tela_history.append(reg)
    contador_packing_tela += 1
    return render_template_string(PACKING_TELA_PRINT, numero=reg['numero'], cliente=reg['cliente'], orden=reg['orden'], seleccionados=sel, total_kilos=tot)

@app.route('/cruda/historial-tela')
def cruda_historial_tela():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-6xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold text-emerald-400 mb-6">Historial Despachos Tela</h1><a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a><table class="w-full text-left text-sm"><thead><tr class="border-b border-slate-800"><th class="p-2">Número</th><th class="p-2">Destino</th><th class="p-2">Kilos</th><th class="p-2">Acción</th></tr></thead><tbody>{% for p in historial %}<tr class="border-b border-slate-800"><td class="p-2 font-bold">N {{ p.numero }}</td><td class="p-2">{{ p.cliente }}</td><td class="p-2">{{ "%.2f"|format(p.total_kilos) }} kg</td><td class="p-2"><a href="/cruda/ver-packing-tela/{{ p.id_historial }}" target="_blank" class="text-teal-400 underline">Ver</a></td></tr>{% endfor %}</tbody></table></div></body></html>""", historial=packing_tela_history)

@app.route('/cruda/ver-packing-tela/<int:hid>')
def cruda_ver_packing_tela(hid):
    r = next((p for p in packing_tela_history if p['id_historial'] == hid), None)
    return render_template_string(PACKING_TELA_PRINT, numero=r['numero'], cliente=r['cliente'], orden=r['orden'], seleccionados=r['seleccionados'], total_kilos=r['total_kilos'])


# ==========================================
# MÓDULO DEVOLUCIONES Y SALIDA (/devo)
# ==========================================
HTML_DEVO = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>DENIER - Devoluciones</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen p-6 flex flex-col justify-between">
    <div class="max-w-7xl mx-auto w-full">
        <div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl">
            <div class="flex items-center gap-4">
                """ + LOGO_DENIER_SVG + """
                <div>
                    <h1 class="text-2xl font-black tracking-wider uppercase text-white">DENIER - TELA DEVUELTA</h1>
                    <p class="text-sm text-slate-400">Devoluciones y Salidas por Venta / Reproceso</p>
                </div>
            </div>
        </div>
        <div class="flex flex-wrap gap-3 mb-6">
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white">← Menú DENIER</a>
            <a href="/devo" class="px-4 py-3 rounded-xl font-semibold bg-rose-700 text-white">Devoluciones Activas</a>
        </div>
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <h2 class="text-xl font-bold mb-4 text-rose-400">Registrar Devolución de Cliente</h2>
            <form action="/devo/guardar-devolucion" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div><label class="block text-sm mb-1 text-slate-300">Cliente</label><input type="text" name="cliente" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Lote</label><input type="text" name="lote" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Tela</label><input type="text" name="tela" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Color</label><input type="text" name="color" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Rollos</label><input type="number" name="rollos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Kilos</label><input type="number" step="0.01" name="kilos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Yardas</label><input type="number" step="0.01" name="yardas" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Calidad</label><select name="calidad" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"><option value="Primera">Primera</option><option value="Segunda">Segunda</option><option value="Rechazada">Rechazada</option></select></div>
                <div><label class="block text-sm mb-1 text-slate-300">Ubicación</label><input type="text" name="ubicacion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Rack 3"></div>
                <div class="md:col-span-3 flex justify-end"><button type="submit" class="bg-rose-700 text-white px-6 py-2.5 rounded-xl font-medium">Guardar Devolución</button></div>
            </form>
        </div>
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
            <table class="w-full text-left">
                <thead><tr class="border-b border-slate-800 text-sm text-slate-400"><th class="p-3">Cliente</th><th class="p-3">Lote</th><th class="p-3">Tela / Color</th><th class="p-3">Rollos</th><th class="p-3">Kilos / Yardas</th><th class="p-3">Calidad</th><th class="p-3 text-center">Despacho</th></tr></thead>
                <tbody class="divide-y divide-slate-800 text-sm">
                    {% for item in resultados %}
                    <tr>
                        <td class="p-3 font-bold text-rose-400">{{ item.cliente }}</td><td class="p-3">{{ item.lote }}</td><td class="p-3">{{ item.tela }} ({{ item.color }})</td>
                        <td class="p-3">{{ item.rollos }} rollos</td><td class="p-3 text-rose-300 font-bold">{{ "%.2f"|format(item.kilos) }} kg</td>
                        <td><span class="px-2 py-0.5 rounded text-xs font-bold bg-rose-950 text-rose-400 border border-rose-800">{{ item.calidad }}</span></td>
                        <td class="p-3 text-center">
                            {% if item.kilos > 0 %}
                            <a href="/devo/despachar-form/{{ item.id }}" class="bg-purple-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold">📦 Despachar / Reproceso</a>
                            {% else %}<span class="text-xs text-slate-500">Agotado</span>{% endif %}
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider mt-8">Creado: Estuardo Alonzo</div>
</body>
</html>
"""

DESPACHO_DEVO_PRINT = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Despacho Devolución</title>
<style>body{font-family:monospace;padding:30px;background:#fff;color:#000;}table{width:100%;border-collapse:collapse;margin-top:20px;}th,td{border:1px solid #000;padding:8px;text-align:left;font-size:13px;}th{background:#f2f2f2;}</style>
</head>
<body>
    <div style="display:flex;justify-content:space-between;border-bottom:2px solid #000;padding-bottom:10px;">
        <div>""" + LOGO_PRINT_SVG + """<h3>DENIER S.A.</h3><p>SALIDA TELA DEVUELTA</p></div>
        <div style="text-align:right;"><h2>NO: {{ numero }}</h2><p>Destino: {{ destino }} | Motivo: {{ motivo }}</p></div>
    </div>
    <table>
        <thead><tr><th>Lote</th><th>Tela / Color</th><th>Calidad</th><th>Rollos</th><th>Kilos</th><th>Yardas</th></tr></thead>
        <tbody>
            <tr><td><strong>{{ item.lote }}</strong></td><td>{{ item.tela }} ({{ item.color }})</td><td>{{ item.calidad }}</td><td>{{ req_rollos }}</td><td>{{ "%.2f"|format(req_kilos) }}</td><td>{{ "%.2f"|format(req_yardas) }}</td></tr>
        </tbody>
    </table>
    <div style="text-align:center;margin-top:40px;"><button onclick="window.print()" style="padding:10px 20px;">🖨 Imprimir</button><br><br><a href="/devo">Volver</a></div>
</body>
</html>
"""

@app.route('/devo')
def devo_index():
    if 'email' not in session or ('devo' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])): return redirect(url_for('index'))
    return render_template_string(HTML_DEVOLUCION, resultados=devoluciones_procesada, mensaje=request.args.get('mensaje',''), error=request.args.get('error',''))

@app.route('/devo/guardar-devolucion', methods=['POST'])
def devo_guardar_devolucion():
    devoluciones_procesada.append({
        'id': len(devoluciones_procesada) + 1, 'cliente': request.form.get('cliente'), 'lote': request.form.get('lote'),
        'tela': request.form.get('tela'), 'color': request.form.get('color'), 'rollos': int(request.form.get('rollos')),
        'kilos': float(request.form.get('kilos')), 'yardas': float(request.form.get('yardas')),
        'calidad': request.form.get('calidad'), 'ubicacion': request.form.get('ubicacion'), 'archivo': None
    })
    return redirect(url_for('devo_index'))

@app.route('/devo/despachar-form/<int:dev_id>')
def devo_despachar_form(dev_id):
    item = next((i for i in devoluciones_procesada if i['id'] == dev_id), None)
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-6"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold mb-2 text-rose-400">Despachar Tela Devuelta</h1><form action="/devo/procesar-despacho/{{ item.id }}" method="POST" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800"><input type="text" name="destino" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded text-white" placeholder="Destino / Venta / Reproceso"><select name="motivo" class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded text-white"><option value="Venta de Saldo">Venta de Saldo</option><option value="Reproceso / Teñido">Reproceso / Teñido</option></select><div class="grid grid-cols-3 gap-4"><input type="number" name="req_rollos" value="{{ item.rollos }}" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded text-white"><input type="number" step="0.01" name="req_kilos" value="{{ item.kilos }}" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded text-white"><input type="number" step="0.01" name="req_yardas" value="{{ item.yardas }}" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded text-white"></div><button type="submit" class="bg-rose-700 text-white px-6 py-3 rounded-xl font-bold w-full">Generar Comprobante</button></form></div></body></html>""", item=item)

@app.route('/devo/procesar-despacho/<int:dev_id>', methods=['POST'])
def devo_procesar_despacho(dev_id):
    global contador_despacho_devo
    item = next((i for i in devoluciones_procesada if i['id'] == dev_id), None)
    rk, ry, rr = float(request.form.get('req_kilos', 0)), float(request.form.get('req_yardas', 0)), int(request.form.get('req_rollos', 0))
    item['kilos'] -= rk
    item['yardas'] -= ry
    item['rollos'] -= rr
    reg = {'numero': f"DES-DEV-{contador_despacho_devo:03d}", 'destino': request.form.get('destino'), 'motivo': request.form.get('motivo'), 'item': item, 'req_kilos': rk, 'req_yardas': ry, 'req_rollos': rr}
    contador_despacho_devo += 1
    return render_template_string(DESPACHO_DEVO_PRINT, **reg)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
