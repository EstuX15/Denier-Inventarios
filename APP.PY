# -*- coding: utf-8 -*-
import io
import os
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from flask import Flask, render_template_string, request, redirect, url_for, session, send_file, send_from_directory
from datetime import datetime
import openpyxl

app = Flask(__name__)
app.secret_key = 'clave_secreta_denier_sistema_total'

# Configuración de carpeta para archivos de Devoluciones
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- BASES DE DATOS EN MEMORIA ---
usuarios_db = [
    { 'id': 1, 'email': 'controlprod_1@deniersa.com', 'password': '123', 'nombre': 'Administrador General', 'modulos': ['admin', 'hilos', 'cruda', 'devo'] }
]

inventario_hilos = [
    { 'id': 1, 'tipo': 'Denier', 'oc': 'P04150', 'codigo': 'H-COT-20', 'calibre': '20/1', 'composicion': 'Algodon 100%', 'proveedor': 'Hilos del Sur', 'lote': 'L-442', 'kilos': 450.5, 'conos': 120, 'ubicacion': 'A-1' }
]
packing_hilos_history = []
contador_packing_hilos = 1

inventario_tela_cruda = [
    { 'id': 1, 'oc': 'OC-9012', 'proveedor': 'Knitopia', 'lote_tela': 'TC-101', 'rollo': 'R-001', 'kilos': 250.5, 'ubicacion': 'Bodega A-1' }
]
packing_tela_history = []
contador_packing_tela = 1

devoluciones_procesada = [
    { 'id': 1, 'cliente': 'Magic Apparel', 'lote': '20452026-8', 'tela': 'Interlock, 100% Algodón', 'color': 'White', 'rollos': 3, 'kilos': 75.50, 'yardas': 240.00, 'calidad': 'Rechazada', 'ubicacion': 'Rack 3', 'archivo': None }
]

# --- RUTAS DE AUTENTICACIÓN Y MENÚ PRINCIPAL ---

@app.route('/')
def index():
    if 'email' not in session:
        return redirect(url_for('login_form'))
    return render_template_string(HOME_TEMPLATE, usuarios_db=usuarios_db)

@app.route('/login-form')
def login_form():
    return render_template_string(LOGIN_TEMPLATE, error=None)

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
    return render_template_string(LOGIN_TEMPLATE, error="Correo o contraseña incorrectos")

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_form'))

# --- MÓDULO HILOS (Prefijo /hilos) ---

@app.route('/hilos')
def hilos_index():
    if 'email' not in session or ('admin' not in session.get('modulos', []) and 'hilos' not in session.get('modulos', [])):
        return redirect(url_for('index'))
    
    query = request.args.get('q', '').lower()
    filtro_tipo = request.args.get('filtro_tipo', '')
    filtro_prov = request.args.get('filtro_prov', '')
    proveedores_unicos = sorted(list(set(i['proveedor'] for i in inventario_hilos)))
    activos = [i for i in inventario_hilos if i['kilos'] > 0 or i['conos'] > 0]
    
    resultados = activos
    if filtro_tipo:
        resultados = [i for i in resultados if i['tipo'] == filtro_tipo]
    if filtro_prov:
        resultados = [i for i in resultados if i['proveedor'] == filtro_prov]
    if query:
        resultados = [i for i in resultados if query in i['lote'].lower() or query in i['composicion'].lower() or query in i['oc'].lower() or query in i['ubicacion'].lower()]
    
    total_kilos = sum(i['kilos'] for i in resultados)
    total_conos = sum(i['conos'] for i in resultados)
    return render_template_string(HILOS_TEMPLATE, resultados=resultados, proveedores_unicos=proveedores_unicos, filtro_tipo=filtro_tipo, filtro_prov=filtro_prov, total_kilos_filtrados=total_kilos, total_conos_filtrados=total_conos, query=query)

@app.route('/hilos/guardar', methods=['POST'])
def hilos_guardar():
    tipo = request.form.get('tipo')
    inventario_hilos.append({
        'id': len(inventario_hilos) + 1,
        'tipo': tipo,
        'oc': 'N/A' if tipo == 'Artitex' else request.form.get('oc'),
        'codigo': request.form.get('codigo'),
        'calibre': request.form.get('calibre'),
        'composicion': request.form.get('composicion'),
        'proveedor': request.form.get('proveedor'),
        'lote': request.form.get('lote'),
        'kilos': float(request.form.get('kilos')),
        'conos': int(request.form.get('conos')),
        'ubicacion': request.form.get('ubicacion')
    })
    return redirect(url_for('hilos_index'))

# --- MÓDULO TELA CRUDA (Prefijo /cruda) ---

@app.route('/cruda')
def cruda_index():
    if 'email' not in session or ('admin' not in session.get('modulos', []) and 'cruda' not in session.get('modulos', [])):
        return redirect(url_for('index'))
    
    query = request.args.get('q', '').lower()
    filtro_prov = request.args.get('filtro_prov', '')
    proveedores_unicos = sorted(list(set(i['proveedor'] for i in inventario_tela_cruda)))
    activos = [i for i in inventario_tela_cruda if i['kilos'] > 0]
    
    resultados = activos
    if filtro_prov:
        resultados = [i for i in resultados if i['proveedor'] == filtro_prov]
    if query:
        resultados = [i for i in resultados if query in i['lote_tela'].lower() or query in i['oc'].lower() or query in i['rollo'].lower() or query in i['proveedor'].lower() or query in i['ubicacion'].lower()]
    
    total_kilos = sum(i['kilos'] for i in resultados)
    return render_template_string(CRUDA_TEMPLATE, resultados=resultados, proveedores_unicos=proveedores_unicos, filtro_prov=filtro_prov, query=query, total_kilos_filtrados=total_kilos)

@app.route('/cruda/guardar', methods=['POST'])
def cruda_guardar():
    inventario_tela_cruda.append({
        'id': len(inventario_tela_cruda) + 1,
        'oc': request.form.get('oc'),
        'proveedor': request.form.get('proveedor'),
        'lote_tela': request.form.get('lote_tela'),
        'rollo': request.form.get('rollo'),
        'kilos': float(request.form.get('kilos')),
        'ubicacion': request.form.get('ubicacion')
    })
    return redirect(url_for('cruda_index'))

# --- MÓDULO DEVOLUCIONES (Prefijo /devoluciones) ---

@app.route('/devoluciones')
def devoluciones_index():
    if 'email' not in session or ('admin' not in session.get('modulos', []) and 'devo' not in session.get('modulos', [])):
        return redirect(url_for('index'))
    
    query = request.args.get('q', '').lower()
    resultados = devoluciones_procesada
    if query:
        resultados = [i for i in resultados if query in i['cliente'].lower() or query in i['lote'].lower() or query in i['tela'].lower()]
    
    total_kilos = sum(i['kilos'] for i in resultados)
    total_yardas = sum(i['yardas'] for i in resultados)
    return render_template_string(DEVO_TEMPLATE, resultados=resultados, query=query, total_kilos_filtrados=total_kilos, total_yardas_filtradas=total_yardas)

@app.route('/devoluciones/guardar', methods=['POST'])
def devoluciones_guardar():
    archivo = request.files.get('archivo_devolucion')
    nombre_archivo = None
    if archivo and archivo.filename:
        nombre_archivo = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{archivo.filename}"
        archivo.save(os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo))
    
    devoluciones_procesada.append({
        'id': len(devoluciones_procesada) + 1,
        'cliente': request.form.get('cliente').strip(),
        'lote': request.form.get('lote').strip(),
        'tela': request.form.get('tela').strip(),
        'color': request.form.get('color').strip(),
        'rollos': int(request.form.get('rollos')),
        'kilos': float(request.form.get('kilos')),
        'yardas': float(request.form.get('yardas')),
        'calidad': request.form.get('calidad'),
        'ubicacion': request.form.get('ubicacion').strip(),
        'archivo': nombre_archivo
    })
    return redirect(url_for('devoluciones_index'))

@app.route('/uploads/<filename>')
def ver_archivo(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

# --- PLANTILLAS HTML UNIFICADAS ---

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>DENIER - Login</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen flex items-center justify-center p-6">
    <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-sm">
        <h1 class="text-2xl font-black text-center text-white mb-6 uppercase">DENIER INVENTARIO</h1>
        {% if error %}<div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded mb-4 text-xs font-bold text-center">{{ error }}</div>{% endif %}
        <form action="/login" method="POST" class="space-y-4">
            <div><label class="block text-xs font-bold mb-1 text-slate-300">Email</label><input type="email" name="email" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white"></div>
            <div><label class="block text-xs font-bold mb-1 text-slate-300">Password</label><input type="password" name="password" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white"></div>
            <button type="submit" class="w-full bg-teal-700 hover:bg-teal-600 text-white py-2.5 rounded-lg font-bold text-sm">Log in</button>
        </form>
    </div>
</body>
</html>
"""

HOME_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>DENIER - Panel Principal</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen p-8">
    <div class="max-w-6xl mx-auto">
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8 flex justify-between items-center">
            <div>
                <h1 class="text-2xl font-black text-white">DENIER - INVENTARIO NUBE</h1>
                <p class="text-sm text-slate-400">Usuario: <b class="text-teal-400">{{ session.get('nombre') }}</b></p>
            </div>
            <a href="/logout" class="bg-red-950 text-red-300 border border-red-800 px-4 py-2 rounded-xl text-sm font-semibold">Cerrar Sesión</a>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow flex flex-col justify-between">
                <div><h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Hilos</h2><p class="text-sm text-slate-400 mb-4">Control de inventario de conos y hilos.</p></div>
                <a href="/hilos" class="bg-teal-700 text-white text-center py-3 rounded-xl font-medium">Entrar a Hilos</a>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow flex flex-col justify-between">
                <div><h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Tela Cruda</h2><p class="text-sm text-slate-400 mb-4">Control de rollos y tejeduría.</p></div>
                <a href="/cruda" class="bg-teal-700 text-white text-center py-3 rounded-xl font-medium">Entrar a Tela Cruda</a>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow flex flex-col justify-between">
                <div><h2 class="text-xl font-bold text-rose-400 mb-2">Módulo Devoluciones</h2><p class="text-sm text-slate-400 mb-4">Control de devoluciones y archivos PDF.</p></div>
                <a href="/devoluciones" class="bg-rose-700 text-white text-center py-3 rounded-xl font-medium">Entrar a Devoluciones</a>
            </div>
        </div>
    </div>
</body>
</html>
"""

HILOS_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Hilos</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-6"><div class="max-w-7xl mx-auto">
    <div class="mb-6"><a href="/" class="px-4 py-2 bg-slate-800 text-white rounded-xl text-sm font-bold">← Menú Principal</a></div>
    <div class="bg-slate-900 p-6 rounded-2xl border border-slate-800 mb-6">
        <h2 class="text-xl font-bold text-teal-400 mb-4">Registrar Nuevo Hilo</h2>
        <form action="/hilos/guardar" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4">
            <select name="tipo" class="p-2.5 bg-black border border-slate-700 rounded text-white"><option value="Denier">Denier</option><option value="Artitex">Artitex</option></select>
            <input type="text" name="oc" placeholder="Orden" class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="codigo" placeholder="Código" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="calibre" placeholder="Calibre" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="composicion" placeholder="Composición" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="proveedor" placeholder="Proveedor" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="lote" placeholder="Lote" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="number" step="0.01" name="kilos" placeholder="Kilos" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="number" name="conos" placeholder="Conos" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="ubicacion" placeholder="Ubicación" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <button type="submit" class="md:col-span-3 bg-teal-700 text-white p-2.5 rounded font-bold">Guardar Hilo</button>
        </form>
    </div>
    <div class="bg-slate-900 p-6 rounded-2xl border border-slate-800">
        <p class="mb-4 text-teal-400 font-bold">Suma Kilos: {{ "%.2f"|format(total_kilos_filtrados) }} kg | Conos: {{ total_conos_filtrados }}</p>
        <table class="w-full text-left text-sm">
            <tr class="border-b border-slate-800 text-slate-400"><th class="p-2">Lote</th><th class="p-2">Composición</th><th class="p-2">Kilos</th><th class="p-2">Ubicación</th></tr>
            {% for i in resultados %}<tr class="border-b border-slate-800"><td class="p-2">{{ i.lote }}</td><td class="p-2">{{ i.composicion }}</td><td class="p-2">{{ i.kilos }} kg</td><td class="p-2">{{ i.ubicacion }}</td></tr>{% endfor %}
        </table>
    </div>
</div></body></html>
"""

CRUDA_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Tela Cruda</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-6"><div class="max-w-7xl mx-auto">
    <div class="mb-6"><a href="/" class="px-4 py-2 bg-slate-800 text-white rounded-xl text-sm font-bold">← Menú Principal</a></div>
    <div class="bg-slate-900 p-6 rounded-2xl border border-slate-800 mb-6">
        <h2 class="text-xl font-bold text-teal-400 mb-4">Registrar Rollo Crudo</h2>
        <form action="/cruda/guardar" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4">
            <input type="text" name="oc" placeholder="OC" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="proveedor" placeholder="Proveedor" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="lote_tela" placeholder="Lote Tela" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="rollo" placeholder="No. Rollo" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="number" step="0.01" name="kilos" placeholder="Kilos" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="ubicacion" placeholder="Ubicación" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <button type="submit" class="md:col-span-3 bg-teal-700 text-white p-2.5 rounded font-bold">Guardar Rollo</button>
        </form>
    </div>
    <div class="bg-slate-900 p-6 rounded-2xl border border-slate-800">
        <p class="mb-4 text-teal-400 font-bold">Suma Kilos Cruda: {{ "%.2f"|format(total_kilos_filtrados) }} kg</p>
        <table class="w-full text-left text-sm">
            <tr class="border-b border-slate-800 text-slate-400"><th class="p-2">OC</th><th class="p-2">Proveedor</th><th class="p-2">Rollo</th><th class="p-2">Kilos</th><th class="p-2">Ubicación</th></tr>
            {% for i in resultados %}<tr class="border-b border-slate-800"><td class="p-2">{{ i.oc }}</td><td class="p-2">{{ i.proveedor }}</td><td class="p-2">{{ i.rollo }}</td><td class="p-2">{{ i.kilos }} kg</td><td class="p-2">{{ i.ubicacion }}</td></tr>{% endfor %}
        </table>
    </div>
</div></body></html>
"""

DEVO_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Devoluciones</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-6"><div class="max-w-7xl mx-auto">
    <div class="mb-6"><a href="/" class="px-4 py-2 bg-slate-800 text-white rounded-xl text-sm font-bold">← Menú Principal</a></div>
    <div class="bg-slate-900 p-6 rounded-2xl border border-slate-800 mb-6">
        <h2 class="text-xl font-bold text-rose-400 mb-4">Registrar Devolución</h2>
        <form action="/devoluciones/guardar" method="POST" enctype="multipart/form-data" class="grid grid-cols-1 md:grid-cols-3 gap-4">
            <input type="text" name="cliente" placeholder="Cliente" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="lote" placeholder="Lote" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="tela" placeholder="Tela" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="text" name="color" placeholder="Color" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="number" name="rollos" placeholder="Rollos" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="number" step="0.01" name="kilos" placeholder="Kilos" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="number" step="0.01" name="yardas" placeholder="Yardas" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <select name="calidad" class="p-2.5 bg-black border border-slate-700 rounded text-white"><option value="Primera">Primera</option><option value="Rechazada">Rechazada</option></select>
            <input type="text" name="ubicacion" placeholder="Ubicación" required class="p-2.5 bg-black border border-slate-700 rounded text-white">
            <input type="file" name="archivo_devolucion" accept=".pdf,.png,.jpg" class="md:col-span-3 p-2 bg-black border border-slate-700 rounded text-white text-sm">
            <button type="submit" class="md:col-span-3 bg-rose-700 text-white p-2.5 rounded font-bold">Guardar Devolución</button>
        </form>
    </div>
    <div class="bg-slate-900 p-6 rounded-2xl border border-slate-800">
        <table class="w-full text-left text-sm">
            <tr class="border-b border-slate-800 text-slate-400"><th class="p-2">Cliente</th><th class="p-2">Lote</th><th class="p-2">Kilos / Yardas</th><th class="p-2">Calidad</th><th class="p-2">Archivo</th></tr>
            {% for i in resultados %}
            <tr class="border-b border-slate-800">
                <td class="p-2">{{ i.cliente }}</td>
                <td class="p-2">{{ i.lote }}</td>
                <td class="p-2">{{ i.kilos }} kg / {{ i.yardas }} yds</td>
                <td class="p-2">{{ i.calidad }}</td>
                <td class="p-2">{% if i.archivo %}<a href="/uploads/{{ i.archivo }}" target="_blank" class="text-teal-400 underline">Ver PDF</a>{% else %}Sin archivo{% endif %}</td>
            </tr>
            {% endfor %}
        </table>
    </div>
</div></body></html>
"""

if __name__ == '__main__':
    app.run(debug=True, port=5000)