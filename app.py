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

# Carpeta para archivos subidos (Devoluciones)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ==========================================
# LOGO SVG OFICIAL DENIER
# ==========================================
LOGO_DENIER_SVG = """
<svg width="70" height="70" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M50 15 L85 50 L50 85 L15 50 Z" stroke="#38bdf8" stroke-width="4" fill="none"/>
    <path d="M50 25 L75 50 L50 75 L25 50 Z" stroke="#38bdf8" stroke-width="3" fill="none"/>
    <path d="M30 30 L70 70 M70 30 L30 70" stroke="#38bdf8" stroke-width="5" stroke-linecap="round"/>
    <path d="M50 15 L50 85 M15 50 L85 50" stroke="#38bdf8" stroke-width="5" stroke-linecap="round"/>
</svg>
"""

LOGO_PRINT_SVG = """
<svg width="60" height="60" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M50 15 L85 50 L50 85 L15 50 Z" stroke="#000" stroke-width="4" fill="none"/>
    <path d="M50 25 L75 50 L50 75 L25 50 Z" stroke="#000" stroke-width="3" fill="none"/>
    <path d="M30 30 L70 70 M70 30 L30 70" stroke="#000" stroke-width="5" stroke-linecap="round"/>
    <path d="M50 15 L50 85 M15 50 L85 50" stroke="#000" stroke-width="5" stroke-linecap="round"/>
</svg>
"""

# ==========================================
# MEMORIA DE DATOS
# ==========================================
usuarios_db = [
    { 'id': 1, 'email': 'controlprod_1@deniersa.com', 'password': '123', 'nombre': 'Administrador General', 'modulos': ['admin', 'hilos', 'cruda', 'devo'] },
    { 'id': 2, 'email': 'bodega.hilos@deniersa.com', 'password': '123', 'nombre': 'Operador Hilos', 'modulos': ['hilos'] },
    { 'id': 3, 'email': 'bodega.cruda@deniersa.com', 'password': '123', 'nombre': 'Operador Tela Cruda', 'modulos': ['cruda'] },
    { 'id': 4, 'email': 'bodega.devoluciones@deniersa.com', 'password': '123', 'nombre': 'Operador Devoluciones', 'modulos': ['devo'] }
]

inventario_hilos = [
    { 'id': 1, 'tipo': 'Denier', 'oc': 'P04150', 'codigo': 'H-COT-20', 'calibre': '20/1', 'composicion': 'Algodon 100%', 'proveedor': 'Hilos del Sur', 'lote': 'L-442', 'kilos': 450.5, 'conos': 120, 'ubicacion': 'A-1' },
    { 'id': 2, 'tipo': 'Denier', 'oc': '4555', 'codigo': 'H-COT-20', 'calibre': '20/1', 'composicion': 'Algodon 100%', 'proveedor': 'Hilos del Sur', 'lote': 'L-442', 'kilos': 35.0, 'conos': 10, 'ubicacion': 'B-2' }
]
packing_history_hilos = []
contador_packing_hilos = 1

inventario_tela = [
    { 'id': 1, 'oc': 'OC-9012', 'proveedor': 'Knitopia', 'lote_tela': 'TC-101', 'rollo': 'R-001', 'kilos': 250.5, 'ubicacion': 'Bodega A-1' },
    { 'id': 2, 'oc': 'OC-9013', 'proveedor': 'Magic Dyeing', 'lote_tela': 'TC-102', 'rollo': 'R-002', 'kilos': 45.0, 'ubicacion': 'Bodega B-2' }
]
packing_tela_history = []
contador_packing_tela = 1

devoluciones_procesada = [
    { 
        'id': 1, 
        'cliente': 'Magic Apparel', 
        'lote': '20452026-8', 
        'tela': 'Interlock, 100% Algodón', 
        'color': 'White, Midnight Navy', 
        'rollos': 3, 
        'kilos': 75.50, 
        'yardas': 240.00, 
        'calidad': 'Rechazada', 
        'ubicacion': 'Rack 3', 
        'archivo': None 
    }
]
despachos_devoluciones_history = []
contador_despacho_devo = 1

# ==========================================
# PLANTILLAS BASE Y LOGIN
# ==========================================
LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>DENIER - INVENTARIO</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen flex flex-col justify-between p-6">
    <div class="flex-grow flex items-center justify-center">
        <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-sm">
            <div class="text-center mb-6">
                <div class="flex justify-center mb-2">
                    """ + LOGO_DENIER_SVG + """
                </div>
                <h1 class="text-2xl font-black tracking-widest text-white uppercase">DENIER</h1>
                <div class="border-b border-slate-800 my-4"></div>
            </div>
            
            {% if error %}
            <div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded-lg mb-4 text-xs font-bold text-center">{{ error }}</div>
            {% endif %}

            <form action="/login" method="POST" class="space-y-4">
                <div>
                    <label class="block text-xs font-bold mb-1.5 text-slate-300">Email</label>
                    <input type="email" name="email" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="controlprod_1@deniersa.com">
                </div>
                <div>
                    <label class="block text-xs font-bold mb-1.5 text-slate-300">Password</label>
                    <input type="password" name="password" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="••••••••">
                </div>
                <button type="submit" class="w-full bg-teal-700 hover:bg-teal-600 text-white py-2.5 rounded-lg font-bold text-sm shadow">Log in</button>
            </form>

            <div class="flex justify-between items-center mt-6 text-xs text-slate-400">
                <span class="text-slate-500">¿Problemas?</span>
                <a href="/reset-password-form" class="text-teal-400 font-semibold hover:underline">Reset Password</a>
            </div>
        </div>
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider">
        Creado: Estuardo Alonzo
    </div>
</body>
</html>
"""

HOME_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>DENIER - Panel Principal</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen p-8 flex flex-col justify-between">
    <div class="max-w-6xl mx-auto w-full">
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8 flex justify-between items-center">
            <div class="flex items-center gap-4">
                """ + LOGO_DENIER_SVG + """
                <div>
                    <h1 class="text-2xl font-black tracking-wide text-white uppercase">DENIER</h1>
                    <p class="text-sm text-slate-400">Usuario: <b class="text-teal-400">{{ session.get('nombre') }}</b> ({{ session.get('email') }})</p>
                </div>
            </div>
            <a href="/logout" class="bg-red-950 text-red-300 border border-red-800 px-4 py-2 rounded-xl text-sm font-semibold">Cerrar Sesión</a>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
            {% if 'admin' in session.get('modulos') or 'hilos' in session.get('modulos') %}
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div>
                    <h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Hilos</h2>
                    <p class="text-sm text-slate-400 mb-4">Control de inventario Denier y Artitex, packing lists y devoluciones.</p>
                </div>
                <a href="/hilos" class="bg-teal-700 hover:bg-teal-600 text-white text-center py-3 rounded-xl font-medium shadow">Entrar a Hilos</a>
            </div>
            {% endif %}

            {% if 'admin' in session.get('modulos') or 'cruda' in session.get('modulos') %}
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div>
                    <h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Tela Cruda</h2>
                    <p class="text-sm text-slate-400 mb-4">Control de rollos, cargas masivas Knitopia y reingresos Magic Dyeing.</p>
                </div>
                <a href="/cruda" class="bg-teal-700 hover:bg-teal-600 text-white text-center py-3 rounded-xl font-medium shadow">Entrar a Tela Cruda</a>
            </div>
            {% endif %}

            {% if 'admin' in session.get('modulos') or 'devo' in session.get('modulos') %}
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div>
                    <h2 class="text-xl font-bold text-rose-400 mb-2">Módulo Devoluciones</h2>
                    <p class="text-sm text-slate-400 mb-4">Control de tela devuelta, gestión de PDF y salidas por venta o reproceso.</p>
                </div>
                <a href="/devo" class="bg-rose-700 hover:bg-rose-600 text-white text-center py-3 rounded-xl font-medium shadow">Entrar a Devoluciones</a>
            </div>
            {% endif %}
        </div>

        {% if 'admin' in session.get('modulos') %}
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl">
            <h2 class="text-xl font-bold text-purple-400 mb-4">Gestión de Usuarios y Accesos - DENIER</h2>
            <form action="/crear-usuario" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8 bg-black p-5 rounded-xl border border-slate-800">
                <div>
                    <label class="block text-xs mb-1 font-semibold text-slate-300">Correo Corporativo</label>
                    <input type="email" name="email" required class="w-full p-2.5 rounded-lg border border-slate-700 bg-slate-900 text-sm text-white" placeholder="usuario@deniersa.com">
                </div>
                <div>
                    <label class="block text-xs mb-1 font-semibold text-slate-300">Nombre Completo</label>
                    <input type="text" name="nombre" required class="w-full p-2.5 rounded-lg border border-slate-700 bg-slate-900 text-sm text-white" placeholder="Nombre...">
                </div>
                <div>
                    <label class="block text-xs mb-1 font-semibold text-slate-300">Contraseña Temporal</label>
                    <input type="password" name="password" required class="w-full p-2.5 rounded-lg border border-slate-700 bg-slate-900 text-sm text-white" placeholder="Clave...">
                </div>
                <div class="md:col-span-3">
                    <label class="block text-xs mb-2 font-semibold text-slate-300">Seleccionar Módulos Permitidos:</label>
                    <div class="flex gap-6">
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300"><input type="checkbox" name="modulos" value="hilos" class="w-4 h-4 text-teal-600 rounded"><span>Bodega Hilos</span></label>
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300"><input type="checkbox" name="modulos" value="cruda" class="w-4 h-4 text-teal-600 rounded"><span>Bodega Tela Cruda</span></label>
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300"><input type="checkbox" name="modulos" value="devo" class="w-4 h-4 text-rose-600 rounded"><span>Devoluciones</span></label>
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300"><input type="checkbox" name="modulos" value="admin" class="w-4 h-4 text-purple-600 rounded"><span>Administrador Total</span></label>
                    </div>
                </div>
                <div class="md:col-span-3 flex justify-end">
                    <button type="submit" class="bg-purple-700 hover:bg-purple-600 text-white px-6 py-2.5 rounded-xl text-sm font-medium">Registrar Nuevo Usuario</button>
                </div>
            </form>

            <h3 class="font-bold text-sm text-slate-300 mb-3">Lista de Usuarios Activos y Permisos</h3>
            <table class="w-full text-left text-sm">
                <thead>
                    <tr class="border-b border-slate-800 text-slate-400">
                        <th class="p-3">Correo</th><th class="p-3">Nombre</th><th class="p-3">Contraseña</th><th class="p-3">Módulos</th><th class="p-3 text-center">Acciones</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-800">
                    {% for u in usuarios_db %}
                    <tr>
                        <td class="p-3 font-medium text-teal-400">{{ u.email }}</td>
                        <td class="p-3 text-slate-200">{{ u.nombre }}</td>
                        <td class="p-3 text-amber-300 font-mono font-bold">{{ u.password }}</td>
                        <td class="p-3">{% for m in u.modulos %}<span class="bg-slate-800 text-slate-300 border border-slate-700 px-2 py-0.5 rounded text-xs font-bold uppercase mr-1">{{ m }}</span>{% endfor %}</td>
                        <td class="p-3 text-center">
                            <div class="flex justify-center gap-2">
                                <a href="/editar-usuario/{{ u.id }}" class="bg-amber-950 text-amber-300 border border-amber-800 px-3 py-1 rounded-lg text-xs font-semibold">Editar</a>
                                {% if u.email != 'controlprod_1@deniersa.com' %}
                                <a href="/eliminar-usuario/{{ u.id }}" onclick="return confirm('¿Seguro deseas eliminar este usuario?');" class="bg-red-950 text-red-300 border border-red-800 px-3 py-1 rounded-lg text-xs font-semibold">Eliminar</a>
                                {% endif %}
                            </div>
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        {% endif %}
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider mt-8">
        Creado: Estuardo Alonzo
    </div>
</body>
</html>
"""

RESET_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Restablecer Contraseña - DENIER</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen flex items-center justify-center p-6">
    <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-sm">
        <h1 class="text-xl font-bold text-white mb-2">Restablecer Contraseña</h1>
        {% if error %}<div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded mb-4 text-xs font-bold text-center">{{ error }}</div>{% endif %}
        {% if mensaje %}<div class="bg-emerald-950 border border-emerald-800 text-emerald-300 p-3 rounded mb-4 text-xs font-bold text-center">{{ mensaje }}</div>{% endif %}
        <form action="/procesar-reset" method="POST" class="space-y-4">
            <input type="email" name="email" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="Correo corporativo">
            <input type="password" name="nueva_password" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white" placeholder="Nueva contraseña">
            <button type="submit" class="w-full bg-teal-700 text-white py-2.5 rounded-lg font-bold text-sm">Actualizar</button>
        </form>
        <div class="text-center mt-6"><a href="/login-form" class="text-xs text-teal-400 font-semibold hover:underline">← Volver al Login</a></div>
    </div>
</body>
</html>
"""

EDITAR_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Editar Permisos</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 min-h-screen flex items-center justify-center p-6">
    <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-md">
        <h1 class="text-xl font-bold text-purple-400 mb-2">Editar Permisos de Usuario</h1>
        <form action="/actualizar-usuario/{{ usuario.id }}" method="POST" class="space-y-4">
            <input type="text" name="nombre" value="{{ usuario.nombre }}" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white">
            <input type="text" name="password" value="{{ usuario.password }}" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white">
            <button type="submit" class="bg-purple-700 text-white px-5 py-2 rounded-xl text-xs font-bold w-full">Guardar Cambios</button>
        </form>
    </div>
</body>
</html>
"""

# ==========================================
# RUTAS AUTENTICACIÓN
# ==========================================
@app.route('/')
def index():
    if 'email' not in session: return redirect(url_for('login_form'))
    return render_template_string(HOME_TEMPLATE, usuarios_db=usuarios_db)

@app.route('/login-form')
def login_form(): return render_template_string(LOGIN_TEMPLATE, error=None)

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

@app.route('/reset-password-form')
def reset_password_form(): return render_template_string(RESET_TEMPLATE, error=None, mensaje=None)

@app.route('/procesar-reset', methods=['POST'])
def procesar_reset():
    email = request.form.get('email', '').strip()
    nueva = request.form.get('nueva_password', '').strip()
    usuario = next((u for u in usuarios_db if u['email'] == email), None)
    if not usuario: return render_template_string(RESET_TEMPLATE, error="El correo no está registrado.", mensaje=None)
    usuario['password'] = nueva
    return render_template_string(RESET_TEMPLATE, error=None, mensaje="¡Contraseña actualizada con éxito!")

@app.route('/crear-usuario', methods=['POST'])
def crear_usuario():
    if 'admin' not in session.get('modulos', []): return redirect(url_for('index'))
    nuevo_email = request.form.get('email', '').strip()
    if any(u['email'] == nuevo_email for u in usuarios_db): return redirect(url_for('index'))
    nuevo = {
        'id': max([u['id'] for u in usuarios_db], default=0) + 1,
        'email': nuevo_email,
        'password': request.form.get('password', '').strip(),
        'nombre': request.form.get('nombre', '').strip(),
        'modulos': request.form.getlist('modulos')
    }
    usuarios_db.append(nuevo)
    return redirect(url_for('index'))

@app.route('/editar-usuario/<int:uid>')
def editar_usuario(uid):
    if 'admin' not in session.get('modulos', []): return redirect(url_for('index'))
    usuario = next((u for u in usuarios_db if u['id'] == uid), None)
    return render_template_string(EDITAR_TEMPLATE, usuario=usuario)

@app.route('/actualizar-usuario/<int:uid>', methods=['POST'])
def actualizar_usuario(uid):
    if 'admin' not in session.get('modulos', []): return redirect(url_for('index'))
    usuario = next((u for u in usuarios_db if u['id'] == uid), None)
    if usuario:
        usuario['nombre'] = request.form.get('nombre', '').strip()
        usuario['password'] = request.form.get('password', '').strip()
        usuario['modulos'] = request.form.getlist('modulos')
    return redirect(url_for('index'))

@app.route('/eliminar-usuario/<int:uid>')
def eliminar_usuario(uid):
    if 'admin' not in session.get('modulos', []): return redirect(url_for('index'))
    global usuarios_db
    usuarios_db = [u for u in usuarios_db if u['id'] != uid or u['email'] == 'controlprod_1@deniersa.com']
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_form'))


# ==========================================
# MÓDULO HILOS (/hilos)
# ==========================================
HTML_TEMPLATE_HILOS = """
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
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white hover:bg-slate-700">← Menú DENIER</a>
            <a href="/hilos" class="px-4 py-3 rounded-xl font-semibold bg-teal-700 text-white">Inventario Activo</a>
            <a href="/hilos/agotados" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Hilos Sin Stock</a>
            <a href="/hilos/historial" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Historial Packing List</a>
            <a href="/hilos/devoluciones-modulo" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Módulo Devoluciones</a>
            <a href="/hilos/exportar-excel" class="px-4 py-3 rounded-xl font-semibold bg-emerald-800 text-white ml-auto">Descargar Excel</a>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <h2 class="text-xl font-bold mb-4 text-teal-400">Registrar Nuevo Hilo</h2>
            <form action="/hilos/guardar" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Tipo de Hilo</label>
                    <select name="tipo" id="tipoSelect" onchange="toggleOC()" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                        <option value="Denier">Denier (Propios)</option>
                        <option value="Artitex">Artitex (Cliente)</option>
                    </select>
                </div>
                <div id="ocDiv"><label class="block text-sm mb-1 text-slate-300">Orden / Ingreso</label><input type="text" name="oc" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Código</label><input type="text" name="codigo" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Calibre</label><input type="text" name="calibre" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Composición</label><input type="text" name="composicion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Proveedor</label><input type="text" name="proveedor" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Lote</label><input type="text" name="lote" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Kilos Iniciales</label><input type="number" step="0.01" name="kilos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Conos Iniciales</label><input type="number" name="conos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"></div>
                <div><label class="block text-sm mb-1 text-slate-300">Ubicación</label><input type="text" name="ubicacion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: A-1"></div>
                <div class="md:col-span-3 flex justify-end"><button type="submit" class="bg-teal-700 hover:bg-teal-600 text-white px-6 py-2.5 rounded-xl font-medium">Guardar en Inventario</button></div>
            </form>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <form method="GET" action="/hilos" class="grid grid-cols-1 md:grid-cols-4 gap-4">
                <select name="filtro_tipo" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"><option value="">-- Todos Tipos --</option><option value="Denier" {{ 'selected' if filtro_tipo == 'Denier' else '' }}>Denier</option><option value="Artitex" {{ 'selected' if filtro_tipo == 'Artitex' else '' }}>Artitex</option></select>
                <select name="filtro_prov" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"><option value="">-- Todos Proveedores --</option>{% for prov in proveedores_unicos %}<option value="{{ prov }}" {{ 'selected' if filtro_prov == prov else '' }}>{{ prov }}</option>{% endfor %}</select>
                <input type="text" name="q" value="{{ query }}" placeholder="Buscar Lote, Ubicación (A-1)..." class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                <button type="submit" class="w-full bg-slate-800 text-white p-2.5 rounded-lg font-medium border border-slate-700">Filtrar y Sumar</button>
            </form>
        </div>

        <form action="/hilos/generar-packing-form" method="POST">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-6 flex justify-between items-center">
                <span class="text-sm text-slate-300">Suma Filtrada: <strong class="text-teal-400 text-base">{{ "%.2f"|format(total_kilos_filtrados) }} kg</strong> / <strong class="text-teal-400 text-base">{{ total_conos_filtrados }} conos</strong></span>
                <button type="submit" class="bg-purple-700 hover:bg-purple-600 text-white px-6 py-2.5 rounded-xl font-medium shadow">Crear Packing List</button>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
                <table class="w-full text-left">
                    <thead><tr class="border-b border-slate-800 text-sm text-slate-400"><th class="p-3">Seleccionar</th><th class="p-3">Tipo</th><th class="p-3">Orden</th><th class="p-3">Lote</th><th class="p-3">Composición</th><th class="p-3">Calibre</th><th class="p-3">Kilos Disponibles</th><th class="p-3">Conos</th><th class="p-3">Ubicación</th></tr></thead>
                    <tbody class="divide-y divide-slate-800 text-sm">
                        {% for item in resultados %}
                        <tr class="{{ 'bg-amber-950/40' if item.kilos < 50 else '' }}">
                            <td class="p-3 text-center"><input type="checkbox" name="hilos_sel" value="{{ item.id }}" class="w-4 h-4 text-purple-600 rounded"></td>
                            <td class="p-3 font-medium text-slate-200">{{ item.tipo }}</td><td class="p-3 font-bold text-teal-400">{{ item.oc }}</td><td class="p-3 font-semibold text-slate-200">{{ item.lote }}</td>
                            <td class="p-3 text-slate-300">{{ item.composicion }}</td><td class="p-3 text-slate-300">{{ item.calibre }}</td>
                            <td class="p-3 font-bold text-teal-300">{{ item.kilos }} kg</td><td class="p-3 text-purple-400 font-bold">{{ item.conos }} conos</td><td class="p-3 font-semibold text-emerald-400">{{ item.ubicacion }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </form>
    </div>
    <div class="text-center text-xs text-slate-500 font-semibold tracking-wider mt-8">Creado: Estuardo Alonzo</div>
    <script>function toggleOC() { const t = document.getElementById('tipoSelect').value; document.getElementById('ocDiv').style.display = t === 'Artitex' ? 'none' : 'block'; }</script>
</body>
</html>
"""

PACKING_VIEW_HILOS_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>PACKING LIST - DENIER</title>
    <style>
        body { font-family: 'Courier New', Courier, monospace; background: #fff; color: #000; padding: 30px; }
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #000; padding: 8px; text-align: left; font-size: 13px; }
        th { background-color: #f2f2f2; }
        .total-box { margin-top: 20px; font-size: 15px; font-weight: bold; text-align: right; }
        .footer { margin-top: 50px; display: flex; justify-content: space-between; font-size: 12px; }
        .signature { border-top: 1px solid #000; width: 200px; text-align: center; padding-top: 5px; }
    </style>
</head>
<body>
    <div class="header">
        <div>
            """ + LOGO_PRINT_SVG + """
            <h1 style="margin: 5px 0 0 0; font-size: 22px;">DENIER S.A.</h1>
            <p style="margin: 0; font-size: 12px;">PACKING LIST - DEPARTAMENTO DE HILOS</p>
        </div>
        <div style="text-align: right;">
            <h2 style="margin: 0; font-size: 18px;">NO: {{ numero }}</h2>
            <p style="margin: 3px 0;"><strong>Cliente:</strong> {{ cliente }}</p>
            <p style="margin: 0;"><strong>Orden:</strong> {{ orden }}</p>
        </div>
    </div>

    <table>
        <thead>
            <tr>
                <th>Lote</th>
                <th>Código / Composición</th>
                <th>Calibre</th>
                <th>Ubicación</th>
                <th>Conos Despachados</th>
                <th>Kilos Despachados</th>
            </tr>
        </thead>
        <tbody>
            {% for item in seleccionados %}
            <tr>
                <td><strong>{{ item.lote }}</strong></td>
                <td>{{ item.codigo }} - {{ item.composicion }}</td>
                <td>{{ item.calibre }}</td>
                <td>{{ item.ubicacion }}</td>
                <td>{{ item.req_conos }} conos</td>
                <td>{{ "%.2f"|format(item.req_kilos) }} kg</td>
            </tr>
            {% endfor %}
        </tbody>
    </table>

    <div class="total-box">
        TOTAL DESPACHO: {{ total_conos }} CONOS / {{ "%.2f"|format(total_kilos) }} KG
    </div>

    <div class="footer">
        <div class="signature">Despachado Por</div>
        <div class="signature">Recibido Conforme</div>
    </div>
    <div style="text-align: center; margin-top: 30px;">
        <button onclick="window.print()" style="padding: 8px 16px; cursor: pointer;">🖨 Imprimir Packing List</button>
        <br><br><a href="/hilos">← Volver a Hilos</a>
    </div>
</body>
</html>
"""

@app.route('/hilos')
def hilos_index():
    if 'email' not in session or ('hilos' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])): return redirect(url_for('index'))
    query = request.args.get('q', '').lower()
    filtro_tipo = request.args.get('filtro_tipo', '')
    filtro_prov = request.args.get('filtro_prov', '')
    proveedores_unicos = sorted(list(set(i['proveedor'] for i in inventario_hilos)))
    activos = [i for i in inventario_hilos if i['kilos'] > 0 or i['conos'] > 0]
    resultados = activos
    if filtro_tipo: resultados = [i for i in resultados if i['tipo'] == filtro_tipo]
    if filtro_prov: resultados = [i for i in resultados if i['proveedor'] == filtro_prov]
    if query: resultados = [i for i in resultados if query in i['lote'].lower() or query in i['composicion'].lower() or query in i['oc'].lower() or query in i['ubicacion'].lower()]
    return render_template_string(HTML_TEMPLATE_HILOS, resultados=resultados, proveedores_unicos=proveedores_unicos, filtro_tipo=filtro_tipo, filtro_prov=filtro_prov, total_kilos_filtrados=sum(i['kilos'] for i in resultados), total_conos_filtrados=sum(i['conos'] for i in resultados), query=query)

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

@app.route('/hilos/exportar-excel')
def hilos_exportar_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Tipo", "OC", "Código", "Calibre", "Composición", "Proveedor", "Lote", "Kilos", "Conos", "Ubicación"])
    for i in inventario_hilos: ws.append([i['tipo'], i['oc'], i['codigo'], i['calibre'], i['composicion'], i['proveedor'], i['lote'], i['kilos'], i['conos'], i['ubicacion']])
    io_b = io.BytesIO()
    wb.save(io_b)
    io_b.seek(0)
    return send_file(io_b, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Hilos_Denier.xlsx')

@app.route('/hilos/agotados')
def hilos_agotados():
    agotados = [i for i in inventario_hilos if i['kilos'] <= 0]
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-6xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold text-amber-400 mb-6">Hilos sin Stock</h1><a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium">Volver</a><table class="w-full mt-4 text-left text-sm"><thead><tr class="border-b border-slate-800"><th class="p-2">Lote</th><th class="p-2">Composición</th><th class="p-2">Proveedor</th></tr></thead><tbody>{% for item in ag %}<tr class="border-b border-slate-800"><td class="p-2">{{ item.lote }}</td><td class="p-2">{{ item.composicion }}</td><td class="p-2">{{ item.proveedor }}</td></tr>{% endfor %}</tbody></table></div></body></html>""", ag=agotados)

@app.route('/hilos/generar-packing-form', methods=['POST'])
def hilos_generar_packing_form():
    ids = request.form.getlist('hilos_sel')
    if not ids: return redirect(url_for('hilos_index'))
    seleccionados = [i for i in inventario_hilos if i['id'] in [int(x) for x in ids]]
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-6"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold mb-6 text-purple-400">Configurar Despacho</h1><form action="/hilos/procesar-packing-multiple" method="POST" class="space-y-6"><div class="grid grid-cols-2 gap-4 pb-4 border-b border-slate-800"><input type="text" name="cliente" class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Cliente..." required><input type="text" name="orden" class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Orden..." required></div>{% for item in seleccionados %}<input type="hidden" name="hilos_sel" value="{{ item.id }}"><div class="p-4 bg-black rounded-xl border border-slate-800"><p class="font-bold text-teal-400">Lote: {{ item.lote }} | Disponibles: {{ item.kilos }} kg / {{ item.conos }} conos</p><div class="grid grid-cols-2 gap-4 mt-2"><input type="number" step="0.01" name="req_kilos_{{ item.id }}" value="{{ item.kilos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white"><input type="number" name="req_conos_{{ item.id }}" value="{{ item.conos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white"></div></div>{% endfor %}<button type="submit" class="bg-purple-700 text-white px-6 py-3 rounded-xl font-medium w-full">Generar Packing List</button></form></div></body></html>""", seleccionados=seleccionados)

@app.route('/hilos/procesar-packing-multiple', methods=['POST'])
def hilos_procesar_packing_multiple():
    global contador_packing_hilos
    ids = request.form.getlist('hilos_sel')
    sel = []
    tot_k = 0
    tot_c = 0
    for sid in ids:
        item = next((i for i in inventario_hilos if i['id'] == int(sid)), None)
        if item:
            rk = float(request.form.get(f'req_kilos_{item["id"]}', 0))
            rc = int(request.form.get(f'req_conos_{item["id"]}', 0))
            item['kilos'] -= rk
            item['conos'] -= rc
            copia = item.copy()
            copia['req_kilos'] = rk
            copia['req_conos'] = rc
            sel.append(copia)
            tot_k += rk
            tot_c += rc
    reg = {'id_historial': len(packing_history_hilos)+1, 'numero': f"{contador_packing_hilos:03d}", 'cliente': request.form.get('cliente') or 'General', 'orden': request.form.get('orden') or 'N/A', 'total_kilos': tot_k, 'total_conos': tot_c, 'seleccionados': sel}
    packing_history_hilos.append(reg)
    contador_packing_hilos += 1
    return render_template_string(PACKING_VIEW_HILOS_TEMPLATE, numero=reg['numero'], cliente=reg['cliente'], orden=reg['orden'], seleccionados=sel, total_kilos=tot_k, total_conos=tot_c)

@app.route('/hilos/historial')
def hilos_historial():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-6xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold text-emerald-400 mb-6">Historial Packing Lists Hilos</h1><a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a><table class="w-full text-left text-sm"><thead><tr class="border-b border-slate-800"><th class="p-2">Número</th><th class="p-2">Cliente</th><th class="p-2">Totales</th><th class="p-2">Acciones</th></tr></thead><tbody>{% for p in historial %}<tr class="border-b border-slate-800"><td class="p-2 font-bold">N {{ p.numero }}</td><td class="p-2">{{ p.cliente }}</td><td class="p-2">{{ p.total_kilos }} kg / {{ p.total_conos }} conos</td><td class="p-2"><a href="/hilos/ver-packing/{{ p.id_historial }}" target="_blank" class="text-teal-400 underline">Ver / Imprimir</a></td></tr>{% endfor %}</tbody></table></div></body></html>""", historial=packing_history_hilos)

@app.route('/hilos/ver-packing/<int:hid>')
def hilos_ver_packing(hid):
    r = next((p for p in packing_history_hilos if p['id_historial'] == hid), None)
    return render_template_string(PACKING_VIEW_HILOS_TEMPLATE, numero=r['numero'], cliente=r['cliente'], orden=r['orden'], seleccionados=r['seleccionados'], total_kilos=r['total_kilos'], total_conos=r['total_conos'])

@app.route('/hilos/devoluciones-modulo')
def hilos_devoluciones_modulo():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-5xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold text-rose-400 mb-6">Módulo de Devolución Parcial - Hilos</h1><a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a>{% if error %}<div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded mb-4 text-xs font-bold">{{ error }}</div>{% endif %}<form action="/hilos/procesar-devolucion-parcial" method="POST" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800"><select name="historial_id" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white"><option value="">-- Elige Packing List --</option>{% for p in historial %}<option value="{{ p.id_historial }}">N {{ p.numero }} | Cliente: {{ p.cliente }} (Max: {{ p.total_kilos }} kg / {{ p.total_conos }} conos)</option>{% endfor %}</select><div class="grid grid-cols-2 gap-4"><input type="number" step="0.01" name="ret_kilos" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Kilos a retornar"><input type="number" name="ret_conos" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Conos a retornar"></div><input type="text" name="motivo" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Motivo de devolución..."><button type="submit" class="bg-rose-700 text-white px-6 py-2 rounded-xl font-medium">Procesar Reingreso</button></form></div></body></html>""", historial=packing_history_hilos, error=request.args.get('error'))

@app.route('/hilos/procesar-devolucion-parcial', methods=['POST'])
def hilos_procesar_devolucion_parcial():
    hid = int(request.form.get('historial_id'))
    ret_k = float(request.form.get('ret_kilos'))
    ret_c = int(request.form.get('ret_conos'))
    p = next((x for x in packing_history_hilos if x['id_historial'] == hid), None)
    if p:
        if ret_k > p['total_kilos'] or ret_c > p['total_conos']:
            return redirect(url_for('hilos_devoluciones_modulo', error="⚠ ERROR: No puedes reingresar más kilos o conos de los despachados en este Packing List."))
        p['total_kilos'] -= ret_k
        p['total_conos'] -= ret_c
    return redirect(url_for('hilos_devoluciones_modulo'))


# ==========================================
# MÓDULO TELA CRUDA (/cruda)
# ==========================================
HTML_TEMPLATE_TELA = """
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
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white hover:bg-slate-700">← Menú DENIER</a>
            <a href="/cruda" class="px-4 py-3 rounded-xl font-semibold bg-teal-700 text-white">Inventario Cruda</a>
            <a href="/cruda/dashboard" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Dashboard KPIs</a>
            <a href="/cruda/carga-masiva-form" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Carga Masiva (Knitopia)</a>
            <a href="/cruda/historial-tela" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Historial Despachos</a>
            <a href="/cruda/reingresos-modulo" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700">Módulo Reingresos</a>
            <a href="/cruda/exportar-excel-tela" class="px-4 py-3 rounded-xl font-semibold bg-emerald-800 text-white ml-auto">Descargar Excel</a>
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
                <div class="md:col-span-3 flex justify-end"><button type="submit" class="bg-teal-700 hover:bg-teal-600 text-white px-6 py-2.5 rounded-xl font-medium">Guardar en Inventario</button></div>
            </form>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <form method="GET" action="/cruda" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <select name="filtro_prov" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white"><option value="">-- Proveedor --</option>{% for prov in proveedores_unicos %}<option value="{{ prov }}" {{ 'selected' if filtro_prov == prov else '' }}>{{ prov }}</option>{% endfor %}</select>
                <input type="text" name="q" value="{{ query }}" placeholder="Buscar Lote, Rollo, Ubicación (A-1)..." class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                <button type="submit" class="bg-slate-800 border border-slate-700 text-white p-2.5 rounded-lg font-medium">Filtrar y Buscar</button>
            </form>
        </div>

        <form action="/cruda/generar-packing-tela-form" method="POST">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-6 flex justify-between items-center">
                <span class="text-sm text-slate-400">Total Kilos Filtrados: <strong class="text-teal-400 text-base">{{ "%.2f"|format(total_kilos_filtrados) }} kg</strong></span>
                <button type="submit" class="bg-purple-700 hover:bg-purple-600 text-white px-6 py-2.5 rounded-xl font-medium shadow">Crear Packing List Tela</button>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
                <table class="w-full text-left">
                    <thead><tr class="border-b border-slate-800 text-sm text-slate-400"><th class="p-3">Sel</th><th class="p-3">OC</th><th class="p-3">Proveedor</th><th class="p-3">Lote</th><th class="p-3">Rollo</th><th class="p-3">Kilos</th><th class="p-3">Ubicación</th></tr></thead>
                    <tbody class="divide-y divide-slate-800 text-sm">
                        {% for item in resultados %}
                        <tr class="{{ 'bg-amber-950/40' if item.kilos < 50 else '' }}">
                            <td class="p-3 text-center"><input type="checkbox" name="telas_sel" value="{{ item.id }}" class="w-4 h-4 text-purple-600 rounded"></td>
                            <td class="p-3 font-bold text-teal-400">{{ item.oc }}</td><td class="p-3">{{ item.proveedor }}</td><td class="p-3">{{ item.lote_tela }}</td><td class="p-3">{{ item.rollo }}</td>
                            <td class="p-3 font-bold text-teal-300">{{ item.kilos }} kg</td><td class="p-3 font-semibold text-emerald-400">{{ item.ubicacion }}</td>
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

PACKING_VIEW_TELA_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>PACKING LIST TELA - DENIER</title>
    <style>
        body { font-family: 'Courier New', Courier, monospace; background: #fff; color: #000; padding: 30px; }
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #000; padding: 8px; text-align: left; font-size: 13px; }
        th { background-color: #f2f2f2; }
        .total-box { margin-top: 20px; font-size: 15px; font-weight: bold; text-align: right; }
        .footer { margin-top: 50px; display: flex; justify-content: space-between; font-size: 12px; }
        .signature { border-top: 1px solid #000; width: 200px; text-align: center; padding-top: 5px; }
    </style>
</head>
<body>
    <div class="header">
        <div>
            """ + LOGO_PRINT_SVG + """
            <h1 style="margin: 5px 0 0 0; font-size: 22px;">DENIER S.A.</h1>
            <p style="margin: 0; font-size: 12px;">DESPACHO TELA CRUDA / TEJEDURÍA</p>
        </div>
        <div style="text-align: right;">
            <h2 style="margin: 0; font-size: 18px;">NO: {{ numero }}</h2>
            <p style="margin: 3px 0;"><strong>Destino:</strong> {{ cliente }}</p>
            <p style="margin: 0;"><strong>Orden:</strong> {{ orden }}</p>
        </div>
    </div>

    <table>
        <thead>
            <tr>
                <th>Orden Compra</th>
                <th>Lote Tela</th>
                <th>Rollo</th>
                <th>Ubicación</th>
                <th>Kilos Despachados</th>
            </tr>
        </thead>
        <tbody>
            {% for item in seleccionados %}
            <tr>
                <td>{{ item.oc }}</td>
                <td><strong>{{ item.lote_tela }}</strong></td>
                <td>{{ item.rollo }}</td>
                <td>{{ item.ubicacion }}</td>
                <td>{{ "%.2f"|format(item.req_kilos) }} kg</td>
            </tr>
            {% endfor %}
        </tbody>
    </table>

    <div class="total-box">
        TOTAL DESPACHO TELA: {{ "%.2f"|format(total_kilos) }} KG
    </div>

    <div class="footer">
        <div class="signature">Despachado Por</div>
        <div class="signature">Recibido Conforme</div>
    </div>
    <div style="text-align: center; margin-top: 30px;">
        <button onclick="window.print()" style="padding: 8px 16px; cursor: pointer;">🖨 Imprimir Packing List</button>
        <br><br><a href="/cruda">← Volver a Tela Cruda</a>
    </div>
</body>
</html>
"""

@app.route('/cruda')
def cruda_index():
    if 'email' not in session or ('cruda' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])): return redirect(url_for('index'))
    query = request.args.get('q', '').lower()
    filtro_prov = request.args.get('filtro_prov', '')
    proveedores_unicos = sorted(list(set(i['proveedor'] for i in inventario_tela)))
    activos = [i for i in inventario_tela if i['kilos'] > 0]
    resultados = activos
    if filtro_prov: resultados = [i for i in resultados if i['proveedor'] == filtro_prov]
    if query: resultados = [i for i in resultados if query in i['lote_tela'].lower() or query in i['oc'].lower() or query in i['rollo'].lower() or query in i['proveedor'].lower() or query in i['ubicacion'].lower()]
    return render_template_string(HTML_TEMPLATE_TELA, resultados=resultados, proveedores_unicos=proveedores_unicos, filtro_prov=filtro_prov, query=query, total_kilos_filtrados=sum(i['kilos'] for i in resultados))

@app.route('/cruda/guardar-tela', methods=['POST'])
def cruda_guardar_tela():
    inventario_tela.append({
        'id': len(inventario_tela) + 1,
        'oc': request.form.get('oc'),
        'proveedor': request.form.get('proveedor'),
        'lote_tela': request.form.get('lote_tela'),
        'rollo': request.form.get('rollo'),
        'kilos': float(request.form.get('kilos')),
        'ubicacion': request.form.get('ubicacion')
    })
    return redirect(url_for('cruda_index'))

@app.route('/cruda/exportar-excel-tela')
def cruda_exportar_excel_tela():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["OC", "Proveedor", "Lote", "Rollo", "Kilos", "Ubicación"])
    for i in inventario_tela:
        if i['kilos'] > 0: ws.append([i['oc'], i['proveedor'], i['lote_tela'], i['rollo'], i['kilos'], i['ubicacion']])
    io_b = io.BytesIO()
    wb.save(io_b)
    io_b.seek(0)
    return send_file(io_b, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Tela_Cruda_Denier.xlsx')

@app.route('/cruda/dashboard')
def cruda_dashboard():
    activos = [i for i in inventario_tela if i['kilos'] > 0]
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-7xl mx-auto"><div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow"><h1 class="text-2xl font-bold text-teal-400">KPIs - Tela Cruda</h1><a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium">Volver</a></div><div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8"><div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow"><p class="text-sm text-slate-400">Total Kilos</p><h3 class="text-3xl font-bold text-white mt-2">{{ tk }} kg</h3></div><div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow"><p class="text-sm text-slate-400">Rollos Activos</p><h3 class="text-3xl font-bold text-white mt-2">{{ tr }}</h3></div></div></div></body></html>""", tk=sum(i['kilos'] for i in activos), tr=len(activos))

@app.route('/cruda/carga-masiva-form')
def cruda_carga_masiva_form():
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-xl"><h1 class="text-2xl font-bold text-amber-400 mb-4">Carga Masiva de Inventario (Knitopia)</h1><div class="mb-6 flex gap-4"><a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium inline-block">Volver</a><a href="/cruda/descargar-plantilla" class="bg-emerald-800 text-white px-4 py-2 rounded-xl text-sm font-medium inline-block">Descargar Plantilla Excel</a></div>{% if msg %}<div class="bg-emerald-950 border border-emerald-800 text-emerald-300 p-3 rounded mb-4 text-sm font-bold">{{ msg }}</div>{% endif %}<form action="/cruda/procesar-carga-masiva" method="POST" enctype="multipart/form-data" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800"><input type="file" name="excel_file" accept=".xlsx" required class="w-full p-2.5 rounded-lg border border-slate-700 bg-slate-900 text-white"><button type="submit" class="bg-amber-700 hover:bg-amber-600 text-white px-6 py-2.5 rounded-xl font-medium w-full">Cargar al Sistema</button></form></div></body></html>""", msg=None)

@app.route('/cruda/descargar-plantilla')
def cruda_descargar_plantilla():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["OC", "Proveedor", "Lote Tela", "Rollo", "Kilos", "Ubicacion"])
    ws.append(["OC-1001", "Knitopia", "LOTE-500", "R-01", 125.50, "A-1"])
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return send_file(excel_io, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Plantilla_Knitopia.xlsx')

@app.route('/cruda/procesar-carga-masiva', methods=['POST'])
def cruda_procesar_carga_masiva():
    archivo = request.files.get('excel_file')
    if archivo:
        wb = openpyxl.load_workbook(archivo)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[4] is not None:
                inventario_tela.append({'id': len(inventario_tela)+1, 'oc': str(row[0]), 'proveedor': str(row[1]), 'lote_tela': str(row[2]), 'rollo': str(row[3]), 'kilos': float(row[4]), 'ubicacion': str(row[5]) if len(row)>5 and row[5] else 'A-1'})
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-8"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-xl"><h1 class="text-2xl font-bold text-amber-400 mb-4">Carga Masiva</h1><div class="mb-6"><a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium">Volver a Tela Cruda</a></div><div class="bg-emerald-950 border border-emerald-800 text-emerald-300 p-3 rounded mb-4 text-sm font-bold">¡Carga masiva procesada con éxito!</div></div></body></html>""")

@app.route('/cruda/generar-packing-tela-form', methods=['POST'])
def cruda_generar_packing_tela_form():
    ids = request.form.getlist('telas_sel')
    if not ids: return redirect(url_for('cruda_index'))
    seleccionados = [i for i in inventario_tela if i['id'] in [int(x) for x in ids]]
    return render_template_string("""<!DOCTYPE html><html lang="es"><head><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-black text-slate-100 p-6"><div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow"><h1 class="text-2xl font-bold mb-6 text-purple-400">Configurar Despacho Tela</h1><form action="/cruda/procesar-packing-tela-multiple" method="POST" class="space-y-6"><input type="text" name="cliente" required class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Cliente destino / Tintorería..."><input
