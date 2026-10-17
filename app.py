# -*- coding: utf-8 -*-
import io
import os
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from flask import Flask, render_template_string, request, redirect, url_for, session, send_file, send_from_directory
from datetime import datetime
import openpyxl

app = Flask(__name__)
app.secret_key = 'clave_secreta_denier_inventario_total_maestro_absoluto'

# Configuración de carpeta para archivos subidos (Devoluciones)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


# =========================================================================
# 1. COMPONENTES GLOBALES Y BASES DE DATOS (ORIGINALES)
# =========================================================================
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
    },
    { 
        'id': 2, 
        'cliente': 'MAGIC DYEING', 
        'lote': '45452026-7', 
        'tela': 'interlock', 
        'color': 'white', 
        'rollos': 18, 
        'kilos': 100.00, 
        'yardas': 500.00, 
        'calidad': 'Primera', 
        'ubicacion': 'A-1', 
        'archivo': None 
    }
]


# =========================================================================
# 2. CÓDIGO ÍNTEGRO DE APP_2.PY (LOGIN, PANEL Y GESTIÓN DE USUARIOS)
# =========================================================================
LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>DENIER - INVENTARIO</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen flex items-center justify-center p-6">
    <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-sm">
        <div class="text-center mb-6">
            <div class="flex justify-center mb-2">
                <svg width="80" height="80" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M30 30L70 70M70 30L30 70" stroke="#38bdf8" stroke-width="6" stroke-linecap="round"/>
                    <path d="M50 20L50 80M20 50L80 50" stroke="#38bdf8" stroke-width="6" stroke-linecap="round"/>
                    <path d="M35 20L65 80M65 20L35 80" stroke="#38bdf8" stroke-width="4" stroke-linecap="round"/>
                </svg>
            </div>
            <h1 class="text-2xl font-black tracking-widest text-white uppercase">DENIER</h1>
            <p class="text-xs font-semibold text-slate-400 tracking-wider uppercase mt-1">INVENTARIO</p>
            <div class="border-b border-slate-800 my-4"></div>
        </div>
        
        {% if error %}
        <div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded-lg mb-4 text-xs font-bold text-center">{{ error }}</div>
        {% endif %}

        <form action="/login" method="POST" class="space-y-4">
            <div>
                <label class="block text-xs font-bold mb-1.5 text-slate-300">Email</label>
                <input type="email" name="email" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white focus:outline-none focus:border-teal-500" placeholder="controlprod_1@deniersa.com">
            </div>
            <div>
                <label class="block text-xs font-bold mb-1.5 text-slate-300">Password</label>
                <input type="password" name="password" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white focus:outline-none focus:border-teal-500" placeholder="••••••••">
            </div>
            <button type="submit" class="w-full bg-teal-700 hover:bg-teal-600 text-white py-2.5 rounded-lg font-bold text-sm shadow transition duration-200">Log in</button>
        </form>

        <div class="flex justify-between items-center mt-6 text-xs text-slate-400">
            <span class="text-slate-500">¿Problemas?</span>
            <a href="/reset-password-form" class="text-teal-400 font-semibold hover:underline">Reset Password</a>
        </div>
    </div>
</body>
</html>
"""

RESET_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Recuperar Contraseña - DENIER</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen flex items-center justify-center p-6">
    <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-sm">
        <div class="text-center mb-6">
            <h1 class="text-xl font-bold text-white">Restablecer Contraseña</h1>
            <p class="text-xs text-slate-400 mt-1">Ingrese su correo corporativo y su nueva clave</p>
            <div class="border-b border-slate-800 my-4"></div>
        </div>

        {% if error %}
        <div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded-lg mb-4 text-xs font-bold text-center">{{ error }}</div>
        {% endif %}
        {% if mensaje %}
        <div class="bg-emerald-950 border border-emerald-800 text-emerald-300 p-3 rounded-lg mb-4 text-xs font-bold text-center">{{ mensaje }}</div>
        {% endif %}

        <form action="/procesar-reset" method="POST" class="space-y-4">
            <div>
                <label class="block text-xs font-bold mb-1.5 text-slate-300">Correo Corporativo</label>
                <input type="email" name="email" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white focus:outline-none focus:border-teal-500" placeholder="controlprod_1@deniersa.com">
            </div>
            <div>
                <label class="block text-xs font-bold mb-1.5 text-slate-300">Nueva Contraseña</label>
                <input type="password" name="nueva_password" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white focus:outline-none focus:border-teal-500" placeholder="••••••••">
            </div>
            <button type="submit" class="w-full bg-teal-700 hover:bg-teal-600 text-white py-2.5 rounded-lg font-bold text-sm shadow transition duration-200">Actualizar Contraseña</button>
        </form>

        <div class="text-center mt-6">
            <a href="/login-form" class="text-xs text-teal-400 font-semibold hover:underline">← Volver al Login</a>
        </div>
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
<body class="bg-black text-slate-100 min-h-screen p-8">
    <div class="max-w-6xl mx-auto">
        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8 flex justify-between items-center">
            <div>
                <h1 class="text-2xl font-black tracking-wide text-white">DENIER - INVENTARIO</h1>
                <p class="text-sm text-slate-400">Usuario: <b class="text-teal-400">{{ session.get('nombre') }}</b> ({{ session.get('email') }})</p>
            </div>
            <a href="/logout" class="bg-red-950 text-red-300 border border-red-800 px-4 py-2 rounded-xl text-sm font-semibold hover:bg-red-900 transition">Cerrar Sesión</a>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
            {% if 'admin' in session.get('modulos') or 'hilos' in session.get('modulos') %}
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div>
                    <h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Hilos</h2>
                    <p class="text-sm text-slate-400 mb-4">Control de inventario Denier y Artitex, packing lists y devoluciones.</p>
                </div>
                <a href="/hilos" class="bg-teal-700 hover:bg-teal-600 text-white text-center py-3 rounded-xl font-medium shadow transition">Entrar a Hilos</a>
            </div>
            {% endif %}

            {% if 'admin' in session.get('modulos') or 'cruda' in session.get('modulos') %}
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div>
                    <h2 class="text-xl font-bold text-teal-400 mb-2">Módulo de Tela Cruda</h2>
                    <p class="text-sm text-slate-400 mb-4">Control de rollos, cargas masivas Knitopia y reingresos Magic Dyeing.</p>
                </div>
                <a href="/cruda" class="bg-teal-700 hover:bg-teal-600 text-white text-center py-3 rounded-xl font-medium shadow transition">Entrar a Tela Cruda</a>
            </div>
            {% endif %}

            {% if 'admin' in session.get('modulos') or 'devo' in session.get('modulos') %}
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl flex flex-col justify-between">
                <div>
                    <h2 class="text-xl font-bold text-rose-400 mb-2">Módulo Devoluciones</h2>
                    <p class="text-sm text-slate-400 mb-4">Control de devoluciones de tela procesada y gestión de archivos PDF.</p>
                </div>
                <a href="/devo" class="bg-rose-700 hover:bg-rose-600 text-white text-center py-3 rounded-xl font-medium shadow transition">Entrar a Devoluciones</a>
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
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                            <input type="checkbox" name="modulos" value="hilos" class="w-4 h-4 text-teal-600 rounded">
                            <span>Bodega Hilos</span>
                        </label>
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                            <input type="checkbox" name="modulos" value="cruda" class="w-4 h-4 text-teal-600 rounded">
                            <span>Bodega Tela Cruda</span>
                        </label>
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                            <input type="checkbox" name="modulos" value="devo" class="w-4 h-4 text-rose-600 rounded">
                            <span>Devoluciones</span>
                        </label>
                        <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                            <input type="checkbox" name="modulos" value="admin" class="w-4 h-4 text-purple-600 rounded">
                            <span>Administrador Total</span>
                        </label>
                    </div>
                </div>
                <div class="md:col-span-3 flex justify-end">
                    <button type="submit" class="bg-purple-700 hover:bg-purple-600 text-white px-6 py-2.5 rounded-xl text-sm font-medium shadow transition">Registrar Nuevo Usuario</button>
                </div>
            </form>

            <h3 class="font-bold text-sm text-slate-300 mb-3">Lista de Usuarios Activos y Permisos</h3>
            <table class="w-full text-left text-sm">
                <thead>
                    <tr class="border-b border-slate-800 text-slate-400">
                        <th class="p-3">Correo</th>
                        <th class="p-3">Nombre</th>
                        <th class="p-3">Contraseña</th>
                        <th class="p-3">Módulos Asignados</th>
                        <th class="p-3 text-center">Acciones</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-800">
                    {% for u in usuarios_db %}
                    <tr>
                        <td class="p-3 font-medium text-teal-400">{{ u.email }}</td>
                        <td class="p-3 text-slate-200">{{ u.nombre }}</td>
                        <td class="p-3 text-amber-300 font-mono font-bold">{{ u.password }}</td>
                        <td class="p-3">
                            {% for m in u.modulos %}
                                <span class="bg-slate-800 text-slate-300 border border-slate-700 px-2 py-0.5 rounded text-xs font-bold uppercase mr-1">{{ m }}</span>
                            {% endfor %}
                        </td>
                        <td class="p-3 text-center">
                            <div class="flex justify-center gap-2">
                                <a href="/editar-usuario/{{ u.id }}" class="bg-amber-950 text-amber-300 border border-amber-800 px-3 py-1 rounded-lg text-xs font-semibold hover:bg-amber-900">Editar Rol</a>
                                {% if u.email != 'controlprod_1@deniersa.com' %}
                                <a href="/eliminar-usuario/{{ u.id }}" onclick="return confirm('¿Seguro deseas eliminar este usuario?');" class="bg-red-950 text-red-300 border border-red-800 px-3 py-1 rounded-lg text-xs font-semibold hover:bg-red-900">Eliminar</a>
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
</body>
</html>
"""

EDITAR_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Editar Permisos - DENIER</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen flex items-center justify-center p-6">
    <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-2xl w-full max-w-md">
        <h1 class="text-xl font-bold text-purple-400 mb-2">Editar Permisos de Usuario</h1>
        <p class="text-xs text-slate-400 mb-6">Usuario: <b class="text-white">{{ usuario.email }}</b></p>

        <form action="/actualizar-usuario/{{ usuario.id }}" method="POST" class="space-y-4">
            <div>
                <label class="block text-xs font-bold mb-1 text-slate-300">Nombre Completo</label>
                <input type="text" name="nombre" value="{{ usuario.nombre }}" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white">
            </div>
            <div>
                <label class="block text-xs font-bold mb-1 text-slate-300">Contraseña</label>
                <input type="text" name="password" value="{{ usuario.password }}" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-sm text-white">
            </div>
            <div>
                <label class="block text-xs font-bold mb-2 text-slate-300">Modificar Módulos Permitidos:</label>
                <div class="space-y-2 bg-black p-4 rounded-lg border border-slate-800">
                    <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                        <input type="checkbox" name="modulos" value="hilos" {% if 'hilos' in usuario.modulos %}checked{% endif %} class="w-4 h-4 text-teal-600 rounded">
                        <span>Bodega Hilos</span>
                    </label>
                    <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                        <input type="checkbox" name="modulos" value="cruda" {% if 'cruda' in usuario.modulos %}checked{% endif %} class="w-4 h-4 text-teal-600 rounded">
                        <span>Bodega Tela Cruda</span>
                    </label>
                    <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                        <input type="checkbox" name="modulos" value="devo" {% if 'devo' in usuario.modulos %}checked{% endif %} class="w-4 h-4 text-rose-600 rounded">
                        <span>Devoluciones</span>
                    </label>
                    <label class="flex items-center space-x-2 text-sm cursor-pointer text-slate-300">
                        <input type="checkbox" name="modulos" value="admin" {% if 'admin' in usuario.modulos %}checked{% endif %} class="w-4 h-4 text-purple-600 rounded">
                        <span>Administrador Total</span>
                    </label>
                </div>
            </div>
            <div class="flex justify-between items-center pt-4">
                <a href="/" class="text-xs text-slate-400 underline">Cancelar</a>
                <button type="submit" class="bg-purple-700 hover:bg-purple-600 text-white px-5 py-2 rounded-xl text-xs font-bold shadow transition">Guardar Cambios</button>
            </div>
        </form>
    </div>
</body>
</html>
"""

@app.route('/')
def index_main():
    if 'email' not in session:
        return redirect(url_for('login_form'))
    return render_template_string(HOME_TEMPLATE, usuarios_db=usuarios_db)

@app.route('/login-form')
def login_form():
    return render_template_string(LOGIN_TEMPLATE, error=None)

@app.route('/login', methods=['POST'])
def login():
    email = request.form.get('email').strip()
    password = request.form.get('password').strip()
    usuario = next((u for u in usuarios_db if u['email'] == email and u['password'] == password), None)
    if usuario:
        session['email'] = usuario['email']
        session['nombre'] = usuario['nombre']
        session['modulos'] = usuario['modulos']
        return redirect(url_for('index_main'))
    return render_template_string(LOGIN_TEMPLATE, error="Correo o contraseña incorrectos")

@app.route('/reset-password-form')
def reset_password_form():
    return render_template_string(RESET_TEMPLATE, error=None, mensaje=None)

@app.route('/procesar-reset', methods=['POST'])
def procesar_reset():
    email = request.form.get('email').strip()
    nueva_password = request.form.get('nueva_password').strip()
    usuario = next((u for u in usuarios_db if u['email'] == email), None)
    if not usuario:
        return render_template_string(RESET_TEMPLATE, error="El correo ingresado no está registrado en el sistema.", mensaje=None)
    usuario['password'] = nueva_password
    return render_template_string(RESET_TEMPLATE, error=None, mensaje="¡Contraseña actualizada con éxito! Ya puedes iniciar sesión.")

@app.route('/crear-usuario', methods=['POST'])
def crear_usuario():
    if 'admin' not in session.get('modulos', []):
        return redirect(url_for('index_main'))
    nuevo_email = request.form.get('email').strip()
    if any(u['email'] == nuevo_email for u in usuarios_db):
        return redirect(url_for('index_main'))
    modulos_seleccionados = request.form.getlist('modulos')
    nuevo_id = max([u['id'] for u in usuarios_db], default=0) + 1
    nuevo = {
        'id': nuevo_id,
        'email': nuevo_email,
        'password': request.form.get('password').strip(),
        'nombre': request.form.get('nombre').strip(),
        'modulos': modulos_seleccionados
    }
    usuarios_db.append(nuevo)
    return redirect(url_for('index_main'))

@app.route('/editar-usuario/<int:uid>')
def editar_usuario(uid):
    if 'admin' not in session.get('modulos', []):
        return redirect(url_for('index_main'))
    usuario = next((u for u in usuarios_db if u['id'] == uid), None)
    if not usuario:
        return redirect(url_for('index_main'))
    return render_template_string(EDITAR_TEMPLATE, usuario=usuario)

@app.route('/actualizar-usuario/<int:uid>', methods=['POST'])
def actualizar_usuario(uid):
    if 'admin' not in session.get('modulos', []):
        return redirect(url_for('index_main'))
    usuario = next((u for u in usuarios_db if u['id'] == uid), None)
    if usuario:
        usuario['nombre'] = request.form.get('nombre').strip()
        usuario['password'] = request.form.get('password').strip()
        usuario['modulos'] = request.form.getlist('modulos')
    return redirect(url_for('index_main'))

@app.route('/eliminar-usuario/<int:uid>')
def eliminar_usuario(uid):
    if 'admin' not in session.get('modulos', []):
        return redirect(url_for('index_main'))
    global usuarios_db
    usuarios_db = [u for u in usuarios_db if u['id'] != uid or u['email'] == 'controlprod_1@deniersa.com']
    return redirect(url_for('index_main'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_form'))


# =========================================================================
# 3. CÓDIGO ÍNTEGRO DE HILO_2.PY (MÓDULO DE HILOS)
# =========================================================================
HTML_TEMPLATE_HILOS = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>DENIER - Hilos</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen p-6">
    <div class="max-w-7xl mx-auto">
        <div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl">
            <div>
                <h1 class="text-2xl font-black tracking-wider uppercase text-white">DENIER - HILOS</h1>
                <p class="text-sm text-slate-400">Módulo de Control de Conos y Hilos</p>
            </div>
            <div>
                {% if session.get('es_admin_hilos') %}
                    <span class="bg-emerald-950 text-emerald-400 border border-emerald-800 text-xs px-3 py-1 rounded-full font-bold">Admin Activo</span>
                    <a href="/hilos/logout-admin" class="text-xs text-red-400 underline ml-2">Salir</a>
                {% else %}
                    <a href="/hilos/login-admin-form" class="text-xs text-teal-400 underline font-medium">Acceso Admin</a>
                {% endif %}
            </div>
        </div>

        <div class="flex flex-wrap gap-3 mb-6">
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white shadow hover:bg-slate-700">← Menú DENIER</a>
            <a href="/hilos" class="px-4 py-3 rounded-xl font-semibold bg-teal-700 text-white shadow">Inventario Activo</a>
            <a href="/hilos/agotados" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700 shadow">Hilos Sin Stock</a>
            <a href="/hilos/historial" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700 shadow">Historial Packing List</a>
            <a href="/hilos/devoluciones-modulo" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700 shadow">Módulo Devoluciones</a>
            <a href="/hilos/exportar-excel" class="px-4 py-3 rounded-xl font-semibold bg-emerald-800 text-white shadow ml-auto">Descargar Excel</a>
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
                <div id="ocDiv">
                    <label class="block text-sm mb-1 text-slate-300">Orden / Ingreso</label>
                    <input type="text" name="oc" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Código</label>
                    <input type="text" name="codigo" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Calibre</label>
                    <input type="text" name="calibre" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Composición</label>
                    <input type="text" name="composicion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Proveedor</label>
                    <input type="text" name="proveedor" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Lote</label>
                    <input type="text" name="lote" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Kilos Iniciales</label>
                    <input type="number" step="0.01" name="kilos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Conos Iniciales</label>
                    <input type="number" name="conos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Ubicación (Ej: A-1, B-2)</label>
                    <input type="text" name="ubicacion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: A-1">
                </div>
                <div class="md:col-span-3 flex justify-end">
                    <button type="submit" class="bg-teal-700 hover:bg-teal-600 text-white px-6 py-2.5 rounded-xl font-medium">Guardar en Inventario</button>
                </div>
            </form>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <form method="GET" action="/hilos" class="grid grid-cols-1 md:grid-cols-4 gap-4">
                <div>
                    <label class="block text-xs mb-1 text-slate-400">Filtrar por Tipo</label>
                    <select name="filtro_tipo" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                        <option value="">-- Todos --</option>
                        <option value="Denier" {{ 'selected' if filtro_tipo == 'Denier' else '' }}>Denier</option>
                        <option value="Artitex" {{ 'selected' if filtro_tipo == 'Artitex' else '' }}>Artitex</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs mb-1 text-slate-400">Filtrar por Proveedor</label>
                    <select name="filtro_prov" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                        <option value="">-- Todos --</option>
                        {% for prov in proveedores_unicos %}
                        <option value="{{ prov }}" {{ 'selected' if filtro_prov == prov else '' }}>{{ prov }}</option>
                        {% endfor %}
                    </select>
                </div>
                <div>
                    <label class="block text-xs mb-1 text-slate-400">Buscar (Lote, OC, Composición o Ubicación ej: A-1)</label>
                    <input type="text" name="q" value="{{ query }}" placeholder="Ej: A-1, L-442..." class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div class="flex items-end">
                    <button type="submit" class="w-full bg-slate-800 hover:bg-slate-700 border border-slate-700 text-white p-2.5 rounded-lg font-medium">Filtrar y Sumar</button>
                </div>
            </form>
        </div>

        <form action="/hilos/generar-packing-form" method="POST">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-6 flex justify-between items-center">
                <span class="text-sm text-slate-300">
                    Suma Filtrada: <strong class="text-teal-400 text-base">{{ "%.2f"|format(total_kilos_filtrados) }} kg</strong> / <strong class="text-teal-400 text-base">{{ total_conos_filtrados }} conos</strong>
                </span>
                <button type="submit" class="bg-purple-700 hover:bg-purple-600 text-white px-6 py-2.5 rounded-xl font-medium shadow">Crear Packing List</button>
            </div>

            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
                <table class="w-full text-left">
                    <thead>
                        <tr class="border-b border-slate-800 text-sm text-slate-400">
                            <th class="p-3">Seleccionar</th>
                            <th class="p-3">Tipo</th>
                            <th class="p-3">Orden</th>
                            <th class="p-3">Lote</th>
                            <th class="p-3">Composición</th>
                            <th class="p-3">Calibre</th>
                            <th class="p-3">Kilos Disponibles</th>
                            <th class="p-3">Conos</th>
                            <th class="p-3">Ubicación</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-800 text-sm">
                        {% for item in resultados %}
                        <tr class="{{ 'bg-amber-950/40' if item.kilos < 50 else '' }}">
                            <td class="p-3 text-center">
                                <input type="checkbox" name="hilos_sel" value="{{ item.id }}" class="w-4 h-4 text-purple-600 rounded">
                            </td>
                            <td class="p-3 font-medium text-slate-200">{{ item.tipo }}</td>
                            <td class="p-3 font-bold text-teal-400">{{ item.oc }}</td>
                            <td class="p-3 font-semibold text-slate-200">{{ item.lote }}</td>
                            <td class="p-3 text-slate-300">{{ item.composicion }}</td>
                            <td class="p-3 text-slate-300">{{ item.calibre }}</td>
                            <td class="p-3 font-bold">
                                {% if item.kilos < 50 %}
                                    <span class="text-amber-400 bg-amber-950 border border-amber-800 px-2 py-1 rounded">ALERTA: {{ item.kilos }} kg</span>
                                {% else %}
                                    <span class="text-teal-300">{{ item.kilos }} kg</span>
                                {% endif %}
                            </td>
                            <td class="p-3 text-purple-400 font-bold">{{ item.conos }} conos</td>
                            <td class="p-3 font-semibold text-emerald-400">{{ item.ubicacion }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </form>
    </div>
    <script>
        function toggleOC() {
            const tipo = document.getElementById('tipoSelect').value;
            document.getElementById('ocDiv').style.display = tipo === 'Artitex' ? 'none' : 'block';
        }
    </script>
</body>
</html>
"""

AGOTADOS_HILOS_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Agotados</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-8">
    <div class="max-w-6xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow">
        <h1 class="text-2xl font-bold text-amber-400 mb-6">Hilos sin Stock</h1>
        <a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium">Volver</a>
        <table class="w-full mt-4 text-left text-sm">
            <thead><tr class="border-b border-slate-800"><th class="p-2">Lote</th><th class="p-2">Composición</th><th class="p-2">Proveedor</th></tr></thead>
            <tbody>
                {% for item in agotados %}
                <tr class="border-b border-slate-800"><td class="p-2">{{ item.lote }}</td><td class="p-2">{{ item.composicion }}</td><td class="p-2">{{ item.proveedor }}</td></tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</body>
</html>
"""

DEVOLUCIONES_HILOS_MODULO_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Devoluciones</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-8">
    <div class="max-w-5xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow">
        <h1 class="text-2xl font-bold text-rose-400 mb-6">Módulo de Devolución Parcial</h1>
        <a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a>
        {% if error %}<div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded mb-4 text-sm font-bold">{{ error }}</div>{% endif %}
        <form action="/hilos/procesar-devolucion-parcial" method="POST" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800">
            <select name="historial_id" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white">
                <option value="">-- Elige Packing List --</option>
                {% for p in historial %}<option value="{{ p.id_historial }}">N {{ p.numero }} | Cliente: {{ p.cliente }}</option>{% endfor %}
            </select>
            <div class="grid grid-cols-2 gap-4">
                <input type="number" step="0.01" name="ret_kilos" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Kilos a retornar">
                <input type="number" name="ret_conos" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Conos a retornar">
            </div>
            <input type="text" name="motivo" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Motivo de devolución...">
            <button type="submit" class="bg-rose-700 text-white px-6 py-2 rounded-xl font-medium">Procesar Reingreso</button>
        </form>
    </div>
</body>
</html>
"""

PACKING_HILOS_CANTIDADES_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Cantidades</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-6">
    <div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow">
        <h1 class="text-2xl font-bold mb-6 text-purple-400">Configurar Despacho</h1>
        <form action="/hilos/procesar-packing-multiple" method="POST" class="space-y-6">
            <div class="grid grid-cols-2 gap-4 pb-4 border-b border-slate-800">
                <input type="text" name="cliente" class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Cliente...">
                <input type="text" name="orden" class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Orden...">
            </div>
            {% for item in seleccionados %}
            <input type="hidden" name="hilos_sel" value="{{ item.id }}">
            <div class="p-4 bg-black rounded-xl border border-slate-800">
                <p class="font-bold text-teal-400">Lote: {{ item.lote }} | Ubicación: {{ item.ubicacion }}</p>
                <div class="grid grid-cols-2 gap-4 mt-2">
                    <input type="number" step="0.01" name="req_kilos_{{ item.id }}" value="{{ item.kilos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white">
                    <input type="number" name="req_conos_{{ item.id }}" value="{{ item.conos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white">
                </div>
            </div>
            {% endfor %}
            <button type="submit" class="bg-purple-700 text-white px-6 py-3 rounded-xl font-medium w-full">Generar Packing List</button>
        </form>
    </div>
</body>
</html>
"""

PACKING_HILOS_VIEW_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Packing List</title></head>
<body style="font-family: monospace; background: #fff; color: #000; padding: 20px;">
    <h2>PACKING LIST - HILO No {{ numero }}</h2>
    <p>Cliente: {{ cliente }} | Orden: {{ orden }}</p>
    <hr>
    {% for item in seleccionados %}
    <p>Lote: {{ item.lote }} | Kilos: {{ item.req_kilos }} | Conos: {{ item.req_conos }} | Ubicación: {{ item.ubicacion }}</p>
    {% endfor %}
    <hr>
    <p><strong>Totales: {{ total_conos }} Conos / {{ "%.2f"|format(total_kilos) }} Kg</strong></p>
    <br><a href="/hilos">Volver al Inventario</a>
</body>
</html>
"""

HISTORIAL_HILOS_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Historial</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-8">
    <div class="max-w-6xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow">
        <h1 class="text-2xl font-bold text-emerald-400 mb-6">Historial Packing Lists</h1>
        <a href="/hilos" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a>
        <table class="w-full text-left text-sm">
            <thead><tr class="border-b border-slate-800"><th class="p-2">Número</th><th class="p-2">Cliente</th><th class="p-2">Totales</th><th class="p-2">Acciones</th></tr></thead>
            <tbody>
                {% for p in historial %}
                <tr class="border-b border-slate-800">
                    <td class="p-2 font-bold">N {{ p.numero }}</td>
                    <td class="p-2">{{ p.cliente }}</td>
                    <td class="p-2">{{ p.total_kilos }} kg / {{ p.total_conos }} conos</td>
                    <td class="p-2"><a href="/hilos/ver-packing/{{ p.id_historial }}" target="_blank" class="text-teal-400 underline">Ver</a></td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</body>
</html>
"""

LOGIN_ADMIN_HILOS_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Admin</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black min-h-screen flex items-center justify-center p-6">
    <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow w-full max-w-md">
        <h1 class="text-xl font-bold mb-4 text-white">Admin Hilos</h1>
        {% if error %}<p class="text-red-400 text-sm mb-2">{{ error }}</p>{% endif %}
        <form action="/hilos/login-admin" method="POST" class="space-y-4">
            <input type="password" name="password" required class="w-full p-2.5 bg-black border border-slate-700 rounded text-white" placeholder="Clave...">
            <button type="submit" class="bg-teal-700 text-white w-full py-2 rounded">Ingresar</button>
        </form>
    </div>
</body>
</html>
"""

@app.route('/hilos')
def hilos_index():
    if 'email' not in session or ('hilos' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])):
        return redirect(url_for('index_main'))
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
    
    total_kilos_filtrados = sum(i['kilos'] for i in resultados)
    total_conos_filtrados = sum(i['conos'] for i in resultados)
    
    return render_template_string(HTML_TEMPLATE_HILOS, resultados=resultados, proveedores_unicos=proveedores_unicos, filtro_tipo=filtro_tipo, filtro_prov=filtro_prov, total_kilos_filtrados=total_kilos_filtrados, total_conos_filtrados=total_conos_filtrados, query=query)

@app.route('/hilos/agotados')
def hilos_agotados(): return render_template_string(AGOTADOS_HILOS_TEMPLATE, agotados=[i for i in inventario_hilos if i['kilos'] <= 0])

@app.route('/hilos/devoluciones-modulo')
def hilos_devoluciones_modulo(): return render_template_string(DEVOLUCIONES_HILOS_MODULO_TEMPLATE, historial=packing_history_hilos, error=None)

@app.route('/hilos/procesar-devolucion-parcial', methods=['POST'])
def hilos_procesar_devolucion_parcial():
    hid = int(request.form.get('historial_id'))
    p = next((x for x in packing_history_hilos if x['id_historial'] == hid), None)
    if p:
        p['total_kilos'] += float(request.form.get('ret_kilos'))
        p['total_conos'] += int(request.form.get('ret_conos'))
    return redirect(url_for('hilos_devoluciones_modulo'))

@app.route('/hilos/exportar-excel')
def hilos_exportar_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Tipo", "OC", "Código", "Calibre", "Composición", "Proveedor", "Lote", "Kilos", "Conos", "Ubicación"])
    for i in inventario_hilos:
        ws.append([i['tipo'], i['oc'], i['codigo'], i['calibre'], i['composicion'], i['proveedor'], i['lote'], i['kilos'], i['conos'], i['ubicacion']])
    io_b = io.BytesIO()
    wb.save(io_b)
    io_b.seek(0)
    return send_file(io_b, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Hilos.xlsx')

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

@app.route('/hilos/generar-packing-form', methods=['POST'])
def hilos_generar_packing_form():
    ids = request.form.getlist('hilos_sel')
    if not ids: return redirect(url_for('hilos_index'))
    return render_template_string(PACKING_HILOS_CANTIDADES_TEMPLATE, seleccionados=[i for i in inventario_hilos if i['id'] in [int(x) for x in ids]])

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
    return render_template_string(PACKING_HILOS_VIEW_TEMPLATE, numero=reg['numero'], cliente=reg['cliente'], orden=reg['orden'], seleccionados=sel, total_kilos=tot_k, total_conos=tot_c)

@app.route('/hilos/historial')
def hilos_historial(): return render_template_string(HISTORIAL_HILOS_TEMPLATE, historial=packing_history_hilos)

@app.route('/hilos/ver-packing/<int:hid>')
def hilos_ver_packing(hid):
    r = next((p for p in packing_history_hilos if p['id_historial'] == hid), None)
    return render_template_string(PACKING_HILOS_VIEW_TEMPLATE, numero=r['numero'], cliente=r['cliente'], orden=r['orden'], seleccionados=r['seleccionados'], total_kilos=r['total_kilos'], total_conos=r['total_conos'])

@app.route('/hilos/login-admin-form')
def hilos_login_admin_form(): return render_template_string(LOGIN_ADMIN_HILOS_TEMPLATE, error=None)

@app.route('/hilos/login-admin', methods=['POST'])
def hilos_login_admin():
    if request.form.get('password') == 'admin123':
        session['es_admin_hilos'] = True
        return redirect(url_for('hilos_historial'))
    return render_template_string(LOGIN_ADMIN_HILOS_TEMPLATE, error="Clave incorrecta")

@app.route('/hilos/logout-admin')
def hilos_logout_admin():
    session.pop('es_admin_hilos', None)
    return redirect(url_for('hilos_index'))


# =========================================================================
# 4. CÓDIGO ÍNTEGRO DE CRUDO_2.PY (MÓDULO TELA CRUDA)
# =========================================================================
HTML_TEMPLATE_TELA = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>DENIER - Tela Cruda</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen p-6">
    <div class="max-w-7xl mx-auto">
        <div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl">
            <div>
                <h1 class="text-2xl font-black tracking-wider uppercase text-white">DENIER - TELA CRUDA</h1>
                <p class="text-sm text-slate-400">Módulo de Tejeduría y Rollos Crudos</p>
            </div>
            <div>
                <span class="bg-emerald-950 text-emerald-400 border border-emerald-800 text-xs px-3 py-1 rounded-full font-bold">Sistema Activo</span>
            </div>
        </div>

        <div class="flex flex-wrap gap-3 mb-6">
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white shadow hover:bg-slate-700">← Menú DENIER</a>
            <a href="/cruda" class="px-4 py-3 rounded-xl font-semibold bg-teal-700 text-white shadow">Inventario Cruda</a>
            <a href="/cruda/dashboard" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700 shadow">Dashboard KPIs</a>
            <a href="/cruda/carga-masiva-form" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700 shadow">Carga Masiva (Knitopia)</a>
            <a href="/cruda/historial-tela" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700 shadow">Historial Despachos</a>
            <a href="/cruda/reingresos-modulo" class="px-4 py-3 rounded-xl font-semibold bg-slate-800 text-slate-200 border border-slate-700 shadow">Módulo Reingresos</a>
            <a href="/cruda/exportar-excel-tela" class="px-4 py-3 rounded-xl font-semibold bg-emerald-800 text-white shadow ml-auto">Descargar Excel</a>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <h2 class="text-xl font-bold mb-4 text-teal-400">Registrar Rollo Individual</h2>
            <form action="/cruda/guardar-tela" method="POST" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Orden de Compra (OC)</label>
                    <input type="text" name="oc" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: OC-4050">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Proveedor / Tejeduría</label>
                    <input type="text" name="proveedor" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: Knitopia">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Lote de Tela</label>
                    <input type="text" name="lote_tela" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">No. de Rollo</label>
                    <input type="text" name="rollo" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: R-001">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Kilos del Rollo</label>
                    <input type="number" step="0.01" name="kilos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Ubicación (Ej: A-1, B-2)</label>
                    <input type="text" name="ubicacion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: A-1">
                </div>
                <div class="md:col-span-3 flex justify-end">
                    <button type="submit" class="bg-teal-700 hover:bg-teal-600 text-white px-6 py-2.5 rounded-xl font-medium">Guardar en Inventario</button>
                </div>
            </form>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <form method="GET" action="/cruda" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                    <label class="block text-xs mb-1 text-slate-400">Filtrar por Proveedor</label>
                    <select name="filtro_prov" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                        <option value="">-- Todos --</option>
                        {% for prov in proveedores_unicos %}
                        <option value="{{ prov }}" {{ 'selected' if filtro_prov == prov else '' }}>{{ prov }}</option>
                        {% endfor %}
                    </select>
                </div>
                <div>
                    <label class="block text-xs mb-1 text-slate-400">Buscar (Lote, OC, Rollo o Ubicación ej: A-1)</label>
                    <input type="text" name="q" value="{{ query }}" placeholder="Ej: A-1, TC-101..." class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div class="flex items-end">
                    <button type="submit" class="w-full bg-slate-800 hover:bg-slate-700 border border-slate-700 text-white p-2.5 rounded-lg font-medium">Filtrar y Buscar</button>
                </div>
            </form>
        </div>

        <form action="/cruda/generar-packing-tela-form" method="POST">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-6 flex justify-between items-center">
                <span class="text-sm text-slate-400">Total Kilos Filtrados: <strong class="text-teal-400 text-base">{{ "%.2f"|format(total_kilos_filtrados) }} kg</strong></span>
                <button type="submit" class="bg-purple-700 hover:bg-purple-600 text-white px-6 py-2.5 rounded-xl font-medium shadow">Crear Packing List Tela</button>
            </div>

            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
                <table class="w-full text-left">
                    <thead>
                        <tr class="border-b border-slate-800 text-sm text-slate-400">
                            <th class="p-3">Seleccionar</th>
                            <th class="p-3">Orden de Compra</th>
                            <th class="p-3">Proveedor</th>
                            <th class="p-3">Lote Tela</th>
                            <th class="p-3">Rollo</th>
                            <th class="p-3">Kilos</th>
                            <th class="p-3">Ubicación Exacta</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-800 text-sm">
                        {% for item in resultados %}
                        <tr class="{{ 'bg-amber-950/40' if item.kilos < 50 else '' }}">
                            <td class="p-3 text-center">
                                <input type="checkbox" name="telas_sel" value="{{ item.id }}" class="w-4 h-4 text-purple-600 rounded">
                            </td>
                            <td class="p-3 font-bold text-teal-400">{{ item.oc }}</td>
                            <td class="p-3 font-semibold text-slate-200">{{ item.proveedor }}</td>
                            <td class="p-3 font-semibold text-slate-200">{{ item.lote_tela }}</td>
                            <td class="p-3 font-bold text-slate-300 bg-slate-800 rounded">{{ item.rollo }}</td>
                            <td class="p-3 font-bold">
                                {% if item.kilos < 50 %}
                                    <span class="text-amber-400 bg-amber-950 border border-amber-800 px-2 py-1 rounded">ALERTA: {{ item.kilos }} kg</span>
                                {% else %}
                                    <span class="text-teal-300">{{ item.kilos }} kg</span>
                                {% endif %}
                            </td>
                            <td class="p-3 font-semibold text-emerald-400">{{ item.ubicacion }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </form>
    </div>
</body>
</html>
"""

CARGA_MASIVA_TELA_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Carga Masiva</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-8">
    <div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow-xl">
        <h1 class="text-2xl font-bold text-amber-400 mb-4">Carga Masiva de Inventario</h1>
        <div class="mb-6 flex gap-4">
            <a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium inline-block">Volver</a>
            <a href="/cruda/descargar-plantilla" class="bg-emerald-800 text-white px-4 py-2 rounded-xl text-sm font-medium inline-block">Descargar Plantilla Excel</a>
        </div>
        {% if mensaje %}<div class="bg-emerald-950 border border-emerald-800 text-emerald-300 p-3 rounded mb-4 text-sm font-bold">{{ mensaje }}</div>{% endif %}
        <form action="/cruda/procesar-carga-masiva" method="POST" enctype="multipart/form-data" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800">
            <input type="file" name="excel_file" accept=".xlsx" required class="w-full p-2.5 rounded-lg border border-slate-700 bg-slate-900 text-white">
            <button type="submit" class="bg-amber-700 hover:bg-amber-600 text-white px-6 py-2.5 rounded-xl font-medium w-full">Cargar al Sistema</button>
        </form>
    </div>
</body>
</html>
"""

DASHBOARD_TELA_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Dashboard KPIs</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-8">
    <div class="max-w-7xl mx-auto">
        <div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow">
            <h1 class="text-2xl font-bold text-teal-400">KPIs - Tela Cruda</h1>
            <a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium">Volver</a>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8">
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow">
                <p class="text-sm text-slate-400">Total Kilos</p>
                <h3 class="text-3xl font-bold text-white mt-2">{{ "%.2f"|format(total_kilos_inventario) }} kg</h3>
            </div>
            <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow">
                <p class="text-sm text-slate-400">Rollos Activos</p>
                <h3 class="text-3xl font-bold text-white mt-2">{{ total_rollos }}</h3>
            </div>
        </div>
    </div>
</body>
</html>
"""

REINGRESOS_TELA_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Reingresos</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-8">
    <div class="max-w-5xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow">
        <h1 class="text-2xl font-bold text-rose-400 mb-6">Módulo de Reingresos</h1>
        <a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a>
        <form action="/cruda/procesar-reingreso-tela" method="POST" class="space-y-4 bg-black p-6 rounded-xl border border-slate-800">
            <select name="historial_id" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white">
                <option value="">-- Elige Despacho --</option>
                {% for p in historial %}<option value="{{ p.id_historial }}">N {{ p.numero }} | Destino: {{ p.cliente }}</option>{% endfor %}
            </select>
            <input type="number" step="0.01" name="ret_kilos" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Kilos a reingresar">
            <input type="text" name="motivo" required class="w-full p-2.5 bg-slate-900 border border-slate-700 rounded-lg text-white" placeholder="Motivo...">
            <button type="submit" class="bg-rose-700 text-white px-6 py-2 rounded-xl font-medium">Reingresar</button>
        </form>
    </div>
</body>
</html>
"""

PACKING_TELA_CANTIDADES_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Cantidades</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-6">
    <div class="max-w-3xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow">
        <h1 class="text-2xl font-bold mb-6 text-purple-400">Configurar Despacho</h1>
        <form action="/cruda/procesar-packing-tela-multiple" method="POST" class="space-y-6">
            <input type="text" name="cliente" required class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Cliente destino...">
            <input type="text" name="orden" required class="w-full p-3 bg-black border border-slate-700 rounded-lg text-white" placeholder="Orden...">
            {% for item in seleccionados %}
            <input type="hidden" name="telas_sel" value="{{ item.id }}">
            <div class="p-4 bg-black rounded-xl border border-slate-800">
                <p class="font-bold text-teal-400">Rollo: {{ item.rollo }} | Ubicación: {{ item.ubicacion }}</p>
                <input type="number" step="0.01" name="req_kilos_{{ item.id }}" value="{{ item.kilos }}" required class="w-full p-2 bg-slate-900 border border-slate-700 rounded text-white mt-2">
            </div>
            {% endfor %}
            <button type="submit" class="bg-purple-700 text-white px-6 py-3 rounded-xl font-medium w-full">Generar Packing List</button>
        </form>
    </div>
</body>
</html>
"""

PACKING_TELA_VIEW_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Packing List</title></head>
<body style="font-family: monospace; background: #fff; color: #000; padding: 20px;">
    <h2>PACKING LIST TELA No {{ numero }}</h2>
    <p>Destino: {{ cliente }} | Orden: {{ orden }}</p>
    <hr>
    {% for item in seleccionados %}
    <p>Rollo: {{ item.rollo }} | Kilos: {{ item.req_kilos }} | Ubicación: {{ item.ubicacion }}</p>
    {% endfor %}
    <hr>
    <p><strong>Total Kilos: {{ "%.2f"|format(total_kilos) }} Kg</strong></p>
    <br><a href="/cruda">Volver al Inventario</a>
</body>
</html>
"""

HISTORIAL_TELA_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Historial</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-black text-slate-100 p-8">
    <div class="max-w-6xl mx-auto bg-slate-900 border border-slate-800 p-8 rounded-2xl shadow">
        <h1 class="text-2xl font-bold text-emerald-400 mb-6">Historial Despachos</h1>
        <a href="/cruda" class="bg-slate-800 text-slate-200 px-4 py-2 rounded-xl text-sm font-medium mb-6 inline-block">Volver</a>
        <table class="w-full text-left text-sm">
            <thead><tr class="border-b border-slate-800"><th class="p-2">Número</th><th class="p-2">Destino</th><th class="p-2">Kilos</th><th class="p-2">Acción</th></tr></thead>
            <tbody>
                {% for p in historial %}
                <tr class="border-b border-slate-800">
                    <td class="p-2 font-bold">N {{ p.numero }}</td>
                    <td class="p-2">{{ p.cliente }}</td>
                    <td class="p-2">{{ "%.2f"|format(p.total_kilos) }} kg</td>
                    <td class="p-2"><a href="/cruda/ver-packing-tela/{{ p.id_historial }}" target="_blank" class="text-teal-400 underline">Ver</a></td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</body>
</html>
"""

@app.route('/cruda')
def cruda_index():
    if 'email' not in session or ('cruda' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])):
        return redirect(url_for('index_main'))
    query = request.args.get('q', '').lower()
    filtro_prov = request.args.get('filtro_prov', '')
    proveedores_unicos = sorted(list(set(i['proveedor'] for i in inventario_tela)))
    activos = [i for i in inventario_tela if i['kilos'] > 0]
    
    resultados = activos
    if filtro_prov:
        resultados = [i for i in resultados if i['proveedor'] == filtro_prov]
    if query:
        resultados = [i for i in resultados if query in i['lote_tela'].lower() or query in i['oc'].lower() or query in i['rollo'].lower() or query in i['proveedor'].lower() or query in i['ubicacion'].lower()]
    
    total_kilos_filtrados = sum(i['kilos'] for i in resultados)
    return render_template_string(HTML_TEMPLATE_TELA, resultados=resultados, proveedores_unicos=proveedores_unicos, filtro_prov=filtro_prov, query=query, total_kilos_filtrados=total_kilos_filtrados)

@app.route('/cruda/dashboard')
def cruda_dashboard():
    activos = [i for i in inventario_tela if i['kilos'] > 0]
    return render_template_string(DASHBOARD_TELA_TEMPLATE, total_kilos_inventario=sum(i['kilos'] for i in activos), total_rollos=len(activos))

@app.route('/cruda/carga-masiva-form')
def cruda_carga_masiva_form(): return render_template_string(CARGA_MASIVA_TELA_TEMPLATE, mensaje=None)

@app.route('/cruda/descargar-plantilla')
def cruda_descargar_plantilla():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["OC", "Proveedor", "Lote Tela", "Rollo", "Kilos", "Ubicacion"])
    ws.append(["OC-1001", "Knitopia", "LOTE-500", "R-01", 125.50, "A-1"])
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return send_file(excel_io, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Plantilla.xlsx')

@app.route('/cruda/procesar-carga-masiva', methods=['POST'])
def cruda_procesar_carga_masiva():
    archivo = request.files.get('excel_file')
    if archivo:
        wb = openpyxl.load_workbook(archivo)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[4] is not None:
                inventario_tela.append({'id': len(inventario_tela)+1, 'oc': str(row[0]), 'proveedor': str(row[1]), 'lote_tela': str(row[2]), 'rollo': str(row[3]), 'kilos': float(row[4]), 'ubicacion': str(row[5]) if len(row)>5 and row[5] else 'A-1'})
    return render_template_string(CARGA_MASIVA_TELA_TEMPLATE, mensaje="¡Carga exitosa!")

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

@app.route('/cruda/reingresos-modulo')
def cruda_reingresos_modulo(): return render_template_string(REINGRESOS_TELA_TEMPLATE, historial=packing_tela_history)

@app.route('/cruda/procesar-reingreso-tela', methods=['POST'])
def cruda_procesar_reingreso_tela():
    hid = int(request.form.get('historial_id'))
    ret_k = float(request.form.get('ret_kilos'))
    p = next((x for x in packing_tela_history if x['id_historial'] == hid), None)
    if p:
        p['total_kilos'] += ret_k
    return redirect(url_for('cruda_reingresos_modulo'))

@app.route('/cruda/generar-packing-tela-form', methods=['POST'])
def cruda_generar_packing_tela_form():
    ids = request.form.getlist('telas_sel')
    if not ids: return redirect(url_for('cruda_index'))
    return render_template_string(PACKING_TELA_CANTIDADES_TEMPLATE, seleccionados=[i for i in inventario_tela if i['id'] in [int(x) for x in ids]])

@app.route('/cruda/procesar-packing-tela-multiple', methods=['POST'])
def cruda_procesar_packing_tela_multiple():
    global contador_packing_tela
    ids = request.form.getlist('telas_sel')
    sel = []
    tot = 0
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
    return render_template_string(PACKING_TELA_VIEW_TEMPLATE, numero=reg['numero'], cliente=reg['cliente'], orden=reg['orden'], seleccionados=sel, total_kilos=tot)

@app.route('/cruda/historial-tela')
def cruda_historial_tela(): return render_template_string(HISTORIAL_TELA_TEMPLATE, historial=packing_tela_history)

@app.route('/cruda/ver-packing-tela/<int:hid>')
def cruda_ver_packing_tela(hid):
    r = next((p for p in packing_tela_history if p['id_historial'] == hid), None)
    return render_template_string(PACKING_TELA_VIEW_TEMPLATE, numero=r['numero'], cliente=r['cliente'], orden=r['orden'], seleccionados=r['seleccionados'], total_kilos=r['total_kilos'])

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
    return send_file(io_b, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Tela.xlsx')


# =========================================================================
# 5. CÓDIGO ÍNTEGRO DE DEVO_2.PY (MÓDULO DEVOLUCIONES)
# =========================================================================
HTML_TEMPLATE_DEVOLUCION = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>DENIER - Devoluciones de Tela Procesada</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-black text-slate-100 min-h-screen p-6">
    <div class="max-w-7xl mx-auto">
        <div class="flex justify-between items-center mb-6 bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl">
            <div>
                <h1 class="text-2xl font-black tracking-wider uppercase text-white">DENIER - TELA PROCESADA DEVUELTA</h1>
                <p class="text-sm text-slate-400">Control de devoluciones y gestión avanzada de archivos PDF/Escáner</p>
            </div>
            <div>
                <span class="bg-rose-950 text-rose-400 border border-rose-800 text-xs px-3 py-1 rounded-full font-bold">Módulo Activo</span>
            </div>
        </div>

        <div class="flex flex-wrap gap-3 mb-6">
            <a href="/" class="px-4 py-3 rounded-xl font-bold bg-slate-800 text-white shadow hover:bg-slate-700">← Menú DENIER</a>
            <a href="/devo" class="px-4 py-3 rounded-xl font-semibold bg-rose-700 text-white shadow">Devoluciones Activas</a>
            <a href="/devo/exportar-excel-devolucion" class="px-4 py-3 rounded-xl font-semibold bg-emerald-800 text-white shadow ml-auto">Descargar Excel</a>
        </div>

        {% if mensaje %}
        <div class="bg-emerald-950 border border-emerald-800 text-emerald-300 p-3 rounded-xl mb-6 text-sm font-bold text-center">{{ mensaje }}</div>
        {% endif %}
        {% if error %}
        <div class="bg-red-950 border border-red-800 text-red-300 p-3 rounded-xl mb-6 text-sm font-bold text-center">{{ error }}</div>
        {% endif %}

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <h2 class="text-xl font-bold mb-4 text-rose-400">Registrar Nueva Tela Devuelta por Cliente</h2>
            <form action="/devo/guardar-devolucion" method="POST" enctype="multipart/form-data" class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Cliente</label>
                    <input type="text" name="cliente" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: Magic Apparel">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">No. de Lote</label>
                    <input type="text" name="lote" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: 20452026-8">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Tela (Tipo y Composición)</label>
                    <input type="text" name="tela" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: Interlock, 100% Algodón">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Color / Tono</label>
                    <input type="text" name="color" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: White, Midnight Navy">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Cantidad de Rollos</label>
                    <input type="number" name="rollos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="0">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Kilos Totales</label>
                    <input type="number" step="0.01" name="kilos" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="0.00">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Yardas Totales</label>
                    <input type="number" step="0.01" name="yardas" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="0.00">
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Calidad</label>
                    <select name="calidad" class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                        <option value="Primera">Primera</option>
                        <option value="Segunda">Segunda</option>
                        <option value="Rechazada">Rechazada</option>
                    </select>
                </div>
                <div>
                    <label class="block text-sm mb-1 text-slate-300">Ubicación</label>
                    <input type="text" name="ubicacion" required class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white" placeholder="Ej: Rack 3 / A-1">
                </div>
                <div class="md:col-span-3">
                    <label class="block text-sm mb-1 text-slate-300">Archivo de Respaldo / Escáner (PDF o Imagen)</label>
                    <input type="file" name="archivo_devolucion" accept=".pdf,.png,.jpg,.jpeg" class="w-full p-2 bg-black border border-slate-700 rounded-lg text-sm text-slate-300">
                </div>
                <div class="md:col-span-3 flex justify-end">
                    <button type="submit" class="bg-rose-700 hover:bg-rose-600 text-white px-6 py-2.5 rounded-xl font-medium transition">Registrar Devolución</button>
                </div>
            </form>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-8">
            <form method="GET" action="/devo" class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                    <label class="block text-xs mb-1 text-slate-400">Buscar (Cliente, Lote, Color, Tela...)</label>
                    <input type="text" name="q" value="{{ query }}" placeholder="Ej: Magic Apparel, 2045..." class="w-full p-2.5 bg-black border border-slate-700 rounded-lg text-white">
                </div>
                <div class="flex items-end">
                    <button type="submit" class="w-full bg-slate-800 hover:bg-slate-700 border border-slate-700 text-white p-2.5 rounded-lg font-medium">Filtrar y Buscar</button>
                </div>
            </form>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl mb-6 flex justify-between items-center">
            <span class="text-sm text-slate-300">
                Suma Filtrada: <strong class="text-rose-400 text-base">{{ "%.2f"|format(total_kilos_filtrados) }} kg</strong> / <strong class="text-rose-400 text-base">{{ "%.2f"|format(total_yardas_filtradas) }} yds</strong>
            </span>
        </div>

        <div class="bg-slate-900 border border-slate-800 p-6 rounded-2xl shadow-xl overflow-x-auto">
            <table class="w-full text-left">
                <thead>
                    <tr class="border-b border-slate-800 text-sm text-slate-400">
                        <th class="p-3">Cliente</th>
                        <th class="p-3">Lote</th>
                        <th class="p-3">Tela / Color</th>
                        <th class="p-3">Rollos</th>
                        <th class="p-3">Kilos / Yardas</th>
                        <th class="p-3">Calidad / Ubicación</th>
                        <th class="p-3">Archivo de Respaldo</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-800 text-sm">
                    {% for item in resultados %}
                    <tr class="{{ 'bg-rose-950/40' if item.calidad == 'Rechazada' else '' }}">
                        <td class="p-3 font-bold text-rose-400">{{ item.cliente }}</td>
                        <td class="p-3 font-semibold text-slate-200">{{ item.lote }}</td>
                        <td class="p-3 text-slate-300">{{ item.tela }}<br><span class="text-xs text-slate-400">{{ item.color }}</span></td>
                        <td class="p-3 text-slate-200 font-bold">{{ item.rollos }} rollos</td>
                        <td class="p-3 font-bold text-rose-300">{{ "%.2f"|format(item.kilos) }} kg<br><span class="text-xs text-slate-400">{{ "%.2f"|format(item.yardas) }} yds</span></td>
                        <td class="p-3">
                            <span class="px-2 py-0.5 rounded text-xs font-bold {{ 'bg-emerald-950 text-emerald-400 border border-emerald-800' if item.calidad == 'Primera' else ('bg-amber-950 text-amber-400 border border-amber-800' if item.calidad == 'Segunda' else 'bg-rose-950 text-rose-400 border border-rose-800') }}">{{ item.calidad }}</span>
                            <div class="text-xs text-emerald-400 mt-1 font-semibold">{{ item.ubicacion }}</div>
                        </td>
                        <td class="p-3">
                            {% if item.archivo %}
                                <div class="mb-2 bg-slate-950 p-2.5 rounded-lg border border-slate-800 flex flex-col gap-1.5">
                                    <div class="flex items-center justify-between">
                                        <a href="/devo/ver-archivo/{{ item.archivo }}" target="_blank" class="text-teal-400 font-semibold hover:underline text-xs flex items-center gap-1">
                                            📄 {{ item.archivo }}
                                        </a>
                                        <a href="/devo/eliminar-archivo/{{ item.id }}" onclick="return confirm('¿Seguro deseas quitar este archivo de respaldo?');" class="text-red-400 hover:text-red-300 text-xs font-bold px-1.5 py-0.5 rounded bg-red-950 border border-red-800" title="Quitar archivo">
                                            🗑 Quitar
                                        </a>
                                    </div>
                                    <div class="flex gap-2 mt-1">
                                        <a href="/devo/ver-archivo/{{ item.archivo }}" target="_blank" class="bg-teal-900 hover:bg-teal-800 text-teal-200 px-2.5 py-1 rounded text-xs font-bold border border-teal-700">
                                            👁 Ver PDF
                                        </a>
                                        <button onclick="imprimirPDF('/devo/ver-archivo/{{ item.archivo }}')" class="bg-blue-900 hover:bg-blue-800 text-blue-200 px-2.5 py-1 rounded text-xs font-bold border border-blue-700">
                                            🖨 Imprimir
                                        </button>
                                    </div>
                                </div>
                            {% else %}
                                <div class="mb-2 text-xs text-amber-400 font-semibold">⚠ Sin archivo adjunto</div>
                            {% endif %}
                            
                            <form action="/devo/subir-archivo-devolucion/{{ item.id }}" method="POST" enctype="multipart/form-data" class="flex gap-2 items-center">
                                <input type="file" name="archivo_devolucion" accept=".pdf,.png,.jpg,.jpeg" required class="w-full text-xs text-slate-400 file:mr-1 file:py-1 file:px-2 file:rounded file:border-0 file:text-xs file:font-semibold file:bg-slate-800 file:text-slate-200">
                                <button type="submit" class="bg-slate-800 hover:bg-slate-700 border border-slate-700 text-white px-3 py-1 rounded text-xs font-bold whitespace-nowrap transition">
                                    {{ 'Reemplazar' if item.archivo else 'Subir' }}
                                </button>
                            </form>
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>

    <script>
        function imprimirPDF(url) {
            const iframe = document.createElement('iframe');
            iframe.style.display = 'none';
            iframe.src = url;
            document.body.appendChild(iframe);
            iframe.onload = function() {
                iframe.contentWindow.focus();
                iframe.contentWindow.print();
            };
        }
    </script>
</body>
</html>
"""

@app.route('/devo')
def devo_index():
    if 'email' not in session or ('devo' not in session.get('modulos', []) and 'admin' not in session.get('modulos', [])):
        return redirect(url_for('index_main'))
    query = request.args.get('q', '').lower()
    mensaje = request.args.get('mensaje', '')
    error = request.args.get('error', '')
    
    resultados = devoluciones_procesada
    if query:
        resultados = [i for i in resultados if query in i['cliente'].lower() or query in i['lote'].lower() or query in i['tela'].lower() or query in i['color'].lower() or query in i['ubicacion'].lower()]
        
    total_kilos_filtrados = sum(i['kilos'] for i in resultados)
    total_yardas_filtradas = sum(i['yardas'] for i in resultados)
    
    return render_template_string(HTML_TEMPLATE_DEVOLUCION, resultados=resultados, query=query, total_kilos_filtrados=total_kilos_filtrados, total_yardas_filtradas=total_yardas_filtradas, mensaje=mensaje, error=error)

@app.route('/devo/guardar-devolucion', methods=['POST'])
def devo_guardar_devolucion():
    archivo = request.files.get('archivo_devolucion')
    nombre_archivo = None
    
    if archivo and archivo.filename:
        nombre_archivo = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{archivo.filename}"
        archivo.save(os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo))
    
    nuevo_id = len(devoluciones_procesada) + 1
    devoluciones_procesada.append({
        'id': nuevo_id,
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
    return redirect(url_for('devo_index', mensaje="¡Devolución registrada con éxito!"))

@app.route('/devo/subir-archivo-devolucion/<int:dev_id>', methods=['POST'])
def devo_subir_archivo_devolucion(dev_id):
    archivo = request.files.get('archivo_devolucion')
    if archivo and archivo.filename:
        item = next((i for i in devoluciones_procesada if i['id'] == dev_id), None)
        if item:
            if item['archivo']:
                viejo_path = os.path.join(app.config['UPLOAD_FOLDER'], item['archivo'])
                if os.path.exists(viejo_path):
                    os.remove(viejo_path)
                    
            nombre_archivo = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{archivo.filename}"
            archivo.save(os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo))
            item['archivo'] = nombre_archivo
            return redirect(url_for('devo_index', mensaje="¡Archivo de respaldo actualizado y guardado correctamente!"))
    return redirect(url_for('devo_index', error="No se seleccionó ningún archivo válido."))

@app.route('/devo/ver-archivo/<filename>')
def devo_ver_archivo(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/devo/eliminar-archivo/<int:dev_id>')
def devo_eliminar_archivo(dev_id):
    item = next((i for i in devoluciones_procesada if i['id'] == dev_id), None)
    if item and item['archivo']:
        archivo_path = os.path.join(app.config['UPLOAD_FOLDER'], item['archivo'])
        if os.path.exists(archivo_path):
            os.remove(archivo_path)
        item['archivo'] = None
        return redirect(url_for('devo_index', mensaje="¡Archivo de respaldo eliminado correctamente!"))
    return redirect(url_for('devo_index', error="No se encontró el archivo a eliminar."))

@app.route('/devo/exportar-excel-devolucion')
def devo_exportar_excel_devolucion():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Devoluciones Procesada"
    ws.append(["Cliente", "Lote", "Tela", "Color", "Rollos", "Kilos", "Yardas", "Calidad", "Ubicación", "Archivo Respaldo"])
    for i in devoluciones_procesada:
        ws.append([i['cliente'], i['lote'], i['tela'], i['color'], i['rollos'], i['kilos'], i['yardas'], i['calidad'], i['ubicacion'], i['archivo'] or 'Sin archivo'])
    io_b = io.BytesIO()
    wb.save(io_b)
    io_b.seek(0)
    return send_file(io_b, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Reporte_Devoluciones_Tela_Procesada.xlsx')


if __name__ == '__main__':
    app.run(debug=True, port=5000)
