"""Smoke test para la API Flask — ejecutar con: py test_api.py"""
import os
import sys
import json
import tempfile

_td = tempfile.mkdtemp()

# Apuntar a BD temporal antes de importar main
os.environ['SGIP_TEST'] = '1'

import importlib.util
spec = importlib.util.spec_from_file_location(
    "main", os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py"))
main_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(main_mod)
main_mod.DB_PATH = os.path.join(_td, "test.db")
main_mod.ASSETS  = os.path.join(_td, "assets", "images")
main_mod._init_db()

client = main_mod.app.test_client()

def call(url, data=None, method='GET'):
    if method == 'POST':
        resp = client.post(url, data=json.dumps(data), content_type='application/json')
    else:
        resp = client.get(url)
    return json.loads(resp.data.decode('utf-8'))

def ok(label, expr):
    assert expr, "FAIL: " + label
    print("  OK  " + label)

print("\n=== SGIP Flask API smoke test ===\n")

# ── Ubicaciones ────────────────────────────────────────────────────────────────
r = call('/api/crear_ubicacion', {"planta_sede":"Planta Norte","modulo_pasillo":"Modulo A","estanteria":"EST-01","nivel_ubicacion":"N-1"}, 'POST')
ok("crear_ubicacion", r["ok"])
ub_id = r["id"]

ubs = call('/api/get_ubicaciones')
ok("get_ubicaciones retorna lista", len(ubs) >= 1)

# ── Materiales ─────────────────────────────────────────────────────────────────
r = call('/api/crear_material', {"codigo":"TST-001","nombre":"Material prueba","descripcion_tecnica":"Desc","categoria":"Herramientas","stock_nuevo":20,"stock_usado":0,"stock_danado":0,"stock_minimo":5,"ubicacion_id":ub_id}, 'POST')
ok("crear_material", r["ok"])
mat_id = r["id"]

r = call('/api/crear_material', {"codigo":"TST-001","nombre":"Dup","descripcion_tecnica":"","categoria":"H","stock_nuevo":1,"stock_usado":0,"stock_danado":0,"stock_minimo":1,"ubicacion_id":ub_id}, 'POST')
ok("crear_material duplicado falla", not r["ok"])

mats = call('/api/get_materiales')
ok("get_materiales devuelve 1", len(mats) == 1)
ok("stock_actual inicial = 20", mats[0]["stock_actual"] == 20)
ok("stock_nuevo inicial = 20", mats[0]["stock_nuevo"] == 20)
ok("bajo_minimo False", not mats[0]["bajo_minimo"])

# ── Movimientos ────────────────────────────────────────────────────────────────
r = call('/api/registrar_movimiento', {"material_id":mat_id,"tipo_movimiento":"Salida","cantidad":18,"responsable_retiro":"Juan","orden_trabajo_area":"OT-001","condicion":"Nuevo"}, 'POST')
ok("registrar salida", r["ok"])
ok("stock despues salida = 2", r["nuevo_stock"] == 2)
ok("bajo_minimo detectado", r["bajo_minimo"] is True)

r = call('/api/registrar_movimiento', {"material_id":mat_id,"tipo_movimiento":"Salida","cantidad":99,"responsable_retiro":"Juan","orden_trabajo_area":"OT-002","condicion":"Nuevo"}, 'POST')
ok("salida con stock insuficiente falla", not r["ok"])

r = call('/api/registrar_movimiento', {"material_id":mat_id,"tipo_movimiento":"Entrada","cantidad":10,"responsable_retiro":"Maria","orden_trabajo_area":"OC-100","condicion":"Nuevo"}, 'POST')
ok("registrar entrada", r["ok"])
ok("stock despues entrada = 12", r["nuevo_stock"] == 12)

r2 = call('/api/crear_ubicacion', {"planta_sede":"Planta Norte","modulo_pasillo":"Modulo B","estanteria":"EST-02","nivel_ubicacion":"N-1"}, 'POST')
ub2 = r2["id"]
r = call('/api/registrar_movimiento', {"material_id":mat_id,"tipo_movimiento":"Transferencia","cantidad":0,"responsable_retiro":"Carlos","orden_trabajo_area":"Reubicacion","ubicacion_destino_id":ub2}, 'POST')
ok("registrar transferencia", r["ok"])

mats2 = call('/api/get_materiales')
ok("ubicacion actualizada por transferencia", mats2[0]["ubicacion_id"] == ub2)

movs = call('/api/get_movimientos')
ok("get_movimientos retorna 3", len(movs) == 3)

movs_f = call('/api/get_movimientos?desde=2020-01-01&hasta=2020-01-02')
ok("filtro por fecha devuelve vacio", len(movs_f) == 0)

# ── Editar ─────────────────────────────────────────────────────────────────────
r = call('/api/editar_material', {"material_id":mat_id,"nombre":"Nombre Nuevo","descripcion_tecnica":"Nueva desc","categoria":"Electricos","stock_nuevo":0,"stock_usado":0,"stock_danado":5,"stock_minimo":10,"ubicacion_id":ub2}, 'POST')
ok("editar_material", r["ok"])
mats3 = call('/api/get_materiales')
ok("nombre editado", mats3[0]["nombre"] == "Nombre Nuevo")
ok("estado editado = Danado", mats3[0]["estado"] == "Dañado")
ok("stock_danado editado = 5", mats3[0]["stock_danado"] == 5)

# ── Info / Respaldo ────────────────────────────────────────────────────────────
info = call('/api/get_db_info')
ok("get_db_info tiene mat_count", "mat_count" in info)
ok("mat_count correcto = 1", info["mat_count"] == 1)
ok("mov_count correcto = 3", info["mov_count"] == 3)

r = call('/api/respaldar_db', {}, 'POST')
ok("respaldar_db crea archivo", r["ok"] and os.path.exists(r["path"]))

# ── Eliminar ───────────────────────────────────────────────────────────────────
r = call('/api/eliminar_material', {"material_id":mat_id}, 'POST')
ok("eliminar_material", r["ok"])
ok("get_materiales vacio", len(call('/api/get_materiales')) == 0)

print("\n=== Todos los tests pasaron ===\n")
