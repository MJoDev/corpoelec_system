"""Quick smoke test — run with: py -3.13 test_api.py"""
# Auth removed: single-user app, no login required.
import os, sys, tempfile

# Override paths BEFORE importing main so webview is never actually started
_td = tempfile.mkdtemp()

# Stub webview so we can test the API layer without launching a window
import types
wv = types.ModuleType("webview")
wv.OPEN_DIALOG = 10
class _FakeWindow:
    def create_file_dialog(self, *a, **k): return None
wv.windows = [_FakeWindow()]
def _create_window(*a, **k): return None
def _start(*a, **k): pass
wv.create_window = _create_window
wv.start = _start
sys.modules["webview"] = wv

# Now we can import main safely
import importlib.util, os
spec = importlib.util.spec_from_file_location(
    "main", os.path.join(os.path.dirname(__file__), "main.py"))
main = importlib.util.module_from_spec(spec)
spec.loader.exec_module(main)

# Redirect DB and assets to temp dir
main.DB_PATH = os.path.join(_td, "test.db")
main.ASSETS  = os.path.join(_td, "assets", "images")
main._init_db()

api = main.Api()

# ── Tests ──────────────────────────────────────────────────────────────────────
def ok(label, expr):
    assert expr, f"FAIL: {label}"
    print(f"  OK  {label}")

print("\n=== SGIP API smoke test ===\n")

# Ubicaciones
r = api.crear_ubicacion("Planta Norte", "Módulo A", "EST-01", "N-1")
ok("crear_ubicacion", r["ok"])
ub_id = r["id"]

ubs = api.get_ubicaciones()
ok("get_ubicaciones retorna lista", len(ubs) >= 1)

# Materiales
r = api.crear_material("TST-001", "Material prueba", "Desc técnica",
                        "Herramientas", "Bueno", 20, 5, ub_id)
ok("crear_material", r["ok"])
mat_id = r["id"]

r = api.crear_material("TST-001", "Dup", "", "Herramientas", "Bueno", 1, 1, ub_id)
ok("crear_material duplicado falla", not r["ok"])

mats = api.get_materiales()
ok("get_materiales", len(mats) == 1)
ok("stock_actual inicial", mats[0]["stock_actual"] == 20)
ok("bajo_minimo False inicial", not mats[0]["bajo_minimo"])

# Movimientos
r = api.registrar_movimiento(mat_id, "Salida", 18, "Juan Pérez", "OT-001")
ok("registrar salida", r["ok"])
ok("stock después salida = 2", r["nuevo_stock"] == 2)
ok("bajo_minimo detectado", r["bajo_minimo"] is True)

r = api.registrar_movimiento(mat_id, "Salida", 99, "Juan", "OT-002")
ok("salida con stock insuficiente falla", not r["ok"])

r = api.registrar_movimiento(mat_id, "Entrada", 10, "María", "OC-100")
ok("registrar entrada", r["ok"])
ok("stock después entrada = 12", r["nuevo_stock"] == 12)

ub2 = api.crear_ubicacion("Planta Norte", "Módulo B", "EST-02", "N-1")["id"]
r = api.registrar_movimiento(mat_id, "Transferencia", 0, "Carlos", "Reubicación",
                              ub_id, ub2)
ok("registrar transferencia", r["ok"])

mats2 = api.get_materiales()
ok("ubicacion actualizada por transferencia", mats2[0]["ubicacion_id"] == ub2)

movs = api.get_movimientos()
ok("get_movimientos retorna 3", len(movs) == 3)

movs_f = api.get_movimientos(desde="2020-01-01", hasta="2020-01-02")
ok("filtro por fecha devuelve vacío", len(movs_f) == 0)

# Editar
r = api.editar_material(mat_id, "Nombre Nuevo", "Nueva desc",
                         "Eléctricos", "Dañado", 5, 10, ub2)
ok("editar_material", r["ok"])
mats3 = api.get_materiales()
ok("nombre editado", mats3[0]["nombre"] == "Nombre Nuevo")
ok("estado editado", mats3[0]["estado"] == "Dañado")

# DB info
info = api.get_db_info()
ok("get_db_info", "mat_count" in info)
ok("mat_count correcto", info["mat_count"] == 1)
ok("mov_count correcto", info["mov_count"] == 3)

# Respaldo
r = api.respaldar_db()
ok("respaldar_db crea archivo", r["ok"] and os.path.exists(r["path"]))

# Eliminar material
r = api.eliminar_material(mat_id)
ok("eliminar_material", r["ok"])
ok("get_materiales vacío", len(api.get_materiales()) == 0)

print("\n=== Todos los tests pasaron ===\n")
