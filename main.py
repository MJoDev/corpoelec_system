import webview
import sqlite3
import os
import sys
import shutil
from datetime import datetime


def _app_dir():
    """Directory next to the .exe (or script) — data persists here."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _static_dir():
    """Directory containing bundled HTML/JS assets."""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR   = _app_dir()
STATIC    = _static_dir()
DB_PATH   = os.path.join(APP_DIR, 'sgip_data.db')
ASSETS    = os.path.join(APP_DIR, 'assets', 'images')
HTML_PATH = os.path.join(STATIC, 'sgip.html')


# ─── Database helpers ─────────────────────────────────────────────────────────

def _conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    return c


def _init_db():
    os.makedirs(ASSETS, exist_ok=True)
    with _conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS Ubicaciones (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            planta_sede      TEXT NOT NULL,
            modulo_pasillo   TEXT NOT NULL,
            estanteria       TEXT NOT NULL,
            nivel_ubicacion  TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS Materiales (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo            TEXT UNIQUE NOT NULL,
            nombre            TEXT NOT NULL,
            descripcion_tecnica TEXT,
            categoria         TEXT,
            estado            TEXT DEFAULT 'Bueno'
                              CHECK(estado IN ('Bueno','Dañado','En Reparación')),
            stock_actual      INTEGER DEFAULT 0,
            stock_minimo      INTEGER DEFAULT 0,
            ruta_imagen       TEXT,
            ubicacion_id      INTEGER REFERENCES Ubicaciones(id)
        );
        CREATE TABLE IF NOT EXISTS Movimientos_Kardex (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            material_id          INTEGER REFERENCES Materiales(id),
            tipo_movimiento      TEXT
                                 CHECK(tipo_movimiento IN ('Entrada','Salida','Transferencia')),
            cantidad             INTEGER,
            fecha_hora           DATETIME DEFAULT CURRENT_TIMESTAMP,
            responsable_retiro   TEXT,
            orden_trabajo_area   TEXT,
            ubicacion_origen_id  INTEGER REFERENCES Ubicaciones(id),
            ubicacion_destino_id INTEGER REFERENCES Ubicaciones(id)
        );
        """)


# ─── PyWebView API class ───────────────────────────────────────────────────────

class Api:
    # ── Ubicaciones ───────────────────────────────────────────────────────────

    def get_ubicaciones(self):
        with _conn() as c:
            rows = c.execute("SELECT * FROM Ubicaciones ORDER BY planta_sede, modulo_pasillo, estanteria, nivel_ubicacion").fetchall()
        return [dict(r) for r in rows]

    def crear_ubicacion(self, planta_sede, modulo_pasillo, estanteria, nivel_ubicacion):
        with _conn() as c:
            cur = c.execute(
                "INSERT INTO Ubicaciones(planta_sede, modulo_pasillo, estanteria, nivel_ubicacion) VALUES(?,?,?,?)",
                (planta_sede, modulo_pasillo, estanteria, nivel_ubicacion))
            c.commit()
        return {"ok": True, "id": cur.lastrowid}

    def eliminar_ubicacion(self, ubicacion_id):
        with _conn() as c:
            in_use = c.execute(
                "SELECT COUNT(*) FROM Materiales WHERE ubicacion_id=?", (ubicacion_id,)).fetchone()[0]
        if in_use:
            return {"ok": False, "error": "Ubicación en uso por materiales existentes"}
        with _conn() as c:
            c.execute("DELETE FROM Ubicaciones WHERE id=?", (ubicacion_id,))
            c.commit()
        return {"ok": True}

    # ── Materiales ────────────────────────────────────────────────────────────

    def get_materiales(self):
        with _conn() as c:
            rows = c.execute("""
                SELECT m.*,
                       u.planta_sede, u.modulo_pasillo, u.estanteria, u.nivel_ubicacion
                FROM   Materiales m
                LEFT JOIN Ubicaciones u ON m.ubicacion_id = u.id
                ORDER  BY m.nombre
            """).fetchall()
        result = []
        for r in rows:
            d = dict(r)
            d['bajo_minimo'] = d['stock_actual'] <= d['stock_minimo']
            result.append(d)
        return result

    def crear_material(self, codigo, nombre, descripcion_tecnica, categoria, estado,
                       stock_actual, stock_minimo, ubicacion_id):
        with _conn() as c:
            try:
                cur = c.execute("""
                    INSERT INTO Materiales(codigo, nombre, descripcion_tecnica, categoria,
                                          estado, stock_actual, stock_minimo, ubicacion_id)
                    VALUES(?,?,?,?,?,?,?,?)
                """, (codigo, nombre, descripcion_tecnica, categoria,
                      estado, int(stock_actual), int(stock_minimo),
                      ubicacion_id if ubicacion_id else None))
                c.commit()
                return {"ok": True, "id": cur.lastrowid}
            except sqlite3.IntegrityError:
                return {"ok": False, "error": f"El código '{codigo}' ya existe"}

    def editar_material(self, material_id, nombre, descripcion_tecnica, categoria,
                        estado, stock_actual, stock_minimo, ubicacion_id):
        with _conn() as c:
            c.execute("""
                UPDATE Materiales
                SET    nombre=?, descripcion_tecnica=?, categoria=?,
                       estado=?, stock_actual=?, stock_minimo=?, ubicacion_id=?
                WHERE  id=?
            """, (nombre, descripcion_tecnica, categoria,
                  estado, int(stock_actual), int(stock_minimo),
                  ubicacion_id if ubicacion_id else None,
                  material_id))
            c.commit()
        return {"ok": True}

    def eliminar_material(self, material_id):
        with _conn() as c:
            c.execute("DELETE FROM Movimientos_Kardex WHERE material_id=?", (material_id,))
            c.execute("DELETE FROM Materiales WHERE id=?", (material_id,))
            c.commit()
        return {"ok": True}

    def subir_imagen(self, material_id, src_path):
        """Copy image file to assets dir, save relative path in DB."""
        if not os.path.isfile(src_path):
            return {"ok": False, "error": "Archivo no encontrado"}
        ext = os.path.splitext(src_path)[1].lower()
        fname = f"mat_{material_id}{ext}"
        dst = os.path.join(ASSETS, fname)
        shutil.copy2(src_path, dst)
        rel = os.path.join("assets", "images", fname).replace("\\", "/")
        with _conn() as c:
            c.execute("UPDATE Materiales SET ruta_imagen=? WHERE id=?", (rel, material_id))
            c.commit()
        return {"ok": True, "ruta": rel}

    def seleccionar_imagen(self):
        """Open a native file dialog and return chosen path."""
        types = ('Imágenes (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)',
                 'Todos los archivos (*.*)')
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False, file_types=types)
        if result:
            return {"ok": True, "path": result[0]}
        return {"ok": False}

    # ── Movimientos / Kárdex ──────────────────────────────────────────────────

    def get_movimientos(self, material_id=None, desde=None, hasta=None, tipo=None):
        where, params = [], []
        if material_id:
            where.append("mk.material_id = ?"); params.append(material_id)
        if tipo:
            where.append("mk.tipo_movimiento = ?"); params.append(tipo)
        if desde:
            where.append("mk.fecha_hora >= ?"); params.append(desde)
        if hasta:
            where.append("mk.fecha_hora <= ?"); params.append(hasta + " 23:59:59")
        clause = ("WHERE " + " AND ".join(where)) if where else ""
        with _conn() as c:
            rows = c.execute(f"""
                SELECT mk.*,
                       m.codigo, m.nombre,
                       uo.planta_sede || ' / ' || uo.modulo_pasillo || ' / ' || uo.estanteria || ' / ' || uo.nivel_ubicacion AS origen_txt,
                       ud.planta_sede || ' / ' || ud.modulo_pasillo || ' / ' || ud.estanteria || ' / ' || ud.nivel_ubicacion AS destino_txt
                FROM   Movimientos_Kardex mk
                JOIN   Materiales m ON mk.material_id = m.id
                LEFT JOIN Ubicaciones uo ON mk.ubicacion_origen_id = uo.id
                LEFT JOIN Ubicaciones ud ON mk.ubicacion_destino_id = ud.id
                {clause}
                ORDER  BY mk.fecha_hora DESC
            """, params).fetchall()
        return [dict(r) for r in rows]

    def registrar_movimiento(self, material_id, tipo_movimiento, cantidad,
                             responsable_retiro, orden_trabajo_area,
                             ubicacion_origen_id=None, ubicacion_destino_id=None):
        cantidad = int(cantidad)
        material_id = int(material_id)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with _conn() as c:
            mat = c.execute(
                "SELECT stock_actual, ubicacion_id FROM Materiales WHERE id=?",
                (material_id,)).fetchone()
            if not mat:
                return {"ok": False, "error": "Material no encontrado"}

            stock = mat["stock_actual"]
            ub_actual = mat["ubicacion_id"]

            if tipo_movimiento == "Entrada":
                nuevo_stock = stock + cantidad
                c.execute("UPDATE Materiales SET stock_actual=? WHERE id=?",
                          (nuevo_stock, material_id))
            elif tipo_movimiento == "Salida":
                if cantidad > stock:
                    return {"ok": False, "error": f"Stock insuficiente (disponible: {stock})"}
                nuevo_stock = stock - cantidad
                c.execute("UPDATE Materiales SET stock_actual=? WHERE id=?",
                          (nuevo_stock, material_id))
            elif tipo_movimiento == "Transferencia":
                if not ubicacion_destino_id:
                    return {"ok": False, "error": "Se requiere ubicación destino"}
                c.execute("UPDATE Materiales SET ubicacion_id=? WHERE id=?",
                          (int(ubicacion_destino_id), material_id))
                nuevo_stock = stock
                ubicacion_origen_id = ubicacion_origen_id or ub_actual
            else:
                return {"ok": False, "error": "Tipo de movimiento inválido"}

            orig = int(ubicacion_origen_id) if ubicacion_origen_id else None
            dest = int(ubicacion_destino_id) if ubicacion_destino_id else None

            cur = c.execute("""
                INSERT INTO Movimientos_Kardex
                    (material_id, tipo_movimiento, cantidad, fecha_hora,
                     responsable_retiro, orden_trabajo_area,
                     ubicacion_origen_id, ubicacion_destino_id)
                VALUES(?,?,?,?,?,?,?,?)
            """, (material_id, tipo_movimiento, cantidad, fecha,
                  responsable_retiro, orden_trabajo_area, orig, dest))
            c.commit()

        with _conn() as c:
            m2 = c.execute("SELECT stock_minimo FROM Materiales WHERE id=?",
                           (material_id,)).fetchone()
        bajo = nuevo_stock <= m2["stock_minimo"] if m2 else False

        return {"ok": True, "id": cur.lastrowid, "nuevo_stock": nuevo_stock, "bajo_minimo": bajo}

    # ── Exportaciones ─────────────────────────────────────────────────────────

    def exportar_excel(self, tipo, desde=None, hasta=None):
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment
        except ImportError:
            return {"ok": False, "error": "openpyxl no instalado"}

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"SGIP_{tipo}_{ts}.xlsx"
        dest = os.path.join(APP_DIR, fname)

        wb = openpyxl.Workbook()
        ws = wb.active
        hdr_font = Font(bold=True, color="FFFFFF")
        hdr_fill = PatternFill("solid", fgColor="2D4A8A")

        if tipo == "inventario":
            ws.title = "Existencias"
            headers = ["Código", "Nombre", "Categoría", "Estado",
                       "Stock Actual", "Stock Mínimo", "Ubicación"]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = hdr_font; cell.fill = hdr_fill
            mats = self.get_materiales()
            for row, m in enumerate(mats, 2):
                ub = " / ".join(filter(None, [
                    m.get("planta_sede"), m.get("modulo_pasillo"),
                    m.get("estanteria"), m.get("nivel_ubicacion")]))
                ws.append([m["codigo"], m["nombre"], m["categoria"], m["estado"],
                           m["stock_actual"], m["stock_minimo"], ub])
        else:  # kardex
            ws.title = "Kárdex"
            headers = ["ID", "Fecha/Hora", "Tipo", "Código", "Material",
                       "Cantidad", "Responsable", "OT/Área", "Origen", "Destino"]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = hdr_font; cell.fill = hdr_fill
            movs = self.get_movimientos(desde=desde, hasta=hasta)
            for m in movs:
                ws.append([m["id"], m["fecha_hora"], m["tipo_movimiento"],
                           m["codigo"], m["nombre"], m["cantidad"],
                           m["responsable_retiro"], m["orden_trabajo_area"],
                           m.get("origen_txt", ""), m.get("destino_txt", "")])

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18
        wb.save(dest)
        return {"ok": True, "path": dest}

    def exportar_pdf(self, tipo, desde=None, hasta=None):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib import colors
        except ImportError:
            return {"ok": False, "error": "reportlab no instalado"}

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"SGIP_{tipo}_{ts}.pdf"
        dest = os.path.join(APP_DIR, fname)

        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(dest, pagesize=landscape(A4),
                                leftMargin=30, rightMargin=30,
                                topMargin=30, bottomMargin=30)
        story = []

        if tipo == "inventario":
            story.append(Paragraph("SGIP — Reporte de Existencias", styles["Title"]))
            story.append(Paragraph(datetime.now().strftime("%d/%m/%Y %H:%M"), styles["Normal"]))
            story.append(Spacer(1, 12))
            headers = [["Código", "Nombre", "Categoría", "Estado", "Stock", "Mínimo", "Ubicación"]]
            mats = self.get_materiales()
            data = headers + [
                [m["codigo"], m["nombre"][:30], m["categoria"], m["estado"],
                 str(m["stock_actual"]), str(m["stock_minimo"]),
                 f"{m.get('modulo_pasillo','')}/{m.get('estanteria','')}"]
                for m in mats
            ]
        else:
            story.append(Paragraph("SGIP — Kárdex de Movimientos", styles["Title"]))
            story.append(Paragraph(datetime.now().strftime("%d/%m/%Y %H:%M"), styles["Normal"]))
            story.append(Spacer(1, 12))
            headers = [["ID", "Fecha", "Tipo", "Código", "Material", "Cant.", "Responsable"]]
            movs = self.get_movimientos(desde=desde, hasta=hasta)
            data = headers + [
                [str(m["id"]), m["fecha_hora"][:16], m["tipo_movimiento"],
                 m["codigo"], m["nombre"][:25], str(m["cantidad"]),
                 m["responsable_retiro"] or ""]
                for m in movs
            ]

        tbl = Table(data)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2D4A8A")),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F0F4FF")]),
            ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
            ("ALIGN",      (0, 0), (-1, -1), "LEFT"),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(tbl)
        doc.build(story)
        return {"ok": True, "path": dest}

    # ── Config / Info ─────────────────────────────────────────────────────────

    def get_db_info(self):
        size = os.path.getsize(DB_PATH) if os.path.exists(DB_PATH) else 0
        img_count = len([f for f in os.listdir(ASSETS)
                         if os.path.isfile(os.path.join(ASSETS, f))]) if os.path.exists(ASSETS) else 0
        img_size = sum(os.path.getsize(os.path.join(ASSETS, f))
                       for f in os.listdir(ASSETS)
                       if os.path.isfile(os.path.join(ASSETS, f))) if os.path.exists(ASSETS) else 0
        with _conn() as c:
            mat_count = c.execute("SELECT COUNT(*) FROM Materiales").fetchone()[0]
            mov_count = c.execute("SELECT COUNT(*) FROM Movimientos_Kardex").fetchone()[0]
        return {
            "db_path": DB_PATH,
            "db_size_kb": round(size / 1024, 1),
            "assets_path": ASSETS,
            "img_count": img_count,
            "img_size_mb": round(img_size / 1024 / 1024, 1),
            "mat_count": mat_count,
            "mov_count": mov_count,
            "version": "1.0.0",
        }

    def respaldar_db(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = os.path.join(APP_DIR, f"sgip_backup_{ts}.db")
        shutil.copy2(DB_PATH, dest)
        return {"ok": True, "path": dest}

    def seleccionar_excel(self):
        """Open a native file dialog and return chosen .xlsx path."""
        types = ('Archivos Excel (*.xlsx;*.xls)', 'Todos los archivos (*.*)')
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False, file_types=types)
        if result:
            return {"ok": True, "path": result[0]}
        return {"ok": False}

    def importar_excel(self, path):
        """Import materials from an Excel file in the standard inventory format.

        Expected columns (row 2 = headers, data from row 3):
          A: CODIGO  D: NOMBRE  F: DESCRIPCION  H: CANTIDAD
          I: USADO   J: DAÑADO  K: NUEVO
          L: ESTANTERIA  N: COLUMNA  O: MODULO
        """
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            importados = 0
            omitidos = 0
            errores = []
            ub_cache = {}

            for row in ws.iter_rows(min_row=3, values_only=True):
                codigo = row[0]
                if not codigo:
                    continue
                codigo = str(codigo).strip()
                nombre = str(row[3] or "").strip() or codigo
                descripcion = str(row[5] or "").strip()
                cantidad = int(row[7] or 0)
                danado = row[9]
                estado = "Dañado" if (danado and int(danado) > 0) else "Bueno"

                def _loc(val, prefix, default):
                    if val is None:
                        return default
                    try:
                        return f"{prefix} {int(val)}"
                    except (ValueError, TypeError):
                        return str(val).strip() or default

                modulo     = _loc(row[14], "Módulo",  "Módulo 1")
                estanteria = _loc(row[11], "Est.",    "Est. 1")
                columna    = _loc(row[13], "Col.",    "Col. 1")
                planta    = "Planta Principal"

                ub_key = (planta, modulo, estanteria, columna)
                if ub_key not in ub_cache:
                    with _conn() as c:
                        r = c.execute(
                            "SELECT id FROM Ubicaciones WHERE planta_sede=? AND modulo_pasillo=? AND estanteria=? AND nivel_ubicacion=?",
                            ub_key).fetchone()
                        if r:
                            ub_cache[ub_key] = r["id"]
                        else:
                            c.execute(
                                "INSERT INTO Ubicaciones(planta_sede, modulo_pasillo, estanteria, nivel_ubicacion) VALUES(?,?,?,?)",
                                ub_key)
                            c.commit()
                            ub_cache[ub_key] = c.execute(
                                "SELECT id FROM Ubicaciones WHERE planta_sede=? AND modulo_pasillo=? AND estanteria=? AND nivel_ubicacion=?",
                                ub_key).fetchone()["id"]

                ub_id = ub_cache[ub_key]
                try:
                    with _conn() as c:
                        c.execute(
                            "INSERT INTO Materiales(codigo, nombre, descripcion_tecnica, categoria, estado, stock_actual, stock_minimo, ubicacion_id) VALUES(?,?,?,?,?,?,?,?)",
                            (codigo, nombre, descripcion, "General", estado, cantidad, 0, ub_id))
                        c.commit()
                    importados += 1
                except Exception as e:
                    omitidos += 1
                    if len(errores) < 5:
                        errores.append(f"{codigo}: {e}")

            return {"ok": True, "importados": importados, "omitidos": omitidos, "errores": errores}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def abrir_carpeta_exportacion(self):
        import subprocess
        subprocess.Popen(f'explorer "{APP_DIR}"')
        return {"ok": True}


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _init_db()
    api = Api()
    window = webview.create_window(
        title="SGIP — Sistema de Gestión de Inventario para Plantas",
        url=HTML_PATH,
        js_api=api,
        width=1440,
        height=900,
        min_size=(1024, 680),
        resizable=True,
    )
    webview.start(debug=False)
