# SGIP — Sistema de Gestión de Inventario para Plantas

Sistema de inventario de materiales y activos para **Corpoelec**, diseñado para funcionar de forma completamente **offline** desde una memoria USB sin instalación.

**Stack:** Python 3.13 · PyWebView · SQLite · React 18 · PyInstaller

---

## Características

- Inventario de materiales con código, nombre, descripción técnica, categoría y estado
- Control de stock con alertas de mínimo
- Registro de movimientos: Entradas, Salidas y Transferencias
- Kárdex por material con historial completo
- Consultas y búsqueda por categoría, estado y ubicación
- Importación desde Excel (formato estándar Corpoelec)
- Exportación a Excel y PDF
- Modo oscuro / claro con persistencia
- Base de datos SQLite local — sin servidor, sin internet
- Portable: corre directamente desde pendrive

---

## Requisitos de desarrollo

- Python 3.13 ([python.org](https://python.org))
- Las dependencias se instalan automáticamente al correr `run_dev.bat`

```
pywebview >= 5.0
openpyxl  >= 3.1
reportlab >= 4.0
```

---

## Ejecutar en modo desarrollo

```bat
run_dev.bat
```

Esto instala dependencias si no están presentes y arranca la aplicación directamente desde el código fuente.

---

## Compilar ejecutable portable

```bat
build.bat
```

Genera `dist\SGIP_Portable\SGIP.exe`. Copiar esa carpeta al pendrive es todo lo necesario.

Al primer arranque se crean automáticamente en la misma carpeta:
- `sgip_data.db` — base de datos SQLite
- `assets\images\` — carpeta para imágenes de materiales

---

## Correr pruebas

```bat
py -3.13 test_api.py
```

Pruebas de humo que cubren toda la capa de API (CRUD materiales, ubicaciones, movimientos, stock, exportación) sin abrir ventana.

---

## Estructura del proyecto

```
corpoelec_system/
├── main.py                  # Backend Python — API expuesta a JS via PyWebView
├── sgip.html                # Frontend React (JSX compilado en browser con Babel)
├── static/                  # Librerías JS vendorizadas (React, ReactDOM, Babel)
├── requirements.txt         # Dependencias Python
├── build.bat                # Script de compilación con PyInstaller
├── run_dev.bat              # Script de ejecución en desarrollo
├── test_api.py              # Pruebas de humo de la API
├── Inventario.xlsx          # Archivo de ejemplo para importación
└── logo-servicio-corpoelec.im1761941416671im.avif  # Logo Corpoelec
```

---

## Formato de importación Excel

El botón **Importar Excel** en Configuración acepta archivos con el siguiente formato (datos desde fila 3):

| Col | Campo         | Descripción                        |
|-----|---------------|------------------------------------|
| A   | CODIGO        | Código único del material          |
| D   | NOMBRE        | Nombre del producto                |
| F   | DESCRIPCIÓN   | Descripción técnica                |
| H   | CANTIDAD      | Stock total                        |
| J   | DAÑADO        | Cantidad dañada (determina estado) |
| L   | ESTANTERIA    | Número de estantería               |
| N   | COLUMNA       | Columna/nivel                      |
| O   | MODULO        | Módulo del almacén                 |

Los materiales con código ya existente en la base de datos se omiten sin error.

---

## Desarrollado por

**Ing. Moises Jimenez** — Corpoelec
