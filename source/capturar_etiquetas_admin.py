#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitor de etiquetas - EJECUTAR COMO ADMINISTRADOR.
Escucha trabajos de impresion y guarda cada etiqueta en etiquetas.json.

Uso:
    pip install pywin32
    python capturar_etiquetas_admin.py
"""

import sys
import os
import time
import datetime
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 no instalado. Ejecuta: pip install pywin32")
    sys.exit(1)


# ============================================================================
# CONFIGURACION - Edita estos valores según tu instalación
# ============================================================================

# Ruta a la carpeta HVETIQ CACHANPESCA (donde están RES00, art2_Sql.txt, etc.)
DATOS_DIR = Path(r"C:\HVETIQ CACHANPESCA")

# Ruta donde se guardarán las etiquetas.json
SALIDA_DIR = DATOS_DIR / "etiquetas_json"

# Nombre de la impresora a monitorizar (parcial, busca.contains())
IMPRESORA_NOMBRE = "Godex"

# ============================================================================

# Crear directorios
SALIDA_DIR.mkdir(parents=True, exist_ok=True)

ETIQUETAS_JSON = SALIDA_DIR / "etiquetas.json"
PROCESADOS_FILE = SALIDA_DIR / ".jobs_procesados.json"


def safe_print(msg):
    """Imprime sin errores de encoding."""
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        print(msg.encode("utf-8", errors="replace").decode("utf-8"), flush=True)


# ============================================================================
# RUTAS DE DATOS LOCALES
# ============================================================================

ART2_SQL = DATOS_DIR / "art2_Sql.txt"
ETIQREG_SQL = DATOS_DIR / "etiqreg_Sql.txt"
MESA_1_TXT = DATOS_DIR / "Mesa_1.txt"
CLIENTES_SQL = DATOS_DIR / "clientes_Sql.txt"


# ============================================================================
# CARGAR JOBS PROCESADOS
# ============================================================================

def cargar_jobs_procesados():
    try:
        if PROCESADOS_FILE.exists():
            with open(PROCESADOS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return set(data.get("jobs", []))
    except:
        pass
    return set()

def guardar_jobs_procesados(jobs_set):
    try:
        with open(PROCESADOS_FILE, 'w', encoding='utf-8') as f:
            json.dump({"jobs": list(jobs_set)}, f)
    except:
        pass


# ============================================================================
# UTILIDADES
# ============================================================================

def cargar_clientes_map(path):
    mapa = {}
    if not path.exists():
        return mapa
    try:
        lineas = path.read_text(encoding="latin-1", errors="replace").splitlines()
    except OSError:
        return mapa
    for ln in lineas:
        raw = ln.strip()
        if not raw or not raw[0].isdigit():
            continue
        row = [c.strip() for c in raw.rstrip(";").split("\t")]
        if len(row) < 2:
            continue
        try:
            codigo = int(row[0])
        except ValueError:
            continue
        nome = row[1].strip()
        if nome:
            mapa[codigo] = nome
    return mapa

def extraer_cliente_de_mesa():
    if not MESA_1_TXT.exists():
        return ""
    try:
        lineas = MESA_1_TXT.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return ""
    if not lineas:
        return ""
    header = lineas[0].strip()
    if not header.startswith("#"):
        return ""
    row = [c.strip() for c in header.split("|")]
    if len(row) <= 4:
        return ""
    try:
        codigo = int(row[4])
    except ValueError:
        return ""
    clientes_map = cargar_clientes_map(CLIENTES_SQL)
    return clientes_map.get(codigo, "")

def cargar_maestro_articulos():
    mapeo = {}
    if not ART2_SQL.exists():
        return mapeo
    try:
        lineas = ART2_SQL.read_text(encoding="latin-1", errors="replace").splitlines()
        for linea in lineas:
            if not linea.strip() or linea.startswith("art2") or linea.startswith("CREATE"):
                continue
            parts = linea.split("\t")
            if len(parts) < 14:
                continue
            codigo = parts[1].strip()
            if not codigo:
                continue
            mapeo[codigo] = {
                "nombre_cientifico": parts[13].strip(),
                "zona_captura": parts[8].strip(),
                "presentacion": parts[9].strip(),
                "arte_pesca": parts[16].strip() if len(parts) > 16 else "",
                "prod_nome": parts[6].strip() if len(parts) > 6 else "",
            }
    except:
        pass
    return mapeo

def _normalizar_especie_3(val):
    if val is None:
        return ""
    s = str(val).strip().upper()
    if len(s) == 3 and s.isalpha():
        return s
    return ""

def _maestro_por_codigo(maestro, codigo_plu):
    if not maestro or not codigo_plu:
        return None
    c = str(codigo_plu).strip()
    if c in maestro:
        return maestro[c]
    if c.isdigit():
        try:
            k = str(int(c, 10))
            if k in maestro:
                return maestro[k]
        except ValueError:
            pass
    return None

def cargar_etiqreg_ultimo():
    if not ETIQREG_SQL.exists():
        return None
    try:
        lineas = ETIQREG_SQL.read_text(encoding="latin-1", errors="replace").splitlines()
        hoy = datetime.date.today().isoformat()
        for linea in reversed(lineas):
            if not linea.strip() or linea.startswith("etiqreg") or linea.startswith("CREATE"):
                continue
            parts = linea.rstrip(";").split("\t")
            if len(parts) < 20:
                continue
            try:
                numero = int(parts[0].strip())
            except ValueError:
                continue
            fecha_str = parts[2].strip()
            if not fecha_str.startswith(hoy):
                continue
            out = {
                "numero": numero,
                "codigo": parts[1].strip(),
                "lote": parts[5].strip(),
                "neto": parts[10].strip(),
                "bruto": parts[9].strip(),
                "prod_nome": parts[6].strip() if len(parts) > 6 else "",
            }
            if len(parts) > 14:
                e3 = _normalizar_especie_3(parts[14])
                if e3:
                    out["codigo_alfa"] = e3
            return out
    except:
        pass
    return None


# ============================================================================
# PARSEAR RES
# ============================================================================

def parsear_res():
    res_files = []
    for i in range(20):
        res_file = DATOS_DIR / f"RES{i:02d}"
        if res_file.exists() and res_file.is_file():
            res_files.append((res_file.stat().st_mtime, res_file))

    if not res_files:
        return {}
    res_files.sort(key=lambda x: x[0], reverse=True)
    res_path = res_files[0][1]

    try:
        contenido = res_path.read_text(encoding="latin-1", errors="replace")
        root = ET.fromstring(contenido)
    except:
        return {}

    qr_texto = ""
    textos_items = []

    for item in root.findall(".//Item"):
        item_type = (item.findtext("Type") or "").strip().lower()
        if item_type == "codbar":
            codbar_txt = (item.findtext("codbar/txt") or "").strip()
            if codbar_txt:
                qr_texto = codbar_txt
        elif item_type == "text":
            txt = (item.findtext("txt") or "").strip()
            if txt:
                textos_items.append(txt)

    datos = {}

    if qr_texto:
        if m := re.search(r"comercial\s+(.+?)\s+cientifico", qr_texto, re.IGNORECASE):
            datos["producto"] = m.group(1).strip()
        if m := re.search(r"cientifico\s+(.+?)\s+Arte\s+pesca", qr_texto, re.IGNORECASE):
            datos["nombre_cientifico"] = m.group(1).strip()
        if m := re.search(r"Arte\s+pesca\s+(.+?)\s+zona\s+captura", qr_texto, re.IGNORECASE):
            datos["arte_pesca"] = m.group(1).strip()
        if m := re.search(r"zona\s+captura\s+(.+?)\s+Presentacion", qr_texto, re.IGNORECASE):
            datos["zona_captura"] = m.group(1).strip()
        if m := re.search(r"Presentacion\s+(.+?)\s+Producto", qr_texto, re.IGNORECASE):
            datos["presentacion"] = m.group(1).strip()
        if m := re.search(r"Producto\s+(.+?)(?:\s+\d{2}-\d{2}-\d{4}|\s*$)", qr_texto, re.IGNORECASE):
            datos["producto_tipo"] = m.group(1).strip()
        if m := re.search(r"(\d{2}-\d{2}-\d{4})\s+Peso:\s*([\d.,]+)", qr_texto, re.IGNORECASE):
            datos["fecha_expedicion"] = m.group(1)
            datos["peso_neto"] = float(m.group(2).replace(",", "."))
        elif m := re.search(r"Peso:\s*([\d.,]+)", qr_texto, re.IGNORECASE):
            datos["peso_neto"] = float(m.group(1).replace(",", "."))

    for i, txt in enumerate(textos_items):
        txt_upper = txt.upper()
        if txt_upper.startswith("LOTE:") and "lote" not in datos:
            datos["lote"] = txt.split(":", 1)[1].strip()
        elif txt_upper.startswith("BUQUE:") and "buque" not in datos:
            datos["buque"] = txt.split(":", 1)[1].strip()

    VALORES_METODO = ("CAPTURADO", "CRIA", "AGUA DULCE", "PECHE", "PECHE", "PÊCHE EXTRAVTIVE")
    if "metodo_produccion" not in datos:
        for txt in textos_items:
            txt_upper = txt.upper().strip()
            for vm in VALORES_METODO:
                if txt_upper == vm:
                    datos["metodo_produccion"] = txt.strip()
                    break
            if "metodo_produccion" in datos:
                break

    return datos


# ============================================================================
# EXTRAER PESO DEL EMF
# ============================================================================

def extraer_peso_del_emf(data):
    if not data or len(data) < 100:
        return None
    try:
        for i in range(len(data) - 10):
            if data[i:i+5] == b'Peso:':
                end = min(i + 5 + 20, len(data))
                fragmento = data[i+5:end]
                numeros = []
                for b in fragmento:
                    if 48 <= b <= 57 or b in (44, 46):
                        numeros.append(chr(b))
                    else:
                        if len(numeros) >= 2:
                            try:
                                valor = float(''.join(numeros).replace(',', '.'))
                                if 0.1 <= valor <= 500:
                                    return valor
                            except:
                                pass
                        numeros = []
    except:
        pass
    return None


# ============================================================================
# PROCESAR TRABAJO
# ============================================================================

def procesar_trabajo(job_id, doc_name, size, pages):
    safe_print(f">>> TRABAJO {job_id}: {doc_name}")

    spool_dir = Path(os.environ['WINDIR']) / 'System32' / 'spool' / 'PRINTERS'
    spl_file = spool_dir / f"FP{job_id:05d}.SPL"

    emf_data = None
    if spl_file.exists():
        try:
            with open(spl_file, 'rb') as f:
                emf_data = f.read()
        except:
            pass

    etiqueta = {
        "job_id": job_id,
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    }

    datos_res = parsear_res()
    reg_etiqreg = cargar_etiqreg_ultimo()
    maestro = cargar_maestro_articulos()

    datos = {}

    peso_emf = extraer_peso_del_emf(emf_data) if emf_data else None
    if peso_emf:
        datos["peso_neto"] = peso_emf

    for k, v in datos_res.items():
        if v:
            datos[k] = v

    if reg_etiqreg:
        if not datos.get("peso_neto") and reg_etiqreg.get("neto") and reg_etiqreg["neto"] != "0":
            try:
                datos["peso_neto"] = float(reg_etiqreg["neto"].replace(",", "."))
            except:
                pass
        if reg_etiqreg.get("lote") and not datos.get("lote"):
            datos["lote"] = reg_etiqreg["lote"]
        if reg_etiqreg.get("codigo"):
            datos["codigo_plu"] = reg_etiqreg["codigo"]
        if reg_etiqreg.get("codigo_alfa") and not datos.get("codigo_alfa"):
            datos["codigo_alfa"] = reg_etiqreg["codigo_alfa"]
        if reg_etiqreg.get("prod_nome"):
            datos["prod_nome"] = reg_etiqreg["prod_nome"]

    m = _maestro_por_codigo(maestro, datos.get("codigo_plu"))
    if m:
        for k, v in m.items():
            if v and not datos.get(k):
                datos[k] = v

    lote = datos.get("lote") or ""
    if len(lote) >= 6:
        e3 = _normalizar_especie_3(lote[-3:])
        if e3:
            datos["codigo_alfa"] = e3

    cliente = extraer_cliente_de_mesa()
    if cliente:
        datos["cliente"] = cliente

    etiqueta["datos"] = datos

    # Guardar en JSON acumulado
    try:
        if ETIQUETAS_JSON.exists():
            with open(ETIQUETAS_JSON, 'r', encoding='utf-8') as f:
                etiquetas = json.load(f)
        else:
            etiquetas = []
        etiquetas.append(etiqueta)
        with open(ETIQUETAS_JSON, 'w', encoding='utf-8') as f:
            json.dump(etiquetas, f, ensure_ascii=False, indent=2)
    except Exception as e:
        safe_print(f"    [WARN] Error guardando: {e}")

    safe_print(f"    [OK] cliente={datos.get('cliente','')}, producto={datos.get('producto','')}")

    return etiqueta


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 50)
    print("MONITOR ETIQUETAS CACHANPESCA")
    print("=" * 50)

    # Verificar permisos
    spool_dir = Path(os.environ['WINDIR']) / 'System32' / 'spool' / 'PRINTERS'
    try:
        test_file = spool_dir / "test.txt"
        with open(test_file, 'w') as f:
            f.write("test")
        test_file.unlink()
        print("[OK] Permisos de administrador OK")
    except:
        print("[ERROR] Necesita permisos de administrador")
        print("  Haz clic derecho > Ejecutar como administrador")
        return 1

    # Verificar que existen los directorios de datos
    if not DATOS_DIR.exists():
        print(f"[ERROR] No existe: {DATOS_DIR}")
        print("  Edita DATOS_DIR en el archivo para apuntar a tu carpeta HVETIQ")
        return 1

    print(f"\nDatos: {DATOS_DIR}")
    print(f"Salida: {ETIQUETAS_JSON}\n")

    # Cargar jobs ya procesados
    seen = cargar_jobs_procesados()
    safe_print(f"[INFO] Jobs procesados previamente: {len(seen)}")

    # WMI query
    wmi_init = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    ns_init = wmi_init.ConnectServer(".", "root\\cimv2")

    while True:
        try:
            jobs = ns_init.ExecQuery(
                f"SELECT * FROM Win32_PrintJob WHERE Name LIKE '%{IMPRESORA_NOMBRE}%' AND JobId > 0"
            )

            for job in jobs:
                try:
                    job_id = job.Properties_.Item('JobId').Value
                    doc = job.Properties_.Item('Document').Value or "Unknown"

                    if job_id in seen:
                        continue

                    spl_file = spool_dir / f"FP{job_id:05d}.SPL"
                    if spl_file.exists():
                        time.sleep(0.5)

                    seen.add(job_id)
                    guardar_jobs_procesados(seen)
                    procesar_trabajo(job_id, doc, 0, 0)
                except:
                    pass

            time.sleep(1)

        except KeyboardInterrupt:
            print("\nDetenido.")
            break
        except Exception as e:
            safe_print(f"[ERROR] {e}")
            time.sleep(2)

    return 0


if __name__ == "__main__":
    sys.exit(main())
