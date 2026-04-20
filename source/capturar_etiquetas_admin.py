#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitor de etiquetas - EJECUTAR COMO ADMINISTRADOR.
Escucha trabajos de impresion y guarda cada etiqueta en etiquetas.json.
"""

import sys
import os
import time
import datetime
import json
import hashlib
import re
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("[ERROR] pywin32 no instalado")
    sys.exit(1)


def obtener_directorio_datos():
    """
    Obtiene el directorio donde están los archivos de datos (RES, SQL, etc).
    Siempre devuelve la carpeta HVETIQ CACHANPESCA, sin importar dónde esté el .exe.
    """
    hv_dirs = [
        Path(r"C:\Users\Pablo\Downloads\rosi\HVETIQ CACHANPESCA"),
        Path(r"C:\Users\Pablo\Downloads\HVETIQ CACHANPESCA"),
        Path(r"C:\HVETIQ CACHANPESCA"),
    ]
    # Buscar la carpeta que contenga los archivos de datos
    for candidate in hv_dirs:
        if candidate.exists() and (candidate / "RES00").exists():
            return candidate
    # Si no encuentra, fallback al directorio del exe
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent.resolve()


# Directorio de datos (siempre la carpeta HVETIQ con archivos RES, SQL, etc)
PROGRAMA_DIR = obtener_directorio_datos()

# ============================================================================
# RUTAS DE ETIQUETAS - Configurable para red local
# ============================================================================
# Para usar en RED: comparte la carpeta etiquetas_json en el PC 1 y pon
# la ruta UNC aquí en el PC 2.
# Formato: r"\\NOMBRE_PC\HVETIQ CACHANPESCA\etiquetas_json"
#
# El programa busca la carpeta HVETIQ CACHANPESCA en varios sitios posibles.
# Para modo red, cambiar REDE_ETIQUETAS_DIR.

def obtener_directorio_etiquetas():
    """Busca CACHANPESCA_Installer y devuelve etiquetas_json dentro."""
    installer_dirs = [
        Path(r"C:\Users\Pablo\Downloads\CACHANPESCA_Installer"),
        Path(r"C:\Users\Pablo\CACHANPESCA_Installer"),
        Path(r"C:\CACHANPESCA_Installer"),
    ]
    for candidate in installer_dirs:
        if candidate.exists():
            return candidate / "etiquetas_json"
    # Fallback: usar el directorio del exe
    if getattr(sys, 'frozen', False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent.resolve()
    return base / "etiquetas_json"

REDE_ETIQUETAS_DIR = Path(r"\\PC-CAPTURA\CACHANPESCA_Installer\etiquetas_json")

SALIDA_DIR = obtener_directorio_etiquetas()
SALIDA_DIR.mkdir(parents=True, exist_ok=True)

ETIQUETAS_JSON = SALIDA_DIR / "etiquetas.json"
PROCESADOS_FILE = SALIDA_DIR / ".jobs_procesados.json"

def cargar_jobs_procesados():
    """Carga el set de jobs ya procesados desde archivo."""
    try:
        if PROCESADOS_FILE.exists():
            with open(PROCESADOS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return set(data.get("jobs", []))
    except:
        pass
    return set()

def guardar_jobs_procesados(jobs_set):
    """Guarda el set de jobs procesados en archivo."""
    try:
        with open(PROCESADOS_FILE, 'w', encoding='utf-8') as f:
            json.dump({"jobs": list(jobs_set)}, f)
    except:
        pass

# Bases de datos del programa (relative to PROGRAMA_DIR)
ART2_SQL = PROGRAMA_DIR / "art2_Sql.txt"
ETIQREG_SQL = PROGRAMA_DIR / "etiqreg_Sql.txt"
MESA_1_TXT = PROGRAMA_DIR / "Mesa_1.txt"
CLIENTES_SQL = PROGRAMA_DIR / "clientes_Sql.txt"


def safe_print(msg, *, flush=True):
    """Imprime en consola; flush=True para ver líneas al momento en cmd (sin esperar Ctrl+C)."""
    try:
        print(msg, flush=flush)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(msg.encode(enc, errors="replace").decode(enc), flush=flush)


def configurar_salida_consola_en_vivo() -> None:
    """Evita que cmd acumule texto en búfer (salida aparente solo al interrumpir)."""
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(line_buffering=True)
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(line_buffering=True)
    except (OSError, ValueError, AttributeError):
        pass


def guardar_en_json_acumulado(etiqueta):
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


# ─── Cliente desde Mesa_1.txt ────────────────────────────────────────────────

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


def _maestro_por_codigo(maestro, codigo_plu):
    """Resuelve fila art2: mismo string, o sin ceros a la izquierda (00051 → 51)."""
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


def _normalizar_especie_3(val):
    """
    Código de especie de 3 letras (p. ej. FAO alpha-3 interno: BRF, HKE, MNZ).
    Solo acepta exactamente tres letras A-Z.
    """
    if val is None:
        return ""
    s = str(val).strip().upper()
    if len(s) == 3 and s.isalpha():
        return s
    return ""


def _buscar_especie_por_producto(maestro, producto, nombre_cientifico):
    """Busca código de 3 letras en art2 por nombre comercial o científico."""
    if not maestro:
        return ""
    prod_upper = (producto or "").upper()
    cient_upper = (nombre_cientifico or "").upper()
    for row in maestro.values():
        esp = row.get("especie") or ""
        nom_ci = row.get("nombre_cientifico") or ""
        if esp and len(esp) == 3 and esp.isalpha():
            if prod_upper in nom_ci.upper() or nom_ci.upper() in prod_upper:
                return esp
    return ""


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


# ─── Maestro de productos ────────────────────────────────────────────────────

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
            esp3 = ""
            if len(parts) > 12:
                esp3 = _normalizar_especie_3(parts[12])
            mapeo[codigo] = {
                "nombre_cientifico": parts[13].strip(),
                "zona_captura": parts[8].strip(),
                "presentacion": parts[9].strip(),
                "arte_pesca": parts[16].strip() if len(parts) > 16 else "",
                "codigo_alfa": esp3,
                "prod_nome": parts[6].strip() if len(parts) > 6 else "",
            }
    except:
        pass
    return mapeo


# ─── etiqreg (solo para peso y lote de hoy) ────────────────────────────────

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
            # etiqreg: txt4 (índice 14) = siglas de especie, p. ej. BRF, HKE, MNZ
            if len(parts) > 14:
                e3 = _normalizar_especie_3(parts[14])
                if e3:
                    out["codigo_alfa"] = e3
            return out
    except:
        pass
    return None


# ─── RES (datos del QR) ─────────────────────────────────────────────────────

def parsear_res():
    res_files = []
    for i in range(20):
        res_file = PROGRAMA_DIR / f"RES{i:02d}"
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
        elif txt_upper.startswith("F.CADUCIDAD:") or txt_upper.startswith("FECHA CADUCIDAD:"):
            if "fecha_caducidad" not in datos:
                # Buscar fecha en este item o en los siguientes
                found = False
                for j in range(i, min(i+3, len(textos_items))):
                    for parte in textos_items[j].split():
                        if re.match(r"\d{2}-\d{2}-\d{4}", parte):
                            datos["fecha_caducidad"] = parte
                            found = True
                            break
                    if found:
                        break

    VALORES_METODO_PRODUCCION = ("CAPTURADO", "CRIA", "AGUA DULCE", "PECHE", "PECHE", "PÊCHE EXTRAVTIVE")
    if "metodo_produccion" not in datos:
        for txt in textos_items:
            txt_upper = txt.upper().strip()
            for vm in VALORES_METODO_PRODUCCION:
                if txt_upper == vm:
                    datos["metodo_produccion"] = txt.strip()
                    break
            if "metodo_produccion" in datos:
                break

    return datos


# ─── Extraer peso del EMF ────────────────────────────────────────────────────

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


# ─── Procesar ───────────────────────────────────────────────────────────────

def procesar_trabajo(job_id, doc_name, size, pages):
    safe_print(f"\n>>> TRABAJO {job_id}")

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

    # Fuentes de datos
    datos_res = parsear_res()
    reg_etiqreg = cargar_etiqreg_ultimo()
    maestro = cargar_maestro_articulos()

    datos = {}

    # Peso del EMF (prioridad 1)
    peso_emf = extraer_peso_del_emf(emf_data) if emf_data else None
    if peso_emf:
        datos["peso_neto"] = peso_emf

    # Datos del QR/RES
    for k, v in datos_res.items():
        if v:
            datos[k] = v

    # Datos de etiqreg (solo para enriquecer)
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

    # Maestro de articulos
    m = _maestro_por_codigo(maestro, datos.get("codigo_plu"))
    if m:
        for k, v in m.items():
            if v and not datos.get(k):
                datos[k] = v

    # Siempre: especie_3 desde las últimas 3 letras del lote (anula cualquier valor previo)
    lote = datos.get("lote") or ""
    e3 = ""
    if len(lote) >= 6:
        e3 = _normalizar_especie_3(lote[-3:])
    if e3:
        datos["codigo_alfa"] = e3
    elif not datos.get("codigo_alfa"):
        # Lote sin especie (solo fecha como 180426), buscar en art2 por producto
        especie_from_maestro = _buscar_especie_por_producto(maestro, datos.get("producto"), datos.get("nombre_cientifico"))
        if especie_from_maestro:
            datos["codigo_alfa"] = especie_from_maestro

    # Cliente desde Mesa_1.txt (prioridad)
    cliente = extraer_cliente_de_mesa()
    if cliente:
        datos["cliente"] = cliente

    etiqueta["datos"] = datos

    guardar_en_json_acumulado(etiqueta)
    safe_print(
        f"    [OK] cliente={datos.get('cliente','')}, producto={datos.get('producto','')}, "
        f"especie_3={datos.get('especie_3','')}"
    )

    return etiqueta


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    configurar_salida_consola_en_vivo()

    print("=" * 50, flush=True)
    print("MONITOR ETIQUETAS CACHANPESCA", flush=True)
    print("=" * 50, flush=True)

    spool_dir = Path(os.environ['WINDIR']) / 'System32' / 'spool' / 'PRINTERS'
    try:
        test_file = spool_dir / "test.txt"
        with open(test_file, 'w') as f:
            f.write("test")
        test_file.unlink()
        print("[OK] Permisos OK", flush=True)
    except:
        print("[ERROR] Sin permisos de administrador", flush=True)
        return 1

    print(f"\nSalida: {ETIQUETAS_JSON}\n", flush=True)

    # Ignorar trabajos existentes en cola
    wmi_init = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    ns_init = wmi_init.ConnectServer(".", "root\\cimv2")
    existing = ns_init.ExecQuery("SELECT * FROM Win32_PrintJob WHERE Name LIKE '%Godex%' AND JobId > 0")
    existing_ids = {j.Properties_.Item('JobId').Value for j in existing}
    safe_print(f"[INFO] Ignorando {len(existing_ids)} trabajos en cola\n")

    # Cargar jobs ya procesados desde archivo (evita duplicados en reinicios)
    seen = cargar_jobs_procesados()
    seen.update(existing_ids)
    safe_print(f"[INFO] Jobs procesados previamente: {len(seen)}\n")

    while True:
        try:
            wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            ns = wmi.ConnectServer(".", "root\\cimv2")
            jobs = ns.ExecQuery("SELECT * FROM Win32_PrintJob WHERE Name LIKE '%Godex%' AND JobId > 0")

            for job in jobs:
                try:
                    job_id = job.Properties_.Item('JobId').Value
                    doc = job.Properties_.Item('Document').Value or "Unknown"
                    size = job.Properties_.Item('Size').Value or 0
                    pages = job.Properties_.Item('TotalPages').Value or 0

                    # Ignorar si ya fue procesado este job_id
                    if job_id in seen:
                        continue

                    # Verificar si el archivo SPL ya fue procesado
                    spl_file = spool_dir / f"FP{job_id:05d}.SPL"
                    if spl_file.exists():
                        # Esperar a que termine de escribirse (mover el archivo indica impresion completa)
                        time.sleep(0.5)

                    seen.add(job_id)
                    guardar_jobs_procesados(seen)
                    procesar_trabajo(job_id, doc, size, pages)
                except:
                    pass

            time.sleep(1)

        except KeyboardInterrupt:
            print("\n\nDetenido.", flush=True)
            break
        except Exception as e:
            safe_print(f"[ERROR] {e}")
            time.sleep(2)

    return 0


if __name__ == "__main__":
    sys.exit(main())
