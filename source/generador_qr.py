#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Logica: filtrar etiquetas.json, generar QR y superponerlo en un PDF.
"""

from __future__ import annotations

import copy
import datetime
import io
import json
import os
import sys
import requests
from pathlib import Path
from typing import Any

import qrcode
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas as rl_canvas

# Texto de expedidor en el QR (mismo estilo que etiquetas comerciales)
EMPRESA_QR_DEFECTO = "CACHANPESCA S.L"

# Entre cada producto en el QR
SEPARADOR_PRODUCTOS_QR = "=========="


def _compactar_espacios(s: str) -> str:
    return " ".join((s or "").split())


def _formato_codigo_plu(codigo_plu: Any) -> str:
    if codigo_plu is None:
        return ""
    s = str(codigo_plu).strip()
    if not s:
        return ""
    if "." in s:
        return s
    if s.isdigit():
        return s.lstrip("0") or "0"
    return s


def _formato_peso(peso: Any) -> str:
    if peso is None:
        return ""
    try:
        f = float(peso)
    except (TypeError, ValueError):
        return str(peso).strip()
    if f == int(f):
        return str(int(f))
    return str(f).replace(".", ",")


def _normalizar_valor_fusion(v: Any) -> Any:
    if isinstance(v, float):
        if v == int(v):
            return int(v)
        return round(v, 10)
    return v


def _clave_mismo_producto(datos: dict[str, Any]) -> tuple[Any, ...]:
    return tuple(
        (k, _normalizar_valor_fusion(datos[k]))
        for k in sorted(datos.keys())
        if k != "peso_neto"
    )


def fusionar_etiquetas_mismo_producto(
    etiquetas: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if not etiquetas:
        return []

    grupos: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    orden_claves: list[tuple[Any, ...]] = []

    for et in etiquetas:
        if not isinstance(et, dict):
            continue
        d = et.get("datos")
        if not isinstance(d, dict):
            d = {}
        clave = _clave_mismo_producto(d)
        if clave not in grupos:
            orden_claves.append(clave)
            grupos[clave] = []
        grupos[clave].append(et)

    resultado: list[dict[str, Any]] = []
    for clave in orden_claves:
        grupo = grupos[clave]
        if len(grupo) == 1:
            resultado.append(grupo[0])
            continue

        base = copy.deepcopy(grupo[0])
        if not isinstance(base.get("datos"), dict):
            base["datos"] = {}

        total = 0.0
        hay_peso = False
        for et in grupo:
            raw = (et.get("datos") or {}).get("peso_neto")
            if raw is None or raw == "":
                continue
            try:
                total += float(raw)
                hay_peso = True
            except (TypeError, ValueError):
                pass

        if hay_peso:
            base["datos"]["peso_neto"] = total
        resultado.append(base)

    return resultado


def texto_qr_etiqueta(
    etiqueta: dict[str, Any],
    *,
    empresa: str | None = None,
    incluir_empresa: bool = True,
) -> str:
    d = etiqueta.get("datos") or {}
    cod = _formato_codigo_plu(d.get("codigo_plu"))
    prod = _compactar_espacios(str(d.get("producto") or ""))
    cient = _compactar_espacios(str(d.get("nombre_cientifico") or ""))
    esp3 = _compactar_espacios(str(d.get("codigo_alfa") or d.get("especie_3") or ""))
    arte = _compactar_espacios(str(d.get("arte_pesca") or ""))
    zona = _compactar_espacios(str(d.get("zona_captura") or ""))
    pres = _compactar_espacios(str(d.get("presentacion") or ""))
    tipo_raw = _compactar_espacios(str(d.get("producto_tipo") or ""))
    metodo = _compactar_espacios(str(d.get("metodo_produccion") or ""))
    fecha = str(d.get("fecha_expedicion") or "").strip()
    peso_s = _formato_peso(d.get("peso_neto"))
    lote = _compactar_espacios(str(d.get("lote") or ""))
    buque = _compactar_espacios(str(d.get("buque") or ""))

    partes = []
    if incluir_empresa:
        partes.append(f"CACHANPESCA S.L / {cod}/PO")
    if prod:
        partes.append(f"Producto: {prod}")
    if cient:
        partes.append(f"Nombre Científico: {cient}")
    if esp3:
        partes.append(f"Código Alfa: {esp3}")
    if zona:
        partes.append(f"Zona Captura: {zona}")
    if arte:
        partes.append(f"Arte Pesca: {arte}")
    if pres:
        partes.append(f"Presentación: {pres}")
    if tipo_raw:
        partes.append(f"Tipo Producto: {tipo_raw}")
    if metodo:
        partes.append(f"Método Producción: {metodo}")
    if fecha:
        partes.append(f"Fecha Expedición: {fecha}")
    if peso_s:
        partes.append(f"Peso: {peso_s} kg")
    if lote:
        partes.append(f"Lote: {lote}")
    if buque:
        partes.append(f"Buque: {buque}")

    return "\n".join(partes)


def texto_qr_varias_etiquetas(etiquetas: list[dict[str, Any]]) -> str:
    unidas = fusionar_etiquetas_mismo_producto(etiquetas)
    partes: list[str] = []
    # Header fijo una sola vez
    partes.append("CACHANPESCA S.L / 12.12598/PO")
    for e in unidas:
        partes.append(texto_qr_etiqueta(e, incluir_empresa=False))
    sep = f"\n{SEPARADOR_PRODUCTOS_QR}\n"
    return sep.join(partes)


def directorio_proyecto() -> Path:
    hv_dirs = [
        Path(r"C:\Users\Pablo\Downloads\rosi\HVETIQ CACHANPESCA"),
        Path(r"C:\Users\Pablo\Downloads\HVETIQ CACHANPESCA"),
        Path(r"C:\HVETIQ CACHANPESCA"),
    ]
    for candidate in hv_dirs:
        if candidate.exists() and (candidate / "clientes_Sql.txt").exists():
            return candidate
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def obtener_directorio_etiquetas() -> Path:
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
        base = Path(__file__).resolve().parent
    return base / "etiquetas_json"

REDE_ETIQUETAS_DIR = Path(r"\\PC-CAPTURA\CACHANPESCA_Installer\etiquetas_json")


def ruta_etiquetas_json() -> Path:
    return obtener_directorio_etiquetas() / "etiquetas.json"


def ruta_clientes_sql() -> Path:
    return directorio_proyecto() / "clientes_Sql.txt"


def cargar_nombres_clientes(path: Path | None = None) -> list[str]:
    path = path or ruta_clientes_sql()
    nombres: set[str] = set()
    if not path.exists():
        return []
    try:
        lineas = path.read_text(encoding="latin-1", errors="replace").splitlines()
    except OSError:
        return []
    for ln in lineas:
        raw = ln.strip()
        if not raw or not raw[0].isdigit():
            continue
        row = [c.strip() for c in raw.rstrip(";").split("\t")]
        if len(row) < 2:
            continue
        nome = row[1].strip()
        if nome:
            nombres.add(nome)
    return sorted(nombres, key=lambda s: s.upper())


def _fecha_etiqueta_dia(etiqueta: dict[str, Any]) -> datetime.date | None:
    ts = etiqueta.get("timestamp") or ""
    if isinstance(ts, str) and len(ts) >= 10:
        try:
            return datetime.date.fromisoformat(ts[:10])
        except ValueError:
            pass
    datos = etiqueta.get("datos") or {}
    fe = datos.get("fecha_expedicion")
    if isinstance(fe, str) and len(fe) >= 10:
        try:
            d, m, y = fe.split("-", 2)
            return datetime.date(int(y), int(m), int(d))
        except (ValueError, IndexError):
            pass
    return None


def cargar_etiquetas_filtradas(
    dia: datetime.date,
    cliente: str,
    json_path: Path | None = None,
) -> list[dict[str, Any]]:
    json_path = json_path or ruta_etiquetas_json()
    if not json_path.exists():
        return []
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            todas = json.load(f)
    except (json.JSONDecodeError, OSError):
        return []
    if not isinstance(todas, list):
        return []

    cli_norm = (cliente or "").strip().lower()
    resultado: list[dict[str, Any]] = []
    for item in todas:
        if not isinstance(item, dict):
            continue
        fd = _fecha_etiqueta_dia(item)
        if fd != dia:
            continue
        datos = item.get("datos") or {}
        c = (datos.get("cliente") or "").strip().lower()
        if c != cli_norm:
            continue
        resultado.append(item)
    return resultado


def texto_resumen_etiqueta(etiqueta: dict[str, Any]) -> str:
    datos = etiqueta.get("datos") or {}
    job = etiqueta.get("job_id", "")
    prod = datos.get("producto", "")
    lote = datos.get("lote", "")
    peso = datos.get("peso_neto", "")
    return f"#{job} | {prod} | Lote: {lote} | Peso: {peso}"


def generar_qr_png(
    etiquetas_seleccionadas: list[dict[str, Any]],
    output_png: Path,
    *,
    box_size: int = 5,
    border: int = 2,
) -> None:
    texto = texto_qr_varias_etiquetas(etiquetas_seleccionadas)

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=box_size,
        border=border,
    )
    qr.add_data(texto)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    img.save(output_png, format="PNG")


def superponer_qr_en_pdf(
    pdf_entrada: Path,
    qr_png: Path,
    pdf_salida: Path,
    *,
    pagina_indice: int = 0,
    tamano_qr_pt: float = 62.0,
    margen_superior_pt: float = 52.0,
    gap_texto_qr_pt: float = 5.0,
    x_tras_logo_pt: float | None = None,
    etiqueta_trazabilidad: str = "trazabilidad",
    fuente_etiqueta_pt: float = 7.0,
    color_etiqueta_rgb: tuple[float, float, float] = (0.45, 0.45, 0.45),
) -> None:
    reader = PdfReader(str(pdf_entrada))
    if not reader.pages:
        raise ValueError("El PDF no tiene paginas")

    n = len(reader.pages)
    idx = max(0, min(pagina_indice, n - 1))
    page = reader.pages[idx]

    mb = page.mediabox
    width = float(mb.width)
    height = float(mb.height)

    if x_tras_logo_pt is None:
        x_tras_logo_pt = max(
            210.0, min(width * 0.42, width - tamano_qr_pt - 28.0)
        )

    packet = io.BytesIO()
    c = rl_canvas.Canvas(packet, pagesize=(width, height))
    x = x_tras_logo_pt

    if etiqueta_trazabilidad.strip():
        texto_etq = etiqueta_trazabilidad.strip().upper()
        y_texto = height - margen_superior_pt
        c.setFont("Helvetica", fuente_etiqueta_pt)
        tw = stringWidth(texto_etq, "Helvetica", fuente_etiqueta_pt)
        x_centrado = x + max(0.0, (tamano_qr_pt - tw) / 2)
        r, g, b = color_etiqueta_rgb
        c.setFillColorRGB(r, g, b)
        c.drawString(x_centrado, y_texto, texto_etq)
        c.setFillColorRGB(0, 0, 0)
        y = y_texto - gap_texto_qr_pt - tamano_qr_pt
    else:
        y = height - margen_superior_pt - tamano_qr_pt
    c.drawImage(str(qr_png), x, y, width=tamano_qr_pt, height=tamano_qr_pt, mask="auto")
    c.save()
    packet.seek(0)
    overlay_reader = PdfReader(packet)
    overlay_page = overlay_reader.pages[0]
    page.merge_page(overlay_page)

    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)

    pdf_salida.parent.mkdir(parents=True, exist_ok=True)
    with open(pdf_salida, "wb") as f:
        writer.write(f)


def generar_txt_productos(etiquetas: list[dict[str, Any]], output_txt: Path) -> None:
    """Genera un archivo de texto con la lista de productos."""
    lineas = []
    lineas.append("=" * 60)
    lineas.append("CACHANPESCA S.L - Albaran de Productos")
    lineas.append(f"Fecha: {datetime.date.today().strftime('%d-%m-%Y')}")
    lineas.append("=" * 60)
    lineas.append("")

    for i, et in enumerate(etiquetas, 1):
        d = et.get("datos") or {}
        lineas.append(f"#{i}")
        lineas.append(f"  Producto:           {d.get('producto', '')}")
        lineas.append(f"  Nombre Cientifico: {d.get('nombre_cientifico', '')}")
        lineas.append(f"  Codigo Alfa:       {d.get('codigo_alfa') or d.get('especie_3', '')}")
        lineas.append(f"  Arte Pesca:        {d.get('arte_pesca', '')}")
        lineas.append(f"  Zona Captura:      {d.get('zona_captura', '')}")
        lineas.append(f"  Producto Tipo:     {d.get('producto_tipo', '')}")
        lineas.append(f"  Fecha Expedicion:  {d.get('fecha_expedicion', '')}")
        lineas.append(f"  Peso Neto (kg):    {d.get('peso_neto', '')}")
        lineas.append(f"  Metodo Produccion: {d.get('metodo_produccion', '')}")
        lineas.append(f"  Lote:              {d.get('lote', '')}")
        lineas.append(f"  Buque:             {d.get('buque', '')}")
        lineas.append(f"  Fecha Caducidad:   {d.get('fecha_caducidad', '')}")
        lineas.append(f"  Cliente:           {d.get('cliente', '')}")
        lineas.append("")

    output_txt.write_text("\n".join(lineas), encoding="utf-8")


def subir_a_wordpress(nombre_archivo: str, contenido: bytes) -> str:
    """Sube contenido a WordPress como entrada privada y devuelve la URL."""
    import uuid
    import json
    import sys
    import os

    # Buscar wp_config.json en varios sitios posibles
    posibles_paths = []

    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
        posibles_paths = [
            exe_dir / "wp_config.json",
            exe_dir.parent / "wp_config.json",
            exe_dir.parent.parent / "wp_config.json",
        ]
    else:
        posibles_paths = [
            Path(__file__).parent.parent / "wp_config.json",
            Path.cwd() / "wp_config.json",
        ]

    config_path = None
    for p in posibles_paths:
        if p.exists():
            config_path = p
            break

    if not config_path:
        raise Exception(f"No se encontro wp_config.json. Buscado en: {posibles_paths}")

    config = json.loads(config_path.read_text(encoding="utf-8"))

    wp_url = config["wp_url"].rstrip("/")
    wp_user = config["wp_user"]
    wp_pass = config["wp_pass"]

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/json",
    }

    # Crear post privado con UUID en el título para que la URL sea unpredictable
    post_uuid = uuid.uuid4().hex[:12]
    nombre_sin_ext = Path(nombre_archivo).stem
    titulo = f"CACHANPESCA – {nombre_sin_ext} – {post_uuid}"

    data = {
        "title": titulo,
        "content": contenido.decode("utf-8"),
        "status": "publish",
    }
    response = requests.post(
        f"{wp_url}/wp-json/wp/v2/posts",
        json=data,
        auth=(wp_user, wp_pass),
        headers=headers
    )
    if response.status_code in (200, 201):
        result = response.json()
        return result.get("link", f"{wp_url}/?p={result.get('id')}")
    else:
        raise Exception(f"Error WordPress: {response.status_code} - {response.text}")