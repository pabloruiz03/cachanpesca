#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
App de escritorio: albarán PDF + filtro día/cliente + QR desde etiquetas.json.
CACHANPESCA - Diseño profesional con identidad corporativa.
"""

from __future__ import annotations

import datetime
import io
import os
import tempfile
import tkinter as tk
import qrcode
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None  # type: ignore[misc, assignment]

from generador_qr import (
    cargar_etiquetas_filtradas,
    cargar_nombres_clientes,
    generar_qr_png,
    superponer_qr_en_pdf,
    texto_resumen_etiqueta,
    generar_txt_productos,
    subir_a_wordpress,
)


# =============================================================================
# ESTILOS CORPORATIVOS CACHANPESCA
# =============================================================================
AZUL_CORPORATIVO = "#1a3a5c"
AZUL_CLARO = "#2d5a87"
AZUL_MUY_CLARO = "#e8f0f8"
BLANCO = "#ffffff"
GRIS_CLARO = "#f5f7fa"
GRIS_MEDIO = "#6b7c93"
GRIS_OSCURO = "#3d4f5f"
VERDE_OK = "#28a745"
ROJO_ERROR = "#dc3545"
DORADO = "#d4a84b"


class EstiloCorporativo:
    """Configuración de estilos visuales corporativos."""

    @staticmethod
    def configurar_ttk(root: tk.Tk) -> None:
        style = ttk.Style(root)
        root.style = style

        style.theme_use("clam")

        style.configure(
            "Header.TFrame",
            background=AZUL_CORPORATIVO,
        )

        style.configure(
            "Body.TFrame",
            background=BLANCO,
        )

        style.configure(
            "Footer.TFrame",
            background=GRIS_CLARO,
        )

        style.configure(
            "Logo.TLabel",
            background=AZUL_CORPORATIVO,
            foreground=BLANCO,
            font=("Segoe UI", 9, "bold"),
        )

        style.configure(
            "TituloApp.TLabel",
            background=AZUL_CORPORATIVO,
            foreground=BLANCO,
            font=("Segoe UI", 18, "bold"),
        )

        style.configure(
            "Subtitulo.TLabel",
            background=AZUL_CORPORATIVO,
            foreground=AZUL_MUY_CLARO,
            font=("Segoe UI", 10),
        )

        style.configure(
            "Seccion.TLabelframe",
            background=BLANCO,
            foreground=AZUL_CORPORATIVO,
            font=("Segoe UI", 11, "bold"),
        )

        style.configure(
            "Seccion.TLabelframe.Label",
            background=BLANCO,
            foreground=AZUL_CORPORATIVO,
            font=("Segoe UI", 11, "bold"),
        )

        style.configure(
            "Campo.TLabel",
            background=BLANCO,
            foreground=GRIS_OSCURO,
            font=("Segoe UI", 10),
        )

        style.configure(
            "CampoObligatorio.TLabel",
            background=BLANCO,
            foreground=ROJO_ERROR,
            font=("Segoe UI", 9),
        )

        style.configure(
            "BotonPrimario.TButton",
            background=AZUL_CLARO,
            foreground=BLANCO,
            font=("Segoe UI", 10, "bold"),
            padding=(20, 10),
        )

        style.map(
            "BotonPrimario.TButton",
            background=[("active", AZUL_CORPORATIVO), ("pressed", AZUL_CORPORATIVO)],
            foreground=[("active", BLANCO)],
        )

        style.configure(
            "BotonSecundario.TButton",
            background=GRIS_CLARO,
            foreground=GRIS_OSCURO,
            font=("Segoe UI", 10),
            padding=(15, 8),
        )

        style.map(
            "BotonSecundario.TButton",
            background=[("active", GRIS_MEDIO), ("pressed", GRIS_MEDIO)],
        )

        style.configure(
            "BotonExcel.TButton",
            background=VERDE_OK,
            foreground=BLANCO,
            font=("Segoe UI", 10, "bold"),
            padding=(20, 10),
        )

        style.map(
            "BotonExcel.TButton",
            background=[("active", "#1e7e34"), ("pressed", "#1e7e34")],
        )

        style.configure(
            "Card.TFrame",
            background=GRIS_CLARO,
            relief="flat",
        )

        style.configure(
            "Linea.TSeparator",
            background=AZUL_MUY_CLARO,
        )


# =============================================================================
# WIDGETS PERSONALIZADOS
# =============================================================================

class HeaderCorporativo(ttk.Frame):
    """Header con logo y título de la aplicación."""

    def __init__(self, parent: tk.Widget, logo_path: Path | None = None) -> None:
        super().__init__(parent, style="Header.TFrame", height=100)
        self.pack_propagate(False)
        self._logo_image = None

        contenido = ttk.Frame(self, style="Header.TFrame")
        contenido.pack(fill=tk.BOTH, expand=True, padx=20, pady=12)

        if logo_path and logo_path.exists():
            self._cargar_logo(contenido, logo_path)

        textos = ttk.Frame(contenido, style="Header.TFrame")
        textos.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(
            textos,
            text="CACHANPESCA",
            style="TituloApp.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            textos,
            text="Generador de QR para Albaranes",
            style="Subtitulo.TLabel",
        ).pack(anchor="w", pady=(2, 0))

    def _cargar_logo(self, parent: tk.Widget, logo_path: Path) -> None:
        try:
            from PIL import Image, ImageTk
            img = Image.open(logo_path)
            img = img.convert("RGBA")
            target_h = 70
            ratio = target_h / img.height
            new_w = int(img.width * ratio)
            img = img.resize((new_w, target_h), Image.LANCZOS)
            self._logo_image = ImageTk.PhotoImage(img)
            lbl_logo = tk.Label(
                parent,
                image=self._logo_image,
                background=AZUL_CORPORATIVO,
            )
            lbl_logo.pack(side=tk.LEFT, padx=(0, 15))
        except Exception:
            try:
                self._logo_image = tk.PhotoImage(file=str(logo_path))
                self._logo_image = self._logo_image.subsample(2, 2)
                lbl_logo = tk.Label(
                    parent,
                    image=self._logo_image,
                    background=AZUL_CORPORATIVO,
                )
                lbl_logo.pack(side=tk.LEFT, padx=(0, 15))
            except Exception:
                pass


class CardSeccion(ttk.Labelframe):
    """Card con sección visually destacada."""

    def __init__(
        self,
        parent: tk.Widget,
        titulo: str,
        icono: str = "",
        **kwargs,
    ) -> None:
        super().__init__(
            parent,
            text=f"  {icono} {titulo}" if icono else f"  {titulo}",
            style="Seccion.TLabelframe",
            **kwargs,
        )


class LineaSeparadora(ttk.Frame):
    """Línea separadora horizontal."""

    def __init__(self, parent: tk.Widget, **kwargs) -> None:
        super().__init__(parent, height=2, **kwargs)
        sep = ttk.Separator(self, orient=tk.HORIZONTAL, style="Linea.TSeparator")
        sep.pack(fill=tk.X, expand=True)


# =============================================================================
# APP PRINCIPAL
# =============================================================================

class AppQRAlbaran(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self._pdf_path: Path | None = None
        self._lbl_pdf = tk.StringVar(value="Ningún archivo seleccionado")

        EstiloCorporativo.configurar_ttk(self)

        self.title("CACHANPESCA - Generador QR Albarán")
        self.minsize(580, 420)

        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True)

        self._construir_header(container)
        self._construir_cuerpo(container)
        self._construir_footer(container)

        self._recargar_clientes()

    def _construir_header(self, parent: ttk.Frame) -> None:
        logo_path = Path("cachan_logotipo-final - etiqueta.bmp")
        if not logo_path.exists():
            logo_path = Path(__file__).parent / "cachan_logotipo-final - etiqueta.bmp"

        header = HeaderCorporativo(parent, logo_path=logo_path)
        header.configure(height=90)
        header.pack(fill=tk.X, padx=0, pady=0)

    def _construir_cuerpo(self, parent: ttk.Frame) -> None:
        cuerpo = ttk.Frame(parent, style="Body.TFrame")
        cuerpo.pack(fill=tk.BOTH, expand=True, padx=25, pady=15)
        cuerpo.grid_columnconfigure(0, weight=1)

        self._seccion_pdf(cuerpo)
        LineaSeparadora(cuerpo).pack(fill=tk.X, pady=15)
        self._seccion_filtros(cuerpo)
        LineaSeparadora(cuerpo).pack(fill=tk.X, pady=15)
        self._seccion_botones(cuerpo)

    def _seccion_pdf(self, parent: ttk.Frame) -> None:
        card = CardSeccion(parent, "1. Documento PDF (Albarán)", width=500)
        card.pack(fill=tk.X, pady=(0, 10))

        interno = ttk.Frame(card, padding=(10, 5))
        interno.pack(fill=tk.X)

        ttk.Button(
            interno,
            text="📄  Seleccionar archivo PDF...",
            style="BotonSecundario.TButton",
            command=self._elegir_pdf,
        ).pack(side=tk.LEFT, padx=(5, 15))

        ttk.Label(
            interno,
            textvariable=self._lbl_pdf,
            style="Campo.TLabel",
            wraplength=350,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _seccion_filtros(self, parent: ttk.Frame) -> None:
        card = CardSeccion(parent, "2. Filtros de Búsqueda", width=500)
        card.pack(fill=tk.X, pady=(0, 10))

        interno = ttk.Frame(card, padding=(10, 8))
        interno.pack(fill=tk.X)
        interno.grid_columnconfigure(1, weight=1)
        interno.grid_columnconfigure(3, weight=1)

        ttk.Label(
            interno,
            text="Cliente:",
            style="Campo.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=(5, 10))

        self._combo_cliente = ttk.Combobox(
            interno,
            width=35,
            state="readonly",
            font=("Segoe UI", 10),
        )
        self._combo_cliente.grid(row=0, column=1, sticky="ew", padx=(0, 25))

        ttk.Label(
            interno,
            text="Fecha:",
            style="Campo.TLabel",
        ).grid(row=0, column=2, sticky="w", padx=(0, 10))

        self._fecha_cal = self._crear_dateentry(interno)
        self._fecha_cal.grid(row=0, column=3, sticky="w", padx=(0, 5))

        ttk.Label(
            interno,
            text="⚠️ Importante: seleccione PRIMERO el cliente y DESPUÉS la fecha",
            style="CampoObligatorio.TLabel",
            font=("Segoe UI", 8),
        ).grid(row=1, column=0, columnspan=4, sticky="w", padx=5, pady=(8, 0))

    def _crear_dateentry(self, parent) -> Any:
        if DateEntry is None:
            messagebox.showwarning(
                "Dependencia",
                "Instale tkcalendar: pip install tkcalendar\n"
                "Se usará un campo de texto AAAA-MM-DD.",
            )
            entry = ttk.Entry(parent, width=12, font=("Segoe UI", 10))
            entry.insert(0, datetime.date.today().isoformat())
            return entry

        cal = DateEntry(
            parent,
            width=12,
            background=AZUL_CLARO,
            foreground=BLANCO,
            borderwidth=0,
            date_pattern="yyyy-mm-dd",
            font=("Segoe UI", 10),
        )
        return cal

    def _seccion_botones(self, parent: ttk.Frame) -> None:
        card = CardSeccion(parent, "3. Acciones", width=500)
        card.pack(fill=tk.X)

        interno = ttk.Frame(card, padding=(10, 8))
        interno.pack(fill=tk.X)

        btn_continuar = ttk.Button(
            interno,
            text="▶  Continuar: Elegir etiquetas y generar QR",
            style="BotonExcel.TButton",
            command=self._continuar,
        )
        btn_continuar.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(
            interno,
            text="* Debe seleccionar un PDF, fecha y cliente para continuar",
            style="CampoObligatorio.TLabel",
        ).pack(anchor="w", pady=(8, 0))

    def _construir_footer(self, parent: ttk.Frame) -> None:
        footer = ttk.Frame(parent, style="Footer.TFrame", height=35)
        footer.pack(fill=tk.X, padx=0, pady=0)

        ttk.Label(
            footer,
            text="CACHANPESCA S.L.  •  Sistema de gestión de etiquetas con QR",
            background=GRIS_CLARO,
            foreground=GRIS_MEDIO,
            font=("Segoe UI", 8),
        ).pack(side=tk.RIGHT, padx=15, pady=8)

    def _fecha_seleccionada(self):
        if DateEntry is None or not hasattr(self._fecha_cal, "get_date"):
            s = self._fecha_cal.get().strip()
            try:
                return datetime.date.fromisoformat(s)
            except ValueError:
                return None
        return self._fecha_cal.get_date()

    def _elegir_pdf(self) -> None:
        p = filedialog.askopenfilename(
            title="Albarán PDF - CACHANPESCA",
            filetypes=[
                ("PDF", "*.pdf"),
                ("Todos", "*.*"),
            ],
        )
        if p:
            self._pdf_path = Path(p)
            nombre = self._pdf_path.name
            self._lbl_pdf.set(nombre)

    def _recargar_clientes(self) -> None:
        nombres = cargar_nombres_clientes()
        self._combo_cliente["values"] = nombres
        if nombres:
            self._combo_cliente.current(0)

    def _continuar(self) -> None:
        if not self._pdf_path or not self._pdf_path.is_file():
            messagebox.showerror("Error", "Seleccione un archivo PDF.")
            return

        dia = self._fecha_seleccionada()
        if dia is None:
            messagebox.showerror("Error", "Fecha no válida. Use AAAA-MM-DD.")
            return

        cliente = self._combo_cliente.get().strip()
        if not cliente:
            messagebox.showerror("Error", "Seleccione un cliente.")
            return

        filtradas = cargar_etiquetas_filtradas(dia, cliente)
        if not filtradas:
            messagebox.showinfo(
                "Sin resultados",
                f"No hay etiquetas para:\n\n📅 Día: {dia}\n👤 Cliente: {cliente}",
            )
            return

        dlg = DialogoEtiquetas(self, filtradas)
        self.wait_window(dlg)
        if not dlg.aceptado or not dlg.etiquetas_elegidas:
            return

        # Generar TXT con todos los productos
        try:
            with tempfile.TemporaryDirectory() as td:
                txt_productos = Path(td) / "albaran_productos.txt"
                generar_txt_productos(dlg.etiquetas_elegidas, txt_productos)

                # Subir a WordPress
                gist_url = subir_a_wordpress(
                    f"albaran_{datetime.date.today().strftime('%Y%m%d_%H%M%S')}.txt",
                    txt_productos.read_bytes()
                )

                # Generar QR con link al Gist
                qr_png = Path(td) / "qr.png"

                # Crear QR con la URL del Gist
                qr = qrcode.QRCode(version=10, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=5, border=2)
                qr.add_data(gist_url)
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
                img.save(qr_png, format="PNG")

                # Guardar PDF final con QR superpuesto
                initial_dir = ""
                if self._pdf_path and self._pdf_path.parent.exists():
                    initial_dir = str(self._pdf_path.parent)

                salida = filedialog.asksaveasfilename(
                    title="Guardar PDF con QR - CACHANPESCA",
                    defaultextension=".pdf",
                    filetypes=[("PDF", "*.pdf")],
                    initialfile=self._pdf_path.stem + "_con_qr.pdf",
                    initialdir=initial_dir if initial_dir else None,
                )
                if not salida:
                    return

                out_pdf = Path(salida)
                superponer_qr_en_pdf(self._pdf_path, qr_png, out_pdf)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el PDF:\n{e}")
            return

        messagebox.showinfo("Listo", f"✅ PDF generado correctamente:\n\n{out_pdf}\n\n📱 Escanea el QR para ver el TXT completo con todos los productos.")


# =============================================================================
# DIALOGO DE ETIQUETAS
# =============================================================================

class DialogoEtiquetas(tk.Toplevel):
    """Muestra todas las etiquetas filtradas con checkboxes."""

    def __init__(self, master: tk.Misc, etiquetas: list[dict[str, Any]]) -> None:
        super().__init__(master)
        self.title("Seleccionar Etiquetas - CACHANPESCA")
        self.transient(master)
        self.grab_set()
        self.geometry("600x500")
        self.minsize(500, 400)

        self.aceptado = False
        self.etiquetas_elegidas: list[dict[str, Any]] = []
        self._vars: list[tuple[tk.BooleanVar, dict[str, Any]]] = []

        self._construir_ui(etiquetas)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

    def _construir_ui(self, etiquetas: list[dict[str, Any]]) -> None:
        # Header
        header = ttk.Frame(self, style="Header.TFrame", height=55)
        header.pack(fill=tk.X, padx=0, pady=0)
        header.pack_propagate(False)

        ttk.Label(
            header,
            text="Seleccionar Etiquetas para el QR",
            style="TituloApp.TLabel",
        ).pack(side=tk.LEFT, padx=20, pady=15)

        # Cuerpo principal
        cuerpo = ttk.Frame(self, padding=15)
        cuerpo.pack(fill=tk.BOTH, expand=True)

        # Instrucciones
        ttk.Label(
            cuerpo,
            text="Marca las etiquetas que quieres incluir en el QR",
            style="Campo.TLabel",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(0, 10))

        # Barra de botones
        barra = ttk.Frame(cuerpo)
        barra.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            barra,
            text="Marcar todas",
            command=self._marcar_todas,
            style="Accent.TButton",
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(
            barra,
            text="Desmarcar todas",
            command=self._desmarcar_todas,
        ).pack(side=tk.LEFT)

        # Lista con scroll
        frame_lista = ttk.Frame(cuerpo, style="Card.TFrame", padding=5)
        frame_lista.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        canvas = tk.Canvas(
            frame_lista,
            background=GRIS_CLARO,
            highlightthickness=0,
        )
        scrollbar = ttk.Scrollbar(frame_lista, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.configure(yscrollcommand=scrollbar.set)

        inner = ttk.Frame(canvas, padding=5)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_window, width=event.width)

        inner.bind("<Configure>", on_configure)
        canvas.bind("<Configure>", on_configure)

        for et in etiquetas:
            var = tk.BooleanVar(value=True)
            self._vars.append((var, et))

            texto = texto_resumen_etiqueta(et)
            cb = ttk.Checkbutton(
                inner,
                text=texto,
                variable=var,
            )
            cb.pack(anchor="w", pady=2, padx=5, fill=tk.X)

        # Botones inferiores
        btn_frame = ttk.Frame(self, style="Footer.TFrame", height=55)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=0)
        btn_frame.pack_propagate(False)

        botones = ttk.Frame(btn_frame)
        botones.pack(side=tk.RIGHT, padx=20, pady=12)

        ttk.Button(
            botones,
            text="Cancelar",
            command=self._cancelar,
        ).pack(side=tk.RIGHT, padx=(10, 0))

        ttk.Button(
            botones,
            text="Generar QR",
            command=self._aceptar,
            style="Accent.TButton",
        ).pack(side=tk.RIGHT)

    def _marcar_todas(self) -> None:
        for var, _ in self._vars:
            var.set(True)

    def _desmarcar_todas(self) -> None:
        for var, _ in self._vars:
            var.set(False)

    def _aceptar(self) -> None:
        elegidas = [et for var, et in self._vars if var.get()]
        if not elegidas:
            messagebox.showwarning(
                "Atencion",
                "Debe seleccionar al menos una etiqueta.",
                parent=self,
            )
            return
        self.etiquetas_elegidas = elegidas
        self.aceptado = True
        self.destroy()

    def _cancelar(self) -> None:
        self.aceptado = False
        self.destroy()


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    app = AppQRAlbaran()
    app.mainloop()


if __name__ == "__main__":
    main()
