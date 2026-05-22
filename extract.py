#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract.py — Extractor de datos para el buscador de la Armada Española del siglo XVIII.

Qué hace:
  1. Lee los cinco Excels (BUQUE, PERSONAL, LUGAR, GRADO, lat) que mantiene el
     investigador.
  2. Normaliza nombres de columnas, formatos de fecha (dd/mm/yyyy → yyyy-mm-dd)
     y elimina filas vacías o sin claves mínimas.
  3. Produce un único dataset JSON con la estructura que espera el template HTML
     del buscador.
  4. Si se le indica una plantilla HTML, inyecta el JSON dentro del
     <script id="dataset">...</script> y escribe el index.html resultante.
  5. Registra en incidencias.txt cualquier dato sospechoso (fecha mal formada,
     hoja faltante, columna ausente) sin abortar.

Uso:
  python extract.py                # busca los Excels y la plantilla en el directorio actual
  python extract.py --excels DIR   # busca los Excels en DIR
  python extract.py --template PATH/template.html --out PATH/index.html

Diseñado para ser robusto: cualquier fallo se registra en incidencias.txt en
lenguaje llano, en español, sin abortar la ejecución. Esto importa porque al
correr en GitHub Actions no hay nadie mirando la consola.
"""

from __future__ import annotations
import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# Helpers de limpieza y conversión
# ─────────────────────────────────────────────────────────────────────────────

INCIDENCIAS: List[str] = []


def aviso(msg: str) -> None:
    """Añade un aviso al registro de incidencias y lo imprime."""
    INCIDENCIAS.append(msg)
    print(f"  · {msg}")


def limpia(v: Any) -> Optional[str]:
    """Devuelve un string sin espacios laterales, o None si está vacío."""
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    return s


def fecha_iso(v: Any) -> Optional[str]:
    """Convierte una fecha (texto dd/mm/yyyy, ISO, o datetime de Excel) a 'yyyy-mm-dd'.

    Si no se puede interpretar, devuelve el texto original (para no perder dato)
    y registra una incidencia.
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y",
                "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    # Caso común en datos del XVIII: solo se conoce el año
    if s.isdigit() and len(s) == 4 and 1500 <= int(s) <= 2100:
        return s  # se conserva como "yyyy", sigue siendo ordenable
    aviso(f"Fecha mal formada, conservada como texto: {s!r}")
    return s


def to_int(v: Any) -> Optional[int]:
    """Convierte a entero si se puede, o None."""
    s = limpia(v)
    if s is None:
        return None
    try:
        return int(float(s.replace(",", ".")))
    except (ValueError, TypeError):
        return None


def to_float(v: Any) -> Optional[float]:
    """Convierte a float si se puede, o None."""
    s = limpia(v)
    if s is None:
        return None
    try:
        return float(s.replace(",", "."))
    except (ValueError, TypeError):
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Lectura tolerante de hojas y filas
# ─────────────────────────────────────────────────────────────────────────────

def carga_excel(path: Path) -> Dict[str, pd.DataFrame]:
    """Lee un Excel entero como dict {nombre_hoja: DataFrame}. Si falla, devuelve {}."""
    if not path.exists():
        aviso(f"Archivo no encontrado: {path.name}")
        return {}
    try:
        return pd.read_excel(path, sheet_name=None, dtype=str)
    except Exception as exc:
        aviso(f"No se pudo leer {path.name}: {exc}")
        return {}


def hoja(xls: Dict[str, pd.DataFrame], archivo: str, *nombres: str) -> pd.DataFrame:
    """Devuelve la primera hoja de `xls` cuyo nombre coincida con alguno de los dados.

    Acepta varios alias porque el usuario ha mantenido distintos nombres a lo
    largo del tiempo (singular vs plural, mayúsculas, etc.).
    """
    for n in nombres:
        if n in xls:
            return xls[n]
    aviso(f"Hoja no encontrada en {archivo}: probadas {nombres}")
    return pd.DataFrame()


def filas(df: pd.DataFrame, mapeo: Dict[str, tuple], *, clave_obligatoria: str = "id") -> List[Dict]:
    """Aplica un mapeo {clave_destino: (columna_origen, transformador)} a cada fila.

    Parámetros:
        df: DataFrame de la hoja Excel.
        mapeo: dict de la forma {"id": ("id_buque", limpia), ...}.
        clave_obligatoria: si la clave indicada sale None tras transformar,
            la fila se descarta (no se emite). Por defecto "id".

    Devuelve una lista de dicts. Las filas completamente vacías o sin la clave
    obligatoria se ignoran silenciosamente.
    """
    out: List[Dict] = []
    if df.empty:
        return out
    for _, row in df.iterrows():
        rec: Dict[str, Any] = {}
        for dest, (orig, transf) in mapeo.items():
            val = row.get(orig) if orig and orig in row.index else None
            rec[dest] = transf(val) if val is not None else None
        # Descartar si toda la fila quedó vacía
        if not any(v is not None for v in rec.values()):
            continue
        # Descartar si falta la clave obligatoria
        if clave_obligatoria and rec.get(clave_obligatoria) is None:
            continue
        out.append(rec)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Extracción de cada tabla
# ─────────────────────────────────────────────────────────────────────────────

def extraer(directorio: Path) -> Dict[str, List[Dict]]:
    """Lee los 5 Excels desde `directorio` y devuelve el dataset normalizado."""

    print(f"\n→ Leyendo Excels desde {directorio.resolve()}")
    xl_buq = carga_excel(directorio / "BUQUE_CLAUDE.xlsx")
    xl_per = carga_excel(directorio / "PERSONAL_CLAUDE.xlsx")
    xl_lug = carga_excel(directorio / "LUGAR_CLAUDE.xlsx")
    xl_gra = carga_excel(directorio / "GRADO_CLAUDE.xlsx")
    xl_pos = carga_excel(directorio / "lat_claude.xlsx")

    print("→ Extrayendo tablas")
    D: Dict[str, List[Dict]] = {}

    # ── PERSONAL ────────────────────────────────────────────────────────────
    D["personal"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "PERSONAL"), {
        "id":                    ("id_persona",           limpia),
        "nombre":                ("nombre_persona",       limpia),
        "fecha_nacimiento":      ("fecha_nacimiento",     fecha_iso),
        "id_lugar_nacimiento":   ("id_lugar_nacimiento",  limpia),
        "fecha_defuncion":       ("fecha_defuncion",      fecha_iso),
        "id_lugar_defuncion":    ("id_lugar_defuncion",   limpia),
        "id_batalla_defuncion":  ("id_batalla_defuncion", limpia),
        "id_rama":               ("id_rama",              limpia),
        "id_profesion":          ("id_profesion",         limpia),  # puede no existir en la hoja
        "id_funcion":            ("id_funcion",           limpia),
        "id_jefatura":           ("id_jefatura",          limpia),
        "imagen":                ("imagen",               limpia),
    })

    # ── RELACIONES (parentesco/vínculos) ────────────────────────────────────
    D["relaciones"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "RELACIONES PERSONALES", "Relaciones"), {
        "id":                  ("id_relacion",          limpia),
        "id_persona":          ("id_persona",           limpia),
        "id_funcion":          ("id_funcion",           limpia),  # aquí codifica el tipo de relación
        "id_persona_relacion": ("id_persona_relacion",  limpia),
        "id_rama":             ("id_rama",              limpia),
    })

    # ── TÍTULOS POR PERSONA ────────────────────────────────────────────────
    D["titulos_persona"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "TITULO PERSONA"), {
        "id":               ("id_titulo_persona", limpia),
        "id_persona":       ("id_persona",        limpia),
        "id_titulo":        ("id_titulo",         limpia),
        "fecha_concesion":  ("fecha_concesion",   fecha_iso),
        "fecha_fin":        ("fecha_fin",         fecha_iso),
    })

    # ── MATRIMONIOS ────────────────────────────────────────────────────────
    D["matrimonios"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "MATRIMONIO", "Matrimonios"), {
        "id":                ("id_matrimonio",      limpia),
        "id_persona_1":      ("id_persona_1",       limpia),
        "id_persona_2":      ("id_persona_2",       limpia),
        "fecha_matrimonio":  ("fecha_matrimonio",   fecha_iso),
        "id_lugar":          ("id_lugar",           limpia),
        "fecha_fin":         ("fecha_fin",          fecha_iso),
        "observaciones":     ("observaciones",      limpia),  # puede no existir
    })

    # ── ASCENSOS ───────────────────────────────────────────────────────────
    D["ascensos"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "ASCENSO", "ASCENSOS"), {
        "id":          ("id_ascenso",     limpia),
        "id_persona":  ("id_persona",     limpia),
        "id_grado":    ("id_grado",       limpia),
        "fecha":       ("fecha_ascenso",  fecha_iso),
        "motivo":      ("motivo",         limpia),
    })

    # ── DESTINOS A BORDO ───────────────────────────────────────────────────
    # Importante: en esta hoja, id_funcion del Excel lleva IDs de funcion_buque
    # (prefijo FBQ), por eso en el JSON se mapea a id_funcion_buque.
    D["destinos_mar"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "DESTINO_MAR", "DESTINOS_MAR"), {
        "id":                    ("id_destino_mar",         limpia),
        "id_persona":            ("id_persona",             limpia),
        "id_grado":              ("id_grado",               limpia),
        "nombre_grado_decl":     ("nombre_grado",           limpia),  # fallback declarativo
        "id_buque":              ("id_buque",               limpia),
        "nombre_buque_decl":     ("nombre_buque",           limpia),
        "id_funcion_buque":      ("id_funcion",             limpia),  # nota: rename intencional
        "nombre_funcion_decl":   ("nombre_funcion",         limpia),
        "fecha_real_orden":      ("fecha_real_orden",       fecha_iso),
        "fecha_toma_mando":      ("fecha_toma_mando",       fecha_iso),
        "fecha_cese_orden":      ("fecha_cese_real_orden",  fecha_iso),
        "fecha_cese_efectiva":   ("fecha_cese_efectiva",    fecha_iso),
    })

    # ── DESTINOS EN TIERRA ─────────────────────────────────────────────────
    D["destinos_tierra"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "DESTINO_TIERRA", "DESTINOS_TIERRA"), {
        "id":                    ("id_destino_tierra",     limpia),
        "id_persona":            ("id_persona",            limpia),
        "id_grado":              ("id_grado",              limpia),
        "nombre_grado_decl":     ("nombre_grado",          limpia),
        "id_jefatura":           ("id_jefatura",           limpia),
        "nombre_jefatura_decl":  ("nombre_jefatura",       limpia),
        "id_funcion":            ("id_funcion",            limpia),
        "nombre_funcion_decl":   ("nombre_funcion",        limpia),
        "fecha_real_orden":      ("fecha_real_orden",      fecha_iso),
        "fecha_toma_mando":      ("fecha_toma_mando",      fecha_iso),
        "fecha_cese_orden":      ("fecha_cese_real_orden", fecha_iso),
        "fecha_cese_efectiva":   ("fecha_cese_efectiva",   fecha_iso),
    })

    # ── MANDOS DE ESCUADRA ─────────────────────────────────────────────────
    # id_funcion distingue Comandante Gral. (FUNC015) de Segundo (FUNC063).
    D["mandos_escuadra"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "Mando Escuadra"), {
        "id":                ("id_mando_escuadra",     limpia),
        "id_persona":        ("id_persona",            limpia),
        "id_funcion":        ("id_funcion",            limpia),
        "id_grado":          ("id_grado",              limpia),
        "id_escuadra":       ("id_escuadra",           limpia),
        "id_buque":          ("id_buque",              limpia),
        "fecha_real_orden":  ("fecha_real_orden",      fecha_iso),
        "fecha_toma_mando":  ("fecha_toma_mando",      fecha_iso),
        "fecha_cese_orden":  ("fecha_cese_real_orden", fecha_iso),
    })

    # ── PLANA MAYOR DE ESCUADRA ────────────────────────────────────────────
    # id_mando_escuadra vincula cada plana mayor a su jefe correspondiente.
    D["plana_mayor_escuadra"] = filas(hoja(xl_per, "PERSONAL_CLAUDE.xlsx", "PLANA_MAYOR_ESCUADRA"), {
        "id":                  ("id_plana_escuadra",     limpia),
        "id_persona":          ("id_persona",            limpia),
        "id_grado":            ("id_grado",              limpia),
        "id_escuadra":         ("id_escuadra",           limpia),
        "id_buque":            ("id_buque",              limpia),
        "id_funcion":          ("id_funcion",            limpia),
        "id_mando_escuadra":   ("id_mando_escuadra",     limpia),
        "fecha_real_orden":    ("fecha_real_orden",      fecha_iso),
        "fecha_toma_mando":    ("fecha_toma_mando",      fecha_iso),
        "fecha_cese_orden":    ("fecha_cese_real_orden", fecha_iso),
        "fecha_cese_efectiva": ("fecha_cese_efectiva",   fecha_iso),
    })

    # ── BUQUES ─────────────────────────────────────────────────────────────
    D["buques"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "BUQUES"), {
        "id":                       ("id_buque",                 limpia),
        "nombre":                   ("nombre_buque",             limpia),
        "tipo":                     ("tipo_buque",               limpia),
        "id_rama":                  ("id_rama",                  limpia),
        "clase":                    ("clase",                    limpia),
        "canones":                  ("n_canones",                limpia),
        "fecha_botadura":           ("fecha_botadura",           fecha_iso),
        "fecha_baja":               ("fecha_baja",               fecha_iso),
        "motivo_baja":              ("motivo_baja",              limpia),
        "id_baja_en_batalla":       ("id_baja_en_batalla",       limpia),
        "id_jefatura_propietaria":  ("id_jefatura_propietaria",  limpia),
        "id_persona_propietario":   ("id_persona_propietario",   limpia),
    })

    # ── RUTAS (travesías) ──────────────────────────────────────────────────
    D["rutas"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "RUTAS"), {
        "id":                ("id_ruta",          limpia),
        "id_buque":          ("id_buque",         limpia),
        "id_escuadra":       ("id_escuadra",      limpia),
        "puerto_partida":    ("puerto_partida",   limpia),
        "id_lugar_partida":  ("id_lugar_partida", limpia),
        "fecha_partida":     ("fecha_partida",    fecha_iso),
        "puerto_destino":    ("puerto_destino",   limpia),
        "id_lugar_destino":  ("id_lugar_destino", limpia),
        "fecha_arribada":    ("fecha_arribada",   fecha_iso),
        "millas":            ("millas_recorridas", to_float),
    })

    # ── PASAJEROS (con campos 6.B: lugar de embarque y desembarque) ───────
    D["pasajeros"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "Pasajeros"), {
        "id":                    ("id_pasajeros",         limpia),
        "id_persona":            ("id_persona",           limpia),
        "id_rama":               ("id_rama",              limpia),
        "id_funcion":            ("id_funcion",           limpia),
        "id_jefatura":           ("id_jefatura",          limpia),
        "id_ruta":               ("id_ruta",              limpia),
        "id_lugar_embarque":     ("id_lugar_embarque",    limpia),     # NUEVO 6.B
        "id_lugar_desembarque":  ("id_lugar_desembarque", limpia),     # NUEVO 6.B
    }, clave_obligatoria="id")

    # ── CARGA (con campos 6.B: lugar de carga y descarga) ──────────────────
    D["carga"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "Carga"), {
        "id":                        ("id_carga",                 limpia),
        "id_ruta":                   ("id_ruta",                  limpia),
        "id_buque":                  ("id_buque",                 limpia),
        "id_mercancia":              ("id_mercancia",             limpia),
        "cantidad":                  ("cantidad",                 to_float),
        "unidad":                    ("unidad",                   limpia),
        "peso":                      ("peso",                     to_float),
        "unidad_peso":               ("unidad.1",                 limpia),  # pandas renombra el dup
        "valor_pesos":               ("valor_pesos_moneda",       to_float),
        "remitente":                 ("remitente",                limpia),
        "id_persona_remitente":      ("id_persona_remitente",     limpia),
        "id_jefatura_remitente":     ("id_jefatura_remitente",    limpia),
        "id_lugar_carga":            ("id_lugar_carga",           limpia),  # NUEVO 6.B
        "id_lugar_descarga":         ("id_lugar_descarga",        limpia),  # NUEVO 6.B
        "destinatario":              ("destinatario",             limpia),
        "id_persona_destinatario":   ("id_persona_destinatario",  limpia),
        "id_jefatura_destinatario":  ("id_jefatura_destinatario", limpia),
    })

    # ── CAUDALES (con campos 6.B) ──────────────────────────────────────────
    D["caudales"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "Caudales"), {
        "id":                        ("id_caudales",              limpia),
        "id_ruta":                   ("id_ruta",                  limpia),
        "id_buque":                  ("id_buque",                 limpia),
        "id_mercancias_caudales":    ("id_mercancias_caudales",   limpia),
        "valor_pesos":               ("valor_pesos_moneda",       to_float),
        "remitente":                 ("remitente",                limpia),
        "id_persona_remitente":      ("id_persona_remitente",     limpia),
        "id_jefatura_remitente":     ("id_jefatura_remitente",    limpia),
        "id_lugar_carga":            ("id_lugar_carga",           limpia),  # NUEVO 6.B
        "id_lugar_descarga":         ("id_lugar_descarga",        limpia),  # NUEVO 6.B
        "destinatario":              ("destinatario",             limpia),
        "id_persona_destinatario":   ("id_persona_destinatario",  limpia),
        "id_jefatura_destinatario":  ("id_jefatura_destinatario", limpia),
    })

    # ── TRANSPORTE DE TROPAS (con campos 6.B) ──────────────────────────────
    D["transporte_tropas"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "Transporte de Tropas"), {
        "id":                    ("id_transporte_tropas",        limpia),
        "id_buque":              ("id_buque",                    limpia),
        "id_escuadra":           ("id_escuadra",                 limpia),
        "id_unidad_embarcada":   ("id_unidad_embarcada",         limpia),
        "nombre_unidad":         ("nombre_unidad",               limpia),
        "integridad":            ("integridad_unidad_embarcada", limpia),
        "oficiales_artilleria":  ("oficiales_artilleria",        to_int),
        "artilleros":            ("artilleros",                  to_int),
        "oficiales_infanteria":  ("oficiales_infanteria",        to_int),
        "soldados":              ("soldados",                    to_int),
        "destino_unidad":        ("destino_unidad_embarcada",    limpia),
        "id_ruta":               ("id_ruta",                     limpia),
        "id_lugar_embarque":     ("id_lugar_embarque",           limpia),   # NUEVO 6.B
        "id_lugar_desembarque":  ("id_lugar_desembarque",        limpia),   # NUEVO 6.B
    })

    # ── MARINERÍA: dotación numérica por travesía/buque (agregados) ────────
    # Datos cuantitativos: cuántos Oficiales Mayores, Tropa de Infantería,
    # Artilleros, Marineros, Grumetes, Pajes, Criados etc. iban a bordo en
    # cada ruta concreta. Distinto de destinos_mar, que individualiza oficiales.
    D["marineria"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "Marineria", "Marinería"), {
        "id":                 ("id_marineria",      limpia),
        "id_ruta":            ("id_ruta",           limpia),
        "id_buque":           ("id_buque",          limpia),
        "oficial_mayor":      ("Oficial Mayor",     to_int),
        "tropa_infanteria":   ("Tropa Infantería",  to_int),
        "tropa_artilleria":   ("Tropa Artillería",  to_int),
        "oficial_de_mar":     ("Oficial de Mar",    to_int),
        "artillero":          ("Artillero",         to_int),
        "marinero":           ("Marinero",          to_int),
        "grumete":            ("Grumete",           to_int),
        "paje":               ("Paje",              to_int),
        "criado":             ("Criado",            to_int),
        "dotacion_total":     ("Dotación Total",    to_int),
    })

    # ── SANIDAD A BORDO: fallecidos y enfermos por travesía ────────────────
    D["sanidad_bordo"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "sanidad a bordo", "Sanidad a bordo", "Sanidad"), {
        "id":                                   ("id_sanidad",                            limpia),
        "id_ruta":                              ("id_ruta",                               limpia),
        "id_buque":                             ("id_buque",                              limpia),
        "id_persona_fallecida":                 ("id_persona_fallecida",                  limpia),
        "dotacion_fallecida":                   ("dotacion_fallecida",                    to_int),
        "dotacion_enferma":                     ("dotacion_enferma",                      to_int),
        "id_enfermedad":                        ("id_enfermedad",                         limpia),
        "id_lugar_desembarco_dotacion_enferma": ("id_lugar_desembarco_dotacion_enferma",  limpia),
    })

    # ── MANTENIMIENTOS ─────────────────────────────────────────────────────
    D["mantenimientos"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "MANTENIMIENTO BUQUES"), {
        "id":                  ("id_mantenimiento",              limpia),
        "id_infraestructura":  ("id_infraestructura",            limpia),
        "id_buque":            ("id_buque",                      limpia),
        "fecha_entrada":       ("fecha_entrada_mantenimiento",   fecha_iso),
        "fecha_salida":        ("fecha_salida_mantenimiento",    fecha_iso),
        "trabajos":            ("trabajos_mantenimiento",        limpia),
    })

    # ── PERTENENCIAS DE BUQUE A ESCUADRA ───────────────────────────────────
    # La columna `insignia` distingue "insignia principal" de "segunda insignia".
    D["pertenencias"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "Pertenencia Escuadra"), {
        "id":                  ("id_pertenencia",      limpia),
        "id_buque":            ("id_buque",            limpia),
        "id_escuadra":         ("id_escuadra",         limpia),
        "funcion_en_escuadra": ("funcion_en_escuadra", limpia),
        "insignia":            ("insignia",            limpia),
        "fecha_alta":          ("fecha_alta",          fecha_iso),
        "fecha_baja":          ("fecha_baja",          fecha_iso),
        "motivo_alta":         ("motivo_alta",         limpia),
        "motivo_baja":         ("motivo_baja",         limpia),
    })

    # ── BATALLAS ───────────────────────────────────────────────────────────
    D["batallas"] = filas(hoja(xl_lug, "LUGAR_CLAUDE.xlsx", "BATALLA"), {
        "id":        ("id_batalla",     limpia),
        "nombre":    ("nombre_batalla", limpia),
        "fecha":     ("fecha",          fecha_iso),
        "guerra":    ("Guerra",         limpia),
        "latitud":   ("Latitud",        to_float),
        "longitud":  ("Longitud",       to_float),
    })

    # ── PARTICIPACIONES EN BATALLA ─────────────────────────────────────────
    D["participaciones_batalla"] = filas(hoja(xl_buq, "BUQUE_CLAUDE.xlsx", "Participación Batalla"), {
        "id":          ("id_participacion_batalla", limpia),
        "id_buque":    ("id_buque",                 limpia),
        "id_escuadra": ("id_escuadra",              limpia),
        "id_batalla":  ("id_batalla",               limpia),
        "fecha":       ("fecha",                    fecha_iso),
    })

    # ── LUGARES ────────────────────────────────────────────────────────────
    D["lugares"] = filas(hoja(xl_lug, "LUGAR_CLAUDE.xlsx", "LUGAR", "LUGARES"), {
        "id":       ("id_lugar",     limpia),
        "nombre":   ("nombre_lugar", limpia),
        "latitud":  ("Latitud",      to_float),
        "longitud": ("Longitud",     to_float),
    })

    # ── INFRAESTRUCTURAS ───────────────────────────────────────────────────
    D["infraestructuras"] = filas(hoja(xl_lug, "LUGAR_CLAUDE.xlsx", "INFRAESTRUCTURA", "Infraestructuras"), {
        "id":           ("id_infraestructura",     limpia),
        "nombre":       ("nombre_infraestructura", limpia),
        "tipo":         ("tipo",                   limpia),
        "id_lugar":     ("id_lugar",               limpia),
        "id_jefatura":  ("id_jefatura",            limpia),
        "latitud":      ("Latitud",                to_float),
        "longitud":     ("Longitud",               to_float),
    })

    # ── ENFERMEDADES (catálogo) ────────────────────────────────────────────
    D["enfermedades"] = filas(hoja(xl_lug, "LUGAR_CLAUDE.xlsx", "Enfermedades"), {
        "id":     ("id_enfermedad",     limpia),
        "nombre": ("nombre_enfermedad", limpia),
    })

    # ── MERCANCÍAS Y CATÁLOGOS ASOCIADOS ───────────────────────────────────
    D["mercancias"] = filas(hoja(xl_lug, "LUGAR_CLAUDE.xlsx", "Mercancias"), {
        "id":           ("id_mercancia",    limpia),
        "nombre":       ("nombre_mercancia", limpia),
        "categoria":    ("categoria",       limpia),
        "unidad_base":  ("unidad_base",     limpia),
    })

    D["mercancias_caudales"] = filas(hoja(xl_lug, "LUGAR_CLAUDE.xlsx", "Mercancias_caudales"), {
        "id":     ("id_mercancias_caudales",     limpia),
        "nombre": ("nombre_mercancias_caudales", limpia),
    })

    # ── GRADOS Y CATÁLOGOS DE PERSONAL ─────────────────────────────────────
    D["grados"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Grado"), {
        "id":      ("id_grado",     limpia),
        "nombre":  ("nombre_grado", limpia),
        "orden":   ("orden",        limpia),
        "rama":    ("rama",         limpia),
        "notas":   ("notas",        limpia),
    })

    D["funciones"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Funcion"), {
        "id":             ("id_funcion",     limpia),
        "nombre":         ("nombre_funcion", limpia),
        "tipo":           ("tipo_funcion",   limpia),
        "observaciones":  ("observaciones",  limpia),
    })

    D["funciones_buque"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Funcion Tripulacion Buque", "Funcion Buque"), {
        "id":     ("id_funcion_buque",     limpia),
        "nombre": ("nombre_funcion_buque", limpia),
        "orden":  ("orden",                limpia),
    })

    D["jefaturas"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Jefatura"), {
        "id":                    ("id_jefatura",            limpia),
        "nombre":                ("nombre_jefatura",        limpia),
        "dependencia_superior":  ("dependencia_superior",   limpia),
        "id_lugar":              ("id_lugar",               limpia),
    })

    D["titulos"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Titulos honores"), {
        "id":              ("id_titulo",             limpia),
        "nombre":          ("nombre_titulo",         limpia),
        "nombre_entidad":  ("nombre_titulo_entidad", limpia),   # Marquesado, Condado, etc.
        "fecha_creacion":  ("fecha_creacion",        fecha_iso),
        "tipo":            ("tipo",                  limpia),
        "institucion":     ("institucion",           limpia),
        "rango":           ("rango",                 limpia),
    })

    D["profesiones"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Profesion"), {
        "id":             ("id_profesion",     limpia),
        "nombre":         ("nombre_profesion", limpia),
    })  # puede estar vacía si la hoja no existe

    D["ramas"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Rama"), {
        "id":     ("id_rama",     limpia),
        "nombre": ("nombre_rama", limpia),
        "tipo":   ("tipo",        limpia),
    })

    D["escuadras"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "Escuadra"), {
        "id":                  ("id_escuadra",         limpia),
        "nombre":              ("nombre_escuadra",     limpia),
        "id_buque_insignia":   ("id_buque_insignia",   limpia),
        "tipo":                ("tipo",                limpia),
        "fecha_constitucion":  ("fecha_constitucion",  fecha_iso),
        "fecha_disolucion":    ("fecha_disolucion",    fecha_iso),
        "descripcion":         ("descripcion",         limpia),
    })

    # ── UNIDADES DE LOS REALES EJÉRCITOS (catálogo) ────────────────────────
    # Unidades militares (batallones, regimientos) y categorías genéricas
    # ("Marinería", "Oficiales navales") referenciables desde transporte_tropas.
    D["reales_ejercitos"] = filas(hoja(xl_gra, "GRADO_CLAUDE.xlsx", "REALES EJERCITOS"), {
        "id":             ("id_reales_ejercitos", limpia),
        "nombre":         ("nombre_unidad",       limpia),
        "entidad":        ("entidad",             limpia),
        "id_rama":        ("id_rama",             limpia),
    })

    # ── POSICIONES DIARIAS ─────────────────────────────────────────────────
    # Una fila se acepta si tiene fecha + (id_buque o id_escuadra).
    # Las filas con solo id_escuadra y sin id_buque son "posiciones de escuadra
    # pura" — de ellas heredan posición todos los buques que en esa fecha
    # pertenecen a la escuadra.
    pos_raw = filas(hoja(xl_pos, "lat_claude.xlsx", "POSICIONES"), {
        "fecha":              ("fecha",              fecha_iso),
        "id_buque":           ("id_buque",           limpia),
        "id_escuadra":        ("id_escuadra",        limpia),
        "id_ruta":            ("id_ruta",            limpia),
        "id_batalla":         ("id_batalla",         limpia),
        "id_infraestructura": ("id_infraestructura", limpia),
        "latitud":            ("Latitud",            to_float),
        "longitud":           ("Longitud",           to_float),
        "situacion":          ("situacion",          limpia),
        "estado":             ("estado",             limpia),
        "ordenes":            ("ordenes",            limpia),
        "navegacion":         ("navegacion",         limpia),
    }, clave_obligatoria="fecha")
    # Filtrado adicional: requerir id_buque o id_escuadra
    D["posiciones"] = [
        p for p in pos_raw if p.get("id_buque") or p.get("id_escuadra")
    ]
    descartadas = len(pos_raw) - len(D["posiciones"])
    if descartadas:
        aviso(f"POSICIONES: descartadas {descartadas} filas sin id_buque ni id_escuadra")

    # ── Cierre implícito de destinos sin fecha de cese ──────────────────────
    # Un destino (mar/tierra/mando/plana mayor) sin fecha_cese_orden ni
    # fecha_cese_efectiva queda "abierto" para siempre, lo que provoca que la
    # persona aparezca a bordo en periodos posteriores que no le corresponden.
    # Aquí calculamos `fecha_cese_implicita` mirando dos pistas:
    #   1) la fecha del siguiente destino documentado de la misma persona
    #      (sea en mar, tierra, mando de escuadra o plana mayor);
    #   2) la fecha de defunción de la persona, si no hay un siguiente destino.
    # El template usa este campo SOLO como tercer fallback, después de
    # fecha_cese_efectiva y fecha_cese_orden, así que los datos explícitos
    # siguen prevaleciendo.
    _calcula_cierre_implicito(D)

    return D


def _calcula_cierre_implicito(D: Dict[str, List[Dict]]) -> None:
    """Rellena `fecha_cese_implicita` en los destinos sin cese explícito."""
    personas_por_id = {p["id"]: p for p in D["personal"] if p.get("id")}
    tablas_destino = ["destinos_mar", "destinos_tierra", "mandos_escuadra", "plana_mayor_escuadra"]

    def inicio(d):
        return d.get("fecha_toma_mando") or d.get("fecha_real_orden") or ""

    def cese_explicito(d):
        return d.get("fecha_cese_efectiva") or d.get("fecha_cese_orden")

    # Construir índice persona → todos sus destinos ordenados por fecha de inicio
    por_persona: Dict[str, List[Dict]] = {}
    for tabla in tablas_destino:
        for d in D.get(tabla, []):
            pid = d.get("id_persona")
            if pid:
                por_persona.setdefault(pid, []).append(d)
    for pid, lst in por_persona.items():
        lst.sort(key=inicio)

    cerrados = 0
    for pid, eventos in por_persona.items():
        for i, d in enumerate(eventos):
            d["fecha_cese_implicita"] = None
            if cese_explicito(d):
                continue
            ini_actual = inicio(d)
            # Buscar el siguiente destino con fecha posterior estricta
            cese_impl = None
            for j in range(i + 1, len(eventos)):
                fn = inicio(eventos[j])
                if fn and (not ini_actual or fn > ini_actual):
                    cese_impl = fn
                    break
            # Si no hay siguiente, mirar fecha de defunción
            if not cese_impl:
                persona = personas_por_id.get(pid)
                if persona and persona.get("fecha_defuncion"):
                    cese_impl = persona["fecha_defuncion"]
            d["fecha_cese_implicita"] = cese_impl
            if cese_impl:
                cerrados += 1
    print(f"  · cierres implícitos calculados: {cerrados}")


# ─────────────────────────────────────────────────────────────────────────────
# Validación: avisos de integridad referencial básica
# ─────────────────────────────────────────────────────────────────────────────

def valida(D: Dict[str, List[Dict]]) -> None:
    """Avisos blandos sobre integridad: IDs huérfanos, fechas invertidas, etc."""

    def index(lista, clave="id"):
        return {r[clave] for r in lista if r.get(clave)}

    personas = index(D["personal"])
    buques = index(D["buques"])
    lugares = index(D["lugares"])
    rutas = index(D["rutas"])
    escuadras = index(D["escuadras"])
    batallas = index(D["batallas"])
    grados = index(D["grados"])
    jefaturas = index(D["jefaturas"])

    def check(lista, campo, indice, etiqueta):
        for r in lista:
            v = r.get(campo)
            if v and v not in indice:
                aviso(f"{etiqueta}: {campo}={v!r} no existe en su catálogo (registro id={r.get('id')})")

    check(D["ascensos"],            "id_persona",  personas,  "ASCENSO")
    check(D["ascensos"],            "id_grado",    grados,    "ASCENSO")
    check(D["destinos_mar"],        "id_persona",  personas,  "DESTINO_MAR")
    check(D["destinos_mar"],        "id_buque",    buques,    "DESTINO_MAR")
    check(D["destinos_tierra"],     "id_persona",  personas,  "DESTINO_TIERRA")
    check(D["destinos_tierra"],     "id_jefatura", jefaturas, "DESTINO_TIERRA")
    check(D["rutas"],               "id_buque",    buques,    "RUTAS")
    check(D["rutas"],               "id_lugar_partida", lugares, "RUTAS")
    check(D["rutas"],               "id_lugar_destino", lugares, "RUTAS")
    check(D["carga"],               "id_buque",    buques,    "CARGA")
    check(D["carga"],               "id_ruta",     rutas,     "CARGA")
    check(D["carga"],               "id_lugar_carga",    lugares, "CARGA")
    check(D["carga"],               "id_lugar_descarga", lugares, "CARGA")
    check(D["caudales"],            "id_buque",    buques,    "CAUDALES")
    check(D["caudales"],            "id_ruta",     rutas,     "CAUDALES")
    check(D["caudales"],            "id_lugar_carga",    lugares, "CAUDALES")
    check(D["caudales"],            "id_lugar_descarga", lugares, "CAUDALES")
    check(D["pasajeros"],           "id_persona",  personas,  "PASAJEROS")
    check(D["pasajeros"],           "id_ruta",     rutas,     "PASAJEROS")
    check(D["pasajeros"],           "id_lugar_embarque",    lugares, "PASAJEROS")
    check(D["pasajeros"],           "id_lugar_desembarque", lugares, "PASAJEROS")
    check(D["transporte_tropas"],   "id_buque",    buques,    "TRANSPORTE_TROPAS")
    check(D["transporte_tropas"],   "id_lugar_embarque",    lugares, "TRANSPORTE_TROPAS")
    check(D["transporte_tropas"],   "id_lugar_desembarque", lugares, "TRANSPORTE_TROPAS")
    check(D["pertenencias"],        "id_buque",    buques,    "PERTENENCIA_ESCUADRA")
    check(D["pertenencias"],        "id_escuadra", escuadras, "PERTENENCIA_ESCUADRA")
    check(D["participaciones_batalla"], "id_buque",   buques,   "PARTICIPACION_BATALLA")
    check(D["participaciones_batalla"], "id_batalla", batallas, "PARTICIPACION_BATALLA")
    check(D["buques"],              "id_baja_en_batalla", batallas, "BUQUES")
    check(D["personal"],            "id_batalla_defuncion", batallas, "PERSONAL")


# ─────────────────────────────────────────────────────────────────────────────
# Salida
# ─────────────────────────────────────────────────────────────────────────────

def inyecta_en_template(template_path: Path, dataset_json: str) -> str:
    """Sustituye __DATASET__ en el template por el JSON. Devuelve el HTML resultante."""
    if not template_path.exists():
        raise FileNotFoundError(f"No se encuentra la plantilla: {template_path}")
    html = template_path.read_text(encoding="utf-8")
    if "__DATASET__" not in html:
        raise ValueError("La plantilla no contiene el marcador __DATASET__ donde inyectar los datos.")
    return html.replace("__DATASET__", dataset_json, 1)


def escribe_incidencias(path: Path) -> None:
    """Escribe el log de incidencias en lenguaje llano."""
    cabecera = (
        "INCIDENCIAS DE LA EXTRACCIÓN DE DATOS\n"
        f"Fecha: {dt.datetime.now().isoformat(timespec='seconds')}\n"
        f"Total de avisos: {len(INCIDENCIAS)}\n"
        "─" * 70 + "\n\n"
    )
    if not INCIDENCIAS:
        cuerpo = "Sin incidencias. Todos los datos parseados sin avisos.\n"
    else:
        cuerpo = "\n".join(f"- {m}" for m in INCIDENCIAS) + "\n"
    path.write_text(cabecera + cuerpo, encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="Extractor de datos para la web de la Armada XVIII.")
    parser.add_argument("--excels",   default=".",            help="Directorio con los 5 Excels (por defecto, el actual).")
    parser.add_argument("--template", default="template.html", help="Plantilla HTML con marcador __DATASET__.")
    parser.add_argument("--out",      default="index.html",   help="HTML final a generar.")
    parser.add_argument("--json",     default="data.json",    help="JSON normalizado a generar (útil para depuración).")
    parser.add_argument("--incidencias", default="incidencias.txt", help="Archivo de log de incidencias.")
    args = parser.parse_args()

    excels_dir = Path(args.excels)
    template = Path(args.template)
    out_html = Path(args.out)
    out_json = Path(args.json)
    out_inc  = Path(args.incidencias)

    D = extraer(excels_dir)
    print("→ Validando integridad referencial")
    valida(D)

    # JSON compacto para inyectar (sin indent) — el archivo data.json sí va indentado
    dataset_json_inline = json.dumps(D, ensure_ascii=False, separators=(",", ":"))
    dataset_json_pretty = json.dumps(D, ensure_ascii=False, indent=2)

    out_json.write_text(dataset_json_pretty, encoding="utf-8")
    print(f"→ Escrito {out_json} ({out_json.stat().st_size:,} bytes)")

    if template.exists():
        html_final = inyecta_en_template(template, dataset_json_inline)
        out_html.write_text(html_final, encoding="utf-8")
        print(f"→ Escrito {out_html} ({out_html.stat().st_size:,} bytes)")
    else:
        aviso(f"No se inyectó en plantilla: {template} no existe.")

    escribe_incidencias(out_inc)
    print(f"→ Escrito {out_inc} ({len(INCIDENCIAS)} incidencia(s))")

    # Resumen final
    totales = {k: len(v) for k, v in D.items()}
    print("\n=== Resumen ===")
    for k, n in sorted(totales.items()):
        print(f"  {k:30s}: {n}")
    print(f"  TOTAL registros: {sum(totales.values())}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
