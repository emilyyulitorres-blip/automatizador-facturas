import re

def _buscar(patron, texto):
    m = re.search(patron, texto, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def _limpiar(valor):
    if not valor:
        return ""
    return valor.replace(",", ".").strip()

def detect_document_type(texto):
    t = texto.upper()
    if "PROFORMA" in t:
        return "NO VÁLIDA - PROFORMA"
    if "FACTURA" in t:
        return "FACTURA"
    return "REVISAR - NO IDENTIFICADO"

def extract_data(texto, archivo):
    tipo = detect_document_type(texto)

    ruc = _buscar(r"R\.?U\.?C\.?[:\s]*([0-9]{13})", texto)
    factura = _buscar(r"([0-9]{3}-[0-9]{3}-[0-9]{9})", texto)

    fecha = _buscar(r"Fecha Emisi[oó]n[:\s]*([0-9]{2}/[0-9]{2}/[0-9]{4})", texto)
    if not fecha:
        fecha = _buscar(r"([0-9]{2}/[0-9]{2}/[0-9]{4})", texto)

    subtotal = _buscar(r"SUBTOTAL SIN IMPUESTOS\s*([0-9]+[\.,][0-9]{2})", texto)
    if not subtotal:
        subtotal = _buscar(r"SUBTOTAL\s*[0-9]*%?\s*([0-9]+[\.,][0-9]{2})", texto)

    iva = _buscar(r"IVA\s*[0-9]*%?\s*([0-9]+[\.,][0-9]{2})", texto)

    total = _buscar(r"VALOR TOTAL\s*([0-9]+[\.,][0-9]{2})", texto)
    if not total:
        total = _buscar(r"TOTAL\s*([0-9]+[\.,][0-9]{2})", texto)

    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    proveedor = ""
    for linea in lineas[:15]:
        if len(linea) > 5 and "RUC" not in linea.upper() and "FACTURA" not in linea.upper():
            proveedor = linea
            break

    estado = "OK" if tipo == "FACTURA" else "REVISAR"

    return {
        "Estado": estado,
        "Tipo documento": tipo,
        "Archivo": archivo,
        "Fecha": fecha,
        "Factura": factura,
        "Proveedor": proveedor,
        "RUC Emisor": ruc,
        "Subtotal": _limpiar(subtotal),
        "IVA": _limpiar(iva),
        "Total": _limpiar(total),
        "Proyecto": "",
        "Consumo": "",
        "Categoría": "",
        "Observación": "" if estado == "OK" else "Documento para revisión"
    }
