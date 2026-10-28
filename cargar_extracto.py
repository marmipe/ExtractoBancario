"""One-click load of a BCP or Interbank statement into SAP B1 bank pages (OBNK) via DTW.

Usage:  python cargar_extracto.py <export.xlsx>     (or drop the file on CARGAR EXTRACTO.BAT)

Steps: read + validate the export -> write the trama DTW reads -> make sure none of
the lines already exist in OBNK -> run DTW in batch mode -> confirm the lines arrived.
The target company and trama path come from BACHERO/SHOUXIN.xml (saved from the DTW wizard).
"""
import subprocess
import sys
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime
from pathlib import Path

from comun import AQUI, ErrorExtracto, escribir_trama, leer_extracto, resumen, sql

DTW = Path(r"C:\Program Files\SAP\Data Transfer Workbench\DTW.exe")
XML = AQUI / "BACHERO" / "SHOUXIN.xml"
BASES_PRUEBA = {"SBO_TEST_02092026", "SBO_28092026", "SBO_PAGOSMASIVOS"}
LOGS = AQUI / "LOGS"

LOGS.mkdir(exist_ok=True)
_log = open(LOGS / f"carga_{datetime.now():%Y%m%d_%H%M%S}.txt", "w", encoding="utf-8")


def say(msg=""):
    print(msg)
    _log.write(msg + "\n")
    _log.flush()


def fallar(msg):
    say(f"\n*** NO SE CARGÓ NADA: {msg}")
    sys.exit(1)


def lineas_obnk(base, cuenta, desde, hasta):
    filas = sql(base, f"SELECT CONVERT(char(8), DueDate, 112), ISNULL(Ref,''), "
                      f"CAST(DebAmount AS decimal(19,2)), CAST(CredAmnt AS decimal(19,2)) "
                      f"FROM OBNK WHERE AcctCode='{cuenta}' AND DueDate BETWEEN '{desde}' AND '{hasta}'")
    return Counter((f[0], f[1].strip(), f"{float(f[2]):.2f}", f"{float(f[3]):.2f}") for f in filas)


def clave(l):
    return (l["fecha"], l["ref"], f"{l['debe']:.2f}", f"{l['haber']:.2f}")


def sin_ref(k):
    # Manual entries often have an empty or different Ref (UTC code, blank for Interbank),
    # so the duplicate check compares date + amounts only.
    return (k[0], k[2], k[3])


def main():
    if len(sys.argv) < 2:
        fallar("arrastre el Excel del banco (BCP o Interbank) sobre CARGAR EXTRACTO.BAT.")
    origen = Path(sys.argv[1])
    say(f"Archivo: {origen.name}")

    # 1. DTW settings
    if not XML.exists():
        fallar(f"falta {XML} (guárdelo desde el asistente DTW con 'Save Parameters').")
    raiz = ET.parse(XML).getroot()
    base = (raiz.findtext("Logon/Company") or "").strip()
    trama = Path((raiz.findtext("FileExtractor/Files/BankPages") or "").strip())
    if raiz.findtext("ObjectCode") != "oBankPages" or not trama.name:
        fallar(f"{XML.name} no es una carga de Bank Pages.")
    say(f"Empresa destino: {base}")
    if base not in BASES_PRUEBA:
        resp = input(f"\n{base} es PRODUCCIÓN. Escriba SI para continuar: ")
        _log.write(f"Confirmación producción: {resp!r}\n")
        if resp.strip().upper() != "SI":
            fallar("cancelado por el usuario.")

    # 2. Read and validate the statement
    try:
        ext = leer_extracto(origen)
    except ErrorExtracto as e:
        fallar(str(e))
    lineas, cuenta = ext["lineas"], ext["cuenta_sap"]
    desde, hasta = lineas[0]["fecha"], lineas[-1]["fecha"]
    say(resumen(ext))

    # 3. Refuse lines already in SAP
    antes = lineas_obnk(base, cuenta, desde, hasta)
    en_sap = Counter(sin_ref(k) for k in antes.elements())
    repetidas = [l for l in lineas if en_sap[sin_ref(clave(l))] > 0]
    if repetidas:
        for l in repetidas[:20]:
            say(f"  ya existe: {l['fecha']} {l['ref']} {l['memo']} {l['debe'] or l['haber']:,.2f}")
        fallar(f"{len(repetidas)} línea(s) ya están en SAP para esas fechas. "
               f"Descargue solo los días que faltan.")

    # 4. Write trama and run DTW
    escribir_trama(ext, trama)
    say(f"Trama: {trama}")
    say("Ejecutando DTW...")
    try:
        subprocess.run([str(DTW), "-s", str(XML)])
    finally:
        # Leave only the headers so a manual wizard import can't reload these lines.
        escribir_trama(dict(ext, lineas=[]), trama)

    # 5. Confirm
    despues = lineas_obnk(base, cuenta, desde, hasta)
    nuevas = despues - antes
    esperadas = Counter(clave(l) for l in lineas)
    if nuevas == esperadas:
        say(f"\nOK: {len(lineas)} líneas cargadas en {base}, cuenta {cuenta}.")
    elif not nuevas:
        fallar("DTW no cargó ninguna línea. Revise el log en DTW (Log Manager).")
    else:
        say(f"\n*** ATENCIÓN: se cargaron {sum(nuevas.values())} de {len(lineas)} líneas.")
        for k in (esperadas - nuevas):
            say(f"  falta: {k}")
        sys.exit(1)


if __name__ == "__main__":
    main()
