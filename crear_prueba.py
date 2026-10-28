"""Build a large BCP-format test export (consultant: "agrégale más filas y valida").

Takes the real export's layout and opening balance after 21/10/2026, then adds
one row per August 2026 OBNK line of 104111 (same memo, ref and amount), dated
22–31/10/2026. Lines with a short ref (0909, 0101, 4611) get operation number
00000000 and the ref as UTC, like BCP's batch charges. Balances are recomputed.

Usage: python crear_prueba.py <export_real.xlsx> <salida.xlsx>
"""
import sys
from datetime import date
from pathlib import Path

import openpyxl

from comun import HEADER_ROW, leer_bcp, sql

origen, destino = Path(sys.argv[1]), Path(sys.argv[2])
saldo = leer_bcp(origen)["saldo_final"]

agosto = sql("SBO_TEST_02092026",
             "SELECT DAY(DueDate), ISNULL(Ref,''), Memo, CAST(DebAmount AS decimal(19,2)), "
             "CAST(CredAmnt AS decimal(19,2)) FROM OBNK WHERE AcctCode='104111' "
             "AND DueDate BETWEEN '20260801' AND '20260831' ORDER BY DueDate, Sequence")

filas = []  # chronological
for dia, ref, memo, deb, cred in agosto:
    ref = ref.strip()
    fecha = date(2026, 10, 22 + (int(dia) - 1) * 10 // 31)
    monto = round(float(cred) - float(deb), 2)
    saldo = round(saldo + monto, 2)
    num, utc = (ref.zfill(8), "") if len(ref) > 4 else ("00000000", ref.ljust(6))
    filas.append([fecha.strftime("%d/%m/%Y"), "", memo.strip(), monto, saldo,
                  "111-065", num, "00:00:00", "PRUEBA", utc, ""])

wb = openpyxl.load_workbook(origen)
ws = wb.worksheets[0]
ws.delete_rows(HEADER_ROW + 1, ws.max_row)
for fila in reversed(filas):  # BCP lists newest first
    ws.append(fila)
wb.save(destino)
print(f"{len(filas)} filas, 22/10–31/10/2026, saldo final {saldo:,.2f} -> {destino}")
