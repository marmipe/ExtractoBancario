"""Convert a BCP or Interbank .xlsx export into a DTW oBankPages trama (no load).

Usage:  python convertir_extracto.py <export.xlsx> [salida.txt]

Rules:
  BCP        Ref = Operación - Número (UTC code if it is all zeros); Memo = Descripción operación;
             signed Monto: < 0 -> DebitAmount (cargo), > 0 -> CreditAmount (abono).
  Interbank  Ref = Nro. de operación (for ITF, the number in 'OPERACION: n'); Memo = Movimiento +
             Descripción; Cargo -> DebitAmount, Abono -> CreditAmount; date = Fecha de operación.
  Lines are written oldest first (banks list newest first).
"""
import sys
from pathlib import Path

from comun import AQUI, ErrorExtracto, escribir_trama, leer_extracto, resumen

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    destino = Path(sys.argv[2]) if len(sys.argv) > 2 else AQUI / "TRAMAS" / "EXTRACTO.txt"
    try:
        ext = leer_extracto(Path(sys.argv[1]))
    except ErrorExtracto as e:
        sys.exit(f"ERROR: {e}")
    escribir_trama(ext, destino)
    print(resumen(ext))
    print(f"-> {destino}")
