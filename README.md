# ExtractoBancario

Carga automática de extractos bancarios de Minera Shouxin a SAP Business One (tabla OBNK) con el Data Transfer Workbench (DTW).

## Uso

Arrastrar la exportación del banco (.xlsx) sobre `CARGAR EXTRACTO.BAT`. El programa:

1. Valida el archivo (cuenta, moneda, fechas y saldo línea por línea).
2. Revisa que ninguna línea esté ya cargada en OBNK.
3. Genera la trama y ejecuta `DTW.exe -s BACHERO\SHOUXIN.xml`.
4. Confirma en OBNK que llegaron todas las líneas y deja un log en `LOGS\`.

Solo convertir, sin tocar la base de datos:

```
set PYTHONUTF8=1
python convertir_extracto.py "<exportación>.xlsx" [salida.txt]
```

## Bancos soportados

- BCP "movimientos históricos": Memo = Descripción operación, Referencia = Operación.
- Interbank "Consulta de Movimientos de Cuenta Corriente": Memo = Descripción (`ITF` en las líneas de ITF), Referencia = Nro. de operación.

Cargos van a DebitAmount y abonos a CreditAmount.

## Archivos que no están en el repositorio

- `BACHERO\SHOUXIN.xml`: parámetros de conexión del DTW (se guardan desde el asistente con "Save Parameters").
- `..\..\mssql_credentials.env`: credenciales de SQL Server para las consultas de solo lectura.
- Extractos, tramas y logs.
