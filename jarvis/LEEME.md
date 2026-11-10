# JARVIS · cálculos del resumen diario

No forma parte de la app Finanzas (como la carpeta `prevision/`): lo usa la tarea
programada «Resumen diario JARVIS» de Claude.

- `jarvis_calc.py` (Python 3 + openpyxl): stock bajo (productos de Ventas y de tus
  fincas) y gasto de cultivo de las fincas propias frente al presupuesto y a la
  campaña anterior a la misma fecha.
- Entradas: las hojas «Stock Fincas» y «ventas FG Agro» exportadas a .xlsx desde
  Drive, y (opcional) la pestaña «Mínimos» de la hoja JARVIS.
- Solo lee ficheros. Aquí no hay ningún dato: los datos están en las hojas de Drive.

Uso:

    pip install openpyxl
    python3 jarvis_calc.py --fincas fincas.xlsx --ventas ventas.xlsx --minimos minimos.json --fecha 2026-11-10
