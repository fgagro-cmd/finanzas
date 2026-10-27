# Previsión financiera mensual (día 2)

Script que usa la tarea programada «Previsión financiera mensual» (día 2 de cada mes). **No forma parte de la app Finanzas**: la app no lo carga ni lo necesita. Está aquí para que la tarea lo descargue siempre en su última versión.

No lleva datos: los lee cada mes de tus hojas de Google.

## Qué hace

Cruza Finanzas con las cuentas de las fincas y el resto del negocio y calcula:

- la previsión de caja mes a mes a 12 meses (entradas, salidas y saldo al final de cada mes);
- el colchón de 3 meses de salidas (personal + pagos de fincas + cuotas) y los meses en que la caja baja de él;
- lo que se podría usar hoy sin bajar del colchón en ningún mes;
- los datos de los préstamos para valorar si conviene amortizar;
- los pagos fraccionados (131) de las fincas.

## Datos que necesita (carpeta de datos)

| Fichero | De dónde sale |
|---|---|
| `fincas.xlsx` | Drive: hoja «Stock Fincas» exportada a Excel (Tabla principal, Presupuestos, Parámetros) |
| `ventas.xlsx` | Drive: hoja «ventas FG Agro» exportada a Excel (Ventas, Comisiones) |
| `fin-resumen.json`, `fin-cuentas.json`, `fin-prestamos.json`, `fin-deudas.json`, `fin-cobros.json`, `fin-historico.json` | Hoja «finanzas»: pestañas Resumen, Cuentas, Préstamos, Deudas, Cobros, Histórico (`get_values`) |
| `asesor-clientes.json` | Hoja «Asesor», pestaña Clientes |
| `agro-trabajos.json` | Hoja «agroservicios», pestaña Trabajos externos |

Cada `.json` es la respuesta de Google Sheets tal cual: `{"range": "...", "values": [[...]]}`. Si falta un `.json`, el script busca la pestaña en `finanzas.xlsx`, `asesor.xlsx` o `agro.xlsx`.

## Uso

```
npm install
node prevision.js <carpeta-datos> [AAAA-MM-DD] > resultado.json
```

La reglas de cálculo están explicadas al principio de `prevision.js`.
