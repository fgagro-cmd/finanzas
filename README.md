# Finanzas

App para el móvil Android que sustituye a `Control_financiero.xlsx`: situación actual y evolución mensual del dinero (cuentas, bienes, deudas y cobros), con los mismos índices del Excel y el presupuesto de gastos personales. Funciona sin conexión.

- **Dashboard:** capital neto, activos, deuda, liquidez y cobros pendientes; índices con semáforo (ratio liquidez/deuda, endeudamiento, independencia financiera, % stock, balance de situación); gráficas mes a mes y el histórico de cierres con su comentario y objetivo.
- **Cuentas:** bancos y cajas de efectivo (grupo Personal / Empresa / Ventas) con el saldo que pones tú; el stock de ventas y de tus fincas, que sale solo de sus apps; y los bienes (casa, finca, vehículos y cada máquina por separado, con el subtotal de maquinaria).
- **Deudas y cobros:** préstamos con el capital pendiente, el tipo de interés (TIN), la cuota mensual y la fecha de la última cuota (los usa la previsión mensual del día 2), deudas y cobros con pagos parciales, y los cobros pendientes que salen solos de tus apps.
- **Presupuesto:** nómina objetivo, categorías (fijo, variable, colchón, ahorro), gastos sueltos y vistas de mes y acumulado. Va por libre: no cambia los saldos.

## Reglas de cálculo (las del Excel)
- Total activos = bancos + efectivo + stock + maquinaria + otros activos.
- Total deuda = préstamos + deudas pendientes (la tarjeta de crédito va en deudas).
- Capital neto = activos − deuda + cobros pendientes.
- Liquidez = bancos + efectivo · Balance de situación = liquidez + cobros − deudas pendientes.
- Ratio liquidez/deuda, endeudamiento (deuda ÷ activos), independencia (capital neto ÷ activos) y % stock (stock ÷ capital neto), con los umbrales del Excel.
- Las variaciones se calculan siempre contra el cierre del mes anterior.

## Cierre de mes
La foto del mes se guarda sola al abrir la app ya en el mes siguiente, con fecha del último día del mes, y queda marcada para revisar y añadir el comentario. También se puede cerrar a mano desde el Dashboard, y rehacer o corregir un cierre.

## Datos de las otras apps
Ventas, Stock Fincas, AgroServicios y Asesor están publicadas en la misma dirección (`fgagro-cmd.github.io`), así que Finanzas lee sus datos directamente del navegador del móvil, sin conexión y sin escribir nada en ellas:
- **Ventas:** valor del stock a precio medio, deudas de clientes (sin tus fincas) y comisiones pendientes (se pueden excluir en Ajustes).
- **Stock Fincas:** valor del stock de tus fincas.
- **AgroServicios:** trabajos externos sin cobrar.
- **Asesor:** meses ya vencidos sin cobrar de asesoría (se cobra por adelantado, así que el mes en curso cuenta como vencido). Desde la versión 1.1.2 se leen de Asesor (base «asesor»); si Asesor no está instalada o aún no tiene tablas de planes y cobros, se usa la app Clientes como respaldo.

Si se abre Finanzas en otro dispositivo, usa el último dato guardado.

## Privacidad
Este repositorio es público y **no contiene datos personales**. Los datos se cargan la primera vez desde un fichero (Ajustes → Cargar datos desde fichero) y viven solo en el móvil y en tu hoja de Google.

## Publicar (GitHub Pages)
Settings → Pages → rama `main`, carpeta `/ (root)`. Después abrir la dirección en Chrome del móvil → ⋮ → «Instalar aplicación».

## Copia automática en Google Sheets
1. Crea una hoja de Google nueva (por ejemplo «Finanzas»).
2. Extensiones → Apps Script. Borra lo que haya, pega el contenido de `google-sheets/Codigo.gs` y cambia `CAMBIA-ESTA-CLAVE` por una clave larga propia. Guarda.
3. Implementar → Nueva implementación → tipo «Aplicación web» → Ejecutar como: **Yo** → Quién tiene acceso: **Cualquier usuario** → Implementar. Autoriza los permisos y copia la dirección (termina en `/exec`).
4. En la app: Ajustes → Copia automática en Google Sheets → pega la dirección y la clave → «Probar y activar».

La hoja es un espejo de un solo sentido: la app reescribe las pestañas Resumen, Histórico, Cuentas, Bienes, Préstamos, Deudas, Cobros y Gastos en cada envío. La clave nunca se sube a GitHub: se guarda solo en el móvil y en el script de tu hoja.

## Actualizaciones
Al cambiar cualquier fichero, sube `VERSION` en `sw.js` para que los móviles recojan la nueva versión.
