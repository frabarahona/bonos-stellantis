# Bonos Stellantis · Automotora Alameda

Dashboard para controlar los bonos que la automotora cobra a Stellantis y a Santander por cada venta de Peugeot, Opel y Citroën.

## Qué hace

- **Ventas y bonos**: carga el informe de ventas (.xls) y cruza cada venta con la lista de precios del mes de facturación. Calcula bono marca (aporte Stellantis), aporte Santander (CC/CI), bono flota (tabla B2B) y bonos adicionales (SELL IN u otro). Cada bono se marca como pagado.
- **Estadísticas**: ventas, facturado, margen estimado, bonos, ventas por marca vs meta, tipo de venta, financiamiento y top vendedores.
- **Listas de precios**: historial por mes de precios, bonos TMP/CC/CI, aportes y descuentos B2B.
- **Cargar archivos**: informes y listas por separado (archivos o carpetas completas), con opción de borrar y volver a cargar.

## Cómo está armado

- Página estática: `index.html` (un solo archivo). Se abre en el navegador o se publica en Vercel / GitHub Pages.
- Datos en Supabase (proyecto `prnrnsgkdjfotkanfwul`). Acceso con usuario y clave; la seguridad está en las políticas RLS de la base (solo los correos de la tabla `usuarios` leen y escriben).
- Librerías por CDN: SheetJS (lectura de Excel) y supabase-js.

## Editar

El código fuente está en `src/`:

- `src/core.js`: lectura de listas, B2B e informe; cálculo de bonos.
- `src/app.html`: interfaz.

Después de editar, genera `index.html` con:

```
python3 build.py
```

## Reglas de cálculo

- Aporte red = mín(bono TMP / 2; 2,5 % del precio lista). Aporte Stellantis = bono TMP − aporte red. Aporte financiera = bono CC/CI − bono TMP.
- Formas de pago: Contado, CC, CI, Flota Tramo 1, Flota Tramo 2, Flota GC, Flota CC, Flota CI.
- Santander se suma cuando el Tipo Venta o la Entidad Crédito es Santander.
- Flota: % Stellantis de la tabla B2B (concesionario Professional) × precio lista neto × 1,19.
- Diferencia = bono del informe − (bonos calculados + adicionales).
