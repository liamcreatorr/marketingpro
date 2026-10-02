# Contrato — Organic Club × Ninro Libre (Kenia)

Contrato formal de prestación de servicios entre **Organic Club** (Liam Libre · Alonzo García) y
**Ninro Libre** por la app Kenia. Sustituye a la propuesta del 30 de agosto donde se contradigan y
cierra los dos desacuerdos que estaban abiertos desde septiembre: **el ASO queda fuera** y **el
lanzamiento es en octubre**, no el 14 de septiembre.

Preparado para la reunión del **domingo 4 de octubre de 2026** con Ninro y Kurt.

## Los archivos

| Archivo | Qué es |
|---|---|
| `Contrato Organic Club - Kenia (plantilla).docx` | El contrato con líneas en blanco para los datos personales. 25 cláusulas + 4 anexos |
| `generar-contrato.js` | Genera el contrato. `node contrato/generar-contrato.js` |
| `datos-partes.ejemplo.json` | Plantilla de datos: partes, contacto técnico y condiciones de pago |
| `datos-partes.json` | **Los datos reales (CI, teléfonos, direcciones). No se sube al repositorio** — está en `.gitignore` |

**Para generar la versión con datos:** copiar `datos-partes.ejemplo.json` como `datos-partes.json`,
rellenarlo y correr el script. Sale `Contrato Organic Club - Kenia-FIRMA.docx`, que tampoco se sube.

## Qué falta rellenar

| Dato | De quién |
|---|---|
| CI, nacionalidad, domicilio, teléfono y correo de Ninro | Lo que pasó Kurt |
| Si Kenia está a nombre de una empresa: razón social y RIF | Ninro |
| CI, domicilio, teléfono y correo de Liam y de Alonzo | Nosotros |
| RIF de Organic Club, si existe | Nosotros |
| Teléfono y correo de Kurt | Kurt |
| **Método de pago y cuenta** (Zelle, transferencia, USDT…) | Nosotros |
| **¿Ya se pagaron los US$ 1.000 de septiembre?** | Nosotros — va en el Anexo B |

## Lo que el contrato fija, y que conviene revisar antes del domingo

Son propuestas nuestras. Cada una se cambia en `datos-partes.json` → `condiciones`, o en el script.

| Condición | Propuesta | Dónde |
|---|---|---|
| Honorarios | US$ 1.000 septiembre (mes de fundador) · US$ 1.200/mes desde octubre | Cl. 6 |
| **Día de pago** | **Por adelantado, hasta el día 5 de cada mes** | Cl. 7 · Anexo B |
| Plazo inicial | 1 sep → 31 dic 2026 (US$ 4.600), luego renovación mes a mes | Cl. 5 |
| Retraso | Suspensión del servicio tras 10 días sin pago; terminación tras 30 | Cl. 8 |
| **Volumen mensual** | Hasta 12 videos + 6 carruseles + stories | Cl. 3.5 |
| Aprobaciones | 2 días hábiles; si no hay respuesta, aprobación tácita un día después | Cl. 12 |
| Revisiones | 2 rondas por pieza | Cl. 12 |
| Preaviso de salida | 30 días | Cl. 21 |
| Jurisdicción | Caracas, ley venezolana | Cl. 25 |

**Las fechas del Anexo A son nuestras y hay que validarlas entre Liam y Alonzo antes de enseñarlas:**
análisis de marketing completo el 13 de octubre, moodboard el 9, calendario de noviembre el 30 de
octubre. Firmarlas compromete. Mejor pedir dos días más ahora que incumplir en noviembre.

## Las cuatro cláusulas que nos protegen, por si Ninro pregunta

1. **Cl. 13 — Dependencias.** Si Kurt o Ninro se retrasan, nuestro plazo se mueve lo mismo. Hoy hay
   cinco respuestas de Kurt pendientes desde septiembre (prueba gratuita, cobro, fecha, plan tras un
   fallo, analítica). El Anexo A.3 les pone fecha: **martes 6 de octubre**.
2. **Cl. 14 — Metas de referencia.** Las 7 métricas de la propuesta pasan a ser objetivos, no
   garantías, y los 90 días cuentan **desde que la app está en tiendas**. Seis de las siete dependen
   de producto, precio y cobro, no de contenido.
3. **Cl. 4.1 — ASO.** Queda por escrito que estaba en la propuesta y que se acordó sacarlo.
4. **Cl. 11.6 — Veracidad del producto.** Lo que se publique sobre prevención de lesiones y salud es
   responsabilidad de quien describe el producto.

## Antes de firmar

- **No es asesoría legal.** Es un contrato de servicios estándar adaptado al proyecto. Si el monto o
  la relación crecen, que lo revise un abogado venezolano.
- **Ojo con la agenda:** el domingo 4 también es **Caracas Rock**, la acción de captación del Anexo A.
- Imprimir **tres ejemplares** (uno por firmante) e inicialar cada página: el pie ya trae el espacio.
