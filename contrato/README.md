# Contrato — Organic Club × Ninro Libre (Kenia)

Contrato formal de prestación de servicios entre **Organic Club** (Liam Libre · Alonzo García) y
**LIBRE G PLATFORMS LLC** (Wyoming, EE. UU.), titular de la app Kenia, representada por Ninro Libre
(representante legal) y Kurt Libre (representante autorizado). Sustituye a la propuesta del 30 de agosto donde se contradigan y
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

Los datos de la empresa, de Ninro y de Kurt ya están en `datos-partes.json` (local, fuera de git). Faltan:

| Dato | De quién |
|---|---|
| CI y domicilio de Liam y de Alonzo (teléfono y correo, opcionales) | Nosotros |
| RIF de Organic Club, si existe | Nosotros |
| **Método de pago y cuenta** (Zelle, transferencia, USDT…) | Nosotros |
| **¿Ya se pagaron los US$ 1.000 de septiembre?** | Nosotros — va en el Anexo B |
| Teléfonos de Ninro y Kurt (opcional) | Ninro · Kurt |

## Lo que el contrato fija, y que conviene revisar antes del domingo

Son propuestas nuestras. Cada una se cambia en `datos-partes.json` → `condiciones`, o en el script.

| Condición | Propuesta | Dónde |
|---|---|---|
| Honorarios | US$ 1.000 septiembre (mes de fundador) · US$ 1.200/mes desde octubre | Cl. 6 |
| **Día de pago** | **Por adelantado, hasta el día 5 de cada mes** | Cl. 7 · Anexo B |
| Plazo inicial | 1 sep → 31 dic 2026 (US$ 4.600), luego renovación mes a mes | Cl. 5 |
| **Inicio de la publicación** | **9 de octubre de 2026** | Cl. 5.3 · Anexo A |
| Análisis de marketing restante + moodboard | Hasta el 31 de octubre | Anexo A.2 |
| Retraso | Suspensión del servicio tras 10 días sin pago; terminación tras 30 | Cl. 8 |
| **Volumen mensual** | Hasta 12 videos + 6 carruseles + stories | Cl. 3.5 |
| Aprobaciones | 2 días hábiles; si no hay respuesta, aprobación tácita un día después | Cl. 12 |
| Revisiones | 2 rondas por pieza | Cl. 12 |
| Preaviso de salida | 30 días | Cl. 21 |
| Jurisdicción | Caracas, ley venezolana | Cl. 25 |

**El Anexo A no es una lista de tareas:** dice qué se entregó, qué falta de la fase estratégica (con un
solo plazo, el 31 de octubre) y qué se entrega cada mes. El seguimiento día a día queda fuera del contrato.

## Las cuatro cláusulas que nos protegen, por si Ninro pregunta

1. **Cl. 13 — Dependencias.** Si Kurt o Ninro se retrasan, nuestro plazo se mueve lo mismo. Hoy hay
   cinco respuestas de Kurt pendientes desde septiembre (prueba gratuita, cobro, fecha, plan tras un
   fallo, analítica). Desde la firma, cada consulta técnica tiene **5 días hábiles** de respuesta (cl. 11.4).
2. **Cl. 14 — Metas de referencia.** Las 7 métricas de la propuesta pasan a ser objetivos, no
   garantías, y los 90 días cuentan **desde que la app está en tiendas**. Seis de las siete dependen
   de producto, precio y cobro, no de contenido.
3. **Cl. 4.1 — ASO.** Queda por escrito que estaba en la propuesta y que se acordó sacarlo.
4. **Cl. 11.6 — Veracidad del producto.** Lo que se publique sobre prevención de lesiones y salud es
   responsabilidad de quien describe el producto.

## Antes de firmar

- **No es asesoría legal.** Es un contrato de servicios estándar adaptado al proyecto. Si el monto o
  la relación crecen, que lo revise un abogado venezolano.
- **El cliente es una LLC de Wyoming**, pero el contrato se rige por la ley venezolana porque el
  servicio se presta en Caracas y los firmantes viven allí. Si Ninro prefiere otra jurisdicción, se
  cambia en la cláusula 25.
- Imprimir **dos ejemplares** (uno por parte); firman los cuatro e inicialar cada página: el pie ya trae el espacio.
