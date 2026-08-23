# Marketing Skills — instalación y análisis

Fuente: https://github.com/coreyhaines31/marketingskills (v2.11.0, MIT, por Corey Haines).
Instalado en `.claude/skills/` (formato nativo de Claude Code — se activan solos según lo que pidas, sin necesidad de invocarlos por nombre) y `.claude/tools/` (registro de integraciones que esos skills usan).

## Qué es esto realmente

No es una app ni un plugin con UI: son **50 archivos `SKILL.md`**, cada uno con instrucciones, frameworks y checklists de un experto en un área de marketing concreta. Claude Code los lee automáticamente cuando detecta que la tarea encaja (por palabras clave en tu mensaje) y los aplica como si tuviera a ese especialista al lado. Es la diferencia entre pedirle a un modelo genérico "mejora esta landing" y pedírselo a un modelo que sigue un framework de CRO específico, con criterios de priorización, checklist de accesibilidad, y sabe qué preguntarte antes de opinar.

## Pieza central: `product-marketing`

Todos los demás skills leen primero `.agents/product-marketing.md` (producto, audiencia, posicionamiento, ICP). Se crea una sola vez — Claude puede redactarlo automáticamente leyendo tu código/landing/README — y a partir de ahí ningún otro skill te vuelve a preguntar "¿a qué se dedica tu producto?". Merece la pena crearlo antes que nada.

## Los 50 skills por categoría

**SEO y contenido** — `seo-audit`, `ai-seo` (optimizar para ser citado por ChatGPT/Perplexity/etc.), `programmatic-seo`, `site-architecture`, `competitors`, `schema`, `content-strategy`, `aso` (App Store/Play).

**CRO (conversión)** — `cro` (páginas y formularios), `signup`, `onboarding`, `popups`, `paywalls`.

**Contenido y copy** — `copywriting`, `copy-editing`, `cold-email`, `emails`, `social`, `image`, `video`, `sms`.

**Pago y distribución** — `ads` (Google/Meta/LinkedIn/X), `ad-creative` (generación masiva de variantes), `events`.

**Medición** — `analytics` (tracking GA4/Mixpanel/etc.), `ab-testing`, `attribution`.

**Retención y crecimiento** — `churn-prevention`, `co-marketing`, `free-tools`, `referrals`, `community-marketing`, `influencer-marketing`, `marketing-loops`.

**Estrategia y monetización** — `marketing-ideas` (140 ideas para SaaS), `marketing-psychology`, `launch`, `pricing`, `marketing-plan`, `marketing-council` (simula un consejo de asesores con varias perspectivas), `offers`, `public-relations`, `customer-research`, `competitor-profiling`.

**Ventas y RevOps** — `revops`, `sales-enablement`, `prospecting`, `directory-submissions`.

## Qué puedes lograr con esto en la práctica

- **"Ayúdame a mejorar esta landing"** → activa `cro`: analiza propuesta de valor, jerarquía visual, fricción del formulario, y te da recomendaciones priorizadas por impacto, no una lista genérica.
- **"Escribe la copy del homepage"** → `copywriting`, que ya conoce tu ICP porque leyó `product-marketing.md`.
- **"Monta un tracking de señales en GA4"** → `analytics` + `.claude/tools/integrations/ga4.md`, con eventos concretos y hasta un CLI (`.claude/tools/clis/ga4.js`) para ejecutarlo si tienes credenciales.
- **"Dame un plan de marketing completo"** → `marketing-plan`, estructura tipo AARRR (Adquisición-Activación-Retención-Referencia-Ingresos).
- **"¿Cómo evito el churn?"** → `churn-prevention`: cancelación con save offers, dunning de pagos fallidos.
- **"Necesito ideas"** → `marketing-ideas` (140 ideas específicas para SaaS) o `marketing-council` si quieres varias perspectivas de expertos simuladas debatiendo.
- Los skills se **encadenan solos**: `customer-research` alimenta a `copywriting`/`cro`/`competitors`; `revops` ↔ `sales-enablement` ↔ `cold-email`; `seo-audit` ↔ `schema` ↔ `ai-seo`.

## El registro de herramientas (`.claude/tools/`)

Cuando un skill necesita ejecutar algo real (no solo redactar), consulta `tools/REGISTRY.md`: más de 50 herramientas (GA4, Mixpanel, Stripe, HubSpot, Mailchimp, Google Ads, Meta Ads, Segment, PostHog, Ahrefs, Semrush...) con su método de integración — API, MCP, CLI propio en Node sin dependencias, o SDK — y una guía paso a paso en `tools/integrations/<herramienta>.md`. Para herramientas sin MCP nativo (HubSpot, Salesforce, Meta Ads, Google Sheets, Slack), usa Composio como capa intermedia (`tools/integrations/composio.md`).

**Importante**: estas integraciones necesitan que tú aportes las credenciales/API keys correspondientes cuando llegue el momento de ejecutar algo real (no vienen incluidas ni configuradas).

## Límites — qué esto NO hace

- No reemplaza cuentas ni acceso real a Google Ads, GA4, tu CRM, etc. — genera el plan/copy/tracking plan; ejecutarlo contra una API real requiere que conectes esas credenciales.
- No es un plugin con comandos de barra tipo `/cro` a menos que también instales el marketplace de Claude Code (`/plugin marketplace add coreyhaines31/marketingskills`); tal como está instalado (copia directa en `.claude/skills/`), se activa por contexto conversacional, que es el modo recomendado por el propio autor.
- Es contenido, no código de producto: no toca tu aplicación ni tu infraestructura salvo que se lo pidas explícitamente.

## Mantenimiento

Versión instalada: **2.11.0**. Para comprobar actualizaciones, compara `VERSIONS.md` del repo origen (https://raw.githubusercontent.com/coreyhaines31/marketingskills/main/VERSIONS.md) contra `.claude/skills/*/SKILL.md` (campo `metadata.version`). Para actualizar, vuelve a copiar los skills que hayan cambiado de versión.
