# marketingpro

Repositorio para trabajar tareas de marketing con Claude Code, usando el paquete de **50 Marketing Skills** de [Corey Haines](https://corey.co) ([coreyhaines31/marketingskills](https://github.com/coreyhaines31/marketingskills)).

## Qué hay instalado

- `.claude/skills/` — 50 skills (archivos `SKILL.md` + referencias) que Claude Code carga automáticamente según el contexto de la conversación.
- `.claude/tools/` — registro de herramientas de marketing (GA4, Mixpanel, Stripe, HubSpot, etc.) con guías de integración y CLIs sin dependencias que los skills usan cuando hace falta ejecutar algo real.

Ver [`MARKETING_SKILLS.md`](MARKETING_SKILLS.md) para el detalle de qué hace cada skill, cómo se relacionan entre sí, y cómo mantenerlas actualizadas.

## Primer paso recomendado

Antes de usar cualquier otro skill, pídele a Claude:

```
Crea mi contexto de producto (product-marketing)
```

Esto genera `.agents/product-marketing.md`, el documento de posicionamiento/audiencia que el resto de los skills leen automáticamente para no repetir contexto en cada tarea.

## Licencia

Los skills se distribuyen bajo licencia MIT — ver `.claude/skills-LICENSE`. © Corey Haines.
