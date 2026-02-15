# ESCUELA SUPERIOR DE INNOVACIÓN Y TECNOLOGÍA

## TSU en Servicios en la Nube

## Estancia Profesional

## Grupo SN-21

### Integrantes y roles
- Francisco Elías Cisneros Castro - Líder del Proyecto Jr. / Analista de Procesos Académicos Jr.
- Anibal Ernesto Martínez Cruz - Desarrollador Fullstack Jr.
- Gustavo Adolfo Reyes Vanegas - QA / Documentador Técnico Jr.
- David Mauricio González Estupinián - Ingeniero Cloud Jr.
- Héctor Benjamín Chipagua Calles - Recurso inactivo (por falta de respuesta/participación)

## Proyecto: Plataforma de Registro y Asistencia en la Nube (Cloud Attendance Lite)

### Duración total de proyecto:
10 semanas, organizadas en Fase 0 a Fase 4, según el Plan de trabajo semanal (visión operativa) que abarca desde el onboarding hasta la entrega profesional con demo.

### Fases o etapas del proyecto
- Fase 0: Onboarding y planificación ✔️
- Fase 1: Análisis del proceso de registro y asistencia ✔️
- Fase 2: Diseño e implementación de la plataforma ✔️
- Fase 3: Pruebas, reportes y refinamiento ✔️
- Fase 4: Cierre, documentación y entrega profesional

### Tipo de aplicación
La solución es una plataforma web/cloud ligera de tipo SaaS académico, orientada a:
- Registrar eventos y sesiones (ej. “Taller Nube 101 – Sesión 1”).
- Registrar participantes (nombre, correo, identificación básica).
- Registrar asistencia por sesión (presente/ausente).
- Consultar reportes simples (lista de asistentes, porcentaje de asistencia por evento/sesión).

### Propósito
Resolver la necesidad de la Dirección de Gestión Académica Digital (DGAD) de contar con una plataforma ligera, centralizada y basada en la nube para:
1. Registrar eventos y sesiones académicas.
2. Registrar y mantener un historial de participantes.
3. Registrar asistencia por sesión.
4. Generar reportes básicos de participación y asistencia sin depender de hojas de cálculo manuales enviadas por correo.

### Objetivo
Gestionar:
- Eventos
- Sesiones
- Participantes
- Asistencias (por sesión)

### Infraestructura propuesta (según sección 9.2.3)
- **Base de datos:** Google Sheets
- **Automatización:** Google Apps Script (Web App + onFormSubmit)

### Estructura
- `/db` -> Base de datos
- `/docs` -> Documentación del proyecto
- `/frontend` -> UI web (prototipo)
- `/tests` -> Base de datos en XLSX
