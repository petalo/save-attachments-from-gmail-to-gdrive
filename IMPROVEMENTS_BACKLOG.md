# Backlog De Mejoras (Idempotencia, Multiusuario y Consumo)

Fecha: 2026-03-10  
Contexto: Script de guardado de adjuntos Gmail -> Shared Drive

## Convención de IDs
- Prefijo: `GAS`
- Formato: `GAS-001`, `GAS-002`, ...
- Estado inicial para todos: `Propuesto`

## P0 (Críticas)

### GAS-001 - Liberar lock siempre con `finally`
- Prioridad: P0
- Objetivo: evitar locks huérfanos tras errores/interrupciones.
- Archivos: `src/Main.gs`, `src/Utils.gs`
- Criterio de aceptación: toda ejecución que adquiera lock lo libera en `finally`.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - `saveAttachmentsToDrive` libera lock en bloque `finally`.
    - El `releaseExecutionLock` se protege con `try/catch` para no romper el cierre.

### GAS-002 - Lock por usuario en vez de lock global
- Prioridad: P0
- Objetivo: reducir bloqueos entre usuarios distintos y mejorar concurrencia segura.
- Archivos: `src/Utils.gs`
- Criterio de aceptación: key de lock por usuario (`EXECUTION_LOCK_<email>`).
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Problema detectado: el lock global (`EXECUTION_LOCK`) serializa usuarios distintos y puede provocar esperas/bloqueos innecesarios.
  - Propuesta: generar lock key por usuario normalizado (p.ej. `EXECUTION_LOCK_diego_example_com`) y usar esa key tanto al adquirir como al liberar.
  - Alcance técnico: actualizar `acquireExecutionLock` y `releaseExecutionLock` para recibir/derivar la key por usuario.
  - Compatibilidad: durante la migración, limpiar lock global legado si existe y está expirado.
  - Riesgo principal: crecimiento de keys por usuario en Script Properties; mitigación: limpieza de locks expirados.
  - Riesgo adicional: si se aplica sin GAS-003, se puede aumentar concurrencia sin aislar correctamente el buzón objetivo.
  - Guardrail propuesto: implementar GAS-002 junto con GAS-003 (una ejecución = un buzón efectivo) para que el lock por usuario sea coherente.
  - Decisión preliminar: recomendable, pero no desplegar de forma aislada.
  - Implementación aplicada:
    - `acquireExecutionLock` y `releaseExecutionLock` ahora usan key por usuario.
    - Se añadió helper de key por usuario y limpieza del lock global legado expirado/inválido.
    - Se usa `LockService.getUserLock()` para no serializar usuarios distintos.

### GAS-003 - Procesar solo usuario actual o cola real de 1 usuario por ejecución
- Prioridad: P0
- Objetivo: evitar reprocesado múltiple del mismo buzón por ejecución.
- Archivos: `src/Main.gs`, `src/UserManagement.gs` (si se usa cola)
- Criterio de aceptación: una ejecución procesa solo un buzón efectivo.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Problema detectado: `saveAttachmentsToDrive` recorre todos los usuarios registrados y llama `processUserEmails(user)`, pero el acceso Gmail sigue en el contexto del ejecutor actual.
  - Evidencia: bucle en `src/Main.gs` (procesa `users.length`) y uso de `GmailApp` en `processUserEmails` sin cambio real de identidad.
  - Riesgo: trabajo duplicado, mayor consumo de cuotas, y falsa sensación de procesamiento multiusuario.
  - Propuesta (opción recomendada): por ejecución, procesar solo el usuario efectivo (`Session.getEffectiveUser().getEmail()`).
  - Propuesta alternativa: usar cola real (`getNextUserInQueue`) solo para control de turnos, pero manteniendo una ejecución = un buzón real del ejecutor.
  - Dependencia: GAS-002 (lock por usuario) se activa junto con esta decisión.
  - Implementación aplicada:
    - `saveAttachmentsToDrive` procesa solo `Session.getEffectiveUser().getEmail()`.
    - Se eliminó el bucle sobre usuarios registrados del flujo principal.

### GAS-004 - Sacar `searchVariations` del flujo normal
- Prioridad: P0
- Objetivo: ahorrar cuota y tiempo en Gmail.
- Archivos: `src/GmailProcessing.gs`
- Criterio de aceptación: búsquedas diagnósticas solo en modo debug.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se eliminaron `searchVariations` y búsquedas diagnósticas del flujo normal en `processUserEmails`.
    - Se añadió `diagnoseSearchVariations()` en `src/Debug.gs` para ejecutar ese diagnóstico manualmente.

### GAS-005 - Paginación de `GmailApp.search`
- Prioridad: P0
- Objetivo: evitar cargar backlog completo en memoria.
- Archivos: `src/GmailProcessing.gs`
- Criterio de aceptación: uso de `search(query, start, max)` y avance por páginas.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - `processUserEmails` usa `GmailApp.search(searchCriteria, 0, pageSize)`.
    - Se procesa exactamente una página por ejecución (sin bucle de lotes interno).
    - El avance se produce por etiquetado de threads procesados (`-label:Processed`), manteniendo runtime predecible.

### GAS-006 - No marcar `threadProcessed` si no hubo guardado real
- Prioridad: P0
- Objetivo: evitar falsos positivos de procesado.
- Archivos: `src/GmailProcessing.gs`
- Criterio de aceptación: `threadProcessed` solo se activa tras `savedFile` válido.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - `threadProcessed` solo se activa cuando `saveAttachmentLegacy` devuelve archivo.
    - Se añadió tracking de fallos de guardado para adjuntos válidos.

### GAS-007 - No etiquetar como `Processed` si hay fallos de guardado
- Prioridad: P0
- Objetivo: permitir reintento limpio de hilos fallidos.
- Archivos: `src/GmailProcessing.gs`
- Criterio de aceptación: hilos con error no terminan en label final de éxito.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Si hubo adjuntos válidos con fallos de guardado, el hilo no se etiqueta como `Processed`.
    - Si hubo adjuntos pero ninguno válido (todos filtrados), se mantiene el marcado como `Processed`.

### GAS-008 - Corte por tiempo (deadline) antes de timeout de Apps Script
- Prioridad: P0
- Objetivo: finalizar en estado consistente y reanudable.
- Archivos: `src/Main.gs`, `src/GmailProcessing.gs`
- Criterio de aceptación: parada controlada con checkpoint antes del límite de ejecución.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se añadió `executionSoftLimitMs` en configuración.
    - `saveAttachmentsToDrive` calcula `deadlineMs` y lo pasa al procesamiento.
    - `processThreadsWithCounting` corta el bucle de threads al alcanzar el límite blando y registra salida temprana.

## P1 (Idempotencia fuerte y recuperación)

### GAS-009 - Flujo por etiquetas `Processing` / `Processed` / `Error`
- Prioridad: P1
- Objetivo: modelar estados explícitos de cada hilo.
- Archivos: `src/GmailProcessing.gs`, `src/Config.gs`
- Criterio de aceptación: transición de estados consistente y trazable.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se añadieron `processingLabelName` y `errorLabelName` a configuración.
    - Al iniciar procesamiento de hilo se aplica label `Processing`.
    - En éxito/filtrado total: se aplica `Processed` y se limpia `Error`.
    - En fallo con adjuntos válidos o excepción: se aplica `Error`.
    - Se limpia `Processing` en bloque `finally` por hilo para evitar estados colgados.

### GAS-010 - Checkpoint por adjunto (id determinista)
- Prioridad: P1
- Objetivo: reanudar sin duplicar aunque se interrumpa en mitad.
- Archivos: `src/GmailProcessing.gs`, `src/AttachmentProcessing.gs`
- Criterio de aceptación: cada adjunto se identifica de forma estable y se salta si ya fue persistido.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se genera `sourceAttachmentId` determinista por adjunto.
    - Se indexa en `ScriptProperties` por `sourceAttachmentId + folderId -> fileId`.
    - Antes de guardar, se consulta índice para saltar adjuntos ya procesados.
    - Se guarda metadata `source_attachment_id` en la descripción del archivo.

### GAS-011 - Dedupe fuerte por hash/clave de origen (no nombre+tamañoKB)
- Prioridad: P1
- Objetivo: reducir duplicados en reintentos y concurrencia.
- Archivos: `src/AttachmentProcessing.gs`
- Criterio de aceptación: dedupe exacto por fingerprint estable.
- Estado: Aplazado por decisión funcional (2026-03-10)
- Notas de discusión:
  - Decisión: no implementar por ahora.
  - Motivo: baja probabilidad de colisión real (mismo nombre + mismo tamaño) en el uso actual.
  - Riesgo aceptado: podría colarse algún duplicado en casos poco frecuentes; se asume limpieza posterior si ocurre.

### GAS-012 - Registrar causa de fallo por hilo/adjunto
- Prioridad: P1
- Objetivo: distinguir errores recuperables de permanentes.
- Archivos: `src/GmailProcessing.gs`, `src/Utils.gs`
- Criterio de aceptación: logs/metadata con motivo de error utilizable en reintentos.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se añade estado persistente por hilo `THREAD_FAILURE_<threadId>` en Script Properties.
    - Se registra contexto/código/mensaje/adjunto y contador de intentos.
    - Se incluye helper de diagnóstico `inspectThreadFailureState(threadId)`.

### GAS-013 - Reintentos selectivos (transitorio vs permanente)
- Prioridad: P1
- Objetivo: evitar bucles infinitos y acelerar recuperación.
- Archivos: `src/Utils.gs`, `src/GmailProcessing.gs`
- Criterio de aceptación: errores permanentes no se reintentan indefinidamente.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Clasificación de fallos (`transient`, `permanent`, `too_large`) con `classifyProcessingFailure`.
    - Escalado automático a `permanent` al superar `maxThreadFailureRetries`.
    - Exclusión de hilos `permanent` de la búsqueda normal mediante label dedicado.

### GAS-014 - Recuperación de estados stale (`Processing`/lock antiguos)
- Prioridad: P1
- Objetivo: auto-recuperar ejecuciones cortadas (incluyendo pausas largas).
- Archivos: `src/Utils.gs`, `src/GmailProcessing.gs`
- Criterio de aceptación: limpieza segura de estados viejos con TTL configurable.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se añade checkpoint por hilo `THREAD_PROCESSING_<threadId>` con `timestamp` y `user`.
    - Al iniciar ejecución, `recoverStaleProcessingThreads` revisa una página (`staleRecoveryBatchSize`) de hilos con label `Processing` y limpia los que superan `processingStateTtlMinutes` o tienen estado inválido/ausente.
    - Se limpia siempre el checkpoint por hilo en `finally`.
    - Recuperación de lock stale ya cubierta por `executionLockTime` en `acquireExecutionLock`.

### GAS-015 - Función `resumeBacklog` para backlog histórico
- Prioridad: P1
- Objetivo: procesar acumulados de meses en lotes seguros.
- Archivos: `src/Main.gs`, `src/GmailProcessing.gs`
- Criterio de aceptación: reanudación incremental sin timeouts masivos.
- Estado: Propuesto
- Notas de discusión:

### GAS-016 - Etiqueta específica para `TooLarge`
- Prioridad: P1
- Objetivo: no mezclar “procesado con éxito” con “saltado por tamaño”.
- Archivos: `src/GmailProcessing.gs`, `src/Config.gs`
- Criterio de aceptación: hilos/adjuntos grandes quedan en estado explícito y auditable.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se añade label `tooLargeLabelName`.
    - Los hilos con adjuntos por encima de `maxFileSize` se marcan `TooLarge` y no `Processed`.
    - La búsqueda normal excluye `TooLarge` para evitar reprocesado infinito.

## P2 (Optimización operativa)

### GAS-017 - Reducir verbosidad de logs en modo normal
- Prioridad: P2
- Objetivo: ahorrar tiempo de ejecución y facilitar observabilidad útil.
- Archivos: `src/Utils.gs`, `src/GmailProcessing.gs`
- Criterio de aceptación: niveles de log configurables (`INFO` normal, `DEBUG` opcional).
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se añade `CONFIG.logLevel` y filtrado por nivel en `logWithUser`.
    - Se mueven logs de alta frecuencia (filtrado/adjuntos) a `DEBUG`.

### GAS-018 - Batch size dinámico por tiempo restante y error rate
- Prioridad: P2
- Objetivo: estabilizar throughput evitando timeouts.
- Archivos: `src/GmailProcessing.gs`
- Criterio de aceptación: ajuste de lote basado en telemetría de ejecución.
- Estado: Propuesto
- Notas de discusión:

### GAS-019 - Reporte por ejecución (KPIs)
- Prioridad: P2
- Objetivo: medir salud operativa y progreso de backlog.
- Archivos: `src/Main.gs`, `src/GmailProcessing.gs`
- Criterio de aceptación: resumen con vistos/guardados/duplicados/errores/too-large/reintentos.
- Estado: Propuesto
- Notas de discusión:

### GAS-020 - Proceso opcional de deduplicación/borrado en Drive
- Prioridad: P2
- Objetivo: limpiar duplicados históricos ya existentes.
- Archivos: nuevo módulo en `src/` (por definir)
- Criterio de aceptación: job seguro con dry-run y métricas antes de borrar.
- Estado: Propuesto
- Notas de discusión:

### GAS-021 - Definir estrategia multiusuario final (modelo único)
- Prioridad: P2
- Objetivo: eliminar ambigüedad entre “por usuario” y “orquestador global”.
- Archivos: `README.md`, `src/Main.gs`, `src/UserManagement.gs`
- Criterio de aceptación: arquitectura documentada e implementada sin mezcla de modelos.
- Estado: Implementado (2026-03-10)
- Notas de discusión:
  - Implementación aplicada:
    - Se valida `executionModel = "effective_user_only"` en configuración.
    - `getNextUserInQueue()` queda alineado al modelo efectivo (sin rotación runtime).
    - README actualizado para reflejar ejecución por usuario efectivo.

## Orden sugerido de implementación
1. GAS-001, GAS-003, GAS-004, GAS-005, GAS-006, GAS-007, GAS-008
2. GAS-002, GAS-009, GAS-010, GAS-011
3. GAS-012, GAS-013, GAS-014, GAS-015, GAS-016
4. GAS-017, GAS-018, GAS-019, GAS-020, GAS-021
