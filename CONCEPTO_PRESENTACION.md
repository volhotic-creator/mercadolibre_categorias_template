# Product Hunting App — Concepto para Junta de Socios y Owners

> Nota: este documento resume un proyecto independiente al árbol de categorías MeLi de este repo (`CONTEXT.md`). Se guarda acá como entregable de presentación a pedido del equipo.

## 1. El problema de negocio

El operador de depósito de Volhotic necesita relevar físicamente el stock real (qué hay, dónde está, en qué estado) contra lo que dice el sistema. Hoy ese relevamiento es manual, lento y sin trazabilidad: no hay registro fotográfico, no hay ubicación estructurada (piso/zona), y no escala a múltiples marcas sin duplicar trabajo.

**Restricción dura:** el operador dispone de **1:30h por día** para esta tarea ("Product Hunting"). Cualquier fricción en la herramienta le come tiempo productivo de caminata y carga real.

## 2. La solución: Product Hunting App

App mobile-first (Chrome en Android, sin instalación de store) que el operador agrega a su pantalla de inicio. Flujo de un solo toque por producto:

1. **Escanea/tipea el SKU** → el sistema muestra Marca + Descripción + Stock del sistema para validación visual.
2. **Confirma** que es el producto correcto (si el SKU colisiona entre marcas, elige cuál de las coincidencias es).
3. Completa **Piso, Zona, Estado (Encontrado/No Encontrado), Stock aproximado, Observación** y dos fotos (frente/dorso) tomadas directo con la cámara del celular.
4. Revisa un "ticket" resumen y confirma una sola vez → el sistema graba con **UPSERT** (no duplica si el operador vuelve a pasar por el mismo producto).
5. La pantalla se resetea sola para el siguiente SKU — cero clics extra.

## 3. Arquitectura (costo de infraestructura: $0 adicional)

- **Backend:** Google Apps Script sobre Google Sheets — mismo ecosistema que ya usa la empresa, sin servidores ni licencias nuevas.
- **Frontend:** página web liviana (HTML/JS) "instalable" desde Chrome, sin pasar por una app store.
- **Imágenes:** se suben a Google Drive, comprimidas en el celular antes de enviar (clave para la red Wi-Fi limitada del depósito).
- **Modelo de datos consolidado** (no una hoja por marca):
  - `Bronze_Catalogo`: catálogo maestro de todas las marcas, solo lectura desde la app.
  - `Silver_Registros`: lo que carga el operador, una fila por SKU+Marca, actualizable por UPSERT.
  - `Imagenes`: fotos indexadas por SKU/Marca, desacopladas del registro principal.

**Por qué importa para escalar:** sumar una marca nueva = pegar filas en una hoja. Cero código nuevo, cero deploy nuevo, cero mantenimiento adicional por marca.

## 4. Validación realizada (MVP probado, no es solo una idea)

- ✅ Probado en vivo por el consultor (Francisco) y el stakeholder (Matías).
- ✅ Proceso validado y declarado apto para producción.
- ✅ Caso real con 60 SKUs de la marca Bilstein, corriendo sobre Google Sheets en producción de prueba.
- ✅ Detección y resolución de problemas reales de calidad de datos (formatos de SKU inconsistentes, duplicados, colisiones de SKU entre marcas) antes de llegar a producción — evitando que el operador cargue datos contra el producto equivocado.

## 5. Extensión a multi-marca (en curso)

Se está procesando el archivo madre `vol_ml_listings.xlsx` (6.313 filas, todas las marcas) con un pipeline de QA que:

- No elimina ningún registro original.
- Normaliza y etiqueta cada fila (`OK`, `SKU_Pendiente`, `Conflicto_SKU_Marca`, `Duplicado_Exacto`, `Colision_Multi_Marca`) sin perder trazabilidad.
- Marca como "Pendiente en bodega" los SKUs vacíos o inválidos (~3.675 filas) — quedan excluidos del flujo operativo masivo por decisión de negocio, pero preservados en la planilla.
- Resultado: catálogo limpio y desambiguado, listo para que la app reconozca automáticamente la marca a partir del SKU, sin que el operador tenga que elegirla a mano.

## 6. Valor para el negocio

- **Tiempo del operador 100% en la tarea**, no en fricción de UI ni en doble carga.
- **Trazabilidad real:** quién cargó qué, cuándo, con foto de respaldo — algo que hoy no existe.
- **Dato accionable de inmediato:** reportes tipo "stock real vs. declarado" o "cobertura de relevamiento por piso/zona" son un JOIN directo, sin reprocesar nada.
- **Costo marginal de escalar a una marca nueva ≈ 0** (sin nuevo desarrollo, sin nueva infraestructura).
- **Costo de infraestructura ya pagado:** corre 100% sobre licencias de Google que la empresa ya tiene.

## 7. Próximos pasos propuestos

1. Cerrar la limpieza del catálogo madre (`vol_ml_listings.xlsx`) y confirmar criterio de exclusión de SKUs inválidos.
2. Reemplazar `Bronze_Catalogo` en la hoja en vivo y repetir el QA-Live con una segunda marca real (ej. Luk).
3. Definir el rollout a más operadores: alta como "test users" en el proyecto de Apps Script para evitar la pantalla de advertencia de Google en el primer uso.
4. Acordar cadencia de revisión de datos (`QA_Resumen`) para sostener la calidad del catálogo a medida que entren marcas nuevas.

---
*Documento de síntesis para presentación interna. Detalle técnico completo (código, README de configuración) disponible en la sesión de desarrollo original.*
