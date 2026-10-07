# CRM Grupo Pressacco v2 — capacidad instalada por departamento

## Cómo probarlo sin riesgo
1. En GitHub creá una **rama nueva** (por ejemplo `v2`) y subí estos archivos. No toques `main`.
2. En Streamlit Cloud: *New app* → rama `v2` → archivo principal `app.py`.
   Copiá los mismos **Secrets** de la app actual (`gcp_service_account` y `passwords`).
3. Podés usar el mismo Google Sheet (`CRM_Estudio_Datos`): la v2 solo **agrega** 4 pestañas
   (`Catalogo`, `Personas`, `Registros`, `Feriados`). `Cargas` y `Competencias` no se tocan.
4. La primera vez que entra, la app crea las pestañas y columnas que falten y carga el catálogo,
   las personas y los feriados 2026.

## Qué revisar después del primer arranque
- **Personas**: Nombre, Rol (Operario/Admin), Activo (SI/NO), HorasDia. Para sumar a alguien: una fila nueva
  y su contraseña en Secrets (`[passwords]`), con el mismo nombre.
- **Catalogo**: se edita acá (Departamento, Tarea, Subtarea, Tipo, Orden, Activo). Tipo: `Trabajo`,
  `Disponible` o `Ausencia`.
- **Feriados**: formato AAAA-MM-DD. Hoy solo están los de 2026; agregá los de 2027.

## Pruebas
    python -m unittest discover -s tests -v

## Todavía no está (próxima etapa)
Migración de las cargas viejas, PDFs, competencias, desvío vs. promedio histórico, tendencia y
distribución semanal, y el objetivo mensual.
