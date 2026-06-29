# Proyecto: MercadoLibre Argentina — Base de Datos de Categorías de Autopartes

## Objetivo

Construir un archivo Excel (`categorias_autopartes_meli.xlsx`) con el árbol completo de categorías de MercadoLibre Argentina para el rubro **Autopartes**, estructurado por niveles jerárquicos. Esta base de datos se cruzará luego contra un inventario de stock propio.

## Contexto de negocio

- **Operador:** E-commerce de autopartes
- **Plataforma destino:** MercadoLibre Argentina (MLA)
- **Fuente de verdad:** Inventario de stock de autopartes interno
- **Proceso:** Categorización manual asistida por capturas de pantalla del árbol de categorías de MeLi

## Estructura del archivo Excel

| Columna | Descripción |
|---|---|
| `Categoria_Nivel_1` | Categoría raíz (ej: "Accesorios para Vehículos") |
| `Categoria_Nivel_2` | Subcategoría (ej: "Autopartes y Accesorios") |
| `Categoria_Nivel_3` | Sub-subcategoría |
| `Categoria_Nivel_4` | Nivel 4 (si aplica) |
| `Categoria_Nivel_5` | Nivel 5 (si aplica) |
| `ID_Categoria_MLA` | ID interno de MeLi (si disponible) |
| `Notas` | Observaciones adicionales |

## Convenciones

- **Archivo principal:** `categorias_autopartes_meli.xlsx`
- **Script generador:** `generar_excel.py` (Python + openpyxl)
- **Fuente de datos:** Capturas de pantalla del árbol de categorías de MeLi Argentina
- **Encoding:** UTF-8
- **Idioma:** Español (Argentina)

## Flujo de trabajo

1. El operador envía capturas de pantalla con las categorías visibles en MeLi
2. Se transcriben las categorías al script Python manteniendo la jerarquía
3. Se regenera el Excel con `python generar_excel.py`
4. Se valida el archivo y se commitea al repositorio

## Archivos del proyecto

```
mercadolibre_categorias_template/
├── CONTEXT.md                    # Este archivo
├── categorias_autopartes_meli.xlsx  # Base de datos Excel (generada)
└── generar_excel.py              # Script Python generador
```

## Notas técnicas

- El script Python es la fuente de verdad del Excel; editar el `.py`, no el `.xlsx` directamente
- Cada ejecución sobreescribe el Excel para mantener consistencia
- Las categorías se ingresan manualmente desde capturas de pantalla ya que MeLi no expone el árbol completo via API pública para este nivel de detalle
