# Clasificador Tributario · Sullair

App para clasificar automáticamente facturas de alquiler de equipos según jurisdicción tributaria.

## Tres categorías

| Categoría | Criterio |
|---|---|
| **Añelo** | GPS ≤ 15 km del centro de Añelo · Sin GPS: Añelo, Bajo Añelo, Bajada de Añelo, BAN, Tratayen, Loma Campana |
| **Neuquén** | GPS dentro del ejido municipal · Sin GPS: barrios, PIN, Zona 1, instalaciones de la ciudad |
| **Neuquén/Otros** | Todo lo demás |

## Límites del ejido de Neuquén Capital

- **Norte**: límite con Centenario (lat -38.830)
- **Sur**: Río Limay (lat -38.990)
- **Oeste**: límite con Plottier (lon -68.155)
- **Este**: Río Neuquén / Cipolletti (lon -68.015)

## Cómo usar

1. Accedé al link de la app
2. Subí el archivo Excel de facturas
3. Descargá el resultado clasificado

## Estructura esperada del Excel

| Col 1 | Col 2 | Col 3 | Col 4 | Col 5 |
|---|---|---|---|---|
| Nº Factura | Lug. Trabajo | Añelo | Ciudad de Neuquén | Otro |
