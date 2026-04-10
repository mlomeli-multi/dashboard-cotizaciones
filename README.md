# Dashboard de Cotizaciones MLTI

Dashboard comercial en HTML para analizar cotizaciones y cierres del Departamento de Agentes de MULTI durante enero y febrero de 2026.

## Contenido

- `index.html`: dashboard interactivo con 6 pestañas.
- Proyecto estático, sin backend ni proceso de instalación.
- Puede publicarse fácilmente en GitHub Pages.

## Contexto del dashboard

Este reporte fue construido para revisar el desempeño comercial del equipo de Agentes, incluyendo:

- volumen total de cotizaciones
- cierres reales por mes
- hit rate global y por servicio
- eficiencia por cliente
- eficiencia por ejecutivo
- pérdidas operativas y tendencias entre enero y febrero 2026

## KPIs principales

- 718 referencias válidas
- 143 cierres reales totales
- hit rate global aproximado de 19.9%
- 421 cotizaciones activas sin respuesta
- 48 pérdidas operativas

## Cómo abrirlo localmente

Abre [index.html](C:\Users\Miguel Lomeli\OneDrive - Multitraslados Internacionales SA de CV\Documentos\New project\dashboard-cotizaciones-mlti\index.html) en tu navegador.

## Cómo actualizar la data

1. Abre el dashboard en tu navegador.
2. En la parte superior, carga el archivo de `cotizaciones`.
3. Si lo tienes, carga también el archivo de `cierres`.
4. Haz clic en `Procesar archivos`.
5. El dashboard recalculará KPIs, tablas, clientes, equipo y tendencia.

### Encabezados recomendados

Archivo de cotizaciones:

- `DIA` o `Fecha`
- `Referencia`
- `Cliente`
- `Servicio`
- `Usuario`
- `Estatus`

Archivo de cierres:

- `Fecha`
- `Cliente`
- `Servicio`
- `Usuario`
- `Referencia` opcional

## Notas de funcionamiento

- El dashboard acepta `.xlsx`, `.xls` y `.csv`.
- Si una celda trae varios servicios o varios usuarios, intenta separarlos automáticamente.
- La última carga exitosa se guarda en el navegador del equipo donde se abrió el dashboard.
- Si no subes el archivo de cierres, el dashboard seguirá funcionando, pero los cierres y hit rates quedarán incompletos.

## Publicación en GitHub Pages

1. Sube el proyecto al branch `main`.
2. En GitHub entra a `Settings > Pages`.
3. En `Build and deployment`, elige `Deploy from a branch`.
4. Selecciona `main` y la carpeta `/root`.
5. Guarda los cambios.

## Nota

El HTML original fue generado localmente y aquí se conserva como base para versionarlo, respaldarlo y publicarlo.
