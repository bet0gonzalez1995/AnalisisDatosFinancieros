# Participaciones Municipales de Hidalgo (2025–2026)

Dashboard interactivo creado como practica para  el análisis, visualización y exploración de la distribución de fondos de participaciones a los 84 municipios del estado de Hidalgo.

---

## 🚀 Características Principales

- **Procesamiento 100% Local (Client-side):** Los datos se leen directamente en el navegador mediante SheetJS. Ninguna información ni archivo Excel se envía a servidores externos.
- **Resumen General y KPIs:** Visualización inmediata del total distribuido, número de municipios, promedio estatal y meses con mayor/menor distribución.
- **Evolución y Comparativas:** Gráficas interactivas con Chart.js para comparar el periodo enero–agosto entre 2025 y 2026.
- **Detalle por Municipio:** Consulta individualizada de los 84 municipios del estado de Hidalgo con ranking completo y variación porcentual interanual.
- **Análisis por Fondos:** Desglose porcentual y evolución temporal de los distintos fondos de participación (FGP, FFM, ISAN, ISR, entre otros).
- **Explorador de Datos y Exportación a PDF:** Tabla filtrable con exportación directa a PDF utilizando `jsPDF` y `jsPDF-AutoTable`.
- **Soporte de Tema Oscuro / Claro:** Adaptación automática a las preferencias del sistema operativo.

---

## 🛠️ Tecnologías y Librerías Utilizadas

Este proyecto funciona como un archivo único autocontenido (*Single-File Architecture*) e integra las siguientes librerías vía CDN:

- [HTML5 / CSS3 / JavaScript (ES6+)](https://developer.mozilla.org/)
- [Chart.js (v4.4.1)](https://www.chartjs.org/) — Para la renderización de gráficas interactivas.
- [SheetJS (xlsx v0.18.5)](https://sheetjs.com/) — Para la lectura y parseo de archivos Excel (`.xlsx`).
- [jsPDF (v2.5.1)](https://artskydj.github.io/jsPDF/) & [jsPDF-AutoTable (v3.5.12)](https://github.com/simonbengtsson/jsPDF-AutoTable) — Para la generación y exportación de reportes en PDF.

---

## 🚀 Cómo Usar el Proyecto

Dado que se trata de una aplicación web de archivo único, puedes ejecutarla de dos formas sencillas:

### Opción 1: Abrir localmente
1. Descarga o clona este repositorio.
2. Abre el archivo `Dashborard.html` directamente en cualquier navegador web moderno (Google Chrome, Mozilla Firefox, Microsoft Edge, Safari).
3. Selecciona y carga tus archivos de bases de datos correspondientes (por ejemplo, `BASE_PARTICIPACIONES_2025_CORRECTA.xlsx` y `BASE_PARTICIPACIONES_2026_CORRECTA.xlsx`) en el panel de inicio del dashboard.

---

## 📝 Decisiones de Metodología y Limpieza

- **Unificación de Nombres:** Se detectó una discrepancia en el municipio de *Juárez* (2025) frente a *Juárez Hidalgo* (2026), unificándose bajo el nombre correcto para permitir comparaciones precisas.
- **Fondo de Fomento Municipal (70%/30%):** Se excluyó de las sumas directas por tratarse de un desglose informativo y no de un concepto adicional, evitando duplicidades en los totales.
- **Periodo de Comparación:** Toda comparación interanual se realiza estrictamente sobre el periodo **enero–agosto** para mantener la congruencia entre el año completo 2025 y el acumulado disponible de 2026.

---
