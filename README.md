# Rappel Automation

Sistema híbrido de automatización diseñado para procesar, categorizar y asignar facturas de acuerdos comerciales (Rapel) mediante visión de IA y cruce de datos con bases de datos de provisiones.

Este proyecto integra un script de procesamiento en Python y un flujo de trabajo en Power Automate para reducir los tiempos de procesamiento de **3 días a 20 minutos** por lote de facturas.

---

## Estructura del Repositorio

* **analisis_rapel.py**: Script principal en Python que automatiza la lectura de facturas en PDF, consume la API de OpenAI para extraer metadatos, aplica reglas de negocio locales y genera un reporte consolidado en Excel.
* **Protocolo Rapel.xlsx**: Archivo maestro de reglas que define los criterios de categorización de facturas, prioridades de asignación según antigüedad y códigos de artículo.
* **Rapel_Automation_Flow_Export.zip**: Flujo de Power Automate exportado listo para importar, que actúa como conector y disparador de los procesos de validación de datos.

---

## Flujo de Funcionamiento (Workflow)

1. **Extracción**: El sistema detecta los documentos en PDF y procesa la primera página extrayendo la glosa del servicio.
2. **Análisis Cognitivo**: Envía la imagen procesada a la API de OpenAI para realizar OCR inteligente y extraer descripciones comerciales complejas.
3. **Reglas de Negocio**: Clasifica la factura según las palabras clave de `Protocolo Rapel.xlsx`.
4. **Cruce y Asignación**: Busca al cliente en la BBDD de provisiones del año en curso y asigna el monto neto al mes correspondiente, controlando límites de provisión, antigüedad y estado.
5. **Output**: Exporta un archivo Excel (`Analisis.xlsx`) con formato condicional de colores indicando el estado del procesamiento (OK, Alerta de Provisión, Alerta de Antigüedad, Error).

---

## Configuración y Requisitos

### Requisitos del Sistema (Python)
Instalar las dependencias necesarias:
```bash
pip install openpyxl openai pypdfium2 pillow
