# FDC GUI — Curva de Duración de Caudales

Interfaz gráfica de escritorio (Tkinter) para calcular la **Curva de
Duración de Caudales (FDC)** a partir de series diarias de caudal en
formato Excel, y obtener los percentiles hidrológicos estándar:

- **Q2.75** — caudal excedido el 2.75 % del tiempo (caudal alto)
- **Q50** — caudal mediano (excedido el 50 % del tiempo)
- **Q97.25** — caudal excedido el 97.25 % del tiempo (caudal bajo/estiaje)

También calcula el caudal promedio de la serie y el promedio de los
caudales máximos y mínimos anuales de los años seleccionados.

## Características

- Selector de archivo Excel (`.xlsx`).
- **Ventana emergente de mapeo de celdas**: como el formato de los archivos
  de origen puede variar, antes de cargar los datos se muestra una vista
  previa de la hoja para que el usuario indique con un clic:
  - la celda con el **nombre de la estación**,
  - la celda de **encabezado de la columna de fechas**,
  - la celda de **encabezado de la columna de caudales**.
- Selección múltiple de años para incluir/excluir en el cálculo.
- Curva de duración (fórmula de Weibull) graficada en escala logarítmica,
  con marcadores de Q2.75 / Q50 / Q97.25 y resumen estadístico.
- Título de la gráfica con el nombre de la estación y los años analizados.

## Requisitos

- Python 3.9 o superior (incluye `tkinter`, disponible en la instalación
  estándar de Python en Windows/Linux/Mac).
- Dependencias listadas en `requirements.txt`.

## Instalación

```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux / Mac
source venv/bin/activate

pip install -r requirements.txt
```

## Uso

```bash
python fdc_gui.py
```

1. Haga clic en **"Abrir archivo..."** y seleccione un archivo `.xlsx`.
2. En la ventana emergente, elija el modo de selección (Nombre / Encabezado
   de Fechas / Encabezado de Caudales) y haga clic en la celda
   correspondiente de la vista previa. Confirme con **Aceptar**.
3. Seleccione los años a incluir en el panel izquierdo (por defecto se
   seleccionan todos).
4. Presione **"Calcular FDC"** para ver los percentiles y la gráfica.

## Formato de datos esperado

El archivo debe tener, en alguna hoja:
- Una celda con el nombre/código de la estación.
- Una columna de fechas diarias con su celda de encabezado.
- Una columna de caudales (m³/s) con su celda de encabezado, alineada por
  fila con la columna de fechas.

Los datos se leen hacia abajo desde la fila siguiente al encabezado hasta
la primera fila vacía.

## Licencia y descargo de responsabilidad

Este proyecto se distribuye bajo licencia [MIT](LICENSE).

**Importante:** lea el [descargo de responsabilidad](DISCLAIMER.md) antes
de usar esta herramienta en estudios técnicos, de diseño o cualquier
aplicación donde los resultados puedan tener implicaciones sobre
decisiones de ingeniería, seguridad o normativas.
