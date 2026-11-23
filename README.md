# tpi-analisis-de-datos

# TPI - Introducción al Análisis de Datos

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Janegro09/tpi-analisis-de-datos/blob/main/notebooks/TPI_Analisis_de_Datos.ipynb)

</div>


##  Descripción del Problema
Este proyecto tiene como objetivo analizar el fenómeno de la **fuga de clientes (*Churn*)** en una empresa de telecomunicaciones utilizando el dataset `WA_Fn-UseC_-Telco-Customer-Churn.csv`. Buscamos identificar los patrones de comportamiento, las variables demográficas y los factores contractuales que incrementan la probabilidad de que un cliente abandone los servicios de la compañía, permitiendo formular recomendaciones estratégicas para la retención.


##  Nómina de Integrantes y Roles
* **Santiago Valentin Toledo** — Redacción de README, fundamentación de negocio y clasificación de variables.
* **Javier** — Setup inicial del repositorio, estructura de carpetas y limpieza de datos.
* **Sebastian Safatle** — Matriz de correlación, heatmap y auditoría de reproducibilidad.
* **Juan Antuñez** — Manipulación de datos, transformaciones, análisis de outliers (IQR/Boxplots) y Análisis Exploratorio de Datos (EDA).


##  Stack Tecnológico
El proyecto se desarrolla íntegramente en **Python**, utilizando las siguientes librerías de análisis y visualización de datos:
* **Pandas & NumPy:** Manipulación, limpieza y estructuración de datos tabulares.
* **Seaborn & Matplotlib:** Generación de gráficos estadísticos (histogramas, KDE, mapas de calor).
* **SciPy:** Cálculo de la significancia estadística (p-valor) de las correlaciones de Pearson.
* **Google Colab:** Entorno de ejecución interactivo para garantizar la reproducibilidad.


##  Estructura del Repositorio
```text
├── data/
│   ├── raw/                 # Dataset original sin procesar (WA_Fn-UseC_-Telco-Customer-Churn.csv)
│   └── processed/           # Dataset depurado y transformado
├── notebooks/               # Jupyter / Colab notebooks oficiales del análisis
├── docs/                    # Documentación formal e informes técnico-ejecutivos (PDF)
├── .gitignore
└── README.md
```
