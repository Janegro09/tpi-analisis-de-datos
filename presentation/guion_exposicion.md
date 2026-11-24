# Guion de Exposición Grupal — TPI Análisis de Datos (Telco Customer Churn)

- **Materia:** Introducción al Análisis de Datos (UTN)
- **Dataset:** `WA_Fn-UseC_-Telco-Customer-Churn.csv` → `data/processed/telco_churn_clean.csv`
- **Duración estimada:** 13 a 14 minutos (ideal para defensa oral de 12–15 minutos + espacio para preguntas)
- **Integrantes (según `README.md` y notebook):** Santiago Valentín Toledo · Javier Sosa · Juan Antuñez · Sebastian Safatle *(Nota: en `Planificación TP integrador Analisis de Datos.xlsx` también figura Alejo en tareas pendientes 7 y 10; si se suma a la defensa oral, puede tomar la Diapositiva 7 o la Diapositiva 11).*

## Distribución sugerida de tiempos y bloques

| Bloque | Integrante | Diapositivas | Tiempo estimado |
|---|---|---|---|
| **1. Negocio e Hipótesis** | Santiago Valentín Toledo | Slides 1, 2 y 3 | ~2 min 45 s |
| **2. Datos y Depuración (ETL)** | Javier Sosa | Slides 4 y 5 | ~3 min 00 s |
| **3. EDA y Hallazgos Clave** | Juan Antuñez | Slides 6, 7 y 8 | ~4 min 30 s |
| **4. Correlación, Riesgo y Cierre** | Sebastian Safatle *(+ cierre grupal)* | Slides 9, 10, 11 y 12 | ~3 min 45 s |
| **Total** | **4 integrantes** | **12 diapositivas** | **~14 min 00 s** |

---

## Diapositiva 1 — Portada
- **Título:** Retención de Clientes y Análisis de Fuga (*Churn*) en Telecomunicaciones
- **Subtítulo:** Trabajo Práctico Integrador — Introducción al Análisis de Datos (UTN)
- **Texto visual en slide:**
  - **Integrantes:** Santiago Valentín Toledo · Javier Sosa · Juan Antuñez · Sebastian Safatle
  - **Fecha:** Noviembre 2026
  - **KPIs de portada:** `7.043 clientes analizados` · `26,54 % tasa de churn` · `US$ 1,67 M ingresos anuales en riesgo`
  - **Repositorio reproducible:** GitHub (`Janegro09/tpi-analisis-de-datos`) & Google Colab
- **Sugerencia visual / Gráfico:** Portada en azul marino oscuro (`#0F2942`) con acentos en cian (`#00A8E8`), barra lateral de 3 tarjetas KPI de alto impacto y tipografía moderna sin serifas.
- **Notas del orador (Speaker Notes — Santiago Valentín Toledo · 0:30):**
  > Buenas noches a todos. Somos Santiago Toledo, Javier Sosa, Juan Antuñez y Sebastian Safatle, y hoy presentamos nuestro Trabajo Práctico Integrador de Análisis de Datos sobre la fuga de clientes —conocida como *Churn*— en una empresa de telecomunicaciones. A lo largo de los próximos minutos vamos a recorrer desde la fundamentación del problema de negocio y la depuración técnica de los datos hasta los hallazgos estadísticos, el perfil exacto del cliente en riesgo y cuatro acciones concretas para revertir la pérdida de ingresos.

---

## Diapositiva 2 — Contexto y Problema de Negocio
- **Título:** ¿Por qué el *Churn* es Crítico en Telecomunicaciones?
- **Subtítulo:** Perder clientes golpea los ingresos recurrentes y encarece la adquisición de nuevos abonados
- **Texto visual en slide:**
  - **KPI 1:** `US$ 456.117 / mes` — Facturación mensual total (MRR)
  - **KPI 2:** `US$ 139.131 / mes` — MRR perdido por *churn* (**30,50 %** de los ingresos)
  - **KPI 3:** `US$ 1,67 M / año` — Impacto anualizado en riesgo por bajas
  - **Problema de negocio:**
    - La retención de clientes define la rentabilidad en mercados altamente competitivos.
    - Cada baja elimina ingresos recurrentes y dispara el costo de adquisición (CAC).
    - El estudio busca detectar los detonantes históricos que anticipan la cancelación.
  - **La paradoja financiera (`Churn`):**
    - El **26,54 %** de los clientes concentra el **30,50 %** de la pérdida mensual.
    - Quienes cancelan pagan tarifas más altas: mediana **US$ 79,65** vs. **US$ 64,43**.
- **Sugerencia visual / Gráfico:** Tres tarjetas KPI superiores con contraste cromático (azul marino, coral de alerta y cian) + dos paneles comparativos inferiores (problema de negocio vs. impacto financiero).
- **Notas del orador (Speaker Notes — Santiago Valentín Toledo · 1:15):**
  > En telecomunicaciones, captar un cliente nuevo es mucho más caro que retener uno existente. Cuando analizamos la cartera de 7.043 clientes, vemos que la empresa factura cerca de 456 mil dólares mensuales, pero está perdiendo más de 139 mil dólares cada mes por cancelaciones. Eso equivale a más del 30 % de su facturación mensual, o 1,67 millones de dólares anualizados. Lo más preocupante es la paradoja financiera: aunque se va el 26,5 % de los abonados, se llevan el 30,5 % de los ingresos, porque quienes abandonan tienen tarifas mensuales sensiblemente más altas que los clientes que se quedan. Por eso elegimos `Churn` como nuestra variable objetivo binaria.

---

## Diapositiva 3 — Objetivos e Hipótesis Iniciales
- **Título:** Objetivos del Estudio e Hipótesis Analíticas
- **Subtítulo:** Tres hipótesis de negocio contrastadas con evidencia estadística sobre datos históricos
- **Texto visual en slide:**
  - **Objetivo general:** Identificar patrones contractuales, financieros, de servicio y demográficos asociados al abandono para proponer acciones de retención.
  - **H1 — Tipo de Contrato *(VALIDADA)*:**
    - Los contratos mes a mes presentan una fuga muy superior a contratos anuales/bianuales.
    - Evidencia: **42,71 %** en *Month-to-month* vs. **11,27 %** (1 año) y **2,83 %** (2 años).
  - **H2 — Cargos Mensuales y Antigüedad *(VALIDADA)*:**
    - Tarifas mensuales altas y baja permanencia incrementan la propensión al abandono.
    - Evidencia: bajas con mediana de **10 meses** (vs. 38) y **US$ 79,65/mes** (vs. US$ 64,43).
  - **H3 — Servicios de Valor Añadido *(VALIDADA)*:**
    - Contratar soporte técnico y seguridad online reduce la probabilidad de cancelación.
    - Evidencia: contar con estos servicios reduce el *churn* a **un tercio** (~42 % → ~15 %).
- **Sugerencia visual / Gráfico:** Franja superior con el objetivo central + 3 tarjetas verticales (H1, H2, H3) con etiqueta verde/turquesa de estado **VALIDADA** y la métrica clave destacada al pie de cada tarjeta.
- **Notas del orador (Speaker Notes — Santiago Valentín Toledo · 1:00):**
  > Antes de explorar los datos planteamos tres hipótesis concretas de negocio, documentadas en `fundamentacion_negocio.md`. Primero, que los contratos mes a mes sufren mucha más fuga que los compromisos a uno o dos años. Segundo, que la combinación de tarifas altas y poca antigüedad dispara las bajas. Y tercero, que los servicios de valor añadido —como seguridad en línea y soporte técnico— actúan como anclas de retención. Como adelantamos en las tarjetas, el análisis validó estadísticamente las tres hipótesis con diferencias muy marcadas. Ahora le doy la palabra a Javier para explicar cómo estructuramos el dataset y el proceso de limpieza.

---

## Diapositiva 4 — Origen, Arquitectura del Dataset y Metodología
- **Título:** Origen del Dataset y Arquitectura Reproducible
- **Subtítulo:** Datos estructurados tabulares gestionados bajo flujo colaborativo en GitHub y Google Colab
- **Texto visual en slide:**
  - **Dataset original (`Telco Customer Churn`):**
    - Formato tabular estructurado: **7.043 clientes únicos** × **21 atributos** originales.
    - **Demográficos (4):** género, adulto mayor, pareja, personas a cargo.
    - **Servicios (9):** telefonía, líneas múltiples, internet, seguridad, backup, soporte y streaming.
    - **Contrato y facturación (8):** antigüedad (`tenure`), contrato, método de pago, cargos y `Churn`.
  - **Arquitectura del proyecto y flujo metodológico:**
    - Jerarquía limpia: `data/raw/` (intacto) → `notebooks/` → `data/processed/` → `docs/`.
    - **Stack:** Python · Pandas · NumPy · Seaborn · Matplotlib · SciPy · Google Colab.
    - **Reproducibilidad:** carga dual (ruta local o URL *raw* de GitHub) y trabajo por roles.
- **Sugerencia visual / Gráfico:** Diseño en dos columnas: izquierda con 4 bloques temáticos de variables (clasificación cualitativa vs. cuantitativa) y derecha con diagrama de arquitectura del repositorio (`raw` → `notebook` → `processed`) y roles del equipo.
- **Notas del orador (Speaker Notes — Javier Sosa · 1:30):**
  > Trabajamos sobre el dataset estructurado *Telco Customer Churn*, compuesto por 7.043 filas —donde cada registro es un cliente único— y 21 columnas originales divididas en cuatro grupos: datos demográficos, servicios contratados, modalidad contractual y facturación, más la variable objetivo `Churn`. Desde el punto de vista ingenieril, organizamos el repositorio en GitHub separando estrictamente los datos crudos en `data/raw/`, que nunca se modifican, de los datos depurados en `data/processed/`. Además, el notebook detecta automáticamente si corre en local o en Google Colab y descarga el CSV desde GitHub si hace falta, garantizando reproducibilidad total para los cuatro roles definidos en la planificación del equipo.

---

## Diapositiva 5 — Data Wrangling y Limpieza (ETL)
- **Título:** Data Wrangling y Calidad de Datos (`raw` → `processed`)
- **Subtítulo:** Depuración justificada, ingeniería de atributos, tipado formal y control de outliers
- **Texto visual en slide:**
  - **Pipeline `raw` (7.043 × 21) → `processed` (7.043 × 27):**
    - **Espacios ocultos en `TotalCharges`:** 11 strings `" "` detectados con `.str.strip() == ""` (0,16 %).
    - **Patrón no aleatorio:** los 11 casos corresponden exactamente a clientes nuevos con `tenure = 0`.
    - **Conversión e imputación:** `pd.to_numeric(errors="coerce")` a `float64` + `fillna(0)` (sin facturación aún).
    - **Integridad:** 0 filas duplicadas, 0 `customerID` repetidos y categorías sin inconsistencias.
    - **Ingeniería (+6 variables):** `Churn_num`, `FacturacionTeorica`, `DesvioFacturacion`, `TotalServiciosAddons`, `RatioCostoAntiguedad`, `ModalidadCobro`.
    - **Tipado y Outliers (IQR):** `Contract` ordinal + `category` redujeron memoria en **86,1 %**; **0 outliers globales** (100 % de filas conservadas).
- **Sugerencia visual / Gráfico:** Diagrama de flujo ETL superior (Entrada Cruda → Depuración → Ingeniería y Tipado → Salida Limpia) + tabla comparativa "Antes vs. Después" + tarjeta destacada de `-86,1 % memoria (7.375 KB → 1.022 KB)`.
- **Notas del orador (Speaker Notes — Javier Sosa · 1:30):**
  > Al inspeccionar el archivo crudo encontramos un problema clásico y silencioso: `df.info()` mostraba cero nulos, pero `TotalCharges` aparecía como texto en lugar de número. Aplicando `.str.strip() == ""` detectamos 11 celdas con espacios en blanco. Al cruzarlas con el resto del dataset descubrimos que el patrón no era aleatorio: las 11 filas eran exactamente los 11 clientes con `tenure = 0`, recién dados de alta sin primera factura. Por eso, tras convertir la columna a `float64` con `pd.to_numeric`, elegimos imputar con `0` en lugar de borrar filas o usar la media, que sería conceptualmente incorrecto. Luego creamos `Churn_num` para habilitar correlaciones, sumamos cinco variables analíticas de negocio, convertimos las categóricas a tipo `category` y `Contract` a categórica ordinal —ahorrando un 86,1 % de memoria— y verificamos con IQR que no existen outliers globales. Así generamos `telco_churn_clean.csv` con 7.043 filas y 27 columnas sin perder un solo cliente. Ahora Juan continúa con el EDA.

---

## Diapositiva 6 — Análisis Exploratorio (EDA): Panorama General
- **Título:** EDA — Panorama General y Distribución de Variables Clave
- **Subtítulo:** Uno de cada cuatro clientes abandona; las bajas concentran menor antigüedad y mayor tarifa
- **Texto visual en slide:**
  - **Distribución del target (`n = 7.043`):**
    - **Clientes activos (`No`):** 5.174 (**73,46 %**) — MRR US$ 316.986.
    - **Clientes perdidos (`Yes`):** 1.869 (**26,54 %**) — MRR US$ 139.131.
  - **Contraste estadístico (Medianas — Activos vs. Bajas):**
    - **Antigüedad (`tenure`):** 38,0 meses vs. **10,0 meses** (moda global: 1 mes).
    - **Cargo mensual (`MonthlyCharges`):** US$ 64,43 vs. **US$ 79,65** (+23,6 %).
    - **Cargo total (`TotalCharges`):** US$ 1.679,52 vs. **US$ 703,55** (alta asimetría: 0,96).
  - **Boxplots e IQR:** sin outliers globales; los puntos altos en `Churn = Yes` son clientes premium antiguos que se fugaron.
- **Sugerencia visual / Gráfico:** A la izquierda, gráfico de dona de tasa de *Churn* (`05_churn_overview_donut.png`) + tabla compacta de medianas; a la derecha, el gráfico real exportado de la Sección 3.2 del notebook (`01_nb_boxplots_outliers.png`, Boxplots globales y segmentados por Churn).
- **Notas del orador (Speaker Notes — Juan Antuñez · 1:30):**
  > Entrando en el Análisis Exploratorio, el panorama general marca una tasa de abandono del 26,54 %, es decir, 1.869 clientes perdidos sobre 7.043. En los boxplots reales de nuestro notebook vemos dos comportamientos clave. Primero, a nivel global las tres variables cuantitativas no tienen outliers según la regla de Tukey, lo que confirma la calidad del dataset. Segundo, al segmentar por `Churn`, las distribuciones se separan con claridad: la mediana de permanencia de quienes se van es de apenas 10 meses frente a 38 meses en los clientes retenidos, mientras que su factura mensual mediana es casi 15 dólares más cara —79,65 contra 64,43—. Por la fuerte asimetría de `TotalCharges`, trabajamos con mediana e IQR como medidas más robustas.

---

## Diapositiva 7 — Hallazgos Clave 1: Factores Contractuales y Financieros
- **Título:** Hallazgos Clave 1 — Contrato, Pago y Presión Tarifaria
- **Subtítulo:** El contrato mes a mes y el cheque electrónico explican la inmensa mayoría de las bajas
- **Texto visual en slide:**
  - **Tipo de Contrato (principal predictor categórico):**
    - *Month-to-month* (55 % de la cartera) alcanza **42,71 %** de *churn* (**88,6 %** del total de bajas).
    - *One year* cae a **11,27 %** y *Two year* retiene al **97,17 %** (facturación histórica US$ 6,28 M).
  - **Modalidad de Pago (`ModalidadCobro`):**
    - *Electronic check* dispara la fuga al **45,29 %** (1.071 bajas = **57,3 %** del total).
    - El débito automático (tarjeta/transferencia) reduce el *churn* al **15,98 %** (vs. **34,67 %** manual).
  - **Facturación mensual:**
    - Clientes que pagan **> US$ 70/mes** presentan entre **32,8 % y 37,8 %** de abandono (vs. 10,9 % en < US$ 35).
- **Sugerencia visual / Gráfico:** Gráfico dual de barras horizontales generado con los datos del notebook (`06_contratos_y_pagos.png`): izquierda Tasa de Churn por Tipo de Contrato, derecha Tasa de Churn por Método de Pago, con línea punteada del promedio global (26,54 %).
- **Notas del orador (Speaker Notes — Juan Antuñez · 1:30):**
  > Cuando cruzamos el abandono con variables contractuales y financieras mediante `groupby` y tablas de contingencia, surgen los dos mayores detonantes del negocio. En contratos, el formato mes a mes genera el mayor flujo mensual —257 mil dólares—, pero pierde al 42,7 % de sus usuarios y concentra casi el 89 % de todas las cancelaciones de la empresa. En el extremo opuesto, los contratos a dos años retienen al 97,2 % y generan la mayor facturación histórica acumulada. Algo muy similar ocurre con el medio de pago: quienes abonan por cheque electrónico registran un 45,3 % de fuga —casi tres veces más que los adheridos a débito automático, que rondan el 16 %— porque cada factura manual renueva la fricción y la decisión de cancelar.

---

## Diapositiva 8 — Hallazgos Clave 2: Servicios y Antigüedad (`tenure`)
- **Título:** Hallazgos Clave 2 — Servicios de Valor y Antigüedad
- **Subtítulo:** El primer año y la fibra óptica sin soporte técnico son los puntos críticos de pérdida
- **Texto visual en slide:**
  - **Curva de Antigüedad (`tenure`):**
    - **0–12 meses:** **47,44 %** de *churn* (1.037 bajas = **55,5 %** de toda la pérdida).
    - Contrato mensual + ≤ 12 meses (1.994 clientes) escala al **51,35 %** de abandono.
    - Superados los 4 años (49–72 meses), la tasa cae al **9,51 %**.
  - **Tecnología y Servicios Protectores:**
    - **Fibra óptica:** **41,89 %** de *churn* (vs. **18,96 %** en DSL); combinada con contrato mensual llega al **54,61 %**.
    - **Seguridad Online:** sin servicio **41,77 %** vs. con servicio **14,61 %**.
    - **Soporte Técnico:** sin soporte **41,64 %** vs. con soporte **15,17 %**.
- **Sugerencia visual / Gráfico:** Gráfico compuesto de dos paneles (`07_servicios_y_antiguedad.png`): izquierda Tasa de Churn por tramo de antigüedad (`0–12 m` a `49–72 m`); derecha comparación directa de Fibra vs. DSL y efecto protector de `OnlineSecurity` y `TechSupport`.
- **Notas del orador (Speaker Notes — Juan Antuñez · 1:30):**
  > El segundo eje de hallazgos explica dónde y cuándo ocurre la fuga. En el gráfico de antigüedad vemos que el primer año es decisivo: el 47,4 % de los clientes con hasta 12 meses cancela, representando más de la mitad de todas las bajas. Si además están en contrato mensual, la fuga supera el 51 %. En cuanto a servicios, detectamos una señal de alerta fuerte en fibra óptica: pese a ser un producto premium con tarifa media de 87 dólares en plan mensual, presenta un 41,9 % de abandono general y 54,6 % en mes a mes. En contraste, los servicios de valor añadido funcionan como un escudo real: tener seguridad online o soporte técnico baja la cancelación del 42 % al 15 %, validando por completo nuestra tercera hipótesis. Le paso la palabra a Sebastian para el análisis de correlaciones.

---

## Diapositiva 9 — Evidencia Estadística: Correlación Lineal y Heatmap
- **Título:** Matriz de Correlación, Heatmap y Robustez Estadística
- **Subtítulo:** Validación cuantitativa con Pearson, p-valores (SciPy) y control no paramétrico con Spearman
- **Texto visual en slide:**
  - **Tabla de relaciones principales (`n = 7.043`, todas con `p < 0,001`):**
    - `tenure` – `TotalCharges`: **r = +0,826** ($R^2 = 68,3\%$) | Spearman **ρ = 0,890** *(Fuerte)*
    - `MonthlyCharges` – `TotalCharges`: **r = +0,651** ($R^2 = 42,4\%$) | Spearman **ρ = 0,638** *(Moderada)*
    - `tenure` – `MonthlyCharges`: **r = +0,248** ($R^2 = 6,1\%$) | Spearman **ρ = 0,276** *(Débil)*
  - **Asociación directa con el target (`Churn_num`):**
    - **`RatioCostoAntiguedad` (`r = +0,412`):** mejor predictor numérico (alto costo inicial).
    - **`tenure` (`r = −0,352`)** vs. **`MonthlyCharges` (`r = +0,193`)**: retener pesa más que subir tarifas.
  - **Rigor metodológico:** Spearman confirma las conclusiones; *correlación ≠ causalidad*; alerta de multicolinealidad en `TotalCharges` para ML.
- **Sugerencia visual / Gráfico:** A la izquierda, el gráfico real exportado de la Sección 6.3 del notebook (`03_nb_heatmap_ampliado.png`, Heatmap ampliado con `Churn_num`, `TotalServiciosAddons` y `RatioCostoAntiguedad`); a la derecha, tabla de significancia (Pearson, $R^2$, Spearman, p-valor) y hallazgos clave.
- **Notas del orador (Speaker Notes — Sebastian Safatle · 1:30):**
  > Para respaldar estos hallazgos con rigor estadístico construimos la matriz de correlación de Pearson y validamos la significancia con `scipy.stats.pearsonr`, obteniendo p-valores menores a 0,001 en todos los pares. En el heatmap ampliado de nuestra notebook se destacan tres ideas de negocio. Primero, `tenure` explica el 68 % de la variabilidad de `TotalCharges` ($r = 0,83$), mucho más que la tarifa mensual ($r = 0,65$): el valor de vida del cliente depende sobre todo de cuánto tiempo permanece. Segundo, nuestra variable calculada `RatioCostoAntiguedad` logra la mayor asociación numérica con el churn ($r = +0,41$), capturando en un solo indicador a los clientes nuevos que pagan caro. Tercero, contrastamos con Spearman para controlar asimetrías y no hubo cambios de magnitud ni signo. También dejamos documentadas dos advertencias: correlación no prueba causalidad, y la redundancia entre `TotalCharges`, `tenure` y `MonthlyCharges` deberá tratarse por multicolinealidad en futuros modelos predictivos.

---

## Diapositiva 10 — Perfil del Cliente en Riesgo
- **Título:** Radiografía del Cliente en Alto Riesgo vs. Cliente Fiel
- **Subtítulo:** La acumulación de factores eleva la tasa de abandono del 26,5 % hasta el 74,4 %
- **Texto visual en slide:**
  - **Escalada acumulada de riesgo (tasa de *churn*):**
    - Cartera global: **26,5 %** → Contrato mensual: **42,7 %** → + Fibra: **54,6 %** → + Cheque electrónico: **60,4 %** → + ≤ 12 meses: **71,2 %** → + Sin seguridad/soporte: **74,4 %** (544 clientes, 405 bajas).
  - **Perfil Crítico (Alto Riesgo):**
    - Contrato mes a mes · Fibra óptica · Pago por cheque electrónico.
    - Antigüedad ≤ 12 meses · Tarifa > US$ 70/mes · Sin seguridad ni soporte.
    - Factor demográfico: adultos mayores (**41,7 %** churn) y sin dependientes (**31,3 %**); género neutro (~26,5 %).
  - **Perfil Retenido (Bajo Riesgo — 2,8 % *churn*):**
    - Contrato a 2 años · Débito automático · > 3 años de antigüedad · Con soporte y seguridad.
- **Sugerencia visual / Gráfico:** A la izquierda, gráfico de escalada de riesgo (`08_escalada_riesgo.png`) con degradado de azul a coral intenso + barra verde benchmark de bajo riesgo (2,8 %); a la derecha, dos tarjetas comparativas ("Cliente en Riesgo" vs. "Cliente Fiel").
- **Notas del orador (Speaker Notes — Sebastian Safatle · 1:00):**
  > Al integrar todas las dimensiones, podemos trazar el retrato exacto del cliente en riesgo. Como muestra la escalada de la izquierda, cuando los factores se combinan el riesgo se multiplica: partimos de un 26,5 % global, saltamos a 42,7 % en contrato mensual, a 54,6 % si tiene fibra óptica, a 60,4 % si paga con cheque electrónico —un segmento que por sí solo explica más de 68 mil dólares mensuales perdidos—, a 71,2 % si está en su primer año y hasta un 74,4 % si tampoco cuenta con seguridad ni soporte. Además, en lo sociodemográfico el género es totalmente neutro, pero los adultos mayores alcanzan un 41,7 % de fuga. En la vereda opuesta, un cliente con contrato a dos años y débito automático apenas registra un 2,8 % de cancelación.

---

## Diapositiva 11 — Recomendaciones Accionables de Negocio
- **Título:** Plan de Acción: 4 Palancas para Reducir el Churn Hoy
- **Subtítulo:** Intervenciones focalizadas en los segmentos que concentran el mayor impacto financiero
- **Texto visual en slide:**
  - **1. Migración Contractual (*Month-to-month* → 1 o 2 años):**
    - Ofrecer descuentos escalonados al mes 3–6 y tarifa congelada al pasar a plan anual.
    - *Impacto:* ataca el segmento que concentra el **88,6 %** de las bajas.
  - **2. Incentivos a Débito Automático:**
    - Bonificación única o 5 % OFF al migrar desde *Electronic check* a tarjeta/transferencia.
    - *Impacto:* potencial de bajar la fuga del **45,3 %** hacia el **16 %**.
  - **3. *Bundle* de Retención en Fibra Óptica:**
    - Incluir *OnlineSecurity* y *TechSupport* bonificados 6 meses y auditar calidad técnica de fibra.
    - *Impacto:* protege la cartera de mayor ARPU (churn de **45,7 %** → **~15 %**).
  - **4. Onboarding Activo (Primeros 12 meses) y Segmento Senior:**
    - Seguimiento proactivo en los días 30/60/90, atención simplificada para adultos mayores y alertas por `RatioCostoAntiguedad`.
- **Sugerencia visual / Gráfico:** Matriz ejecutiva 2×2 con cuatro tarjetas numeradas (acento cian/marino), cada una estructurada en: **Palanca · Acción concreta · Impacto esperado**.
- **Notas del orador (Speaker Notes — Sebastian Safatle o Santiago Toledo · 1:15):**
  > Tradujimos estos hallazgos en cuatro acciones que la empresa puede activar de inmediato. Primero, una campaña de migración contractual para convertir clientes mes a mes a planes anuales a partir del tercer mes, atacando casi el 89 % del churn. Segundo, incentivar el paso de cheque electrónico a débito automático para eliminar la fricción de pago mensual. Tercero, empaquetar seguridad online y soporte técnico gratis los primeros seis meses en todas las altas de fibra óptica, además de revisar posibles fallas técnicas o de precio/calidad en ese servicio. Y cuarto, implementar un programa de onboarding durante el primer año con atención dedicada para adultos mayores y alertas tempranas basadas en el `RatioCostoAntiguedad`.

---

## Diapositiva 12 — Conclusiones, Reproducibilidad y Próximos Pasos
- **Título:** Conclusiones y Hoja de Ruta hacia Machine Learning
- **Subtítulo:** Proyecto 100 % reproducible validado en Colab, listo para escalar a predicción de fuga
- **Texto visual en slide:**
  - **Conclusiones de negocio:**
    - Las **3 hipótesis** fueron validadas estadísticamente; el *churn* es altamente concentrado y accionable.
    - Retener apenas al **20 %** de los clientes fugados recuperaría **~US$ 334.000/año** en ingresos.
  - **Auditoría de reproducibilidad (Sección 7):**
    - **12/12 controles automáticos (`assert`)** superados en kernel limpio y Google Colab.
    - Sin rutas absolutas ni `pip install`; dataset limpio (`7.043 × 27`) versionado en GitHub.
  - **Próximos pasos (Machine Learning):**
    - Entrenar modelos clasificadores (**Regresión Logística, Random Forest, XGBoost**) sobre `telco_churn_clean.csv`.
    - Tratar multicolinealidad (`TotalCharges`) y desbalance de clases (73/27) usando métricas *Recall / F1 / ROC-AUC*.
    - Desplegar un **scoring mensual de riesgo** integrado al CRM comercial.
- **Sugerencia visual / Gráfico:** Layout de 3 columnas limpias con íconos/encabezados diferenciados (Negocio en cian, Auditoría en turquesa con sello `12/12 ✔`, Hoja de Ruta ML en azul marino) + banda inferior de cierre para preguntas.
- **Notas del orador (Speaker Notes — Sebastian Safatle / Cierre Grupal · 1:00):**
  > Para cerrar, destacamos tres puntos. En lo analítico, validamos las tres hipótesis y demostramos que el abandono se concentra en patrones muy claros: reducir apenas un 20 % la fuga actual significaría recuperar más de 330 mil dólares anuales en ingresos. En lo técnico, realizamos una auditoría completa de reproducibilidad en Google Colab —superando 12 controles automáticos desde kernel limpio sin errores— y dejamos el dataset procesado listo en el repositorio. Como próximo paso natural, este dataset limpio ya permite entrenar modelos supervisados de Machine Learning —como Regresión Logística o XGBoost— manejando la multicolinealidad de `TotalCharges` y el desbalance de clases para generar un puntaje mensual de riesgo en el CRM y actuar antes de que el cliente pida la baja. Muchas gracias por su atención; quedamos a disposición para sus preguntas.
