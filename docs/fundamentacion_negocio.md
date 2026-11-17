## 1. Fundamentación de Negocio e Hipótesis Analíticas

### Definición del Problema de Negocio
En el competitivo mercado de las telecomunicaciones, la retención de clientes (*Customer Retention*) es un factor crítico para la rentabilidad de la compañía. La pérdida de abonados (*Churn*) no solo representa una pérdida directa de ingresos recurrentes, sino también un incremento en los costos de adquisición de nuevos usuarios. El objetivo central de este estudio es analizar el comportamiento histórico de los clientes de la compañía de telecomunicaciones a partir de sus características sociodemográficas, servicios contratados y modalidades de pago, con el fin de identificar los principales detonantes del abandono del servicio.

### Variable Objetivo (Target)
* **Variable:** `Churn` (Abandono del cliente).
* **Justificación:** Es la variable clave del estudio. Permite clasificar de forma binaria si un cliente canceló sus servicios durante el período analizado o si continúa activo, constituyendo el eje sobre el cual se construirán los análisis estadísticos y las correlaciones de negocio.

### Hipótesis Analíticas Preliminares
1. **Tipo de Contrato:** Los clientes con contratos mes a mes (*Month-to-month*) presentan una tasa de abandono significativamente mayor en comparación con aquellos que poseen contratos anuales o bianuales.
2. **Cargos Mensuales y Antigüedad:** Los clientes con cargos mensuales elevados (*MonthlyCharges*) y baja antigüedad (*tenure*) muestran una mayor propensión al *churn*.
3. **Servicios de Valor Añadido:** La contratación de servicios adicionales de soporte técnico y seguridad en línea reduce la probabilidad de cancelación del servicio.

---

## 2. Clasificación del Dataset y Relevamiento de Variables

### Estructura de los Datos
El dataset **Telco Customer Churn** se clasifica como un conjunto de **datos estructurados**, organizado en formato tabular (filas y columnas) donde cada fila representa un cliente único y cada columna un atributo o característica observable.

### Relevamiento y Contraste con Tipos de Datos (dtypes)
A nivel conceptual, las variables se dividen en:
* **Variables Cualitativas (Categóricas):**
  * *Nominales:* Género (`gender`), servicio telefónico (`PhoneService`), tipo de pago (`PaymentMethod`), etc.
  * *Ordinales:* Nivel de permanencia o categorías derivadas.
* **Variables Cuantitativas (Numéricas):**
  * *Discretas:* Antigüedad en meses (`tenure`).
  * *Continuas:* Cargos mensuales (`MonthlyCharges`) y cargos totales (`TotalCharges`).

*Nota metodológica para el Notebook:* Durante la ingesta con Pandas, se validará el contraste entre la naturaleza lógica de las variables y los tipos de datos asignados por defecto (`dtypes`), prestando especial atención a la conversión de la columna `TotalCharges` (que suele requerir tratamiento por espacios en blanco y formato de tipo objeto a numérico).
