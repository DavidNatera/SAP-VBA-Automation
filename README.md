# SAP-VBA-Automation
Desarrollo de macros en Excel y scripts de VBA integrados con SAP GUI Scripting para la extracción automatizada de datos, procesamiento de flujos de trabajo y optimización de reportes de producción en SAP ECC 6.0 (EHP7)
---

## 🛠️ Decisiones de Arquitectura y Diseño Técnico

Durante el desarrollo y la evaluación con agentes de IA autónomos (Google Jules), se analizaron propuestas de optimización para el ciclo principal de notificaciones en la transacción `CO11N`. A continuación, se detalla el criterio de ingeniería aplicado:

> ### ⚠️ Mitigación de Riesgos Operativos vs. Velocidad de Ejecución
> * **Persistencia de Datos en la GUI:** Se evaluó la posibilidad de omitir el reinicio de la transacción (`Session.StartTransaction "CO11N"`) en cada iteración del bucle para ganar velocidad. Sin embargo, en entornos productivos de **SAP ECC 6.0**, la interfaz suele retener datos remanentes en los buffers de pantalla (ej. números de personal, textos o variantes previas). 
> * **Decisión:** Se optó por mantener el refresco explícito por cada fila del ciclo. Esto garantiza un entorno visual 100% estéril para cada orden de fabricación, eliminando el riesgo de "falsos positivos" o cruce involuntario de datos entre órdenes consecutivas, priorizando la **integridad de los datos** sobre la velocidad pura.
>
> * **Manejo de Advertencias del Sistema (`"W"`):** El sistema SAP emite avisos amarillos de advertencia (tolerancias de fechas, excesos de cantidad, etc.) que requieren validación analítica. Automatizar una confirmación ciega (`.sendVKey 0`) puede forzar el registro de datos erróneos que el planificador de producción debe retener para revisión manual.
