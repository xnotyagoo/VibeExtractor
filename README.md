VibeExtractor: Extractor de Grupos para Exchange
Un script de PowerShell interactivo para auditar y extraer miembros de grupos de distribucion (estaticos y dinamicos) en entornos Microsoft 365 y Exchange.

Nota: Este codigo ha sido 100% vibecodeado. Pura intuicion, flujo directo en la terminal y cero estres.

Funcionalidades principales
Busqueda inteligente: No necesitas el nombre exacto. Escribe una palabra clave y selecciona el grupo correcto de una lista numerada.

Expansion recursiva: Detecta si hay grupos dentro de otros grupos y los abre todos recursivamente hasta extraer unicamente a los usuarios finales, eliminando duplicados.

Enriquecimiento de datos a la carta: Antes de exportar, te permite elegir si quieres anadir informacion extra al reporte: Estado de la cuenta (Activo/Deshabilitado), Manager, Empresa y Oficina.

Extraccion de codigo: Permite descargar las reglas de filtrado (OPath) de los grupos dinamicos en un archivo de texto.

Reportes pulidos: Genera un CSV en el escritorio con los datos limpios (sin espacios residuales) y ordenados alfabeticamente, listo para entregar.

Interfaz amigable: Muestra barras de progreso visuales durante las busquedas masivas para que sepas exactamente en que punto se encuentra la extraccion.

Requisitos
Solo necesitas tener una sesion iniciada contra tu entorno de Exchange (por ejemplo, usando Connect-ExchangeOnline) en tu consola de PowerShell antes de ejecutar la herramienta.

Uso rapido (Sin instalacion)
Puedes ejecutar este script directamente en la memoria de tu consola sin tener que descargar el archivo .ps1. Ejecuta el siguiente comando, reemplazando la URL por tu enlace RAW de GitHub:
