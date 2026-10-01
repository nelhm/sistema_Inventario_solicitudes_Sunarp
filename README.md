Resumen del Sistema de Control e Inventario de Títulos - SUNARP
1. Visión General del Sistema
El Sistema  web del Sistema de Control e Inventario de Títulos (Mesa de Partes) de la Superintendencia Nacional de los Registros Públicos (SUNARP) (Presentación del Servicio Web). Esta herramienta está diseñada para automatizar la recepción, inventario, seguimiento en custodia, generación de reportes y proceso de eliminación documentaria de expedientes registrales dentro del dominio institucional (sunarp.gob.pe).
2. Arquitectura y Tecnología
Entorno de Ejecución: Google Apps Script desplegado como Web App en producción (Código Backend).
Frontend: Interfaz de usuario dinámica en una sola página (SPA) desarrollada con HTML, Tailwind CSS, FontAwesome y Chart.js para visualización de indicadores (Interfaz Web).
Base de Datos: Hojas de cálculo de Google Sheets (Base de Datos SUNARP) con las estructuras:
INVENTARIO_TITULOS (16 columnas de datos registrales).
USUARIOS (gestión de acceso y roles).
HISTORIAL (bitácora de auditoría).
Rendimiento e Integridad: Operaciones de lectura/escritura atómicas en bloque, uso de LockService para prevenir duplicados por concurrencia y CacheService para agilizar la autenticación.
3. Módulos y Funcionalidades Principales
A. Autenticación y Control de Usuarios
Sistema de login con validación de estado activo/inactivo.
Roles diferenciados (Administrador y Ventanilla/Usuario) con asignación por área.
B. Dashboard de Indicadores (KPIs)
Métricas en tiempo real sobre títulos registrados, inscritos, tachados, en custodia y eliminados.
Gráficos interactivos de distribución mensual y estado del inventario.
Alerta automática de títulos vencidos que superan los 240 días hábiles de custodia.
C. Registro y Gestión de Títulos
Captura rápida de datos: Año, Número de Título, Tipo (Inscrito, Tachado, Tacha Especial, Tacha Sustantiva), Fechas clave (Presentación, Inscripción/Tachado, Recepción), Área Remitente y Servidor Recepcionista.
Navegación optimizada mediante teclado (tecla ENTER).
D. Consulta, Edición y Ficha Informativa
Buscador avanzado con filtros por año, tipo, estado y texto.
Ficha detallada por registro e historial de modificaciones.
E. Módulo de Eliminación Documentaria (Días Hábiles Perú)
Cálculo automático de días hábiles de custodia considerando sábados, domingos y feriados nacionales de Perú.
Procesamiento masivo de eliminación documentaria respaldado por número de oficio, fecha y autorizador.
F. Reportes y Auditoría
Generación de reportes detallados y resumidos exportables a formato imprimible/PDF.
Registro de auditoría continua (HISTORIAL) de cada acción del sistema, exportable a CSV.
Módulo de migración masiva de datos desde archivos Excel/CSV.
