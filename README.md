# Avermant Solutions - Sistema de Gestión de Flota (VBA)

Este proyecto es una herramienta de **Gestión de Mantenimiento Aeronáutico (CAMO)** desarrollada en Excel VBA. Permite el control riguroso de las horas de vuelo de una flota ficticia, gestionando alertas de mantenimiento preventivo y bloqueos de seguridad por normativa **AOG (Aircraft On Ground)**.

## 🚀 Funcionalidades Principales

* **Control de Aeronavegabilidad**: Seguimiento preciso de horas de vuelo con un límite de ciclo de vida de 2000 horas por componente/motor.
* **Gestión de Estados Inteligente**: Clasificación automática mediante lógica de colores y estados:
    * ✅ **OPERATIVO**: Funcionamiento normal.
    * ⚠️ **REVISIÓN CERCANA**: Aviso preventivo (margen < 200h).
    * 🛑 **URGENTE**: Mantenimiento inmediato (margen < 50h).
    * ❌ **AOG (Aircraft On Ground)**: Aeronave no apta para vuelo (margen superado).
* **Bloqueo de Seguridad AOG**: El sistema impide legalmente el registro de nuevos vuelos si la aeronave está en AOG. Solo se permite el registro si las horas se resetean a **0**, simulando un cambio de motor o mantenimiento mayor.
* **Validación de Datos**: Protección contra errores de entrada (impide registrar horas menores a las actuales, detectando incoherencias en el historial).
* **Generación de Informes PDF**: Exportación automática de historiales con formato profesional, datos centrados y resaltado visual de estados críticos en rojo y negrita.
* **Interfaz Dinámica (UX)**: Carga de imágenes en tiempo real de cada aeronave y visualización de historial con filtrado dinámico por matrícula.

## 🛠️ Instalación y Configuración

Para que el sistema funcione correctamente en tu equipo local, sigue estos pasos:

1.  **Descarga el repositorio**: Descarga todos los archivos en una misma carpeta.
2.  **Estructura de archivos**: El código utiliza rutas relativas (`ThisWorkbook.Path`), por lo que la carpeta debe mantener esta estructura:
    * `Sistema_Gestion_Flota.xlsm` (Archivo principal).
    * `EC-ABC.jpg`, `EC-XYZ.jpg`, `EC-123.jpg` (Fotos de las aeronaves).
    * `/Reportes_Avermant/` (Carpeta de destino para los informes PDF).
3.  **Habilitar Macros**: Al abrir el archivo Excel, debes activar "Habilitar contenido" para permitir la ejecución del código VBA.

## ⚙️ Requisitos del Sistema

* **Microsoft Excel**: Versión 2010 o superior (necesaria para la exportación nativa a PDF).
* **Sistema Operativo**: Windows (optimizado para el uso de objetos `Shell`).

## 📝 Ejemplo de Uso

1.  **Consulta**: Selecciona una matrícula y pulsa **Consultar Historial**. Se cargará la foto y los registros previos.
2.  **Registro**: Introduce las nuevas horas totales. Si el avión está en AOG, el sistema te denegará el registro.
3.  **Mantenimiento**: Para reactivar un avión en AOG, registra **0 horas** (Mantenimiento Mayor). El estado cambiará automáticamente a **OPERATIVO**.
4.  **Reporte**: Pulsa **Generar Informe PDF** para obtener un documento técnico listo para impresión o envío.


## 🚀 Roadmap: Próximas Versiones (Cloud Expansion)

Para las siguientes fases del proyecto, se está trabajando en la migración de la interfaz de consulta a un entorno 100% móvil y multiplataforma:

### 📱 Consulta en Tablets y Dispositivos Móviles
* **Integración con Google Apps Script (GAS)**: Implementación de un backend en la nube para conectar la base de datos de Excel (vía Office 365 o Google Sheets API) con dispositivos externos.
* **Interfaz Web HTML5/CSS3**: Creación de un dashboard responsive diseñado específicamente para operarios de pista y mecánicos en hangar.
* **Sincronización en Tiempo Real**: Permitir que el personal técnico consulte el estado de aeronavegabilidad (OPERATIVO/AOG) desde una tablet sin necesidad de acceder a la terminal fija de escritorio.
* **Firma Digital**: Incorporación de validación HTML para la confirmación de inspecciones directamente desde el dispositivo móvil.
  ## 🎮 Compatibilidad con Simuladores (DCS World / MSFS)

Aunque es una herramienta de gestión técnica real, el sistema puede ser utilizado tambien por **Escuadrones Virtuales** en simuladores de alta fidelidad como **DCS World** o **Microsoft Flight Simulator**:

* **Gestión de Flota en Campañas**: Ideal para comandantes de escuadrón que necesiten llevar un registro real de las horas de vuelo de sus pilotos tras cada salida de misión.
* **Simulación de Fatiga**: Permite aplicar una capa de realismo extra, obligando a los aviones que han sufrido daños o han superado las horas de combate a entrar en estado **AOG (AIRCRAFT ON GROUND)** hasta que el equipo de mantenimiento virtual "reseteé" la aeronave.
* **Inmersión en Briefing**: Los reportes PDF generados por **Avermant** pueden ser utilizados en los briefings de misión para informar sobre la disponibilidad de aeronaves operativas.
---

## 👤 Autor

**Jorge Martín** 
*| Developer*






---
*Desarrollado como solución técnica para la gestión de activos aeronáuticos bajo estándares de ingeniería.*
