# SLPA
Proyecto VBA basado en estándar API MPMS para cálculo de producción de hidrocarburos.

## Descripción General

El proyecto SLPA es un sistema basado en VBA (Visual Basic for Applications) que implementa estándares de cálculo de producción de hidrocarburos según la norma API MPMS (American Petroleum Institute Measurement Standards). Se enfoca en el cálculo de volúmenes de producción de crudos, productos refinados y lubricantes en instalaciones de procesamiento y almacenamiento de hidrocarburos.

## Estructura del Proyecto

### Clases Principales

1. **clsComponent** (clsComponent.cls)
   - Clase base que define las propiedades comunes a todos los componentes de la instalación
   - Implementa la interfaz IComponent
   - Propiedades: Tag, Descripción, Tipo, Sistema, Servicio, Estado, Estación

2. **clsStation** (clsStation.cls)
   - Representa una estación de procesamiento o almacenamiento
   - Contiene una colección de componentes (tanques, pozos, líneas)
   - Métodos para agregar, remover y buscar componentes

3. **clsTank** (clsTank.cls)
   - Representa un tanque de almacenamiento
   - Propiedades específicas: Capacidad, Diámetro, Material, Techo flotante, etc.
   - Métodos de cálculo avanzados para volúmenes de almacenamiento

4. **clsWell** (clsWell.cls)
   - Representa un pozo productor o inyector
   - Propiedades: Profundidad, Tasa de flujo, Fluidos

5. **clsLine** (clsLine.cls)
   - Representa una línea de flujo o tubería
   - Propiedades: Longitud, Diámetro, Material

6. **clsFluid** (clsFluid.cls)
   - Representa las propiedades termofísicas de fluidos
   - Sincronización automática entre API y densidad
   - Métodos de cálculo para factores de corrección

### Interfaces

1. **IComponent** (IComponent.cls)
   - Define el contrato para todos los componentes de la instalación
   - Propiedades: Tag, Descripción, Tipo, Sistema, Servicio, Estado, Estación
   - Método ShowProperties

### Módulos de Funciones Auxiliares

1. **mdGlobals** (mdGlobals.bas)
   - Definición de enumeraciones (tipos de fluidos, materiales, estados)
   - Constantes globales para cálculos API MPMS
   - Estructuras de datos personalizadas

2. **mdConversion** (mdConversion.bas)
   - Funciones de conversión entre unidades de medida
   - Conversión de API a densidad y viceversa
   - Conversión de temperaturas entre escalas

3. **mdAPICalcs** (mdAPICalcs.bas)
   - Implementación de cálculos según API MPMS 11.1
   - Factores de corrección por temperatura (CTL) y presión (CPL)
   - Coeficientes de expansión térmica (Alfa60)
   - Corrección de hidrómetros (HYC)
   - Cálculo de volúmenes de tanques (FRA, CTSH)

4. **mdHelpers** (mdHelpers.bas)
   - Funciones de validación de datos
   - Funciones auxiliares para cálculos numéricos
   - Funciones de manejo de errores

5. **mdStationService** (mdStationService.bas)
   - Servicio para acceso a la tabla de estaciones
   - Mapeo de datos de Excel a objetos clsStation

6. **mdTankService** (mdTankService.bas)
   - Servicio para acceso a tablas de aforo
   - Caché de tablas de datos para mejor rendimiento
   - Carga de configuraciones de tanques desde Excel

7. **mdTestClasses** (mdTestClasses.bas)
   - Pruebas unitarias para componentes del sistema
   - Escenarios de prueba para verificación de funcionamiento

8. **mdTestCalcs** (mdTestCalcs.bas)
   - Pruebas de cálculo específicos del motor API
   - Validación de algoritmos de cálculo

## Características Principales

### 1. Arquitectura Orientada a Objetos
- Uso de composición y herencia para evitar duplicación de código
- Implementación de interfaces para polimorfismo
- Diseño modular y reutilizable

### 2. Estándares API MPMS
- Implementación completa de los estándares API MPMS 11.1 y 12.1.1
- Cálculos precisos de volúmenes con correcciones por temperatura y presión
- Manejo de diferentes tipos de fluidos (crudos, refinados, lubricantes)

### 3. Gestión de Datos
- Integración con Excel (tablas dinámicas, list objects)
- Caché de datos para mejorar el rendimiento
- Validación de datos y manejo de errores robusto

### 4. Funcionalidades Avanzadas
- Cálculo de volúmenes de almacenamiento (TOV, GOV, GSV, NSV)
- Corrección por techo flotante (FRA)
- Corrección por expansión térmica del tanque (CTSH)
- Manejo de fluidos con propiedades termofísicas completas

## Estructura de Datos

### Enumeraciones
- Tipos de fluidos (CRD, REF, LUB)
- Estados operativos (OP, OF, MT, MC, FS)
- Materiales (Acero al Carbono, Inoxidable)
- Unidades de medida

### Constantes
- Valores físicos fundamentales (densidad del agua, constantes de expansión)
- Rangos normativos de API para diferentes tipos de fluidos
- Constantes de cálculo según estándares API

## Funcionalidades Implementadas

1. **Cálculo de Volúmenes**
   - TOV (Total Observado)
   - GOV (Gross Observado)
   - GSV (Gross Estándar)
   - NSV (Neto Estándar)

2. **Correcciones de Medición**
   - Corrección por temperatura (CTL)
   - Corrección por presión (CPL)
   - Corrección por hidrómetro (HYC)
   - Corrección por techo flotante (FRA)
   - Corrección por expansión térmica del tanque (CTSH)

3. **Gestión de Componentes**
   - Creación y manipulación de estaciones
   - Gestión de componentes (tanques, pozos, líneas)
   - Polimorfismo mediante interfaces

4. **Validación de Datos**
   - Verificación de rangos de API
   - Validación de temperaturas
   - Control de errores y excepciones

## Pruebas y Validación

El proyecto incluye un conjunto completo de pruebas unitarias que verifican:
- Funcionamiento de componentes individuales
- Cálculos de volúmenes con correcciones
- Integración entre diferentes clases
- Manejo de errores y casos extremos

## Consideraciones de Implementación

1. **Rendimiento**: Uso de caché para tablas de aforo y cálculos recurrentes
2. **Escalabilidad**: Diseño modular que permite extensión fácil
3. **Mantenimiento**: Código bien documentado y estructurado
4. **Integración**: Facilita la conexión con Excel y sistemas de datos existentes

## Próximos Pasos

Según el archivo mdImplementar.bas, algunos elementos aún están pendientes:
- Interfaz de usuario (formularios)
- Importación/exportación via Power Query
- Seguimiento de producción de pozos
- Pre-cálculo de tablas CTL/CPL para rendimiento masivo

Este proyecto representa una implementación completa de una solución de cálculo de producción de hidrocarburos siguiendo estándares internacionales, ideal para aplicaciones en industrias petroleras y de gas.