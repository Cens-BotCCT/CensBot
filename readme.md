# Sistema de Gestión de PQRS - CENS

## Descripción
Automatización realizad en Python con selenium integrando la api de GPT 4o de Azure OpenAI

## Características Principales
- Automatización del proceso de login, consulta y descarga de radicados en la plataforma Mercurio
- Extracción y procesamiento de correos electrónicos
- Análisis de contenido mediante Azure OpenAI
- Generación automática de cartas de respuesta
- Gestión de traslados por competencia

## Tecnologías Utilizadas
- Python 3.12.8
- Selenium: cualquier verion que sea compatible con la verson de Python utilizada 
- Webdriver: de acuerdo a que navegador se vaya a utilziar, en este caso se realizo con 
  Edge, se requiere instalar el Edgedriver de acuerdo a la version del navegador 
- Azure OpenAI
- python-docx
- PyPDF2

## Estructura del Proyecto
```
proyecto/
│
├── main.py                 # Punto de entrada principal
├── docx_convert.py         # Procesamiento de documentos
├── flujo_correo.py        # Manejo de flujo de correos
├── rad_utils.py           # Utilidades para radicados
└── instrucciones.md       # archivo con instruccions para el modelo de IA utilizado 
```

## Requisitos del Sistema
- Python 3.x
- Microsoft Edge WebDriver (de acuerdo a la version del navegador)
- Conexión a Internet
- Acceso a la plataforma Mercurio
- Acceso a Azure OpenAI 

## Instalación

1. Clonar el repositorio:
```bash
git clone [url-del-repositorio]
```

2. Instalar dependencias:
```bash
pip install -r requirements.txt
```

3. Configurar las credenciales de Azure OpenAI en docx_convert.py (usar un archivo .env para mayor seguridad) 

4. Configurar la ruta del WebDriver en main.py (usar variable de entorno del sistema o definir la ruta donde se encontrar el ejecutable del edge)


## Uso
1. Ejecutar el script principal:
```bash
python main.py
```

2. El sistema realizará automáticamente:
   - Login en la plataforma Mercurio
   - Búsqueda de radicados
   - Descarga de PQR's
   - Procesamiento de correos
   - Generación de respuestas

## Flujos de Trabajo
El sistema maneja tres tipos principales de flujos:
1. **Correo Electrónico**: Procesa correos y genera respuestas automáticas
2. **Página Web**: Maneja solicitudes web (no implementado)
3. **Oficina**: Gestiona documentos físicos digitalizados (no implentado)


## Configuración de Azure OpenAI
El proyecto utiliza el servicio de Azure OpenAI para realizar la parte del procesamiento del contenido del PQR y para la generacion de las cartas de respuesta.
Para poder utilizar este servicio se necesitan las siguientes credenciales 

-API version
- Endpoint
- Deployment name
- API Key

## Mantenimiento
- Actualizar regularmente el WebDriver de Edge de acuerdo con la version del navegador 
- Verificar las credenciales de acceso
- Monitorear los límites de la API de Azure

## Solución de Problemas (Troubleshooting)

### Errores Comunes

#### 1. Errores de WebDriver
- **Error**: `SessionNotCreatedException: Message: unknown error: cannot find Edge binary`
  - **Solución**: Verificar que Microsoft Edge esté instalado y la versión del WebDriver coincida
- **Error**: `WebDriverException: Message: 'msedgedriver' executable needs to be in PATH`
  - **Solución**: Asegurarse que la ruta del WebDriver está correctamente configurada

#### 2. Errores de Autenticación
- **Error**: `TimeoutException: Message: timeout while waiting for login`
  - **Solución**: Verificar credenciales y conexión a la red corporativa
- **Error**: `ElementNotInteractableException: Message: element not interactable`
  - **Solución**: Esperar a que los elementos sean completamente cargados

#### 3. Errores de Azure OpenAI
- **Error**: `AuthenticationError: Invalid API key provided`
  - **Solución**: Verificar las credenciales en el archivo .env
- **Error**: `RateLimitError: Rate limit reached`
  - **Solución**: Implementar sistema de reintentos o esperar el reseteo del límite

#### 4. Errores de Procesamiento de Documentos
- **Error**: `PermissionError: Permission denied`
  - **Solución**: Verificar permisos de escritura en la carpeta de destino
- **Error**: `FileNotFoundError: No such file or directory`
  - **Solución**: Verificar rutas de archivos y crear directorios necesarios




