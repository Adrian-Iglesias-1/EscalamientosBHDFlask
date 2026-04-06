# EscalamientosApp

Sistema de gestión de escalamientos de ATMs. Aplicación web desarrollada con Flask que permite gestionar fallas de ATMs, generar scripts, administrar registros de Closed & Block, y enviar correos mediante integración con Microsoft Outlook.

## Características

- **Gestión de Fallas**: Recepción y procesamiento de datos de fallas desde Excel
- **Generar Scripts**: Creación automática de scripts de escalamiento para tickets
- **XOLUSAT**: Registro y seguimiento de escalamientos con proveedor
- **Closed & Block**: Gestión de ATMs bloqueados con control de tiempo (48h)
- **Integración Outlook**: Envío de correos directamente desde la aplicación
- **RCU**: Actualización automática de datos de ATMs desde archivos RCU
- **Modo Domingo/Feriado**: Detección automática de domingos y modo feriado manual

## Requisitos

- **Python 3.8 o superior** (debe estar instalado y en el PATH)
- Microsoft Windows (para integración con Outlook)
- Microsoft Outlook instalado
- Microsoft Excel (para archivos .xlsx)

## Instalación Rápida (Nuevo Usuario)

### Paso 1: Descargar y Descomprimir

1. Descargar el archivo ZIP del proyecto
2. **Descomprimir completamente** en cualquier carpeta (ej: Descargas, Documentos, etc.)
3. **IMPORTANTE**: No ejecutar desde dentro del ZIP, debe estar descomprimido

> **Nota**: La aplicación funciona desde cualquier ubicación. No necesita estar en `C:\` específicamente.

### Paso 2: Verificar Python

Abrir **Símbolo del sistema** (CMD) y ejecutar:
```bash
py --version
```

Si muestra la versión (ej: `Python 3.11.0`), Python está listo.

**Si no está instalado:**
1. Descargar desde portal empresa
2. Instalar marcando **"Add Python to PATH"**
3. Cerrar y volver a abrir CMD

### Paso 3: Primera Ejecución

1. Abrir la carpeta donde descomprimiste el ZIP
2. Hacer doble clic en **`iniciar.bat`**
3. La primera vez:
   - Creará el entorno virtual (venv)
   - Instalará las dependencias (puede tardar 2-3 minutos)
   - Creará un **acceso directo en el escritorio** con icono
   - Abrirá el navegador automáticamente

### Paso 4: Uso Diario

Después de la primera vez, usar el **acceso directo en el escritorio** llamado **"EscalamientosApp"**.

---

## Instalación Manual (Desarrolladores)

1. Clonar el repositorio:
```bash
git clone https://github.com/Adrian-Iglesias-1/EscalamientosApp.git
cd EscalamientosApp
```

2. Crear entorno virtual:
```bash
python -m venv backend\venv
```

3. Instalar dependencias:
```bash
cd backend
venv\Scripts\pip install -r requirements.txt
```

4. Iniciar manualmente:
```bash
venv\Scripts\python.exe app.py
```

La aplicación estará disponible en: http://localhost:5000

---

## Estructura del Proyecto

```
EscalamientosApp/
├── backend/
│   ├── app.py                    # Aplicación principal Flask
│   ├── excel_handler.py          # Manejo de archivos Excel
│   ├── closed_and_block_handler.py  # Lógica de Closed & Block
│   ├── outlook_handler.py        # Integración con Outlook
│   ├── static/
│   │   ├── style.css            # Estilos personalizados
│   │   └── script.js            # JavaScript frontend
│   ├── templates/
│   │   └── index.html           # Template principal
│   ├── PlanillaEscalamientos.xlsx  # Base de datos principal
│   └── ClosedAndBlock.xlsx      # Registros de bloqueos
├── requirements.txt             # Dependencias Python
├── iniciar_app.bat              # Script de inicio Windows
├── detener_servidor.bat         # Script de detención
└── README.md                    # Este archivo
```

## Uso

### Gestión de Fallas

1. Pegar datos desde Excel (copiar celdas SIN encabezados)
2. Click en "Procesar Datos"
3. Verificar vista previa de ATMs encontrados
4. Click en "Enviar" para generar correos en Outlook

### Closed & Block

1. Pegar IDs de ATMs bloqueados
2. Completar Asunto y Reportado Por
3. Click en "Agregar"
4. Los registros se eliminan automáticamente después de 48 horas

### XOLUSAT

1. Buscar ATM por ID
2. Completar datos del incidente
3. Click en "Registrar" o "Registrar y Enviar Correo"

### Actualizar RCU

1. Click en "Seleccionar archivo" en la sidebar
2. Seleccionar archivo RCU (.xlsx)
3. El sistema actualizará automáticamente los datos de ATMs

## Archivos Excel Requeridos

### PlanillaEscalamientos.xlsx

Debe contener las siguientes hojas:
- **CONTACTOS SEMANA**: Emails de contacto por custodio
- **CONTACTOS_SUC**: Emails para sucursales (columnas: ID, ..., Email, CC)
- **CONTACTOS FINDE**: Contactos para fines de semana
- **UNIFICADO**: Base de datos de ATMs con columnas: ID, Nombre, Custodio, SLA, etc.
- **RCU**: (Opcional) Hoja para actualización de datos

### ClosedAndBlock.xlsx

Se crea automáticamente si no existe. Estructura:
- ID
- NOMBRE
- CUSTODIO
- FECHA_INGRESO
- ESTADO
- ASUNTO
- REPORTADO_POR

## Configuración

### Variables de Entorno (opcional)

Crear archivo `.env` en la carpeta `backend/`:

```
FLASK_DEBUG=False
FLASK_PORT=5000
EXCEL_PATH=PlanillaEscalamientos.xlsx
CLOSED_BLOCK_PATH=ClosedAndBlock.xlsx
```

### Modo Feriado

Activar el switch "Modo Feriado" en la sidebar para usar los contactos de fin de semana.

## Solución de Problemas

### Error "pywin32 no instalado"

```bash
pip install pywin32 --force-reinstall
```

### Error "Planilla no encontrada"

Verificar que el archivo `PlanillaEscalamientos.xlsx` exista en la carpeta `backend/`.

### Error al enviar correos

- Verificar que Outlook esté instalado y configurado
- Verificar que la cuenta de correo esté activa en Outlook
- Ejecutar la aplicación en el mismo usuario donde está configurado Outlook

## Contribuir

1. Fork el repositorio
2. Crear una rama (`git checkout -b feature/nueva-funcionalidad`)
3. Commitear cambios (`git commit -am 'Agregar nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Crear Pull Request

## Licencia

Este proyecto es privado y confidencial.

## Contacto

Para soporte técnico contactar al administrador del sistema.

---

**Nota**: Esta aplicación está diseñada específicamente para funcionar en entornos Windows con Microsoft Outlook instalado.
