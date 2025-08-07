# 🚀 Herramienta de Automatización de Reportes

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://python.org)
[![Flask](https://img.shields.io/badge/Flask-2.0+-green.svg)](https://flask.palletsprojects.com/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## ¿Por qué creé esta herramienta?

Hace unos meses me encontraba pasando 3-4 horas cada lunes generando reportes manualmente para el equipo de ventas. Copiando datos de Excel, creando gráficos uno por uno, y enviando emails con archivos adjuntos. Era tedioso, propenso a errores, y me impedía enfocarme en análisis más importantes.

Esta herramienta nació de esa frustración. Ahora puedo generar reportes completos en minutos, no horas. Y lo mejor: es tan fácil de usar que cualquier miembro del equipo puede hacerlo sin conocimientos técnicos.

## ✨ Características principales

- 📊 **Subir archivos de datos** (CSV o Excel) a través de una interfaz web amigable
- 🧹 **Limpiar automáticamente** los datos (eliminar duplicados, normalizar fechas, etc.)
- 📈 **Generar reportes** en Excel y PDF con gráficos y tablas resumen
- 📧 **Enviar automáticamente** los reportes por email a tu lista de destinatarios
- 🌐 **Interfaz web intuitiva** que no requiere conocimientos técnicos
- 🔒 **Seguridad** con variables de entorno para credenciales sensibles

## 🚀 Instalación Rápida

### Prerrequisitos
- Python 3.8 o superior
- Git (para clonar el repositorio)

### 1. Clonar el repositorio
```bash
git clone https://github.com/BrianDevCo/-Automatizaci-n-de-reportes-con-Python-Flask.git
cd -Automatizaci-n-de-reportes-con-Python-Flask
```

### 2. Crear y activar el entorno virtual

**Windows:**
```bash
# Crear el entorno virtual
python -m venv venv

# Activar el entorno
venv\Scripts\activate
```

**macOS/Linux:**
```bash
# Crear el entorno virtual
python3 -m venv venv

# Activar el entorno
source venv/bin/activate
```

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 4. Configurar variables de entorno

Crea un archivo `.env` en la raíz del proyecto:

```env
EMAIL_USER=tu_correo@gmail.com
EMAIL_PASSWORD=tu_contraseña_de_aplicacion
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
```

**⚠️ Importante:** Para Gmail, necesitas crear una "contraseña de aplicación":
1. Ve a [Configuración de tu cuenta Google](https://myaccount.google.com/)
2. Seguridad → Verificación en dos pasos (actívala si no está activa)
3. Contraseñas de aplicación → Generar nueva contraseña
4. Usa esa contraseña en el archivo `.env`

### 5. Ejecutar la aplicación
```bash
python app.py
```

Luego abre tu navegador en `http://localhost:5000`

## 📖 Cómo usar la herramienta

### Paso 1: Preparar tus datos
- Asegúrate de que tu archivo CSV/Excel tenga columnas con nombres claros
- Incluye al menos una columna de fechas y una de valores numéricos
- Guarda el archivo en formato CSV o Excel (.xlsx)

### Paso 2: Subir y procesar
- Abre la aplicación en tu navegador
- Haz clic en "Seleccionar archivo" y elige tu archivo de datos
- Completa la información del reporte (título, destinatarios, etc.)
- Haz clic en "Generar Reporte"

### Paso 3: Revisar y enviar
- Revisa el reporte generado en el navegador
- Descárgalo si lo necesitas
- El reporte se enviará automáticamente a los destinatarios especificados

## 💡 Ejemplo de uso real

La semana pasada, María del equipo de ventas me pidió un reporte semanal de ventas por región. Antes me hubiera tomado 2 horas. Con esta herramienta:

1. Subí el archivo CSV de ventas (30 segundos)
2. Configuré el reporte con gráficos de barras por región (1 minuto)
3. Agregué los emails del equipo de ventas (30 segundos)
4. ¡Listo! El reporte se generó y envió automáticamente en menos de 2 minutos.

## 📁 Estructura del proyecto

```
📁 proyecto/
├── 📄 app.py              # Aplicación principal con Flask
├── 📄 data_processor.py   # Lógica de procesamiento de datos
├── 📄 report_generator.py # Generación de reportes Excel/PDF
├── 📄 email_sender.py     # Envío automático de emails
├── 📄 requirements.txt    # Dependencias del proyecto
├── 📄 .env               # Variables de entorno (crear tú)
├── 📄 .gitignore         # Archivos a ignorar en Git
├── 📁 templates/         # Plantillas HTML
├── 📁 static/           # Archivos CSS/JS
├── 📁 reports/          # Reportes generados
└── 📁 uploads/          # Archivos subidos temporalmente
```

## ❓ Preguntas frecuentes

### ¿Qué formatos de archivo acepta?
- CSV (.csv)
- Excel (.xlsx, .xls)

### ¿Puedo personalizar los gráficos?
¡Por supuesto! La herramienta detecta automáticamente el tipo de datos y sugiere el gráfico más apropiado. También puedes especificar qué columnas usar para cada gráfico.

### ¿Los reportes se guardan?
Sí, todos los reportes se guardan en la carpeta `reports/` con fecha y hora para que puedas acceder a ellos después.

### ¿Qué pasa si el email no llega?
- Revisa tu carpeta de spam
- Verifica que la contraseña de aplicación esté correcta
- Asegúrate de que el servidor SMTP esté bien configurado

## 🔧 Posibles errores y cómo resolverlos

### Error: "No se pudo conectar al servidor SMTP"
- Verifica que `EMAIL_USER` y `EMAIL_PASSWORD` estén correctos
- Para Gmail, usa una contraseña de aplicación, no tu contraseña normal
- Asegúrate de que el puerto 587 esté abierto

### Error: "No se pudo leer el archivo"
- Verifica que el archivo no esté corrupto
- Asegúrate de que sea CSV o Excel
- Revisa que tenga datos en las primeras filas

### Error: "No se encontraron columnas numéricas"
- Asegúrate de que tu archivo tenga al menos una columna con números
- Verifica que no haya texto en columnas que deberían ser numéricas

### La aplicación no inicia
- Verifica que el entorno virtual esté activado
- Asegúrate de que todas las dependencias estén instaladas
- Revisa que no haya otro proceso usando el puerto 5000

## 🤝 Contribuir

Si encuentras un bug o tienes una idea para mejorar la herramienta, ¡me encantaría escucharla! Esta herramienta la hice para mí, pero si te sirve a ti también, ¡mejor aún!

### Cómo contribuir:
1. Haz un fork del repositorio
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Haz commit de tus cambios (`git commit -am 'Agregar nueva funcionalidad'`)
4. Haz push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto es de código abierto bajo la licencia MIT. Úsalo, modifícalo, compártelo. Solo pido que si lo mejoras, comparte las mejoras con la comunidad.

---

**¡Espero que esta herramienta te ahorre tanto tiempo como me ha ahorrado a mí!** 🎉

---
*Desarrollado con ❤️ para automatizar tareas repetitivas y hacer la vida más fácil.*
